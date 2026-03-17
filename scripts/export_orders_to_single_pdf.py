#!/usr/bin/env python3
"""
Export Gmail invoice emails for a list of order IDs into a single combined PDF.

Workflow:
1) Read order IDs from a text file (one per line)
2) For each ID, find a Gmail message whose subject contains "#<orderId>"
3) Fetch message, extract HTML (fallback to text), render to SINGLE-PAGE A4 PDF
4) Merge all order PDFs into one long PDF, in the same order as the input list

Example:
  python export_orders_to_single_pdf.py --orders-file lost-orders-woo.txt --after 2025/11/01
  python export_orders_to_single_pdf.py --orders-file lost-orders-woo.txt --after 2025/11/01 --limit 3
"""

from __future__ import annotations

import argparse
import base64
import os
import pickle
import re
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
TOKEN_FILE = "token.pickle"
CREDENTIALS_FILE = "credentials.json"


def log(msg: str) -> None:
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


def authenticate_gmail():
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "rb") as token:
            creds = pickle.load(token)

    if not creds or not getattr(creds, "valid", False):
        if creds and getattr(creds, "expired", False) and getattr(creds, "refresh_token", None):
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                raise FileNotFoundError(f"Missing {CREDENTIALS_FILE} in current directory.")
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, "wb") as token:
            pickle.dump(creds, token)

    return build("gmail", "v1", credentials=creds)


def urlsafe_b64decode(data: str) -> bytes:
    return base64.urlsafe_b64decode(data.encode("utf-8"))


def extract_header(payload: dict, header_name: str) -> str:
    headers = payload.get("headers", []) or []
    for h in headers:
        if (h.get("name") or "").lower() == header_name.lower():
            return h.get("value") or ""
    return ""


def find_html_part(payload: dict) -> Optional[str]:
    """Return decoded HTML body if present, else None."""
    if not payload:
        return None
    mime = payload.get("mimeType")
    body = (payload.get("body") or {}).get("data")
    if mime == "text/html" and body:
        return urlsafe_b64decode(body).decode("utf-8", errors="replace")
    for p in (payload.get("parts") or []):
        found = find_html_part(p)
        if found:
            return found
    return None


def find_text_part(payload: dict) -> Optional[str]:
    """Return decoded plain text body if present, else None."""
    if not payload:
        return None
    mime = payload.get("mimeType")
    body = (payload.get("body") or {}).get("data")
    if mime == "text/plain" and body:
        return urlsafe_b64decode(body).decode("utf-8", errors="replace")
    for p in (payload.get("parts") or []):
        found = find_text_part(p)
        if found:
            return found
    return None


def wrap_as_html(subject: str, inner_html: str) -> str:
    return f"""<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <title>{subject}</title>
    <style>
      body {{ font-family: Arial, sans-serif; margin: 24px; }}
      img {{ max-width: 100%; height: auto; }}
      table {{ max-width: 100%; }}
    </style>
  </head>
  <body>
    <h2>{subject}</h2>
    <hr />
    {inner_html}
  </body>
</html>
"""


def wrap_text_as_html(subject: str, text: str) -> str:
    safe = (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )
    return wrap_as_html(subject, f"<pre>{safe}</pre>")


def read_order_ids(path: Path) -> List[str]:
    ids: List[str] = []
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        v = line.strip()
        if not v:
            continue
        if not re.fullmatch(r"\d+", v):
            continue
        ids.append(v)
    # preserve order, remove duplicates
    seen: Set[str] = set()
    ordered_unique: List[str] = []
    for oid in ids:
        if oid in seen:
            continue
        seen.add(oid)
        ordered_unique.append(oid)
    return ordered_unique


def build_subject_query_chunks(
    order_ids: List[str],
    email_address: str,
    after_date: str,
    *,
    include_plain_id: bool,
    include_address_filter: bool,
    chunk_size: int = 50,
    max_query_len: int = 2000,
) -> List[str]:
    # Gmail supports: in:inbox, in:anywhere, etc.
    # We include a conservative default and keep queries short.
    chunks: List[str] = []
    current: List[str] = []
    addr = f"(from:{email_address} OR to:{email_address}) " if include_address_filter else ""
    after = f"after:{after_date} " if after_date else ""
    for oid in order_ids:
        current.append(oid)
        def terms(x: str) -> List[str]:
            return [f"#{x}", x] if include_plain_id else [f"#{x}"]

        subject_terms: List[str] = []
        for x in current:
            subject_terms.extend(terms(x))
        subject_part = " OR ".join(subject_terms)
        q = f'{addr}{after}subject:({subject_part})'
        if len(current) >= chunk_size or len(q) > max_query_len:
            current.pop()
            if current:
                subject_terms = []
                for x in current:
                    subject_terms.extend(terms(x))
                subject_part = " OR ".join(subject_terms)
                q = f'{addr}{after}subject:({subject_part})'
                chunks.append(q)
            current = [oid]
    if current:
        def terms(x: str) -> List[str]:
            return [f"#{x}", x] if include_plain_id else [f"#{x}"]

        subject_terms = []
        for x in current:
            subject_terms.extend(terms(x))
        subject_part = " OR ".join(subject_terms)
        q = f'{addr}{after}subject:({subject_part})'
        chunks.append(q)
    return chunks


def search_messages_paginated(service, query: str, batch_size: int = 500) -> List[dict]:
    messages: List[dict] = []
    page_token = None
    while True:
        params = {"userId": "me", "q": query, "maxResults": batch_size}
        if page_token:
            params["pageToken"] = page_token
        resp = service.users().messages().list(**params).execute()
        messages.extend(resp.get("messages") or [])
        page_token = resp.get("nextPageToken")
        if not page_token:
            return messages


def get_message_metadata_subject(service, msg_id: str) -> str:
    msg = (
        service.users()
        .messages()
        .get(userId="me", id=msg_id, format="metadata", metadataHeaders=["Subject"])
        .execute()
    )
    payload = msg.get("payload") or {}
    return extract_header(payload, "Subject") or ""


def extract_order_id_from_subject(subject: str) -> Optional[str]:
    if not subject:
        return None
    m = re.search(r"order\s*#\s*(\d+)", subject, re.IGNORECASE)
    if m:
        return m.group(1)
    m = re.search(r"#(\d+)", subject)
    if m:
        return m.group(1)
    return None


def map_order_to_message_id(
    service,
    order_ids: List[str],
    email_address: str,
    after_date: str,
    scope: str,
    *,
    include_plain_id: bool,
    include_address_filter: bool,
) -> Dict[str, str]:
    wanted = set(order_ids)
    mapping: Dict[str, str] = {}

    # scope: inbox or anywhere (archived/read emails are often not in inbox)
    scope_query = f"in:{scope}"
    chunks = [
        f"{scope_query} {q}"
        for q in build_subject_query_chunks(
            order_ids,
            email_address=email_address,
            after_date=after_date,
            include_plain_id=include_plain_id,
            include_address_filter=include_address_filter,
        )
    ]
    log(f"Built {len(chunks)} Gmail queries (subject OR chunks)")

    all_msg_ids: List[str] = []
    for idx, q in enumerate(chunks, 1):
        log(f"[Search {idx}/{len(chunks)}] listing message ids...")
        msgs = search_messages_paginated(service, q)
        ids = [m["id"] for m in msgs if "id" in m]
        all_msg_ids.extend(ids)
        log(f"  found {len(ids)} messages (total so far: {len(all_msg_ids)})")

    # De-dupe message ids while preserving order
    seen_msg: Set[str] = set()
    msg_ids_unique: List[str] = []
    for mid in all_msg_ids:
        if mid in seen_msg:
            continue
        seen_msg.add(mid)
        msg_ids_unique.append(mid)

    log(f"Resolving subjects for {len(msg_ids_unique)} messages to map order->message...")
    for i, mid in enumerate(msg_ids_unique, 1):
        if i == 1 or i % 100 == 0:
            log(f"  subject fetch {i}/{len(msg_ids_unique)}")
        try:
            subj = get_message_metadata_subject(service, mid)
        except HttpError:
            continue
        oid = extract_order_id_from_subject(subj)
        if oid and oid in wanted and oid not in mapping:
            mapping[oid] = mid
            if len(mapping) == len(wanted):
                break

    log(f"Mapped {len(mapping)}/{len(order_ids)} order IDs to message IDs")
    return mapping


def fetch_message_full(service, msg_id: str) -> dict:
    return service.users().messages().get(userId="me", id=msg_id, format="full").execute()


def render_many_one_page_pdfs(html_by_order: List[Tuple[str, Path]], out_dir: Path) -> List[Tuple[str, Path]]:
    """
    Render many HTML files to single-page PDFs using one shared Chromium instance.
    Returns list of (order_id, pdf_path) for successfully rendered PDFs.
    """
    try:
        from playwright.sync_api import sync_playwright
    except Exception as e:
        raise RuntimeError("Playwright missing. Run: pip install -r requirements.txt && python -m playwright install chromium") from e

    from math import floor

    # A4 at ~96 DPI
    a4_width_px = 794
    a4_height_px = 1123
    margin_mm = 8
    margin_px = int(margin_mm * 96 / 25.4)
    printable_w = max(1, a4_width_px - 2 * margin_px)
    printable_h = max(1, a4_height_px - 2 * margin_px)

    rendered: List[Tuple[str, Path]] = []
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": a4_width_px, "height": a4_height_px})
        page.emulate_media(media="print")

        for i, (oid, html_path) in enumerate(html_by_order, 1):
            if i == 1 or i % 25 == 0:
                log(f"PDF render {i}/{len(html_by_order)} (order {oid})")

            pdf_path = out_dir / f"order-{oid}.pdf"
            page.goto(html_path.resolve().as_uri(), wait_until="load")

            # Scale-to-fit one page (best effort)
            scroll_w = page.evaluate("() => document.documentElement.scrollWidth")
            scroll_h = page.evaluate("() => document.documentElement.scrollHeight")
            scale = 1.0
            if scroll_w and scroll_h:
                w_scale = printable_w / float(scroll_w)
                h_scale = printable_h / float(scroll_h)
                scale = min(1.0, w_scale, h_scale)
                scale = max(0.2, scale)

            page.pdf(
                path=str(pdf_path),
                format="A4",
                print_background=True,
                scale=scale,
                margin={"top": f"{margin_mm}mm", "bottom": f"{margin_mm}mm", "left": f"{margin_mm}mm", "right": f"{margin_mm}mm"},
            )
            rendered.append((oid, pdf_path))

        browser.close()

    return rendered


def merge_pdfs(pdfs_in_order: List[Path], out_pdf: Path) -> None:
    try:
        from pypdf import PdfReader, PdfWriter
    except Exception as e:
        raise RuntimeError("Missing pypdf. Install with: pip install pypdf") from e

    writer = PdfWriter()
    for p in pdfs_in_order:
        reader = PdfReader(str(p))
        for page in reader.pages:
            writer.add_page(page)
    with open(out_pdf, "wb") as f:
        writer.write(f)


def append_to_existing_pdf(existing_pdf: Path, extra_pdfs: List[Path]) -> None:
    """Append pages from extra_pdfs to an existing PDF, in-place (via temp file)."""
    from pypdf import PdfReader, PdfWriter

    writer = PdfWriter()
    base = PdfReader(str(existing_pdf))
    for page in base.pages:
        writer.add_page(page)
    for p in extra_pdfs:
        r = PdfReader(str(p))
        for page in r.pages:
            writer.add_page(page)

    tmp = existing_pdf.with_suffix(existing_pdf.suffix + ".tmp")
    with open(tmp, "wb") as f:
        writer.write(f)
    tmp.replace(existing_pdf)


def main() -> int:
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass

    parser = argparse.ArgumentParser()
    parser.add_argument("--orders-file", default="lost-orders-woo.txt")
    parser.add_argument("--after", default="2025/11/01")
    parser.add_argument("--email", default="customerinvoices@particleformen.com")
    parser.add_argument("--out-dir", default="out-batch")
    parser.add_argument("--combined-pdf", default="orders-combined.pdf")
    parser.add_argument("--limit", type=int, default=0, help="Process only first N orders (for testing)")
    parser.add_argument("--scope", choices=["inbox", "anywhere"], default="inbox", help="Gmail search scope (use 'anywhere' to include archived).")
    parser.add_argument("--append-to", default="", help="If set, append the newly generated order PDFs to this existing PDF path.")
    parser.add_argument("--include-plain-id", action=argparse.BooleanOptionalAction, default=False, help="Also search subject for the plain numeric ID (not just #ID).")
    parser.add_argument("--address-filter", action=argparse.BooleanOptionalAction, default=True, help="Include (from:email OR to:email) filter in Gmail query (default: true). Use --no-address-filter to search regardless of recipients/senders (useful for forwarded/BCC).")
    args = parser.parse_args()

    start = time.time()
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    html_dir = out_dir / "html"
    html_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir = out_dir / "pdf"
    pdf_dir.mkdir(parents=True, exist_ok=True)

    order_ids = read_order_ids(Path(args.orders_file))
    if args.limit and args.limit > 0:
        order_ids = order_ids[: args.limit]
    if not order_ids:
        log("No order IDs found.")
        return 2

    log(f"Loaded {len(order_ids)} order IDs from {args.orders_file}")
    log("Authenticating...")
    service = authenticate_gmail()
    log("[OK] Authenticated.")

    mapping = map_order_to_message_id(
        service,
        order_ids,
        email_address=args.email,
        after_date=args.after,
        scope=args.scope,
        include_plain_id=args.include_plain_id,
        include_address_filter=args.address_filter,
    )

    missing = [oid for oid in order_ids if oid not in mapping]
    if missing:
        (out_dir / "missing-orders.txt").write_text("\n".join(missing) + "\n", encoding="utf-8")
        log(f"Missing {len(missing)} orders (written to {out_dir / 'missing-orders.txt'})")

    # Fetch + write HTML per order (only those mapped)
    html_by_order: List[Tuple[str, Path]] = []
    for i, oid in enumerate(order_ids, 1):
        if oid not in mapping:
            continue
        if i == 1 or i % 25 == 0:
            log(f"HTML fetch {i}/{len(order_ids)} (order {oid})")
        mid = mapping[oid]
        try:
            msg = fetch_message_full(service, mid)
        except HttpError as e:
            continue
        payload = msg.get("payload") or {}
        subject = extract_header(payload, "Subject") or f"Order #{oid}"
        html = find_html_part(payload)
        if html:
            full_html = wrap_as_html(subject, html)
        else:
            text = find_text_part(payload) or ""
            full_html = wrap_text_as_html(subject, text)
        html_path = html_dir / f"order-{oid}.html"
        html_path.write_text(full_html, encoding="utf-8", errors="replace")
        html_by_order.append((oid, html_path))

    log(f"Wrote {len(html_by_order)} HTML files. Rendering PDFs (1 page per order)...")
    rendered = render_many_one_page_pdfs(html_by_order, pdf_dir)

    # Keep in original order and only for successfully rendered PDFs
    pdf_map = {oid: p for oid, p in rendered}
    pdfs_in_order = [pdf_map[oid] for oid in order_ids if oid in pdf_map]

    combined_pdf_path = out_dir / args.combined_pdf
    if args.append_to:
        existing_pdf = Path(args.append_to)
        log(f"Appending {len(pdfs_in_order)} PDFs into existing {existing_pdf} ...")
        append_to_existing_pdf(existing_pdf, pdfs_in_order)
        combined_pdf_path = existing_pdf
    else:
        log(f"Merging {len(pdfs_in_order)} PDFs into {combined_pdf_path} ...")
        merge_pdfs(pdfs_in_order, combined_pdf_path)

    elapsed = time.time() - start
    log(f"[OK] Done. Combined PDF pages: {len(pdfs_in_order)}")
    log(f"Elapsed: {elapsed:.1f}s ({elapsed/60:.1f} min)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


