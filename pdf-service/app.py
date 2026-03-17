#!/usr/bin/env python3
"""
PDF microservice — wraps Gmail invoice extraction + Playwright PDF rendering.
Exposes a single POST /render endpoint consumed by the Node.js web service.
"""

from __future__ import annotations

import base64
import os
import pickle
import re
import tempfile
import time
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from flask import Flask, jsonify, make_response, request
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
TOKEN_FILE = os.getenv("TOKEN_FILE", "token.pickle")
CREDENTIALS_FILE = os.getenv("CREDENTIALS_FILE", "credentials.json")
GMAIL_EMAIL = os.getenv("GMAIL_EMAIL", "customerinvoices@particleformen.com")
GMAIL_AFTER_DATE = os.getenv("GMAIL_AFTER_DATE", "2024/01/01")

app = Flask(__name__)


# ── Logging ───────────────────────────────────────────────────────────────────

def log(msg: str) -> None:
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


# ── Gmail auth ────────────────────────────────────────────────────────────────

def authenticate_gmail():
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "rb") as f:
            creds = pickle.load(f)

    if not creds or not getattr(creds, "valid", False):
        if creds and getattr(creds, "expired", False) and getattr(creds, "refresh_token", None):
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                raise FileNotFoundError(f"Missing {CREDENTIALS_FILE}")
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, "wb") as f:
            pickle.dump(creds, f)

    return build("gmail", "v1", credentials=creds)


# ── Gmail helpers (ported from export_orders_to_single_pdf.py) ────────────────

def urlsafe_b64decode(data: str) -> bytes:
    return base64.urlsafe_b64decode(data.encode("utf-8"))


def extract_header(payload: dict, header_name: str) -> str:
    for h in (payload.get("headers") or []):
        if (h.get("name") or "").lower() == header_name.lower():
            return h.get("value") or ""
    return ""


def find_html_part(payload: dict) -> Optional[str]:
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
    safe = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return wrap_as_html(subject, f"<pre>{safe}</pre>")


def build_subject_query_chunks(
    order_ids: List[str],
    email_address: str,
    after_date: str,
    *,
    include_address_filter: bool,
    chunk_size: int = 50,
    max_query_len: int = 2000,
) -> List[str]:
    addr = f"(from:{email_address} OR to:{email_address}) " if include_address_filter else ""
    after = f"after:{after_date} " if after_date else ""
    chunks: List[str] = []
    current: List[str] = []

    for oid in order_ids:
        current.append(oid)
        subject_part = " OR ".join(f"#{x}" for x in current)
        q = f"{addr}{after}subject:({subject_part})"
        if len(current) >= chunk_size or len(q) > max_query_len:
            current.pop()
            if current:
                subject_part = " OR ".join(f"#{x}" for x in current)
                chunks.append(f"{addr}{after}subject:({subject_part})")
            current = [oid]

    if current:
        subject_part = " OR ".join(f"#{x}" for x in current)
        chunks.append(f"{addr}{after}subject:({subject_part})")

    return chunks


def search_messages_paginated(service, query: str) -> List[dict]:
    messages: List[dict] = []
    page_token = None
    while True:
        params = {"userId": "me", "q": query, "maxResults": 500}
        if page_token:
            params["pageToken"] = page_token
        resp = service.users().messages().list(**params).execute()
        messages.extend(resp.get("messages") or [])
        page_token = resp.get("nextPageToken")
        if not page_token:
            return messages


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
    scope: str = "anywhere",
    include_address_filter: bool = True,
) -> Dict[str, str]:
    wanted = set(order_ids)
    mapping: Dict[str, str] = {}

    chunks = [
        f"in:{scope} {q}"
        for q in build_subject_query_chunks(
            order_ids,
            email_address=email_address,
            after_date=after_date,
            include_address_filter=include_address_filter,
        )
    ]
    log(f"Built {len(chunks)} Gmail queries")

    all_msg_ids: List[str] = []
    for idx, q in enumerate(chunks, 1):
        log(f"[Search {idx}/{len(chunks)}] querying Gmail...")
        msgs = search_messages_paginated(service, q)
        ids = [m["id"] for m in msgs if "id" in m]
        all_msg_ids.extend(ids)
        log(f"  found {len(ids)} messages (total: {len(all_msg_ids)})")

    seen: Set[str] = set()
    unique_ids = [mid for mid in all_msg_ids if not (mid in seen or seen.add(mid))]  # type: ignore[func-returns-value]

    log(f"Resolving subjects for {len(unique_ids)} messages...")
    for i, mid in enumerate(unique_ids, 1):
        if i == 1 or i % 100 == 0:
            log(f"  subject fetch {i}/{len(unique_ids)}")
        try:
            msg = (
                service.users()
                .messages()
                .get(userId="me", id=mid, format="metadata", metadataHeaders=["Subject"])
                .execute()
            )
        except HttpError:
            continue
        subj = extract_header(msg.get("payload") or {}, "Subject")
        oid = extract_order_id_from_subject(subj)
        if oid and oid in wanted and oid not in mapping:
            mapping[oid] = mid
            if len(mapping) == len(wanted):
                break

    log(f"Mapped {len(mapping)}/{len(order_ids)} orders to messages")
    return mapping


def fetch_message_full(service, msg_id: str) -> dict:
    return service.users().messages().get(userId="me", id=msg_id, format="full").execute()


def render_many_one_page_pdfs(
    html_by_order: List[Tuple[str, Path]], out_dir: Path
) -> List[Tuple[str, Path]]:
    from playwright.sync_api import sync_playwright

    a4_w, a4_h = 794, 1123
    margin_mm = 8
    margin_px = int(margin_mm * 96 / 25.4)
    printable_w = max(1, a4_w - 2 * margin_px)
    printable_h = max(1, a4_h - 2 * margin_px)

    rendered: List[Tuple[str, Path]] = []
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": a4_w, "height": a4_h})
        page.emulate_media(media="print")

        for i, (oid, html_path) in enumerate(html_by_order, 1):
            if i == 1 or i % 25 == 0:
                log(f"PDF render {i}/{len(html_by_order)} (order {oid})")

            pdf_path = out_dir / f"order-{oid}.pdf"
            page.goto(html_path.resolve().as_uri(), wait_until="load")

            scroll_w = page.evaluate("() => document.documentElement.scrollWidth")
            scroll_h = page.evaluate("() => document.documentElement.scrollHeight")
            scale = 1.0
            if scroll_w and scroll_h:
                scale = max(0.2, min(1.0, printable_w / float(scroll_w), printable_h / float(scroll_h)))

            page.pdf(
                path=str(pdf_path),
                format="A4",
                print_background=True,
                scale=scale,
                margin={
                    "top": f"{margin_mm}mm",
                    "bottom": f"{margin_mm}mm",
                    "left": f"{margin_mm}mm",
                    "right": f"{margin_mm}mm",
                },
            )
            rendered.append((oid, pdf_path))

        browser.close()
    return rendered


def merge_pdfs(pdfs: List[Path], out_pdf: Path) -> None:
    from pypdf import PdfReader, PdfWriter

    writer = PdfWriter()
    for p in pdfs:
        for page in PdfReader(str(p)).pages:
            writer.add_page(page)
    with open(out_pdf, "wb") as f:
        writer.write(f)


# ── Flask routes ──────────────────────────────────────────────────────────────

@app.get("/health")
def health():
    return jsonify({"ok": True})


@app.post("/render")
def render():
    data = request.get_json(silent=True) or {}
    raw_ids = data.get("order_ids", [])
    after_date = data.get("after") or GMAIL_AFTER_DATE

    order_ids = [str(oid).strip() for oid in raw_ids if re.fullmatch(r"\d+", str(oid).strip())]
    if not order_ids:
        return jsonify({"error": "No valid order IDs provided"}), 400

    log(f"Render request: {len(order_ids)} orders, after={after_date}")

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        html_dir = tmp / "html"
        pdf_dir = tmp / "pdf"
        html_dir.mkdir()
        pdf_dir.mkdir()

        try:
            service = authenticate_gmail()
        except Exception as exc:
            log(f"Gmail auth error: {exc}")
            return jsonify({"error": f"Gmail authentication failed: {exc}"}), 500

        mapping = map_order_to_message_id(
            service,
            order_ids,
            email_address=GMAIL_EMAIL,
            after_date=after_date,
            scope="anywhere",
            include_address_filter=True,
        )

        missing = [oid for oid in order_ids if oid not in mapping]

        html_by_order: List[Tuple[str, Path]] = []
        for oid in order_ids:
            if oid not in mapping:
                continue
            try:
                msg = fetch_message_full(service, mapping[oid])
            except HttpError as exc:
                log(f"Failed to fetch message for order {oid}: {exc}")
                missing.append(oid)
                continue

            payload = msg.get("payload") or {}
            subject = extract_header(payload, "Subject") or f"Order #{oid}"
            html = find_html_part(payload)
            full_html = wrap_as_html(subject, html) if html else wrap_text_as_html(subject, find_text_part(payload) or "")
            html_path = html_dir / f"order-{oid}.html"
            html_path.write_text(full_html, encoding="utf-8", errors="replace")
            html_by_order.append((oid, html_path))

        if not html_by_order:
            return jsonify({"error": "No invoices found", "missing": missing}), 404

        rendered = render_many_one_page_pdfs(html_by_order, pdf_dir)
        pdf_map = {oid: p for oid, p in rendered}
        pdfs_in_order = [pdf_map[oid] for oid in order_ids if oid in pdf_map]

        combined_path = tmp / "combined.pdf"
        merge_pdfs(pdfs_in_order, combined_path)
        pdf_bytes = combined_path.read_bytes()

    response = make_response(pdf_bytes)
    response.headers["Content-Type"] = "application/pdf"
    response.headers["X-Missing-Orders"] = ",".join(missing)
    response.headers["X-Found-Count"] = str(len(pdfs_in_order))
    response.headers["X-Total-Count"] = str(len(order_ids))
    log(f"Done: {len(pdfs_in_order)} PDFs, {len(missing)} missing")
    return response


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, threaded=False)
