#!/usr/bin/env python3
"""
Export one Gmail message (matched by order ID in subject) to HTML + PDF.

Example:
  python export_one_order.py --order 3536793 --after 2025/11/01
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
from typing import Optional

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

    parts = payload.get("parts") or []
    for p in parts:
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

    parts = payload.get("parts") or []
    for p in parts:
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


def search_message_id_for_order(service, order_id: str, after_date: str, email_address: str) -> Optional[str]:
    """
    Find a message whose subject contains #<order_id> (common: 'Your Particle order #3536793 receipt').
    """
    q = (
        f'in:inbox (from:{email_address} OR to:{email_address}) '
        f"after:{after_date} "
        f'subject:("#{order_id}" OR "#{order_id} ")'
    )
    try:
        resp = service.users().messages().list(userId="me", q=q, maxResults=5).execute()
    except HttpError as e:
        log(f"ERROR searching Gmail: {e}")
        return None

    msgs = resp.get("messages") or []
    if not msgs:
        return None
    return msgs[0].get("id")


def fetch_message_full(service, msg_id: str) -> dict:
    return (
        service.users()
        .messages()
        .get(userId="me", id=msg_id, format="full")
        .execute()
    )


def render_pdf_with_playwright(html_file: Path, pdf_file: Path) -> None:
    try:
        from playwright.sync_api import sync_playwright
    except Exception as e:
        raise RuntimeError(
            "Playwright is not installed. Install with: pip install playwright && python -m playwright install chromium"
        ) from e

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        page.goto(html_file.resolve().as_uri(), wait_until="load")
        page.pdf(path=str(pdf_file), format="A4", print_background=True)
        browser.close()

def render_pdf_one_page_with_playwright(
    html_file: Path,
    pdf_file: Path,
    mode: str = "fit",
    margin_mm: int = 8,
) -> None:
    """
    Render to a single A4 PDF page.

    mode:
      - fit: auto-scale down to fit width/height into one page (best effort, keeps all content)
      - clip: force one page by clipping overflow (may cut off bottom)
      - first: force only the first page (guaranteed 1 page; may omit content)
    """
    try:
        from playwright.sync_api import sync_playwright
    except Exception as e:
        raise RuntimeError(
            "Playwright is not installed. Install with: pip install playwright && python -m playwright install chromium"
        ) from e

    # A4 at ~96 DPI
    a4_width_px = 794
    a4_height_px = 1123
    margin_px = int(margin_mm * 96 / 25.4)
    printable_w = max(1, a4_width_px - 2 * margin_px)
    printable_h = max(1, a4_height_px - 2 * margin_px)

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": a4_width_px, "height": a4_height_px})
        page.emulate_media(media="print")
        page.goto(html_file.resolve().as_uri(), wait_until="load")

        if mode == "clip":
            # Force a single page by clipping overflow.
            page.add_style_tag(
                content=f"""
@page {{ size: A4; margin: {margin_mm}mm; }}
html, body {{ height: 297mm; overflow: hidden !important; }}
"""
            )

        scale = 1.0
        if mode == "fit":
            # Best-effort fit: measure content size and scale down so it fits inside one page.
            scroll_w = page.evaluate("() => document.documentElement.scrollWidth")
            scroll_h = page.evaluate("() => document.documentElement.scrollHeight")
            if scroll_w and scroll_h:
                w_scale = printable_w / float(scroll_w)
                h_scale = printable_h / float(scroll_h)
                scale = min(1.0, w_scale, h_scale)
                # Avoid unreadably tiny PDFs; clamp to a sane minimum
                scale = max(0.2, scale)

        pdf_kwargs = dict(
            path=str(pdf_file),
            format="A4",
            print_background=True,
            margin={"top": f"{margin_mm}mm", "bottom": f"{margin_mm}mm", "left": f"{margin_mm}mm", "right": f"{margin_mm}mm"},
        )

        if mode == "fit":
            pdf_kwargs["scale"] = scale
        if mode == "first":
            pdf_kwargs["page_ranges"] = "1"

        page.pdf(**pdf_kwargs)
        browser.close()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--order", required=True, help="Order ID (digits)")
    parser.add_argument("--after", default="2025/11/01", help="Gmail after: date, e.g. 2025/11/01")
    parser.add_argument("--email", default="customerinvoices@particleformen.com", help="Mailbox address")
    parser.add_argument("--out", default="out", help="Output directory")
    parser.add_argument(
        "--one-page",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Render a single-page A4 PDF (default: true). Use --no-one-page for normal multi-page PDF.",
    )
    parser.add_argument(
        "--one-page-mode",
        choices=["fit", "clip", "first"],
        default="fit",
        help="Single-page strategy: fit (scale down), clip (crop overflow), first (only first page).",
    )
    args = parser.parse_args()

    order_id = str(args.order).strip()
    if not re.fullmatch(r"\d+", order_id):
        log("ERROR: --order must be digits only")
        return 2

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    html_path = out_dir / f"order-{order_id}.html"
    pdf_path = out_dir / f"order-{order_id}.pdf"

    try:
        log("Authenticating...")
        service = authenticate_gmail()
        log("[OK] Authenticated.")
    except Exception as e:
        log(f"ERROR: auth failed: {e}")
        return 1

    log(f"Searching for order #{order_id} after {args.after}...")
    msg_id = search_message_id_for_order(service, order_id, args.after, args.email)
    if not msg_id:
        log("No matching message found.")
        return 1

    log(f"Fetching message {msg_id}...")
    msg = fetch_message_full(service, msg_id)
    payload = msg.get("payload") or {}
    subject = extract_header(payload, "Subject") or f"Order #{order_id}"

    html = find_html_part(payload)
    if html:
        full_html = wrap_as_html(subject, html)
    else:
        text = find_text_part(payload) or ""
        full_html = wrap_text_as_html(subject, text)

    html_path.write_text(full_html, encoding="utf-8", errors="replace")
    log(f"[OK] Wrote HTML: {html_path}")

    try:
        log("Rendering PDF via Playwright (Chromium)...")
        if args.one_page:
            render_pdf_one_page_with_playwright(html_path, pdf_path, mode=args.one_page_mode)
        else:
            render_pdf_with_playwright(html_path, pdf_path)
        log(f"[OK] Wrote PDF: {pdf_path}")
    except Exception as e:
        log(f"PDF render failed: {e}")
        log("You can still open the HTML file and print-to-PDF manually.")
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())


