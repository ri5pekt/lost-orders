#!/usr/bin/env python3
"""
Gmail Order Extractor
Extracts order IDs from emails in customerinvoices@particleformen.com inbox
"""

import os
import json
import re
import pickle
import sys
import time
from pathlib import Path
from typing import List, Set
from datetime import datetime

try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except ImportError:
    print("Missing required packages. Please install:")
    print("pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client")
    exit(1)

# Gmail API scopes
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

# Token file to store credentials
TOKEN_FILE = 'token.pickle'
CREDENTIALS_FILE = 'credentials.json'
RUN_LOG_FILE = 'run.log'


def log(message: str) -> None:
    """Log to stdout (flush) and to a local file (run.log)."""
    ts = time.strftime('%Y-%m-%d %H:%M:%S')
    line = f"[{ts}] {message}"
    print(line, flush=True)
    try:
        with open(RUN_LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        # Never fail the run due to logging
        pass


def authenticate_gmail():
    """Authenticate and return Gmail service object"""
    creds = None

    # Load existing token if available
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                print(f"Error: {CREDENTIALS_FILE} not found!")
                print("Please download OAuth2 credentials from Google Cloud Console:")
                print("1. Go to https://console.cloud.google.com/")
                print("2. Create/select a project")
                print("3. Enable Gmail API")
                print("4. Create OAuth 2.0 Client ID credentials")
                print("5. Download as JSON and save as 'credentials.json'")
                return None

            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        # Save credentials for next run
        with open(TOKEN_FILE, 'wb') as token:
            pickle.dump(creds, token)

    try:
        service = build('gmail', 'v1', credentials=creds)
        return service
    except HttpError as error:
        print(f'An error occurred: {error}')
        return None


def search_emails_paginated(service, query: str, batch_size: int = 500):
    """Search for emails with pagination support"""
    messages = []
    page_token = None
    page_num = 0

    try:
        while True:
            page_num += 1
            # Print as normal lines (no carriage return) so it reliably shows in terminals
            if page_num == 1 or page_num % 10 == 0:
                log(f"    Fetching page {page_num} (batch_size={batch_size})")

            request_params = {
                'userId': 'me',
                'q': query,
                'maxResults': batch_size
            }

            if page_token:
                request_params['pageToken'] = page_token

            results = service.users().messages().list(**request_params).execute()
            batch = results.get('messages', [])
            messages.extend(batch)

            page_token = results.get('nextPageToken')
            if not page_token:
                log(f"    Fetched {len(messages)} total emails from {page_num} pages")
                break

    except HttpError as error:
        log(f"    ERROR searching: {error}")
    except Exception as e:
        log(f"    ERROR searching (unexpected): {e}")

    return messages


def get_email_content(service, msg_id: str, metadata_only: bool = False) -> dict:
    """Get email content or metadata only"""
    try:
        params = {
            'userId': 'me',
            'id': msg_id,
        }

        if metadata_only:
            params['format'] = 'metadata'
            params['metadataHeaders'] = ['Subject']
        else:
            params['format'] = 'full'

        message = service.users().messages().get(**params).execute()
        return message
    except HttpError as error:
        print(f'An error occurred while fetching email: {error}')
        return {}


def get_email_subject(email_content: dict) -> str:
    """Extract subject from email"""
    payload = email_content.get('payload', {})
    headers = payload.get('headers', [])

    for header in headers:
        if header.get('name', '').lower() == 'subject':
            return header.get('value', '')
    return ''


def extract_order_id_from_subject(subject: str) -> str:
    """Extract order ID from subject line (e.g., 'Your Particle order #3536793 receipt')"""
    if not subject:
        return None

    # Pattern to match order ID in subject: "order #123456" or "order # 123456"
    patterns = [
        r'order\s*#\s*(\d+)',  # order #123456
        r'order\s+#(\d+)',     # order #123456 (no space)
        r'#(\d+)',             # #123456 (fallback)
    ]

    for pattern in patterns:
        match = re.search(pattern, subject, re.IGNORECASE)
        if match:
            return match.group(1)

    return None


def load_existing_order_ids(filename: str = 'lost-orders-woo.txt') -> Set[str]:
    """Load existing order IDs from file"""
    if not os.path.exists(filename):
        return set()

    with open(filename, 'r') as f:
        return {line.strip() for line in f if line.strip()}


def build_order_id_query_chunks(
    order_ids: Set[str],
    email_address: str,
    after_date: str,
    chunk_size: int = 50,
    max_query_length: int = 2000,
) -> List[str]:
    """Build Gmail search query chunks with order IDs in subject filter"""
    log(f"  Building query chunks from {len(order_ids)} order IDs...")
    order_ids_list = sorted(order_ids)
    chunks = []
    current_chunk = []
    processed_count = 0

    for order_id in order_ids_list:
        processed_count += 1
        if processed_count % 500 == 0:
            log(f"    Processing order IDs: {processed_count}/{len(order_ids_list)}")

        # Add order ID to current chunk
        current_chunk.append(order_id)

        # Build query to check length
        subject_part = ' OR '.join([f'#{oid}' for oid in current_chunk])
        query = f'in:inbox (from:{email_address} OR to:{email_address}) after:{after_date} subject:({subject_part})'

        # If chunk is full or query too long, save current chunk and start new one
        if len(current_chunk) >= chunk_size or len(query) > max_query_length:
            # Remove the last added ID and save the chunk without it
            current_chunk.pop()
            if current_chunk:
                subject_part = ' OR '.join([f'#{oid}' for oid in current_chunk])
                query = f'in:inbox (from:{email_address} OR to:{email_address}) after:{after_date} subject:({subject_part})'
                chunks.append(query)
            # Start new chunk with the ID we just popped
            current_chunk = [order_id]

    # Add the last chunk
    if current_chunk:
        subject_part = ' OR '.join([f'#{oid}' for oid in current_chunk])
        query = f'in:inbox (from:{email_address} OR to:{email_address}) after:{after_date} subject:({subject_part})'
        chunks.append(query)

    log(f"    Built {len(chunks)} query chunks from {len(order_ids_list)} order IDs")
    return chunks


def filter_emails_by_order_ids(service, email_address: str, existing_order_ids: Set[str], after_date: str = '2025/11/01') -> Set[str]:
    """Filter emails that contain order IDs from our list in their subjects using Gmail search"""
    log(f"Searching emails from/to {email_address} after {after_date}...")
    log(f"Building search queries for {len(existing_order_ids)} order IDs...")

    # Build query chunks with order IDs in subject filter
    query_chunks = build_order_id_query_chunks(existing_order_ids, email_address, after_date)
    log(f"Created {len(query_chunks)} query chunks to search")

    all_messages = []

    # Execute each query chunk
    log(f"Executing {len(query_chunks)} search queries...")
    for i, query in enumerate(query_chunks, 1):
        log(f"[Chunk {i}/{len(query_chunks)}] Searching...")
        messages = search_emails_paginated(service, query, batch_size=500)
        all_messages.extend(messages)
        log(f"  [OK] Found {len(messages)} emails in this chunk (total so far: {len(all_messages)})")

    if not all_messages:
        log("No emails found.")
        return set()

    log(f"Found {len(all_messages)} total matching emails. Extracting order IDs from subjects...")

    matched_order_ids = set()
    error_count = 0

    log(f"Processing {len(all_messages)} emails to extract order IDs...")
    for i, msg in enumerate(all_messages, 1):
        if i == 1 or i % 200 == 0:
            progress_pct = (i / len(all_messages)) * 100
            log(f"  Progress {i}/{len(all_messages)} ({progress_pct:.1f}%) | matched={len(matched_order_ids)} | errors={error_count}")

        # Get email metadata (subject only, not full content for performance)
        try:
            email_content = get_email_content(service, msg['id'], metadata_only=True)
            subject = get_email_subject(email_content)

            # Extract order ID from subject
            order_id = extract_order_id_from_subject(subject)

            if order_id and order_id in existing_order_ids:
                matched_order_ids.add(order_id)

        except HttpError as e:
            error_count += 1
            if error_count % 100 == 0:
                log(f"  WARNING: {error_count} HTTP errors encountered. Continuing...")
            continue
        except Exception as e:
            error_count += 1
            if error_count % 100 == 0:
                log(f"  WARNING: {error_count} errors encountered. Continuing...")
            continue

    log(f"Finished processing all {len(all_messages)} emails")
    log(f"Matched {len(matched_order_ids)} unique order IDs from our list")

    return matched_order_ids


def save_order_ids(order_ids: Set[str], filename: str = 'gmail-orders.txt'):
    """Save order IDs to a file"""
    sorted_ids = sorted(order_ids, key=lambda x: int(x) if x.isdigit() else 0, reverse=True)

    with open(filename, 'w') as f:
        for order_id in sorted_ids:
            f.write(f"{order_id}\n")

    print(f"\nSaved {len(sorted_ids)} unique order IDs to {filename}")


def main():
    """Main function"""
    start_time = time.time()

    # Make stdout more reliable on Windows consoles
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass

    log("=" * 60)
    log("Gmail Order Extractor")
    log("=" * 60)
    log(f"Started at: {time.strftime('%Y-%m-%d %H:%M:%S')}")

    email_address = 'customerinvoices@particleformen.com'

    # Load existing order IDs
    log("[1/4] Loading existing order IDs from lost-orders-woo.txt...")
    existing_order_ids = load_existing_order_ids('lost-orders-woo.txt')
    log(f"  [OK] Loaded {len(existing_order_ids)} existing order IDs")

    if not existing_order_ids:
        log("  [WARNING] No existing order IDs found. Will process all emails.")

    log("[2/4] Authenticating with Gmail API...")
    service = authenticate_gmail()

    if not service:
        log("  [ERROR] Authentication failed. Please check your credentials.")
        return

    log("  [OK] Authentication successful!")

    # Filter emails that contain order IDs from our list (orders created after Nov 1, 2025)
    log("[3/4] Searching and filtering emails...")
    matched_order_ids = filter_emails_by_order_ids(service, email_address, existing_order_ids, after_date='2025/11/01')

    log("[4/4] Saving results...")
    if matched_order_ids:
        log(f"  [OK] Found {len(matched_order_ids)} order IDs in email subjects")

        # Save matched order IDs
        output_file = 'gmail-matched-orders.txt'
        save_order_ids(matched_order_ids, output_file)

        # These are already in our list, so we just confirm they exist
        log("  [OK] All matched order IDs are already in lost-orders-woo.txt")
        log(f"  [OK] This confirms {len(matched_order_ids)} orders have corresponding emails in Gmail")
    else:
        log("  [ERROR] No matching emails found.")

    elapsed_time = time.time() - start_time
    log("=" * 60)
    log(f"Completed in {elapsed_time:.1f} seconds ({elapsed_time/60:.1f} minutes)")
    log(f"Finished at: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    log("=" * 60)


if __name__ == '__main__':
    main()

