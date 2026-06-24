#!/usr/bin/env python3
"""Interactive Outlook attachment downloader via IMAP.

Works with outlook.com personal accounts and Microsoft 365 accounts that have
IMAP enabled. Requires an app password (not your regular login password):
  - Personal: https://account.live.com/proofs/AppPassword
  - M365:     account admin must enable IMAP + app passwords for your user

Standard library only.
"""

from __future__ import annotations

import email
import getpass
import imaplib
import os
import re
import sys
from datetime import date, datetime
from email.header import decode_header
from email.message import Message
from pathlib import Path


IMAP_HOST = "outlook.office365.com"
IMAP_PORT = 993


def prompt(label: str, default: str = "") -> str:
    suffix = f" [{default}]" if default else ""
    value = input(f"{label}{suffix}: ").strip()
    return value or default


def parse_date(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()


def imap_date(d: date) -> str:
    # IMAP wants e.g. 01-Jan-2026
    return d.strftime("%d-%b-%Y")


def decode_mime(value: str | None) -> str:
    if not value:
        return ""
    parts = decode_header(value)
    out = []
    for text, charset in parts:
        if isinstance(text, bytes):
            try:
                out.append(text.decode(charset or "utf-8", errors="replace"))
            except LookupError:
                out.append(text.decode("utf-8", errors="replace"))
        else:
            out.append(text)
    return "".join(out)


SAFE_NAME = re.compile(r"[^A-Za-z0-9._\- ]+")


def safe_filename(name: str) -> str:
    name = SAFE_NAME.sub("_", name).strip().strip(".")
    return name or "attachment"


def unique_path(dest: Path, name: str) -> Path:
    path = dest / name
    if not path.exists():
        return path
    stem, ext = os.path.splitext(name)
    i = 1
    while True:
        candidate = dest / f"{stem} ({i}){ext}"
        if not candidate.exists():
            return candidate
        i += 1


def matches_extension(filename: str, exts: set[str]) -> bool:
    if not exts:
        return True
    ext = os.path.splitext(filename)[1].lower().lstrip(".")
    return ext in exts


def iter_attachments(msg: Message):
    for part in msg.walk():
        if part.is_multipart():
            continue
        disp = (part.get("Content-Disposition") or "").lower()
        filename = part.get_filename()
        # Treat anything with an explicit filename or "attachment" disposition.
        if not filename and "attachment" not in disp:
            continue
        if filename:
            filename = decode_mime(filename)
        else:
            filename = "attachment.bin"
        yield filename, part


def build_search(since: date | None, before: date | None, sender: str) -> list[str]:
    parts: list[str] = []
    if since:
        parts += ["SINCE", imap_date(since)]
    if before:
        # BEFORE is exclusive, so add one day's worth of inclusive intent by
        # using the day AFTER the user's "to" date.
        from datetime import timedelta
        parts += ["BEFORE", imap_date(before + timedelta(days=1))]
    if sender:
        parts += ["FROM", f'"{sender}"']
    if not parts:
        parts = ["ALL"]
    return parts


def main() -> int:
    print("Outlook attachment downloader (IMAP)")
    print("-" * 40)

    email_addr = prompt("Email address")
    if not email_addr:
        print("Email is required.", file=sys.stderr)
        return 1

    password = getpass.getpass("App password (input hidden): ")
    if not password:
        print("Password is required.", file=sys.stderr)
        return 1

    folder = prompt("Folder", "INBOX")

    since_s = prompt("From date (YYYY-MM-DD, blank for no lower bound)")
    before_s = prompt("To date   (YYYY-MM-DD, blank for no upper bound)")
    sender = prompt("Sender filter (e.g. boss@example.com or substring; blank = any)")
    exts_s = prompt("File extensions (comma-separated, e.g. pdf,xlsx; blank = all)")
    dest_s = prompt("Download destination", str(Path.home() / "Downloads" / "outlook-attachments"))

    since = parse_date(since_s) if since_s else None
    before = parse_date(before_s) if before_s else None
    exts = {e.strip().lower().lstrip(".") for e in exts_s.split(",") if e.strip()}
    dest = Path(dest_s).expanduser()
    dest.mkdir(parents=True, exist_ok=True)

    search_parts = build_search(since, before, sender)

    print(f"\nConnecting to {IMAP_HOST}...")
    with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT) as M:
        try:
            M.login(email_addr, password)
        except imaplib.IMAP4.error as e:
            print(f"Login failed: {e}", file=sys.stderr)
            print("Notes: outlook.com personal needs an app password; M365 tenants must "
                  "have IMAP enabled.", file=sys.stderr)
            return 2

        typ, _ = M.select(folder, readonly=True)
        if typ != "OK":
            print(f"Cannot select folder {folder!r}", file=sys.stderr)
            return 3

        typ, data = M.search(None, *search_parts)
        if typ != "OK":
            print(f"Search failed: {data}", file=sys.stderr)
            return 4

        ids = data[0].split()
        print(f"Matched {len(ids)} message(s). Scanning for attachments...\n")

        saved = 0
        for n, msg_id in enumerate(ids, 1):
            typ, msg_data = M.fetch(msg_id, "(RFC822)")
            if typ != "OK" or not msg_data or not msg_data[0]:
                continue
            raw = msg_data[0][1]
            if not isinstance(raw, (bytes, bytearray)):
                continue
            msg = email.message_from_bytes(raw)
            subject = decode_mime(msg.get("Subject", ""))[:80]
            from_hdr = decode_mime(msg.get("From", ""))

            for filename, part in iter_attachments(msg):
                if not matches_extension(filename, exts):
                    continue
                payload = part.get_payload(decode=True)
                if payload is None:
                    continue
                out = unique_path(dest, safe_filename(filename))
                out.write_bytes(payload)
                saved += 1
                print(f"  [{n}/{len(ids)}] {from_hdr} | {subject!r}")
                print(f"        -> {out}")

        print(f"\nDone. Saved {saved} file(s) to {dest}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nInterrupted.", file=sys.stderr)
        sys.exit(130)
