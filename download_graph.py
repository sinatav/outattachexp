#!/usr/bin/env python3
"""Interactive Outlook attachment downloader via Microsoft Graph (device code).

Uses the public 'Microsoft Graph Command Line Tools' client ID, which most
M365 tenants have pre-consented. On first run you'll get a code to paste at
https://microsoft.com/devicelogin in a browser.

Standard library only.
"""

from __future__ import annotations

import base64
import getpass  # noqa: F401  (kept for parity with imap version; unused)
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta, timezone
from pathlib import Path


# Public, pre-consented in most tenants. No secret needed.
# Override with env var OUTLOOK_DL_CLIENT_ID to try a different first-party app.
# Known candidates:
#   04b07795-8ddb-461a-bbee-02f9e1bf7b46  Azure CLI         (default below)
#   14d82eec-204b-4c2f-b7e8-296a70dab67e  Microsoft Graph CLI
#   d3590ed6-52b3-4102-aeff-aad2292ab01c  Microsoft Office  (very widely allowed)
CLIENT_ID = os.environ.get(
    "OUTLOOK_DL_CLIENT_ID",
    "04b07795-8ddb-461a-bbee-02f9e1bf7b46",
)
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = "Mail.Read offline_access"
GRAPH = "https://graph.microsoft.com/v1.0"

TOKEN_CACHE = Path.home() / ".cache" / "outlook-attachments" / "token.json"


# ----- HTTP helpers ---------------------------------------------------------


def http_post_form(url: str, data: dict) -> dict:
    body = urllib.parse.urlencode(data).encode()
    req = urllib.request.Request(url, data=body, method="POST")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")
    try:
        with urllib.request.urlopen(req) as r:
            return json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        try:
            return json.loads(e.read().decode())
        except Exception:
            raise


def http_get(url: str, token: str) -> dict:
    req = urllib.request.Request(url, method="GET")
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Accept", "application/json")
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read().decode())


# ----- Auth -----------------------------------------------------------------


def load_cached_token() -> dict | None:
    if not TOKEN_CACHE.exists():
        return None
    try:
        return json.loads(TOKEN_CACHE.read_text())
    except Exception:
        return None


def save_token(tok: dict) -> None:
    TOKEN_CACHE.parent.mkdir(parents=True, exist_ok=True)
    TOKEN_CACHE.write_text(json.dumps(tok))
    try:
        os.chmod(TOKEN_CACHE, 0o600)
    except OSError:
        pass


def refresh_token(refresh: str) -> dict | None:
    resp = http_post_form(
        f"{AUTHORITY}/oauth2/v2.0/token",
        {
            "client_id": CLIENT_ID,
            "grant_type": "refresh_token",
            "refresh_token": refresh,
            "scope": SCOPES,
        },
    )
    if "access_token" in resp:
        resp["obtained_at"] = int(time.time())
        return resp
    return None


def device_code_login() -> dict:
    init = http_post_form(
        f"{AUTHORITY}/oauth2/v2.0/devicecode",
        {"client_id": CLIENT_ID, "scope": SCOPES},
    )
    if "user_code" not in init:
        raise RuntimeError(f"Device code init failed: {init}")

    print("\n" + "=" * 60)
    print(init.get("message", ""))
    print("=" * 60 + "\n")

    interval = int(init.get("interval", 5))
    expires_in = int(init.get("expires_in", 900))
    deadline = time.time() + expires_in

    while time.time() < deadline:
        time.sleep(interval)
        resp = http_post_form(
            f"{AUTHORITY}/oauth2/v2.0/token",
            {
                "client_id": CLIENT_ID,
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                "device_code": init["device_code"],
            },
        )
        if "access_token" in resp:
            resp["obtained_at"] = int(time.time())
            return resp
        err = resp.get("error")
        if err == "authorization_pending":
            continue
        if err == "slow_down":
            interval += 5
            continue
        if err in ("authorization_declined", "expired_token", "bad_verification_code"):
            raise RuntimeError(f"Auth failed: {resp}")
        # Anything else (e.g. consent_required) — surface it.
        raise RuntimeError(f"Auth failed: {resp}")

    raise RuntimeError("Device code expired before sign-in.")


def get_access_token() -> str:
    cached = load_cached_token()
    if cached:
        age = int(time.time()) - int(cached.get("obtained_at", 0))
        # Use cached access token if it's still fresh (Graph default ~1h).
        if age < int(cached.get("expires_in", 3600)) - 300 and "access_token" in cached:
            return cached["access_token"]
        # Try refresh.
        if "refresh_token" in cached:
            refreshed = refresh_token(cached["refresh_token"])
            if refreshed:
                save_token(refreshed)
                return refreshed["access_token"]

    tok = device_code_login()
    save_token(tok)
    return tok["access_token"]


# ----- Input prompts --------------------------------------------------------


def prompt(label: str, default: str = "") -> str:
    suffix = f" [{default}]" if default else ""
    value = input(f"{label}{suffix}: ").strip()
    return value or default


def parse_date(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()


# ----- Graph helpers --------------------------------------------------------


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


def build_filter(since: date | None, before: date | None, sender: str) -> str:
    parts: list[str] = ["hasAttachments eq true"]
    if since:
        dt = datetime.combine(since, datetime.min.time(), tzinfo=timezone.utc)
        parts.append(f"receivedDateTime ge {dt.isoformat().replace('+00:00', 'Z')}")
    if before:
        # inclusive: bump by a day
        dt = datetime.combine(before + timedelta(days=1), datetime.min.time(), tzinfo=timezone.utc)
        parts.append(f"receivedDateTime lt {dt.isoformat().replace('+00:00', 'Z')}")
    if sender:
        # If it looks like an address, exact match. Otherwise substring on name.
        if "@" in sender:
            parts.append(f"from/emailAddress/address eq '{sender}'")
        else:
            parts.append(f"contains(from/emailAddress/name, '{sender}')")
    return " and ".join(parts)


def iter_messages(token: str, folder: str, filt: str):
    params = {
        "$filter": filt,
        "$select": "id,subject,from,receivedDateTime,hasAttachments",
        "$top": "50",
        "$orderby": "receivedDateTime desc",
    }
    url = f"{GRAPH}/me/mailFolders/{folder}/messages?{urllib.parse.urlencode(params)}"
    while url:
        page = http_get(url, token)
        for m in page.get("value", []):
            yield m
        url = page.get("@odata.nextLink")


def list_attachments(token: str, message_id: str) -> list[dict]:
    url = f"{GRAPH}/me/messages/{message_id}/attachments?$select=id,name,contentType,size,isInline,@odata.type"
    resp = http_get(url, token)
    return resp.get("value", [])


def fetch_attachment_bytes(token: str, message_id: str, attachment_id: str) -> bytes:
    # $value returns raw bytes for fileAttachment
    url = f"{GRAPH}/me/messages/{message_id}/attachments/{attachment_id}/$value"
    req = urllib.request.Request(url, method="GET")
    req.add_header("Authorization", f"Bearer {token}")
    with urllib.request.urlopen(req) as r:
        return r.read()


# ----- Main -----------------------------------------------------------------


def main() -> int:
    print("Outlook attachment downloader (Microsoft Graph)")
    print("-" * 50)

    folder = prompt("Folder (well-known: inbox, sentitems, archive)", "inbox")
    since_s = prompt("From date (YYYY-MM-DD, blank for no lower bound)")
    before_s = prompt("To date   (YYYY-MM-DD, blank for no upper bound)")
    sender = prompt("Sender filter (email = exact, text = substring on name; blank = any)")
    exts_s = prompt("File extensions (comma-separated, e.g. pdf,xlsx; blank = all)")
    dest_s = prompt("Download destination", str(Path.home() / "Downloads" / "outlook-attachments"))

    since = parse_date(since_s) if since_s else None
    before = parse_date(before_s) if before_s else None
    exts = {e.strip().lower().lstrip(".") for e in exts_s.split(",") if e.strip()}
    dest = Path(dest_s).expanduser()
    dest.mkdir(parents=True, exist_ok=True)

    print("\nSigning in...")
    try:
        token = get_access_token()
    except Exception as e:
        print(f"Auth failed: {e}", file=sys.stderr)
        return 2

    filt = build_filter(since, before, sender)
    print(f"\nQuery filter: {filt}")
    print(f"Folder: {folder}\n")

    saved = 0
    scanned = 0
    try:
        for msg in iter_messages(token, folder, filt):
            scanned += 1
            subject = (msg.get("subject") or "")[:80]
            sender_name = ((msg.get("from") or {}).get("emailAddress") or {}).get("address", "")
            if not msg.get("hasAttachments"):
                continue
            for att in list_attachments(token, msg["id"]):
                if att.get("isInline"):
                    continue
                otype = att.get("@odata.type", "")
                if "fileAttachment" not in otype:
                    # itemAttachment / referenceAttachment — skip
                    continue
                name = att.get("name") or "attachment.bin"
                if not matches_extension(name, exts):
                    continue
                data = fetch_attachment_bytes(token, msg["id"], att["id"])
                out = unique_path(dest, safe_filename(name))
                out.write_bytes(data)
                saved += 1
                print(f"  [{scanned}] {sender_name} | {subject!r}")
                print(f"        -> {out}")
    except urllib.error.HTTPError as e:
        body = e.read().decode(errors="replace")
        print(f"\nGraph error {e.code}: {body}", file=sys.stderr)
        return 3

    print(f"\nDone. Scanned {scanned} message(s), saved {saved} file(s) to {dest}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nInterrupted.", file=sys.stderr)
        sys.exit(130)
