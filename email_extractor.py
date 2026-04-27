"""
Outlook Email Extractor for HR Recruitment (Volibits)
Reads emails via Microsoft Graph API and inserts candidate data into PostgreSQL.

recruiter        = local part of sender email   — sender is always @volibits.com
client_recruiter = local part of receiver email — can be any domain

Duplicate handling:
  1. Same recruiter + same key fields → auto-fill only NULL/empty fields in the
     existing record; never overwrite populated values.
  2. Different recruiter + same candidate/JR → insert record immediately and
     send informational notification to both recruiters.

OneDrive auth:
  - Uses a DELEGATED refresh token (OD_REFRESH_TOKEN) obtained once via
    generate_refresh_token.py and stored as a GitHub Secret.
  - This avoids requiring Files.ReadWrite.All (application permission).
  - The token is scoped to ONEDRIVE_USER's personal OneDrive only.
  - Refresh tokens expire after 90 days — re-run generate_refresh_token.py
    before they expire to obtain a new one.

OneDrive notes:
  - Only .pdf / .doc / .docx attachments are uploaded.
  - Drive uploads use ONEDRIVE_USER (your personal email).
  - The HR Resumes folder is auto-created on first run if it doesn't exist.
"""

import os
import re
import logging
import requests
import psycopg2
from psycopg2 import sql as pgsql
from html import unescape
from datetime import datetime, date
from typing import Optional
from msal import ConfidentialClientApplication, PublicClientApplication
from notifier import (
    send_notification_email,
    send_diff_recruiter_notification_email,
)

# ─── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("extractor.log"),
    ],
)
log = logging.getLogger(__name__)


# ─── Configuration ─────────────────────────────────────────────────────────────
AZURE_TENANT_ID     = os.environ["AZURE_TENANT_ID"]
AZURE_CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
AZURE_CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]

OD_TENANT_ID        = os.environ["OD_TENANT_ID"]
OD_CLIENT_ID        = os.environ["OD_CLIENT_ID"]
OD_CLIENT_SECRET    = os.environ.get("OD_CLIENT_SECRET", "")
OD_REFRESH_TOKEN    = os.environ["OD_REFRESH_TOKEN"]

TARGET_MAILBOX      = os.environ["TARGET_MAILBOX"]
ONEDRIVE_USER       = os.environ["ONEDRIVE_USER"]

DB_DSN              = os.environ["DB_DSN"]
DB_TABLE            = os.environ.get("DB_TABLE_NAME", "hrvolibit")
ONEDRIVE_FOLDER     = os.environ.get("ONEDRIVE_FOLDER", "HR Resumes")

# Subfolder inside the target mailbox to read candidate emails from.
# Set to empty string "" to read from the root inbox instead.
INBOX_SUBFOLDER     = os.environ.get("INBOX_SUBFOLDER", "Company Profiles")
Limit               = int(os.environ.get("LIMIT", "100"))
VOLIBITS_DOMAIN     = "volibits.com"
RESUME_EXTENSIONS   = {".pdf", ".doc", ".docx"}

# Number of months of history to consider when checking for duplicates.
DUPLICATE_CHECK_MONTHS = int(os.environ.get("DUPLICATE_CHECK_MONTHS", "3"))

# Optional end-date filter: only emails received on or before this date are
# processed.  Format: YYYY-MM-DD.  Defaults to today when not set.
# Example: END_DATE=2025-04-20 → reads emails from today back to 2025-04-20.
_end_date_raw = os.environ.get("END_DATE", "").strip()
try:
    END_DATE: date = datetime.strptime(_end_date_raw, "%Y-%m-%d").date() if _end_date_raw else date.today()
except ValueError:
    raise SystemExit(
        f"END_DATE env var '{_end_date_raw}' is not a valid YYYY-MM-DD date."
    )

# Fields that may be auto-filled when NULL/empty in an existing record.
UPDATABLE_FIELDS = [
    "jr_no",
    "total_experience", "relevant_experience",
    "current_ctc", "expected_ctc",
    "notice_period", "current_org",
    "current_location", "preferred_location",
    "attachment", "remarks",
]


# ─── Fuzzy column-header aliases ───────────────────────────────────────────────
COLUMN_ALIASES: dict[str, list[str]] = {
    "jr_no": [
        "jr no", "jr_no", "jr no.", "jr", "jrno", "req_id", "req id",
        "requisition id", "requisition_id", "jr_number", "jr number",
        "job req", "job requisition", "job_req", "reqid",
        "requirement id", "requirement_id",
    ],
    "name_of_candidate": [
        "candidate name", "candidate_name", "name", "applicant name",
        "applicant_name", "full name", "full_name", "name of candidate",
        "candidate", "person name",
    ],
    "contact_number": [
        "contact number", "contact_number", "phone", "mobile", "cell",
        "phone number", "phone_number", "mobile number", "mobile_number",
        "contact no", "contact_no", "ph no", "ph_no", "contact",
    ],
    "email_id": [
        "email id", "email_id", "email", "e-mail", "mail id", "mail_id",
        "email address", "email_address", "e mail",
    ],
    "current_org": [
        "current company", "current_company", "company", "employer",
        "current employer", "current_employer", "current org", "current_org",
        "organisation", "organization", "curr company", "present company", "curr org"
    ],
    "total_experience": [
        "total experience", "total_experience", "total exp", "total_exp",
        "experience", "exp", "total yrs", "total years", "tot exp"
    ],
    "relevant_experience": [
        "relevant experience", "relevant_experience", "relevant exp",
        "relevant_exp", "rel exp", "rel_exp", "relevant yrs",
    ],
    "current_ctc": [
        "current ctc", "current_ctc", "ctc", "current salary",
        "current_salary", "present ctc", "present_ctc", "curr ctc", "Rates"
    ],
    "expected_ctc": [
        "expected ctc", "expected_ctc", "expected salary", "expected_salary",
        "exp ctc", "exp_ctc", "desired ctc", "desired_ctc",
    ],
    "notice_period": [
        "notice period", "notice_period", "notice", "np", "joining time",
        "notice days", "notice_days",
    ],
    "current_location": [
        "current location", "current_location", "location", "city",
        "current city", "current_city", "present location", "curr loc",
    ],
    "preferred_location": [
        "preferred location", "preferred_location", "preferred city",
        "preferred_city", "pref location", "pref_location",
        "willing to relocate", "target location", "exc loc"
    ],
    "date": [
        "date", "submission date", "submission_date", "applied date",
        "applied_date", "profile date",
    ],
    "remarks": [
        "remarks", "comments", "comment", "note", "notes",
    ],
}


# ─── Build alias lookup once at startup ────────────────────────────────────────
def _normalize(s: str) -> str:
    return re.sub(r"\s+", " ", s.strip().lower())


ALIAS_MAP: dict[str, str] = {
    _normalize(alias): canonical
    for canonical, aliases in COLUMN_ALIASES.items()
    for alias in aliases
}


def _resolve_header(raw_header: str) -> Optional[str]:
    return ALIAS_MAP.get(_normalize(raw_header))


# ─── Microsoft Graph authentication ────────────────────────────────────────────
def _get_app_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    """Client-credentials (app-only) token — used for mail reading."""
    result = ConfidentialClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
    ).acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(
            f"App token acquisition failed: {result.get('error_description')}"
        )
    return result["access_token"]


def get_mail_token() -> str:
    """App-only token for reading TARGET_MAILBOX via Mail.Read application permission."""
    return _get_app_token(AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET)


def get_onedrive_token() -> str:
    """
    Delegated token for ONEDRIVE_USER's personal OneDrive.

    Uses a stored refresh token (OD_REFRESH_TOKEN) so no interactive login
    is needed at runtime. The PublicClientApplication.acquire_token_by_refresh_token
    call exchanges the refresh token for a fresh access token each run.

    Refresh tokens last ~90 days. Re-run generate_refresh_token.py before
    expiry and update the OD_REFRESH_TOKEN GitHub Secret.
    """
    app = PublicClientApplication(
        OD_CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{OD_TENANT_ID}",
    )
    result = app.acquire_token_by_refresh_token(
        OD_REFRESH_TOKEN,
        scopes=["https://graph.microsoft.com/.default"],
    )
    if "access_token" not in result:
        error_desc = result.get("error_description", str(result))
        raise RuntimeError(
            f"OneDrive token refresh failed: {error_desc}\n"
            "If the refresh token has expired, re-run generate_refresh_token.py "
            "and update the OD_REFRESH_TOKEN GitHub Secret."
        )
    log.info("OneDrive delegated token obtained successfully.")
    return result["access_token"]


# ─── Folder resolution (mail) ──────────────────────────────────────────────────
def resolve_folder_id(token: str, folder_name: str) -> Optional[str]:
    """
    Return the Graph API folder ID for a named subfolder inside the target mailbox.
    Checks top-level mailFolders first; if not found, searches one level of
    child folders (handles Inbox > Company Profiles nesting).
    Returns None if the folder cannot be found (caller falls back to root inbox).
    """
    if not folder_name:
        return None

    headers = {"Authorization": f"Bearer {token}"}
    url = (
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/mailFolders"
        f"?$select=id,displayName&$top=50"
    )
    resp = requests.get(url, headers=headers, timeout=30)
    if not resp.ok:
        log.warning(f"resolve_folder_id: mailFolders fetch failed {resp.status_code}")
        return None

    for folder in resp.json().get("value", []):
        if folder.get("displayName", "").strip().lower() == folder_name.strip().lower():
            log.info(f"Resolved folder {folder_name!r} → id={folder['id']}")
            return folder["id"]

        child_url = (
            f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}"
            f"/mailFolders/{folder['id']}/childFolders"
            f"?$select=id,displayName&$top=50"
        )
        child_resp = requests.get(child_url, headers=headers, timeout=30)
        if not child_resp.ok:
            continue
        for child in child_resp.json().get("value", []):
            if child.get("displayName", "").strip().lower() == folder_name.strip().lower():
                log.info(
                    f"Resolved folder {folder_name!r} under {folder['displayName']!r} "
                    f"→ id={child['id']}"
                )
                return child["id"]

    log.warning(f"resolve_folder_id: folder {folder_name!r} not found — using root inbox")
    return None


# ─── Fetch emails ──────────────────────────────────────────────────────────────
def fetch_emails(token: str, top: int = 50, folder_id: Optional[str] = None, end_date: Optional[date] = None) -> list[dict]:
    """
    Fetch emails page by page WITHOUT expanding attachments inline.

    Expanding attachments ($expand=attachments) causes response payloads to
    balloon to tens of MB on large mailboxes, which triggers the Graph API
    gateway to cancel the request with "The operation was canceled."

    Attachments are fetched separately per message on demand inside
    upload_attachments() → _fetch_attachment_content().

    end_date: if provided, only emails received on or before this date (up to
    23:59:59 UTC of that day) are fetched via a Graph API $filter.  Emails are
    always ordered newest-first, so pagination stops as soon as an email older
    than the window is encountered.
    """
    headers = {"Authorization": f"Bearer {token}"}

    if folder_id:
        base = (
            f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}"
            f"/mailFolders/{folder_id}/messages"
        )
    else:
        base = f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/messages"

    # Page size capped at 50 — keeps individual responses small and avoids
    # gateway timeouts that occur with larger pages on busy mailboxes.
    # $expand=attachments is intentionally omitted — see docstring above.
    #
    # Date filter: Graph API uses ISO-8601 datetimes in UTC.
    # We filter receivedDateTime between end_date 00:00:00 UTC and today
    # 23:59:59 UTC so the full date range is included.
    _today_iso    = datetime.utcnow().strftime("%Y-%m-%dT23:59:59Z")
    _end_date_iso = (
        f"{end_date.strftime('%Y-%m-%d')}T00:00:00Z"
        if end_date else "1970-01-01T00:00:00Z"
    )
    _date_filter = (
        f"receivedDateTime ge {_end_date_iso}"
        f" and receivedDateTime le {_today_iso}"
    )
    url = (
        f"{base}"
        f"?$top=50"
        f"&$select=id,subject,from,toRecipients,ccRecipients,"
        f"body,receivedDateTime,isRead,hasAttachments"
        f"&$filter={requests.utils.quote(_date_filter)}"
        f"&$orderby=receivedDateTime desc"
    )

    all_emails: list[dict] = []
    page = 0
    while url:
        page += 1
        log.info(f"Fetching page {page}…")
        try:
            resp = requests.get(url, headers=headers, timeout=90)
            resp.raise_for_status()
        except requests.exceptions.Timeout:
            log.error(
                f"Timeout on page {page} — stopping pagination early. "
                f"Got {len(all_emails)} email(s) so far."
            )
            break
        except requests.exceptions.RequestException as exc:
            log.error(
                f"Request error on page {page}: {exc} — stopping pagination early. "
                f"Got {len(all_emails)} email(s) so far."
            )
            break

        data  = resp.json()
        batch = data.get("value", [])
        all_emails.extend(batch)
        log.info(f"  Page {page}: {len(batch)} email(s) (total so far: {len(all_emails)})")

        if len(all_emails) >= top:
            log.info(f"Reached LIMIT={top} — stopping pagination.")
            break

        url = data.get("@odata.nextLink")

    log.info(f"Fetched {len(all_emails)} email(s) total across {page} page(s).")
    return all_emails[:top]


def mark_email_read(token: str, message_id: str) -> None:
    try:
        requests.patch(
            f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/messages/{message_id}",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"isRead": True},
            timeout=30,
        )
    except Exception as exc:
        log.warning(f"mark_email_read failed for {message_id}: {exc} — continuing")


# ─── Subject parsing ───────────────────────────────────────────────────────────
SUBJECT_RE = re.compile(r"^(?P<code>[A-Z]{2,5})\s*:\s*(?P<rest>.+)$", re.IGNORECASE)
IGNORE_RE  = re.compile(r"^(fw|fwd|re|aw)\s*:", re.IGNORECASE)

_JR_PATTERNS = [
    re.compile(r"[-|]\s*(?:jr\.?\s*(?:no\.?\s*)?)(?P<jr>\d{4,})\s*$",   re.IGNORECASE),
    re.compile(r"^\s*(?:jr\.?\s*(?:no\.?\s*)?)(?P<jr>\d{4,})\s*[-|]",   re.IGNORECASE),
    re.compile(r"[-|]\s*(?P<jr>\d{4,})\s*$",                              re.IGNORECASE),
    re.compile(r"^\s*(?P<jr>\d{4,})\s*[-|]",                              re.IGNORECASE),
    re.compile(r"\(\s*jr\.?\s*(?:no\.?\s*)?(?P<jr>\d{4,})\s*\)",          re.IGNORECASE),
    re.compile(r"\(\s*(?P<jr>\d{4,})\s*\)",                                re.IGNORECASE),
    re.compile(r"\bjr\.?\s*(?:no\.?\s*)?(?P<jr>\d{4,})\b",                re.IGNORECASE),
    re.compile(r"\b(?P<jr>\d{4,})\b",                                      re.IGNORECASE),
]

_JR_STUB_RE = re.compile(r"\bjr\.?\s*(?:no\.?)?\s*$", re.IGNORECASE)

_SKILL_FILLER_RE = re.compile(
    r"\b(?:profiles?\s+for\s+(?:the\s+)?"
    r"|profiles?\s+of\s+(?:the\s+)?"
    r"|sharing\s+profiles?\s+for\s+(?:the\s+)?"
    r"|sharing\s+profiles?\s*"
    r"|cv\s+for\s+(?:the\s+)?"
    r"|resume\s+for\s+(?:the\s+)?"
    r"|candidate\s+for\s+(?:the\s+)?"
    r"| \|\| \|\| .*"
    r")",
    re.IGNORECASE,
)


def _extract_jr_and_skill(rest: str) -> tuple[Optional[str], str]:
    rest = rest.strip()
    jr_no: Optional[str] = None

    for pattern in _JR_PATTERNS:
        m = pattern.search(rest)
        if m:
            jr_no = m.group("jr")
            start, end = m.span()
            rest = (rest[:start] + " " + rest[end:]).strip()
            rest = re.sub(r"^[-|.\s]+|[-|.\s]+$", "", rest).strip()
            rest = _JR_STUB_RE.sub("", rest).strip()
            rest = re.sub(r"[-|.\s]+$", "", rest).strip()
            break

    rest = _SKILL_FILLER_RE.sub("", rest).strip()
    rest = re.sub(r"^[-|.\s]+|[-|.\s]+$", "", rest).strip()
    return jr_no, re.sub(r"\s{2,}", " ", rest).strip()


def parse_subject(subject: str) -> Optional[tuple[str, str, Optional[str]]]:
    """
    Parse a subject line of the form "CODE: <skill> [- JR12345]".

    Returns (company_name, general_skill, jr_no).

    company_name is the raw subject code exactly as written (e.g. "TCS", "BS",
    "NEWCO"). We do NOT look up a hardcoded code→name table because the list of
    client companies is open-ended; an unknown code would otherwise silently
    produce the wrong value or get swallowed.

    Returns None for FW/RE threads or unrecognised subject formats.
    """
    subject = subject.strip()
    if IGNORE_RE.match(subject):
        log.debug(f"Skip — FW/RE: {subject!r}")
        return None
    m = SUBJECT_RE.match(subject)
    if not m:
        log.debug(f"Skip — no pattern match: {subject!r}")
        return None

    # Use the raw code as-is — no lookup table
    company_name = m.group("code").upper()
    jr_no, skill = _extract_jr_and_skill(m.group("rest"))
    log.debug(f"Parsed subject → company={company_name!r} skill={skill!r} jr_no={jr_no!r}")
    return company_name, skill, jr_no


# ─── HTML table parser ─────────────────────────────────────────────────────────
def _strip_html(text: str) -> str:
    # Remove style/script blocks so their content isn't treated as cell text
    text = re.sub(r"<(style|script)[^>]*>.*?</\1>", "", text, flags=re.DOTALL | re.IGNORECASE)
    return re.sub(r"<[^>]+>", "", text)

def _clean_cell(text: str) -> str:
    text = _strip_html(text)
    text = unescape(text)                                                 # converts &nbsp; → \xa0, &amp; → &, &#160; → \xa0
    text = re.sub(r"[\xa0\u00a0\u200b\u200c\u200d\ufeff]+", " ", text)  # replaces ALL non-breaking space variants → regular space
    return re.sub(r"\s+", " ", text).strip()


def parse_html_table(html: str) -> list[dict]:
    rows: list[dict] = []
    for table_html in re.findall(r"<table[^>]*>(.*?)</table>", html, re.DOTALL | re.IGNORECASE):
        all_rows = re.findall(r"<tr[^>]*>(.*?)</tr>", table_html, re.DOTALL | re.IGNORECASE)
        if not all_rows:
            continue
        raw_headers = re.findall(
            r"<t[hd][^>]*>(.*?)</t[hd]>", all_rows[0], re.DOTALL | re.IGNORECASE
        )
        headers = [_clean_cell(h) for h in raw_headers]
        if not headers:
            continue
        col_map: dict[int, str] = {}
        for idx, h in enumerate(headers):
            c = _resolve_header(h)
            if c:
                col_map[idx] = c
            else:
                log.debug(f"Unmapped header: {h!r}")
        for row_html in all_rows[1:]:
            cells = re.findall(
                r"<t[hd][^>]*>(.*?)</t[hd]>", row_html, re.DOTALL | re.IGNORECASE
            )
            if not cells:
                continue
            cleaned = [_clean_cell(c) for c in cells]
            if all(v == "" for v in cleaned):
                continue
            record = {col_map[i]: v for i, v in enumerate(cleaned) if i in col_map}
            if record:
                rows.append(record)
    return rows


# ─── OneDrive folder & attachment upload ───────────────────────────────────────
def _is_resume(filename: str) -> bool:
    return os.path.splitext(filename)[1].lower() in RESUME_EXTENSIONS


def _ensure_onedrive_folder(od_token: str, folder_name: str) -> None:
    """Create the target folder in OneDrive if it doesn't already exist."""
    headers = {
        "Authorization": f"Bearer {od_token}",
        "Content-Type": "application/json",
    }
    check = requests.get(
        f"https://graph.microsoft.com/v1.0/me/drive/root:/{folder_name}",
        headers=headers,
        timeout=30,
    )
    if check.status_code == 200:
        log.info(f"OneDrive folder '{folder_name}' already exists.")
        return

    resp = requests.post(
        "https://graph.microsoft.com/v1.0/me/drive/root/children",
        headers=headers,
        json={
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename",
        },
        timeout=30,
    )
    if resp.status_code in (200, 201):
        log.info(f"Created OneDrive folder '{folder_name}'.")
    else:
        log.warning(
            f"Could not create OneDrive folder '{folder_name}': "
            f"{resp.status_code} — {resp.text[:200]}"
        )


def _fetch_attachment_content(mail_token: str, message_id: str, attachment_id: str) -> bytes:
    """Download raw attachment bytes using the mail token and TARGET_MAILBOX."""
    url = (
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}"
        f"/messages/{message_id}/attachments/{attachment_id}/$value"
    )
    resp = requests.get(url, headers={"Authorization": f"Bearer {mail_token}"}, timeout=120)
    resp.raise_for_status()
    return resp.content


def _upload_to_onedrive(od_token: str, filename: str, content: bytes) -> Optional[str]:
    """Upload a resume file to ONEDRIVE_USER's personal OneDrive."""
    remote_path = f"{ONEDRIVE_FOLDER}/{filename}" if ONEDRIVE_FOLDER else filename
    upload_url  = f"https://graph.microsoft.com/v1.0/me/drive/root:/{remote_path}:/content"

    try:
        resp = requests.put(
            upload_url,
            headers={
                "Authorization": f"Bearer {od_token}",
                "Content-Type": "application/octet-stream",
            },
            data=content,
            timeout=120,
        )
    except requests.exceptions.Timeout:
        log.error(f"OneDrive upload timed out for {filename!r}")
        return None

    if resp.status_code not in (200, 201):
        log.error(
            f"OneDrive upload failed for {filename!r}: "
            f"{resp.status_code} — {resp.text[:300]}"
        )
        return None

    item    = resp.json()
    item_id = item.get("id")

    try:
        link_resp = requests.post(
            f"https://graph.microsoft.com/v1.0/me/drive/items/{item_id}/createLink",
            headers={
                "Authorization": f"Bearer {od_token}",
                "Content-Type": "application/json",
            },
            json={"type": "view", "scope": "organization"},
            timeout=30,
        )
        if link_resp.status_code in (200, 201):
            web_url = link_resp.json().get("link", {}).get("webUrl", "")
            if web_url:
                log.info(f"  ↑ Uploaded {filename!r} → {web_url}")
                return web_url
    except Exception as exc:
        log.warning(f"createLink failed for {filename!r}: {exc}")

    fallback = item.get("webUrl", "")
    if fallback:
        log.warning(f"createLink failed for {filename!r}; using item webUrl: {fallback}")
        return fallback

    return None


def upload_attachments(od_token: str, mail_token: str, msg: dict) -> str:
    """
    Fetch resume attachments on demand and upload them to OneDrive.

    Attachments are NOT expanded inline during the email list fetch (which
    would balloon response sizes). Instead they are fetched here per message
    using the /attachments endpoint, only when hasAttachments is True.
    """
    if not msg.get("hasAttachments"):
        return ""

    url = (
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}"
        f"/messages/{msg['id']}/attachments?$select=id,name,contentType"
    )
    try:
        resp = requests.get(
            url,
            headers={"Authorization": f"Bearer {mail_token}"},
            timeout=60,
        )
        resp.raise_for_status()
        attachments = resp.json().get("value", [])
    except Exception as exc:
        log.warning(f"Could not fetch attachment list for {msg['id']}: {exc}")
        return ""

    links: list[str] = []
    for att in attachments:
        filename = att.get("name", "")
        if not filename or not _is_resume(filename):
            log.debug(f"  ↷ Skipping non-resume attachment: {filename!r}")
            continue
        att_id = att.get("id", "")
        if not att_id:
            links.append(filename)
            continue
        try:
            content = _fetch_attachment_content(mail_token, msg["id"], att_id)
            link    = _upload_to_onedrive(od_token, filename, content)
            links.append(link if link else filename)
        except Exception as exc:
            log.error(f"Failed to upload {filename!r}: {exc}")
            links.append(filename)

    return ", ".join(links)


# ─── DB schema helpers ──────────────────────────────────────────────────────────
def ensure_tables(cur) -> None:
    cur.execute("""
        CREATE TABLE IF NOT EXISTS hr_processed_emails (
            message_id    TEXT PRIMARY KEY,
            processed_at  TIMESTAMP DEFAULT NOW(),
            subject       TEXT,
            from_addr     TEXT,
            outcome       TEXT,
            rows_inserted INTEGER DEFAULT 0,
            rows_updated  INTEGER DEFAULT 0,
            rows_errors   INTEGER DEFAULT 0
        )
    """)


# ─── Processed-email tracking ──────────────────────────────────────────────────
def is_email_processed(cur, message_id: str) -> bool:
    cur.execute(
        "SELECT 1 FROM hr_processed_emails WHERE message_id = %s LIMIT 1",
        (message_id,),
    )
    return cur.fetchone() is not None


def mark_email_processed(
    cur,
    message_id: str,
    subject: str,
    from_addr: str,
    outcome: str,
    rows_inserted: int = 0,
    rows_updated: int = 0,
    rows_errors: int = 0,
) -> None:
    cur.execute(
        """
        INSERT INTO hr_processed_emails
            (message_id, subject, from_addr, outcome,
             rows_inserted, rows_updated, rows_errors)
        VALUES (%s, %s, %s, %s, %s, %s, %s)
        ON CONFLICT (message_id) DO NOTHING
        """,
        (message_id, subject, from_addr, outcome,
         rows_inserted, rows_updated, rows_errors),
    )


# ─── Smart duplicate detection & update ───────────────────────────────────────
def _select_cols_for_update() -> list[str]:
    return ["id"] + UPDATABLE_FIELDS


def find_same_key_record(
    cur,
    row_date,
    contact_number: str,
    email_id: str,
    jr_no: str,
    general_skill: str,
    client_recruiter: str,
    recruiter: str,
    delivery_type: str,
) -> Optional[dict]:
    if not contact_number and not email_id:
        return None

    cols  = _select_cols_for_update()
    query = pgsql.SQL("""
        SELECT {cols} FROM {table}
        WHERE date = %s
          AND date >= CURRENT_DATE - (INTERVAL '1 month' * %s)
          AND LOWER(COALESCE(contact_number,   '')) = LOWER(COALESCE(%s, ''))
          AND LOWER(COALESCE(email_id,         '')) = LOWER(COALESCE(%s, ''))
          AND (LOWER(COALESCE(jr_no,           '')) = LOWER(COALESCE(%s, ''))
               OR LOWER(COALESCE(general_skill,'')) = LOWER(COALESCE(%s, '')))
          AND LOWER(COALESCE(client_recruiter, '')) = LOWER(COALESCE(%s, ''))
          AND LOWER(COALESCE(recruiter,        '')) = LOWER(COALESCE(%s, ''))
          AND LOWER(COALESCE(delivery_type,    '')) = LOWER(COALESCE(%s, ''))
        LIMIT 1
    """).format(
        cols=pgsql.SQL(", ").join(pgsql.Identifier(c) for c in cols),
        table=pgsql.Identifier(DB_TABLE),
    )
    cur.execute(query, (
        row_date, DUPLICATE_CHECK_MONTHS,
        contact_number, email_id,
        jr_no, general_skill,
        client_recruiter, recruiter, delivery_type,
    ))
    row = cur.fetchone()
    return dict(zip(cols, row)) if row else None


def find_different_recruiter_record(
    cur,
    row_date,
    contact_number: str,
    email_id: str,
    jr_no: str,
    general_skill: str,
    recruiter: str,
    client_recruiter: str,
) -> Optional[dict]:
    if not contact_number and not email_id:
        return None

    cols  = ["id", "recruiter", "client_recruiter", "name_of_candidate", "email_from"] + UPDATABLE_FIELDS
    query = pgsql.SQL("""
        SELECT {cols} FROM {table}
        WHERE date >= CURRENT_DATE - (INTERVAL '1 month' * %s)
          AND LOWER(COALESCE(contact_number,   '')) = LOWER(COALESCE(%s, ''))
          AND LOWER(COALESCE(email_id,         '')) = LOWER(COALESCE(%s, ''))
          AND (LOWER(COALESCE(jr_no,           '')) = LOWER(COALESCE(%s, ''))
               OR LOWER(COALESCE(general_skill,'')) = LOWER(COALESCE(%s, '')))
          AND (LOWER(COALESCE(recruiter,        '')) != LOWER(COALESCE(%s, ''))
               OR LOWER(COALESCE(client_recruiter,'')) != LOWER(COALESCE(%s, '')))
        LIMIT 1
    """).format(
        cols=pgsql.SQL(", ").join(pgsql.Identifier(c) for c in cols),
        table=pgsql.Identifier(DB_TABLE),
    )
    cur.execute(query, (
        DUPLICATE_CHECK_MONTHS,
        contact_number, email_id,
        jr_no, general_skill,
        recruiter, client_recruiter,
    ))
    row = cur.fetchone()
    return dict(zip(cols, row)) if row else None


def find_contact_email_only_duplicate(
    cur,
    contact_number: str,
    email_id: str,
) -> Optional[dict]:
    if not contact_number and not email_id:
        return None

    cols  = ["id", "recruiter", "name_of_candidate", "jr_no", "general_skill", "date"]
    query = pgsql.SQL("""
        SELECT {cols} FROM {table}
        WHERE date >= CURRENT_DATE - (INTERVAL '1 month' * %s)
          AND (
            (LOWER(COALESCE(contact_number, '')) = LOWER(COALESCE(%s, '')) AND %s != '')
            OR
            (LOWER(COALESCE(email_id,       '')) = LOWER(COALESCE(%s, '')) AND %s != '')
          )
        LIMIT 1
    """).format(
        cols=pgsql.SQL(", ").join(pgsql.Identifier(c) for c in cols),
        table=pgsql.Identifier(DB_TABLE),
    )
    cur.execute(query, (
        DUPLICATE_CHECK_MONTHS,
        contact_number, contact_number,
        email_id,       email_id,
    ))
    row = cur.fetchone()
    return dict(zip(cols, row)) if row else None


def compute_updates(existing_row: dict, new_data: dict) -> dict:
    updates: dict = {}
    for field in UPDATABLE_FIELDS:
        existing_val   = existing_row.get(field)
        new_val        = new_data.get(field)
        existing_empty = existing_val is None or str(existing_val).strip() == ""
        new_has_value  = new_val is not None and str(new_val).strip() != ""
        if existing_empty and new_has_value:
            updates[field] = new_val
    return updates


def apply_field_updates(cur, record_id: int, updates: dict) -> None:
    if not updates:
        return
    set_parts = [
        pgsql.SQL("{col} = %s").format(col=pgsql.Identifier(k))
        for k in updates
    ]
    set_parts.append(pgsql.SQL("modified_date = NOW()"))
    query = pgsql.SQL("UPDATE {table} SET {sets} WHERE id = %s").format(
        table=pgsql.Identifier(DB_TABLE),
        sets=pgsql.SQL(", ").join(set_parts),
    )
    cur.execute(query, list(updates.values()) + [record_id])


# ─── Standard duplicate flag ───────────────────────────────────────────────────
def check_duplicate(cur, contact_number: str, email_id: str) -> str:
    if not contact_number and not email_id:
        return ""

    tbl = pgsql.Identifier(DB_TABLE)

    cell_dup = email_dup = False
    if contact_number:
        cur.execute(
            pgsql.SQL(
                "SELECT 1 FROM {t} WHERE contact_number = %s"
                " AND date >= CURRENT_DATE - (INTERVAL '1 month' * %s) LIMIT 1"
            ).format(t=tbl),
            (contact_number, DUPLICATE_CHECK_MONTHS),
        )
        cell_dup = bool(cur.fetchone())
    if email_id:
        cur.execute(
            pgsql.SQL(
                "SELECT 1 FROM {t} WHERE email_id = %s"
                " AND date >= CURRENT_DATE - (INTERVAL '1 month' * %s) LIMIT 1"
            ).format(t=tbl),
            (email_id, DUPLICATE_CHECK_MONTHS),
        )
        email_dup = bool(cur.fetchone())

    if cell_dup and email_dup:
        return "Duplicate"
    if cell_dup:
        return "Duplicate Cell"
    if email_dup:
        return "Duplicate Email"
    return ""


# ─── Email address helpers ─────────────────────────────────────────────────────
def _extract_address(addr_obj: dict) -> str:
    try:
        return addr_obj["emailAddress"]["address"].strip().lower()
    except (KeyError, TypeError):
        return ""


def _normalize_name_from_email(email: str) -> str:
    if not email or "@" not in email:
        return ""
    local = email.split("@", 1)[0]
    local = re.sub(r"\d+", "", local)
    local = re.sub(r"[._\-]+", " ", local)
    local = re.sub(r"\s+", " ", local).strip()
    return local.title()


def _recruiter_name(email: str) -> str:
    return _normalize_name_from_email(email)


def _username_from_email(email: str) -> str:
    return _normalize_name_from_email(email)


def _delivery_type(from_addr: str, to_addr: str) -> str:
    return (
        "Internal"
        if (VOLIBITS_DOMAIN in from_addr and VOLIBITS_DOMAIN in to_addr)
        else "External"
    )


# ─── Record validation ─────────────────────────────────────────────────────────
_REQUIRED_FIELDS = (
    "name_of_candidate", "contact_number", "email_id",
    "general_skill", "company_name",
    "recruiter", "email_from", "email_to",
)


def _record_status(data: dict) -> str:
    missing = [f for f in _REQUIRED_FIELDS if not data.get(f)]
    if missing:
        log.warning(f"record_status=Fail — missing: {missing}")
        return "Fail"
    return "Pass"


# ─── DB insertion ──────────────────────────────────────────────────────────────
_SAFE_FIELDS = (
    "recruiter", "client_recruiter", "general_skill", "company_name",
    "email_from", "email_to", "delivery_type", "email_id",
    "name_of_candidate", "contact_number", "is_duplicate", "status",
    "created_by", "created_date", "modified_by", "modified_date",
    "record_status", "attachment", "remarks",
)


def _t(val: Optional[str]) -> Optional[str]:
    if val is None:
        return None
    v = val.strip()
    return v or None


def _build_insert_sql() -> pgsql.Composable:
    return pgsql.SQL("""
        INSERT INTO {table} (
            recruiter, date, jr_no, client_recruiter, general_skill,
            name_of_candidate, contact_number, email_id,
            total_experience, relevant_experience,
            current_ctc, expected_ctc, notice_period,
            current_org, current_location, preferred_location,
            email_from, email_to, delivery_type, company_name,
            attachment, is_duplicate,
            created_by, created_date, modified_by, modified_date,
            record_status, status, remarks
        ) VALUES (
            %(recruiter)s, %(date)s, %(jr_no)s, %(client_recruiter)s, %(general_skill)s,
            %(name_of_candidate)s, %(contact_number)s, %(email_id)s,
            %(total_experience)s, %(relevant_experience)s,
            %(current_ctc)s, %(expected_ctc)s, %(notice_period)s,
            %(current_org)s, %(current_location)s, %(preferred_location)s,
            %(email_from)s, %(email_to)s, %(delivery_type)s, %(company_name)s,
            %(attachment)s, %(is_duplicate)s,
            %(created_by)s, %(created_date)s, %(modified_by)s, %(modified_date)s,
            %(record_status)s, %(status)s, %(remarks)s
        )
    """).format(table=pgsql.Identifier(DB_TABLE))


def insert_record(cur, data: dict) -> bool:
    insert_sql = _build_insert_sql()
    try:
        cur.execute(insert_sql, data)
        return True
    except Exception as exc:
        log.error(f"Insert error: {exc} | candidate={data.get('name_of_candidate')}")
        try:
            cur.connection.rollback()
            fallback = {k: None for k in data}
            for key in _SAFE_FIELDS:
                fallback[key] = data.get(key)
            fallback["record_status"] = "Fail"
            cur.execute(insert_sql, fallback)
            return True
        except Exception as exc2:
            log.error(f"Fallback insert also failed: {exc2}")
            return False


# ─── Date parsing ──────────────────────────────────────────────────────────────
_DATE_FORMATS = (
    "%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y",
    "%d-%b-%y", "%d-%b-%Y",
    "%b-%d-%Y", "%b-%d-%y",
    "%d/%b/%Y", "%d/%b/%y",
    "%d %b %Y", "%d %b %y",
    "%b %d %Y", "%b %d %y",
    "%d-%B-%y", "%d-%B-%Y",
    "%B-%d-%Y", "%B-%d-%y",
    "%d/%B/%Y", "%d/%B/%y",
    "%d %B %Y", "%d %B %y",
    "%B %d %Y", "%B %d %y",
)


def _parse_date(raw: str) -> Optional[date]:
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(raw.strip(), fmt).date()
        except ValueError:
            continue
    return None


# ─── Main pipeline ─────────────────────────────────────────────────────────────
def process_emails() -> None:
    log.info("=== HR Email Extractor starting ===")
    log.info(f"Target mailbox : {TARGET_MAILBOX}")
    log.info(f"OneDrive user  : {ONEDRIVE_USER}")
    log.info(f"OneDrive folder: {ONEDRIVE_FOLDER}")
    log.info(f"Inbox subfolder: {INBOX_SUBFOLDER or '(root inbox)'}")
    log.info(f"Date filter    : {END_DATE} → today ({date.today()})")

    token    = get_mail_token()
    od_token = get_onedrive_token()
    log.info("Tokens obtained.")

    _drive_check = requests.get(
        "https://graph.microsoft.com/v1.0/me/drive",
        headers={"Authorization": f"Bearer {od_token}"},
        timeout=30,
    )
    if _drive_check.status_code == 200:
        log.info(f"OneDrive access confirmed for {ONEDRIVE_USER}.")
        if ONEDRIVE_FOLDER:
            _ensure_onedrive_folder(od_token, ONEDRIVE_FOLDER)
    else:
        log.warning(
            f"OneDrive access check returned {_drive_check.status_code} for "
            f"{ONEDRIVE_USER} — attachment uploads may fail. "
            f"Response: {_drive_check.text[:200]}"
        )

    inbox_folder_id = resolve_folder_id(token, INBOX_SUBFOLDER) if INBOX_SUBFOLDER else None

    conn = psycopg2.connect(DB_DSN)
    conn.autocommit = False
    cur  = conn.cursor()

    ensure_tables(cur)
    conn.commit()

    emails = fetch_emails(token, top=Limit, folder_id=inbox_folder_id, end_date=END_DATE)

    processed = skipped = inserted = updated = conflicts = errors = 0

    for msg in emails:
        subject    = msg.get("subject", "").strip()
        message_id = msg["id"]

        if is_email_processed(cur, message_id):
            log.info(f"Already processed — skipping {message_id} | {subject!r}")
            mark_email_read(token, message_id)
            continue

        log.info(f"Subject: {subject!r}")

        parsed = parse_subject(subject)
        if parsed is None:
            skipped += 1
            mark_email_read(token, msg["id"])
            continue

        # parse_subject returns (company_name, general_skill, jr_no)
        # company_name is the raw subject code — no lookup table applied
        company_name, general_skill, subject_jr_no = parsed

        from_addr = _extract_address(msg.get("from", {}))
        to_list   = msg.get("toRecipients", [])
        to_addr   = _extract_address(to_list[0]) if to_list else ""

        recruiter        = _recruiter_name(from_addr)
        client_recruiter = _username_from_email(to_addr)
        delivery_type    = _delivery_type(from_addr, to_addr)

        attachment_str = upload_attachments(od_token, token, msg)

        body_html = (msg.get("body") or {}).get("content", "")
        rows      = parse_html_table(body_html)
        if not rows:
            log.warning("No table rows found — inserting skeleton record.")
            rows = [{}]

        email_date: Optional[date] = None
        raw_dt = msg.get("receivedDateTime", "")
        if raw_dt:
            try:
                email_date = datetime.fromisoformat(raw_dt.replace("Z", "+00:00")).date()
            except ValueError:
                pass

        now = datetime.now()

        rows_summary: list[dict]   = []
        email_inserted = email_skipped_rows = email_errors = 0
        email_updated  = email_conflicts = 0

        for row in rows:
            row_date = _parse_date(row["date"]) if row.get("date") else None
            if row_date is None:
                row_date = date.today()
                if row.get("date"):
                    log.warning(
                        f"Unparseable date value {row['date']!r} — "
                        f"defaulting to today ({row_date})"
                    )
                else:
                    log.debug(f"No date in row — defaulting to today ({row_date})")

            contact_number  = _t(row.get("contact_number"))
            email_id_val    = _t(row.get("email_id"))
            effective_jr_no = _t(row.get("jr_no")) or subject_jr_no
            candidate_name  = _t(row.get("name_of_candidate"))

            # ── TIER 1: Same recruiter — auto-update empty fields ─────────────
            same_key = find_same_key_record(
                cur, row_date,
                contact_number   or "",
                email_id_val     or "",
                effective_jr_no  or "",
                general_skill    or "",
                client_recruiter or "",
                recruiter        or "",
                delivery_type,
            )

            if same_key is not None:
                record_updates = compute_updates(same_key, {
                    "jr_no":               effective_jr_no,
                    "total_experience":    _t(row.get("total_experience")),
                    "relevant_experience": _t(row.get("relevant_experience")),
                    "current_ctc":         _t(row.get("current_ctc")),
                    "expected_ctc":        _t(row.get("expected_ctc")),
                    "notice_period":       _t(row.get("notice_period")),
                    "current_org":         _t(row.get("current_org")),
                    "current_location":    _t(row.get("current_location")),
                    "preferred_location":  _t(row.get("preferred_location")),
                    "attachment":          _t(attachment_str),
                    "remarks":             _t(row.get("remarks")),
                })
                if record_updates:
                    apply_field_updates(cur, same_key["id"], record_updates)
                    outcome = "updated"
                    email_updated += 1
                    updated += 1
                    log.info(
                        f"  ↑ Updated record {same_key['id']} for "
                        f"{candidate_name} — filled: {list(record_updates.keys())}"
                    )
                else:
                    outcome = "no_change"
                    email_skipped_rows += 1
                    skipped += 1
                    log.info(
                        f"  ⟳ No new data to fill for {candidate_name} "
                        f"(record {same_key['id']})"
                    )

                rows_summary.append({
                    "name":           candidate_name or "(unknown)",
                    "outcome":        outcome,
                    "record_data":    {},
                    "updated_fields": list(record_updates.keys()) if record_updates else [],
                    "missing":        [],
                    "dup_flag":       "",
                    "db_error":       None,
                })
                continue

            # ── TIER 2: Different recruiter — insert immediately, notify ───────
            diff_key = find_different_recruiter_record(
                cur, row_date,
                contact_number  or "",
                email_id_val    or "",
                effective_jr_no or "",
                general_skill   or "",
                recruiter       or "",
                client_recruiter or "",
            )

            if diff_key is not None:
                new_record_data = {
                    "recruiter":           _t(recruiter),
                    "date":                row_date,
                    "jr_no":               effective_jr_no,
                    "client_recruiter":    _t(client_recruiter),
                    "general_skill":       _t(general_skill),
                    "name_of_candidate":   candidate_name,
                    "contact_number":      contact_number,
                    "email_id":            email_id_val,
                    "total_experience":    _t(row.get("total_experience")),
                    "relevant_experience": _t(row.get("relevant_experience")),
                    "current_ctc":         _t(row.get("current_ctc")),
                    "expected_ctc":        _t(row.get("expected_ctc")),
                    "notice_period":       _t(row.get("notice_period")),
                    "current_org":         _t(row.get("current_org")),
                    "current_location":    _t(row.get("current_location")),
                    "preferred_location":  _t(row.get("preferred_location")),
                    "email_from":          _t(from_addr),
                    "email_to":            _t(to_addr),
                    "delivery_type":       delivery_type,
                    "company_name":        company_name,
                    "attachment":          _t(attachment_str),
                    "is_duplicate":        "Duplicate Recruiter",
                    "created_by":          _t(from_addr),
                    "created_date":        now,
                    "modified_by":         _t(from_addr),
                    "modified_date":       now,
                    "remarks":             _t(row.get("remarks")),
                    "record_status":       _record_status({
                        "name_of_candidate": candidate_name,
                        "contact_number":    contact_number,
                        "email_id":          email_id_val,
                        "general_skill":     _t(general_skill),
                        "company_name":      company_name,
                        "recruiter":         _t(recruiter),
                        "email_from":        _t(from_addr),
                        "email_to":          _t(to_addr),
                    }),
                    "status": "Screen Pending",
                }

                ok = insert_record(cur, new_record_data)
                if ok:
                    inserted += 1
                    email_inserted += 1
                    log.info(
                        f"  ✓ [DIFF-RECRUITER] Inserted {candidate_name or '(unknown)'} "
                        f"by recruiter={recruiter}; also exists under {diff_key.get('recruiter')}"
                    )
                else:
                    errors += 1
                    email_errors += 1
                    log.error(f"  ✗ [DIFF-RECRUITER] Insert failed for {candidate_name}")

                send_diff_recruiter_notification_email(
                    token=token,
                    new_recruiter_addr=from_addr,
                    existing_recruiter_addr=(
                        diff_key.get("email_from")
                        or (diff_key.get("recruiter", "") + "@" + VOLIBITS_DOMAIN)
                    ),
                    existing_row=diff_key,
                    new_record_data=new_record_data,
                    original_subject=subject,
                    jr_no=effective_jr_no or general_skill or "",
                    inserted_ok=ok,
                )

                rows_summary.append({
                    "name":           candidate_name or "(unknown)",
                    "outcome":        "inserted" if ok else "error",
                    "record_data":    new_record_data,
                    "updated_fields": [],
                    "missing":        [f for f in _REQUIRED_FIELDS if not new_record_data.get(f)],
                    "dup_flag":       "Duplicate Recruiter",
                    "db_error":       None if ok else "DB insert failed — check extractor.log",
                    "diff_recruiter": diff_key.get("recruiter"),
                })
                continue

            # ── TIER 2.5: Same contact/email, different JR — flag only ─────────
            contact_dup = find_contact_email_only_duplicate(
                cur,
                contact_number or "",
                email_id_val   or "",
            )

            # ── TIER 3: No matching record — standard insert path ──────────────
            dup_flag = check_duplicate(cur, contact_number or "", email_id_val or "")

            record_data = {
                "recruiter":           _t(recruiter),
                "date":                row_date,
                "jr_no":               effective_jr_no,
                "client_recruiter":    _t(client_recruiter),
                "general_skill":       _t(general_skill),
                "name_of_candidate":   candidate_name,
                "contact_number":      contact_number,
                "email_id":            email_id_val,
                "total_experience":    _t(row.get("total_experience")),
                "relevant_experience": _t(row.get("relevant_experience")),
                "current_ctc":         _t(row.get("current_ctc")),
                "expected_ctc":        _t(row.get("expected_ctc")),
                "notice_period":       _t(row.get("notice_period")),
                "current_org":         _t(row.get("current_org")),
                "current_location":    _t(row.get("current_location")),
                "preferred_location":  _t(row.get("preferred_location")),
                "email_from":          _t(from_addr),
                "email_to":            _t(to_addr),
                "delivery_type":       delivery_type,
                "company_name":        company_name,
                "attachment":          _t(attachment_str),
                "is_duplicate":        _t(dup_flag),
                "created_by":          _t(from_addr),
                "created_date":        now,
                "modified_by":         _t(from_addr),
                "modified_date":       now,
                "remarks":             _t(row.get("remarks")),
                "record_status":       _record_status({
                    "name_of_candidate": candidate_name,
                    "contact_number":    contact_number,
                    "email_id":          email_id_val,
                    "general_skill":     _t(general_skill),
                    "company_name":      company_name,
                    "recruiter":         _t(recruiter),
                    "email_from":        _t(from_addr),
                    "email_to":          _t(to_addr),
                }),
                "status": "Screen Pending",
            }

            ok = insert_record(cur, record_data)
            if ok:
                inserted += 1
                email_inserted += 1
                log.info(
                    f"  ✓ {candidate_name or '(unknown)'} | "
                    f"recruiter={recruiter} | dup={dup_flag or 'none'}"
                )
                if contact_dup is not None:
                    send_diff_recruiter_notification_email(
                        token=token,
                        new_recruiter_addr=from_addr,
                        existing_recruiter_addr=(
                            contact_dup.get("recruiter", "") + "@" + VOLIBITS_DOMAIN
                        ),
                        existing_row=contact_dup,
                        new_record_data=record_data,
                        original_subject=subject,
                        jr_no=effective_jr_no or general_skill or "",
                        inserted_ok=ok,
                    )
                    log.info(
                        f"  ✉ [DUP-CONTACT] Notified manager for {candidate_name} — "
                        f"existing recruiter: {contact_dup.get('recruiter')}"
                    )
                rows_summary.append({
                    "name":           candidate_name or "(unknown)",
                    "outcome":        "inserted",
                    "record_data":    record_data,
                    "updated_fields": [],
                    "missing":        [f for f in _REQUIRED_FIELDS if not record_data.get(f)],
                    "dup_flag":       dup_flag or "",
                    "db_error":       None,
                })
            else:
                errors += 1
                email_errors += 1
                rows_summary.append({
                    "name":           candidate_name or "(unknown)",
                    "outcome":        "error",
                    "record_data":    record_data,
                    "updated_fields": [],
                    "missing":        [f for f in _REQUIRED_FIELDS if not record_data.get(f)],
                    "dup_flag":       dup_flag or "",
                    "db_error":       "DB insert failed — check extractor.log",
                })

        if email_errors > 0 and email_inserted == 0 and email_updated == 0:
            _outcome = "all_failed"
        elif email_errors > 0:
            _outcome = "partial"
        else:
            _outcome = "ok"

        mark_email_processed(
            cur,
            message_id=message_id,
            subject=subject,
            from_addr=from_addr,
            outcome=_outcome,
            rows_inserted=email_inserted,
            rows_updated=email_updated,
            rows_errors=email_errors,
        )

        try:
            conn.commit()
        except Exception as e:
            log.error(f"Commit failed: {e}")
            conn.rollback()

        send_notification_email(
            token=token,
            from_addr=from_addr,
            original_subject=subject,
            rows_summary=rows_summary,
            email_inserted=email_inserted,
            email_skipped=email_skipped_rows,
            email_errors=email_errors,
            email_updated=email_updated,
            email_conflicts=email_conflicts,
        )

        mark_email_read(token, message_id)
        processed += 1

    cur.close()
    conn.close()

    log.info(
        f"=== Done — processed={processed} skipped={skipped} "
        f"inserted={inserted} updated={updated} "
        f"conflicts={conflicts} errors={errors} ==="
    )


if __name__ == "__main__":
    process_emails()