"""
Outlook Email Extractor for HR Recruitment (Volibits)
Reads emails via Microsoft Graph API and inserts candidate data into PostgreSQL.

recruiter        = local part of sender email   — sender is always @volibits.com
client_recruiter = local part of receiver email — can be any domain

Duplicate handling (three-tier):
  1. Same recruiter + same key fields → auto-fill only NULL/empty fields in the
     existing record; never overwrite populated values.
  2. Different recruiter + same candidate/JR → store conflict in
     hr_pending_conflicts, send urgent notification with mailto action links.
  3. [HR-ACTION] reply emails → processed at the start of each run to resolve
     pending conflicts (UPDATE / NEW / SKIP).

OneDrive fix notes:
  - Only .pdf / .doc / .docx attachments are uploaded
  - Requires Files.ReadWrite (Application) permission in Azure
"""

import os
import re
import json
import uuid
import logging
import requests
import psycopg2
from psycopg2 import sql as pgsql
from datetime import datetime, date
from typing import Optional
from msal import ConfidentialClientApplication
from notifier import send_notification_email, send_conflict_notification_email

# ─── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("extractor.log"),
    ],
)
log = logging.getLogger(__name__)


# ─── Configuration ───────────────────────────────────────────────────────────────
AZURE_TENANT_ID     = os.environ["AZURE_TENANT_ID"]
AZURE_CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
AZURE_CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]

OD_TENANT_ID        = os.environ["OD_TENANT_ID"]
OD_CLIENT_ID        = os.environ["OD_CLIENT_ID"]
OD_CLIENT_SECRET    = os.environ["OD_CLIENT_SECRET"]

TARGET_MAILBOX      = os.environ["TARGET_MAILBOX"]
DB_DSN              = os.environ["DB_DSN"]
DB_TABLE            = os.environ.get("DB_TABLE_NAME", "hrvolibit")
ONEDRIVE_FOLDER     = os.environ.get("ONEDRIVE_FOLDER", "HR Resumes")

VOLIBITS_DOMAIN     = "volibits.com"
RESUME_EXTENSIONS   = {".pdf", ".doc", ".docx"}

# Fields that may be auto-filled when NULL/empty in an existing record.
# Identity and routing fields are intentionally excluded — they are never changed.
UPDATABLE_FIELDS = [
    "jr_no",
    "total_experience", "relevant_experience",
    "current_ctc", "expected_ctc",
    "notice_period", "current_org",
    "current_location", "preferred_location",
    "attachment", "remarks",
]

# Subject prefix that marks a recruiter's action-reply email
ACTION_REPLY_RE = re.compile(
    r"^\[HR-ACTION\]\s+(UPDATE|NEW|SKIP)\s+([\w-]+)\s*$", re.IGNORECASE
)


# ─── Company code → full name ────────────────────────────────────────────────────
COMPANY_CODES: dict[str, str] = {
    "BS":  "Birlasoft",
    "BW":  "BeWealthy",
    "TCS": "TCS",
    "INF": "Infosys",
    "WIP": "Wipro",
    "HCL": "HCL Technologies",
    "ACC": "Accenture",
    "CAP": "Capgemini",
    "COG": "Cognizant",
    "IBM": "IBM",
    "SAP": "SAP",
    "ORC": "Oracle",
    "MS":  "Microsoft",
    "RS":  "RS Software",
}


# ─── Fuzzy column-header aliases ────────────────────────────────────────────────
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
        "organisation", "organization", "curr company", "present company",
    ],
    "total_experience": [
        "total experience", "total_experience", "total exp", "total_exp",
        "experience", "exp", "total yrs", "total years",
    ],
    "relevant_experience": [
        "relevant experience", "relevant_experience", "relevant exp",
        "relevant_exp", "rel exp", "rel_exp", "relevant yrs",
    ],
    "current_ctc": [
        "current ctc", "current_ctc", "ctc", "current salary",
        "current_salary", "present ctc", "present_ctc", "curr ctc",
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
        "current city", "current_city", "present location",
    ],
    "preferred_location": [
        "preferred location", "preferred_location", "preferred city",
        "preferred_city", "pref location", "pref_location",
        "willing to relocate", "target location",
    ],
    "date": [
        "date", "submission date", "submission_date", "applied date",
        "applied_date", "profile date",
    ],
    "remarks": [
        "remarks", "comments", "comment", "note", "notes",
    ],
}


# ─── Build alias lookup once at startup ─────────────────────────────────────────
def _normalize(s: str) -> str:
    return re.sub(r"\s+", " ", s.strip().lower())


ALIAS_MAP: dict[str, str] = {
    _normalize(alias): canonical
    for canonical, aliases in COLUMN_ALIASES.items()
    for alias in aliases
}


def _resolve_header(raw_header: str) -> Optional[str]:
    return ALIAS_MAP.get(_normalize(raw_header))


# ─── Microsoft Graph authentication ─────────────────────────────────────────────
def _get_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    result = ConfidentialClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
    ).acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(
            f"Token acquisition failed: {result.get('error_description')}"
        )
    return result["access_token"]


def get_mail_token() -> str:
    return _get_token(AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET)


def get_onedrive_token() -> str:
    return _get_token(OD_TENANT_ID, OD_CLIENT_ID, OD_CLIENT_SECRET)


# ─── Fetch & mark emails ─────────────────────────────────────────────────────────
def fetch_emails(token: str, top: int = 50) -> list[dict]:
    headers = {"Authorization": f"Bearer {token}"}
    url = (
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/messages"
        f"?$top={top}"
        f"&$select=id,subject,from,toRecipients,ccRecipients,"
        f"body,receivedDateTime,isRead,hasAttachments"
        f"&$expand=attachments($select=id,name,contentType,size)"
        f"&$filter=isRead eq false"
        f"&$orderby=receivedDateTime asc"
    )
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.json().get("value", [])


def mark_email_read(token: str, message_id: str) -> None:
    requests.patch(
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/messages/{message_id}",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"isRead": True},
        timeout=10,
    )


# ─── Subject parsing ─────────────────────────────────────────────────────────────
SUBJECT_RE = re.compile(r"^(?P<code>[A-Z]{2,5})\s*:\s*(?P<rest>.+)$", re.IGNORECASE)
IGNORE_RE  = re.compile(r"^(fw|fwd|re|aw)\s*:", re.IGNORECASE)

_JR_PATTERNS = [
    re.compile(r"[-|]\s*(?:jr\s*)?(?P<jr>\d{4,})\s*$",   re.IGNORECASE),
    re.compile(r"^\s*(?:jr\s*)?(?P<jr>\d{4,})\s*[-|]",   re.IGNORECASE),
    re.compile(r"\(\s*jr\s*(?P<jr>\d{4,})\s*\)",          re.IGNORECASE),
    re.compile(r"\(\s*(?P<jr>\d{4,})\s*\)",               re.IGNORECASE),
    re.compile(r"\bjr\s*(?P<jr>\d{4,})\b",                re.IGNORECASE),
    re.compile(r"\b(?P<jr>\d{4,})\b",                     re.IGNORECASE),
]


def _extract_jr_and_skill(rest: str) -> tuple[Optional[str], str]:
    rest = rest.strip()
    jr_no: Optional[str] = None
    for pattern in _JR_PATTERNS:
        m = pattern.search(rest)
        if m:
            jr_no = m.group("jr")
            start, end = m.span()
            rest = (rest[:start] + " " + rest[end:]).strip()
            rest = re.sub(r"^[-|\s]+|[-|\s]+$", "", rest).strip()
            break
    return jr_no, re.sub(r"\s{2,}", " ", rest).strip()


def parse_subject(subject: str) -> Optional[tuple[str, str, Optional[str]]]:
    subject = subject.strip()
    if IGNORE_RE.match(subject):
        log.debug(f"Skip — FW/RE: {subject!r}")
        return None
    m = SUBJECT_RE.match(subject)
    if not m:
        log.debug(f"Skip — no pattern match: {subject!r}")
        return None
    code  = m.group("code").upper()
    jr_no, skill = _extract_jr_and_skill(m.group("rest"))
    log.debug(f"Parsed subject → code={code} skill={skill!r} jr_no={jr_no!r}")
    return code, skill, jr_no


# ─── HTML table parser ───────────────────────────────────────────────────────────
def _strip_html(text: str) -> str:
    return re.sub(r"<[^>]+>", "", text)


def _clean_cell(text: str) -> str:
    return re.sub(r"\s+", " ", _strip_html(text)).strip()


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


# ─── OneDrive attachment upload ──────────────────────────────────────────────────
def _is_resume(filename: str) -> bool:
    return os.path.splitext(filename)[1].lower() in RESUME_EXTENSIONS


def _fetch_attachment_content(token: str, message_id: str, attachment_id: str) -> bytes:
    url = (
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}"
        f"/messages/{message_id}/attachments/{attachment_id}/$value"
    )
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    resp.raise_for_status()
    return resp.content


def _upload_to_onedrive(token: str, filename: str, content: bytes) -> Optional[str]:
    remote_path = f"{ONEDRIVE_FOLDER}/{filename}"
    resp = requests.put(
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}"
        f"/drive/root:/{remote_path}:/content",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/octet-stream"},
        data=content,
        timeout=120,
    )
    if resp.status_code not in (200, 201):
        log.error(f"OneDrive upload failed for {filename!r}: {resp.status_code} — {resp.text[:300]}")
        return None

    item    = resp.json()
    item_id = item.get("id")
    link_resp = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}"
        f"/drive/items/{item_id}/createLink",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"type": "view", "scope": "organization"},
        timeout=30,
    )
    if link_resp.status_code in (200, 201):
        web_url = link_resp.json().get("link", {}).get("webUrl", "")
        if web_url:
            log.info(f"  ↑ Uploaded {filename!r} → {web_url}")
            return web_url

    fallback = item.get("webUrl", "")
    if fallback:
        log.warning(f"createLink failed for {filename!r}; using webUrl: {fallback}")
        return fallback
    return None


def upload_attachments(od_token: str, mail_token: str, msg: dict) -> str:
    attachments = msg.get("attachments") or []
    if not attachments:
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


# ─── DB schema helpers ────────────────────────────────────────────────────────────
def ensure_tables(cur) -> None:
    """Create ancillary tables if they don't exist yet."""
    cur.execute("""
        CREATE TABLE IF NOT EXISTS hr_pending_conflicts (
            conflict_id         TEXT PRIMARY KEY,
            created_at          TIMESTAMP DEFAULT NOW(),
            status              TEXT DEFAULT 'pending',
            action_taken        TEXT,
            resolved_at         TIMESTAMP,
            existing_record_id  INTEGER,
            existing_recruiter  TEXT,
            existing_email_from TEXT,
            new_recruiter       TEXT,
            candidate_name      TEXT,
            contact_number      TEXT,
            email_id_val        TEXT,
            new_record_data     TEXT
        )
    """)


# ─── Smart duplicate detection & update ─────────────────────────────────────────
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
    """
    Find an existing record that matches all key identity fields AND was
    submitted by the same recruiter/client pair with the same delivery type.

    Match criteria:
      date, contact_number, email_id,
      jr_no OR general_skill,
      client_recruiter, recruiter, delivery_type

    Returns a dict of {col: value} including 'id', or None.
    """
    if not contact_number and not email_id:
        return None

    cols  = _select_cols_for_update()
    query = pgsql.SQL("""
        SELECT {cols} FROM {table}
        WHERE date = %s
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
        row_date, contact_number, email_id,
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
    """
    Find an existing record with the same candidate/JR identity but submitted
    by a DIFFERENT recruiter or client pair.
    Returns a dict including 'id', 'recruiter', 'client_recruiter',
    'name_of_candidate', 'email_from', plus all UPDATABLE_FIELDS, or None.
    """
    if not contact_number and not email_id:
        return None

    cols  = ["id", "recruiter", "client_recruiter", "name_of_candidate", "email_from"] + UPDATABLE_FIELDS
    query = pgsql.SQL("""
        SELECT {cols} FROM {table}
        WHERE date = %s
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
        row_date, contact_number, email_id,
        jr_no, general_skill,
        recruiter, client_recruiter,
    ))
    row = cur.fetchone()
    return dict(zip(cols, row)) if row else None


def compute_updates(existing_row: dict, new_data: dict) -> dict:
    """
    Return {field: new_value} for every UPDATABLE_FIELD where:
      - the existing record value is NULL or blank, AND
      - the new record has a non-blank value.
    Never overwrites populated fields.
    """
    updates: dict = {}
    for field in UPDATABLE_FIELDS:
        existing_val = existing_row.get(field)
        new_val      = new_data.get(field)
        existing_empty = existing_val is None or str(existing_val).strip() == ""
        new_has_value  = new_val is not None and str(new_val).strip() != ""
        if existing_empty and new_has_value:
            updates[field] = new_val
    return updates


def apply_field_updates(cur, record_id: int, updates: dict) -> None:
    """Apply a dict of {field: value} to the existing record (UPDATE only those cols)."""
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


def store_pending_conflict(
    cur,
    conflict_id: str,
    existing_row: dict,
    new_record_data: dict,
) -> None:
    cur.execute(
        """
        INSERT INTO hr_pending_conflicts
            (conflict_id, existing_record_id, existing_recruiter, existing_email_from,
             new_recruiter, candidate_name, contact_number, email_id_val, new_record_data)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
        ON CONFLICT (conflict_id) DO NOTHING
        """,
        (
            conflict_id,
            existing_row.get("id"),
            existing_row.get("recruiter"),
            existing_row.get("email_from"),
            new_record_data.get("recruiter"),
            new_record_data.get("name_of_candidate"),
            new_record_data.get("contact_number"),
            new_record_data.get("email_id"),
            json.dumps(new_record_data, default=str),
        ),
    )


# ─── Action-reply processor ───────────────────────────────────────────────────────
def _restore_record_dates(data: dict) -> dict:
    """Parse date strings back to Python objects after JSON round-trip."""
    for key in ("date",):
        val = data.get(key)
        if isinstance(val, str):
            parsed = _parse_date(val)
            data[key] = parsed
    for key in ("created_date", "modified_date"):
        val = data.get(key)
        if isinstance(val, str):
            try:
                data[key] = datetime.fromisoformat(val)
            except ValueError:
                data[key] = datetime.now()
    return data


def process_action_replies(token: str, cur, conn) -> None:
    """
    Scan inbox for [HR-ACTION] reply emails (oldest first) and resolve the
    matching pending conflict record before normal email processing begins.
    """
    headers = {"Authorization": f"Bearer {token}"}
    url = (
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/messages"
        f"?$top=50"
        f"&$select=id,subject,from,receivedDateTime"
        f"&$filter=isRead eq false"
        f"&$orderby=receivedDateTime asc"
    )
    resp = requests.get(url, headers=headers, timeout=30)
    if not resp.ok:
        log.warning(f"process_action_replies: fetch failed {resp.status_code}")
        return

    for msg in resp.json().get("value", []):
        subject = (msg.get("subject") or "").strip()
        m       = ACTION_REPLY_RE.match(subject)
        if not m:
            continue   # Not an action reply — leave for normal processing

        action      = m.group(1).upper()
        conflict_id = m.group(2)
        sender      = _extract_address(msg.get("from", {}))
        log.info(f"Action reply: {action} | id={conflict_id} | from={sender}")

        cur.execute(
            "SELECT existing_record_id, new_record_data "
            "FROM hr_pending_conflicts "
            "WHERE conflict_id = %s AND status = 'pending'",
            (conflict_id,),
        )
        conflict = cur.fetchone()
        if not conflict:
            log.warning(f"No pending conflict for id={conflict_id} — may be already resolved")
            mark_email_read(token, msg["id"])
            continue

        existing_id, new_record_json = conflict
        new_record_data = _restore_record_dates(json.loads(new_record_json))

        try:
            if action == "UPDATE":
                cols = _select_cols_for_update()
                q = pgsql.SQL("SELECT {cols} FROM {tbl} WHERE id = %s").format(
                    cols=pgsql.SQL(", ").join(pgsql.Identifier(c) for c in cols),
                    tbl=pgsql.Identifier(DB_TABLE),
                )
                cur.execute(q, (existing_id,))
                row = cur.fetchone()
                if row:
                    existing_row = dict(zip(cols, row))
                    updates      = compute_updates(existing_row, new_record_data)
                    if updates:
                        apply_field_updates(cur, existing_id, updates)
                        log.info(
                            f"  ✓ [ACTION:UPDATE] record {existing_id} "
                            f"— filled: {list(updates.keys())}"
                        )
                    else:
                        log.info(f"  ↷ [ACTION:UPDATE] no empty fields to fill in record {existing_id}")
                else:
                    log.warning(f"  [ACTION:UPDATE] record {existing_id} not found in DB")

            elif action == "NEW":
                new_record_data["created_date"]  = datetime.now()
                new_record_data["modified_date"] = datetime.now()
                insert_record(cur, new_record_data)
                log.info(f"  ✓ [ACTION:NEW] inserted new record for conflict {conflict_id}")

            else:  # SKIP
                log.info(f"  ↷ [ACTION:SKIP] conflict {conflict_id} skipped by {sender}")

            cur.execute(
                "UPDATE hr_pending_conflicts "
                "SET status='resolved', action_taken=%s, resolved_at=NOW() "
                "WHERE conflict_id=%s",
                (action, conflict_id),
            )
            conn.commit()

        except Exception as exc:
            log.error(f"Action reply processing failed for {conflict_id}: {exc}")
            conn.rollback()

        mark_email_read(token, msg["id"])


# ─── Standard duplicate flag (for is_duplicate column) ──────────────────────────
def check_duplicate(cur, contact_number: str, email_id: str) -> str:
    """Returns 'Duplicate', 'Duplicate Cell', 'Duplicate Email', or ''."""
    if not contact_number and not email_id:
        return ""

    tbl = pgsql.Identifier(DB_TABLE)

    if contact_number and email_id:
        cur.execute(
            pgsql.SQL("SELECT 1 FROM {t} WHERE jr_no=%s OR general_skill=%s AND contact_number=%s AND email_id=%s LIMIT 1").format(t=tbl),
            (contact_number, email_id),
        )
        if cur.fetchone():
            return "Duplicate"

    cell_dup = email_dup = False
    if contact_number:
        cur.execute(
            pgsql.SQL("SELECT 1 FROM {t} WHERE contact_number=%s LIMIT 1").format(t=tbl),
            (contact_number,),
        )
        cell_dup = bool(cur.fetchone())
    if email_id:
        cur.execute(
            pgsql.SQL("SELECT 1 FROM {t} WHERE email_id=%s LIMIT 1").format(t=tbl),
            (email_id,),
        )
        email_dup = bool(cur.fetchone())

    if cell_dup and email_dup:
        return "Duplicate"
    if cell_dup:
        return "Duplicate Cell"
    if email_dup:
        return "Duplicate Email"
    return ""


# ─── Email address helpers ───────────────────────────────────────────────────────
def _extract_address(addr_obj: dict) -> str:
    try:
        return addr_obj["emailAddress"]["address"].strip().lower()
    except (KeyError, TypeError):
        return ""


def _recruiter_name(email: str) -> str:
    return email.split("@", 1)[0].strip() if "@" in email else email.strip()


def _username_from_email(email: str) -> str:
    return email.split("@", 1)[0].strip() if "@" in email else email.strip()


def _delivery_type(from_addr: str, to_addr: str) -> str:
    return (
        "Internal"
        if (VOLIBITS_DOMAIN in from_addr and VOLIBITS_DOMAIN in to_addr)
        else "External"
    )


# ─── Record validation ───────────────────────────────────────────────────────────
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


# ─── DB insertion ────────────────────────────────────────────────────────────────
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


# ─── Date parsing ────────────────────────────────────────────────────────────────
_DATE_FORMATS = ("%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y")


def _parse_date(raw: str) -> Optional[date]:
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(raw.strip(), fmt).date()
        except ValueError:
            continue
    return None


# ─── Main pipeline ───────────────────────────────────────────────────────────────
def process_emails() -> None:
    log.info("=== HR Email Extractor starting ===")
    log.info(f"Target table   : {DB_TABLE}")
    log.info(f"OneDrive folder: {ONEDRIVE_FOLDER}")

    token    = get_mail_token()
    od_token = get_onedrive_token()
    log.info("Tokens obtained.")

    conn = psycopg2.connect(DB_DSN)
    conn.autocommit = False
    cur  = conn.cursor()

    # Ensure ancillary tables exist
    ensure_tables(cur)
    conn.commit()

    # ── Resolve any pending conflict action-replies first ────────────────────
    process_action_replies(token, cur, conn)

    emails = fetch_emails(token)
    log.info(f"Fetched {len(emails)} unread email(s).")

    processed = skipped = inserted = updated = conflicts = errors = 0

    for msg in emails:
        subject = msg.get("subject", "").strip()

        # Skip [HR-ACTION] emails — already handled above
        if ACTION_REPLY_RE.match(subject):
            mark_email_read(token, msg["id"])
            continue

        log.info(f"Subject: {subject!r}")

        parsed = parse_subject(subject)
        if parsed is None:
            skipped += 1
            mark_email_read(token, msg["id"])
            continue

        company_name, general_skill, subject_jr_no = parsed

        from_addr = _extract_address(msg.get("from", {}))
        to_list   = msg.get("toRecipients", [])
        to_addr   = _extract_address(to_list[0]) if to_list else ""

        recruiter        = _recruiter_name(from_addr)
        client_recruiter = _username_from_email(to_addr)
        delivery_type    = _delivery_type(from_addr, to_addr)
        attachment_str   = upload_attachments(od_token, token, msg)

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
            row_date        = _parse_date(row["date"]) if row.get("date") else email_date
            contact_number  = _t(row.get("contact_number"))
            email_id_val    = _t(row.get("email_id"))
            effective_jr_no = _t(row.get("jr_no")) or subject_jr_no
            candidate_name  = _t(row.get("name_of_candidate"))

            # ── TIER 1: Same recruiter — auto-update empty fields ────────────
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

            # ── TIER 2: Different recruiter — conflict resolution required ───
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
                conflict_id = str(uuid.uuid4())

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
                    "is_duplicate":        "Conflict",
                    "created_by":          _t(from_addr),
                    "created_date":        now,
                    "modified_by":         _t(from_addr),
                    "modified_date":       now,
                    "remarks":             _t(row.get("remarks")),
                    "record_status":       "Pending",
                    "status":              "Screen Pending",
                }

                store_pending_conflict(cur, conflict_id, diff_key, new_record_data)

                # Reconstruct full email for existing recruiter (always @volibits.com)
                existing_email_from = diff_key.get("email_from") or (
                    diff_key.get("recruiter", "") + "@" + VOLIBITS_DOMAIN
                )

                send_conflict_notification_email(
                    token=token,
                    new_recruiter_addr=from_addr,
                    existing_recruiter_addr=existing_email_from,
                    conflict_id=conflict_id,
                    existing_row=diff_key,
                    new_record_data=new_record_data,
                    original_subject=subject,
                )

                email_conflicts += 1
                conflicts += 1
                log.info(
                    f"  ⚠ Conflict raised for {candidate_name} — "
                    f"existing recruiter: {diff_key.get('recruiter')} | "
                    f"conflict_id: {conflict_id}"
                )

                rows_summary.append({
                    "name":           candidate_name or "(unknown)",
                    "outcome":        "conflict",
                    "record_data":    new_record_data,
                    "updated_fields": [],
                    "missing":        [],
                    "dup_flag":       "",
                    "db_error":       None,
                    "conflict_id":    conflict_id,
                    "existing_recruiter": diff_key.get("recruiter"),
                })
                continue

            # ── TIER 3: No matching record — standard insert path ────────────
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

        mark_email_read(token, msg["id"])
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