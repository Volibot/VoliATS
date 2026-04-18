"""
Outlook Email Extractor for HR Recruitment (Volibits)
Reads emails via Microsoft Graph API and inserts candidate data into PostgreSQL.

recruiter        = local part of sender email   — sender is always @volibits.com
client_recruiter = local part of receiver email — can be any domain
"""

import os
import re
import logging
import requests
import psycopg2
from psycopg2 import sql as pgsql          # ← NEW: safe identifier quoting
from datetime import datetime, date
from typing import Optional
from msal import ConfidentialClientApplication

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


# ─── Configuration — loaded from environment variables ──────────────────────────
AZURE_TENANT_ID     = os.environ["AZURE_TENANT_ID"]
AZURE_CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
AZURE_CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
TARGET_MAILBOX      = os.environ["TARGET_MAILBOX"]          # HRvolibot@volibits.com
DB_DSN              = os.environ["DB_DSN"]

# ── NEW: table name and OneDrive folder from secrets ────────────────────────────
DB_TABLE            = os.environ["DB_TABLE_NAME"]           # e.g. hrvolibit
ONEDRIVE_FOLDER     = os.environ.get("ONEDRIVE_FOLDER", "HR Resumes")
# ────────────────────────────────────────────────────────────────────────────────

VOLIBITS_DOMAIN = "volibits.com"


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
    # Add more as needed — KEY must match the prefix in the email subject
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
    ]
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
def get_access_token() -> str:
    authority = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}"
    app = ConfidentialClientApplication(
        AZURE_CLIENT_ID,
        authority=authority,
        client_credential=AZURE_CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise RuntimeError(
            f"Token acquisition failed: {result.get('error_description')}"
        )
    return result["access_token"]


# ─── Fetch & mark emails ─────────────────────────────────────────────────────────
def fetch_emails(token: str, top: int = 50) -> list[dict]:
    """Fetch unread emails from the target mailbox, newest first."""
    headers = {"Authorization": f"Bearer {token}"}
    url = (
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/messages"
        f"?$top={top}"
        f"&$select=id,subject,from,toRecipients,ccRecipients,"
        f"body,receivedDateTime,isRead,hasAttachments"
        f"&$expand=attachments($select=id,name,contentType,size)"  # ← NEW: include id & contentType
        f"&$filter=isRead eq false"
        f"&$orderby=receivedDateTime desc"
    )
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.json().get("value", [])


def mark_email_read(token: str, message_id: str) -> None:
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    url = (
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}"
        f"/messages/{message_id}"
    )
    requests.patch(url, headers=headers, json={"isRead": True}, timeout=10)


# ─── Subject parsing ─────────────────────────────────────────────────────────────
SUBJECT_RE = re.compile(
    r"^(?P<code>[A-Z]{2,5})\s*:\s*(?P<rest>.+)$", re.IGNORECASE
)
IGNORE_RE = re.compile(r"^(fw|fwd|re|aw)\s*:", re.IGNORECASE)

_JR_PATTERNS = [
    re.compile(r"[-|]\s*(?:jr\s*)?(?P<jr>\d{4,})\s*$",          re.IGNORECASE),
    re.compile(r"^\s*(?:jr\s*)?(?P<jr>\d{4,})\s*[-|]",          re.IGNORECASE),
    re.compile(r"\(\s*jr\s*(?P<jr>\d{4,})\s*\)",               re.IGNORECASE),
    re.compile(r"\(\s*(?P<jr>\d{4,})\s*\)",                     re.IGNORECASE),
    re.compile(r"\bjr\s*(?P<jr>\d{4,})\b",                       re.IGNORECASE),
    re.compile(r"\b(?P<jr>\d{4,})\b",                             re.IGNORECASE),
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

    skill = re.sub(r"\s{2,}", " ", rest).strip()
    return jr_no, skill


def parse_subject(subject: str) -> Optional[tuple[str, str, str, Optional[str]]]:
    subject = subject.strip()
    if IGNORE_RE.match(subject):
        log.debug(f"Skip — FW/RE: {subject!r}")
        return None
    m = SUBJECT_RE.match(subject)
    if not m:
        log.debug(f"Skip — no pattern match: {subject!r}")
        return None
    code    = m.group("code").upper()
    rest    = m.group("rest")
    company = COMPANY_CODES.get(code, code)
    jr_no, skill = _extract_jr_and_skill(rest)
    log.debug(f"Parsed subject → code={code} skill={skill!r} jr_no={jr_no!r}")
    return code, company, skill, jr_no


# ─── HTML table parser ───────────────────────────────────────────────────────────
def _strip_html(text: str) -> str:
    return re.sub(r"<[^>]+>", "", text)


def _clean_cell(text: str) -> str:
    return re.sub(r"\s+", " ", _strip_html(text)).strip()


def parse_html_table(html: str) -> list[dict]:
    rows: list[dict] = []
    tables = re.findall(
        r"<table[^>]*>(.*?)</table>", html, re.DOTALL | re.IGNORECASE
    )
    for table_html in tables:
        all_rows = re.findall(
            r"<tr[^>]*>(.*?)</tr>", table_html, re.DOTALL | re.IGNORECASE
        )
        if not all_rows:
            continue

        raw_headers = re.findall(
            r"<t[hd][^>]*>(.*?)</t[hd]>",
            all_rows[0],
            re.DOTALL | re.IGNORECASE,
        )
        headers = [_clean_cell(h) for h in raw_headers]
        if not headers:
            continue

        col_map: dict[int, str] = {}
        for idx, h in enumerate(headers):
            canonical = _resolve_header(h)
            if canonical:
                col_map[idx] = canonical
            else:
                log.debug(f"Unmapped header: {h!r}")

        for row_html in all_rows[1:]:
            cells = re.findall(
                r"<t[hd][^>]*>(.*?)</t[hd]>",
                row_html,
                re.DOTALL | re.IGNORECASE,
            )
            if not cells:
                continue
            cleaned = [_clean_cell(c) for c in cells]
            if all(v == "" for v in cleaned):
                continue
            record: dict = {
                col_map[i]: v
                for i, v in enumerate(cleaned)
                if i in col_map
            }
            if record:
                rows.append(record)
    return rows


# ─── NEW: OneDrive attachment upload ────────────────────────────────────────────

def _fetch_attachment_content(token: str, message_id: str, attachment_id: str) -> bytes:
    """Download raw bytes of a single email attachment via Graph API."""
    url = (
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}"
        f"/messages/{message_id}/attachments/{attachment_id}/$value"
    )
    resp = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        timeout=60,
    )
    resp.raise_for_status()
    return resp.content


def _upload_to_onedrive(token: str, filename: str, content: bytes) -> Optional[str]:
    """
    Upload *content* to  ONEDRIVE_FOLDER/<filename>  in TARGET_MAILBOX's OneDrive.
    Returns the SharePoint sharing link on success, None on failure.

    Graph PUT endpoint handles files up to ~4 MB in a single request.
    Larger files would need an upload session — add that path if needed.
    """
    # URL-encode only the slashes in the folder path; Graph handles the rest
    remote_path = f"{ONEDRIVE_FOLDER}/{filename}"
    upload_url = (
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}"
        f"/drive/root:/{remote_path}:/content"
    )
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/octet-stream",
    }
    resp = requests.put(upload_url, headers=headers, data=content, timeout=120)
    if resp.status_code not in (200, 201):
        log.error(
            f"OneDrive upload failed for {filename!r}: "
            f"{resp.status_code} — {resp.text[:200]}"
        )
        return None

    item = resp.json()
    item_id = item.get("id")

    # Request an org-scoped view link so the URL is stable and human-readable
    link_url = (
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}"
        f"/drive/items/{item_id}/createLink"
    )
    link_resp = requests.post(
        link_url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json={"type": "view", "scope": "organization"},
        timeout=30,
    )
    if link_resp.status_code in (200, 201):
        share_link = link_resp.json().get("link", {}).get("webUrl", "")
        if share_link:
            log.info(f"  ↑ Uploaded {filename!r} → {share_link}")
            return share_link

    # Fall back to the raw webUrl from the upload response if createLink fails
    web_url = item.get("webUrl", "")
    log.warning(
        f"createLink failed for {filename!r} ({link_resp.status_code}); "
        f"using webUrl: {web_url}"
    )
    return web_url or None


def upload_attachments(token: str, msg: dict) -> str:
    """
    Upload all file attachments in *msg* to OneDrive.
    Returns a comma-separated string of SharePoint links (or filenames on failure).
    """
    attachments = msg.get("attachments") or []
    if not attachments:
        return ""

    links: list[str] = []
    msg_id = msg["id"]

    for att in attachments:
        att_id   = att.get("id", "")
        filename = att.get("name", "attachment")

        if not att_id:
            links.append(filename)
            continue

        try:
            content = _fetch_attachment_content(token, msg_id, att_id)
            link    = _upload_to_onedrive(token, filename, content)
            links.append(link if link else filename)
        except Exception as exc:
            log.error(f"Failed to upload attachment {filename!r}: {exc}")
            links.append(filename)   # degrade gracefully — store filename at minimum

    return ", ".join(links)

# ────────────────────────────────────────────────────────────────────────────────


# ─── Duplicate detection ─────────────────────────────────────────────────────────
def is_exact_duplicate(
    cur,
    row_date,
    contact_number: str,
    email_id: str,
    name: str,
    general_skill: str,
    company_name: str,
    current_ctc: str,
    expected_ctc: str,
    client_recruiter: str,
    email_to: str,
    jr_no: str,
    recruiter: str,
) -> bool:
    if not contact_number and not email_id:
        return False

    query = pgsql.SQL("""
        SELECT 1 FROM {table}
        WHERE date = %s
          AND LOWER(COALESCE(name_of_candidate,  '')) = LOWER(COALESCE(%s, ''))
          AND LOWER(COALESCE(general_skill,       '')) = LOWER(COALESCE(%s, ''))
          AND LOWER(COALESCE(company_name,        '')) = LOWER(COALESCE(%s, ''))
          AND LOWER(COALESCE(contact_number,      '')) = LOWER(COALESCE(%s, ''))
          AND LOWER(COALESCE(email_id,            '')) = LOWER(COALESCE(%s, ''))
          AND LOWER(COALESCE(current_ctc,         '')) = LOWER(COALESCE(%s, ''))
          AND LOWER(COALESCE(expected_ctc,        '')) = LOWER(COALESCE(%s, ''))
          AND LOWER(COALESCE(client_recruiter,    '')) = LOWER(COALESCE(%s, ''))
          AND LOWER(COALESCE(email_to,            '')) = LOWER(COALESCE(%s, ''))
          AND LOWER(COALESCE(jr_no,               '')) = LOWER(COALESCE(%s, ''))
          AND LOWER(COALESCE(recruiter,           '')) = LOWER(COALESCE(%s, ''))
        LIMIT 1
    """).format(table=pgsql.Identifier(DB_TABLE))

    cur.execute(query, (
        row_date, name, general_skill, company_name,
        contact_number, email_id, current_ctc, expected_ctc,
        client_recruiter, email_to, jr_no, recruiter,
    ))
    return cur.fetchone() is not None


def check_duplicate(cur, contact_number: str, email_id: str) -> str:
    if not contact_number and not email_id:
        return ""

    tbl = pgsql.Identifier(DB_TABLE)

    if contact_number and email_id:
        cur.execute(
            pgsql.SQL(
                "SELECT 1 FROM {t} WHERE contact_number=%s AND email_id=%s LIMIT 1"
            ).format(t=tbl),
            (contact_number, email_id),
        )
        if cur.fetchone():
            return "Duplicate"

    cell_dup = email_dup = False

    if contact_number:
        cur.execute(
            pgsql.SQL(
                "SELECT 1 FROM {t} WHERE contact_number=%s LIMIT 1"
            ).format(t=tbl),
            (contact_number,),
        )
        cell_dup = bool(cur.fetchone())

    if email_id:
        cur.execute(
            pgsql.SQL(
                "SELECT 1 FROM {t} WHERE email_id=%s LIMIT 1"
            ).format(t=tbl),
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
    if "@" in email:
        return email.split("@", 1)[0].strip()
    return email.strip()


def _username_from_email(email: str) -> str:
    if "@" in email:
        return email.split("@", 1)[0].strip()
    return email.strip()


def _delivery_type(from_addr: str, to_addr: str) -> str:
    from_internal = VOLIBITS_DOMAIN in from_addr
    to_internal   = VOLIBITS_DOMAIN in to_addr
    return "Internal" if (from_internal and to_internal) else "External"


# ─── DB insertion ────────────────────────────────────────────────────────────────
# INSERT_SQL is built at call time using pgsql.SQL so the table name is
# a proper quoted identifier, never raw string interpolation in the query itself.
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


_SAFE_FIELDS = (
    "recruiter", "client_recruiter", "general_skill", "company_name",
    "email_from", "email_to", "delivery_type", "email_id",
    "name_of_candidate", "contact_number", "is_duplicate", "status",
    "created_by", "created_date", "modified_by", "modified_date",
    "record_status", "attachment", "remarks"
)


def _t(val: Optional[str]) -> Optional[str]:
    if val is None:
        return None
    v = val.strip()
    return v or None


# ─── Record validation ──────────────────────────────────────────────────────────
_REQUIRED_FIELDS = (
    "name_of_candidate",
    "contact_number",
    "email_id",
    "general_skill",
    "company_name",
    "recruiter",
    "email_from",
    "email_to",
)


def _record_status(data: dict) -> str:
    missing = [f for f in _REQUIRED_FIELDS if not data.get(f)]
    if missing:
        log.warning(f"record_status=Fail — missing fields: {missing}")
        return "Fail"
    return "Pass"


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


# ─── Date parsing helper ─────────────────────────────────────────────────────────
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
    log.info(f"Target table : {DB_TABLE}")
    log.info(f"OneDrive folder: {ONEDRIVE_FOLDER}")

    token = get_access_token()
    log.info("Access token obtained.")

    emails = fetch_emails(token)
    log.info(f"Fetched {len(emails)} unread email(s).")

    conn = psycopg2.connect(DB_DSN)
    conn.autocommit = False
    cur  = conn.cursor()

    processed = skipped = inserted = errors = 0

    for msg in emails:
        subject = msg.get("subject", "").strip()
        log.info(f"Subject: {subject!r}")

        parsed = parse_subject(subject)
        if parsed is None:
            skipped += 1
            mark_email_read(token, msg["id"])
            continue

        _code, company_name, general_skill, subject_jr_no = parsed

        # ── Sender / receiver ────────────────────────────────────────────────
        from_addr = _extract_address(msg.get("from", {}))
        to_list   = msg.get("toRecipients", [])
        to_addr   = _extract_address(to_list[0]) if to_list else ""

        recruiter        = _recruiter_name(from_addr)
        client_recruiter = _username_from_email(to_addr)
        delivery_type    = _delivery_type(from_addr, to_addr)

        # ── NEW: upload attachments → SharePoint links ───────────────────────
        attachment_str = upload_attachments(token, msg)
        # ────────────────────────────────────────────────────────────────────

        # ── Parse table from email body ──────────────────────────────────────
        body_html = (msg.get("body") or {}).get("content", "")
        rows      = parse_html_table(body_html)

        if not rows:
            log.warning(f"No table rows found — inserting skeleton record.")
            rows = [{}]

        email_date: Optional[date] = None
        raw_dt = msg.get("receivedDateTime", "")
        if raw_dt:
            try:
                email_date = datetime.fromisoformat(
                    raw_dt.replace("Z", "+00:00")
                ).date()
            except ValueError:
                pass

        now = datetime.now()

        for row in rows:
            row_date = _parse_date(row["date"]) if row.get("date") else email_date

            contact_number = _t(row.get("contact_number"))
            email_id_val   = _t(row.get("email_id"))
            effective_jr_no = _t(row.get("jr_no")) or subject_jr_no

            if is_exact_duplicate(
                cur,
                row_date,
                contact_number or "",
                email_id_val   or "",
                _t(row.get("name_of_candidate")) or "",
                _t(general_skill)                or "",
                company_name                     or "",
                _t(row.get("current_ctc"))       or "",
                _t(row.get("expected_ctc"))      or "",
                _t(client_recruiter)             or "",
                _t(to_addr)                      or "",
                effective_jr_no                  or "",
                _t(recruiter)                    or "",
            ):
                log.info(
                    f"  ⟳ Exact duplicate on same day — skipped: "
                    f"{row.get('name_of_candidate', '?')} | {contact_number} | "
                    f"jr={effective_jr_no} | recruiter={recruiter} | "
                    f"client={client_recruiter} | to={to_addr}"
                )
                skipped += 1
                continue

            dup_flag = check_duplicate(cur, contact_number or "", email_id_val or "")

            record_data = {
                "recruiter":           _t(recruiter),
                "date":                row_date,
                "jr_no":               effective_jr_no,
                "client_recruiter":    _t(client_recruiter),
                "general_skill":       _t(general_skill),
                "name_of_candidate":   _t(row.get("name_of_candidate")),
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
                "attachment":          _t(attachment_str),   # ← SharePoint link(s)
                "is_duplicate":        _t(dup_flag),
                "created_by":          _t(from_addr),
                "created_date":        now,
                "modified_by":         _t(from_addr),
                "modified_date":       now,
                "remarks":             _t(row.get("remarks")),
                "record_status":       _record_status({
                    "name_of_candidate": _t(row.get("name_of_candidate")),
                    "contact_number":    contact_number,
                    "email_id":          email_id_val,
                    "general_skill":     _t(general_skill),
                    "company_name":      company_name,
                    "recruiter":         _t(recruiter),
                    "email_from":        _t(from_addr),
                    "email_to":          _t(to_addr),
                }),
                "status":              "Screen Pending",
            }

            ok = insert_record(cur, record_data)
            if ok:
                inserted += 1
                log.info(
                    f"  ✓ {row.get('name_of_candidate', '(unknown)')} | "
                    f"recruiter={recruiter} | client={client_recruiter} | "
                    f"dup={dup_flag or 'none'} | attachment={_t(attachment_str) or 'none'}"
                )
            else:
                errors += 1

        try:
            conn.commit()
        except Exception as e:
            log.error(f"Commit failed: {e}")
            conn.rollback()

        mark_email_read(token, msg["id"])
        processed += 1

    cur.close()
    conn.close()

    log.info(
        f"=== Done — processed={processed} skipped={skipped} "
        f"inserted={inserted} errors={errors} ==="
    )


if __name__ == "__main__":
    process_emails()