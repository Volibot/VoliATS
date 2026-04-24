"""
Ad-hoc extractor: pulls BS candidate emails from "Company Profiles" subfolder
received on or after 2026-01-01 and writes all parsed fields to an Excel file.

Subject filters accepted:
  - BS: ...
  - [External]: BS: ...

No database writes. No notifications. No OneDrive uploads. Read-only run.

Output: bs_candidates_<timestamp>.xlsx in the current working directory.

Required env vars:
  AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET
  TARGET_MAILBOX

Optional:
  INBOX_SUBFOLDER   (default: "Company Profiles")

── Local run ────────────────────────────────────────────────────────────────────
1. Install dependencies:
       pip install msal openpyxl requests python-dotenv

2. Create a .env file next to this script (⚠️ never commit it to git):

       AZURE_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
       AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
       AZURE_CLIENT_SECRET=your-secret-here
       TARGET_MAILBOX=mailbox@yourdomain.com

3. Run:
       python -m dotenv run -- python bs_extractor.py

   Or export vars manually:
       export $(cat .env | xargs) && python bs_extractor.py

   Output Excel is written to the current directory as bs_candidates_<timestamp>.xlsx
────────────────────────────────────────────────────────────────────────────────
"""

import os
import re
import logging
import requests
from datetime import datetime, date, timezone
from typing import Optional

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from msal import ConfidentialClientApplication
from dotenv import load_dotenv
load_dotenv()
# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(), logging.FileHandler("bs_extract.log")],
)
log = logging.getLogger(__name__)

# ── Config ────────────────────────────────────────────────────────────────────
AZURE_TENANT_ID     = os.environ["AZURE_TENANT_ID"]
AZURE_CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
AZURE_CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]

TARGET_MAILBOX  = os.environ["TARGET_MAILBOX"]
INBOX_SUBFOLDER = os.environ.get("INBOX_SUBFOLDER", "Company Profiles")
VOLIBITS_DOMAIN = "volibits.com"
RESUME_EXTENSIONS = {".pdf", ".doc", ".docx"}

# Only pull emails from this date onwards
SINCE_DATE = date(2026, 1, 1)

# Excel columns — exactly matching the DB fields
EXCEL_COLUMNS = [
    "email_subject",
    "recruiter", "client_recruiter", "email_from", "email_to",
    "delivery_type", "company_name", "general_skill",
    "date", "jr_no",
    "name_of_candidate", "contact_number", "email_id",
    "total_experience", "relevant_experience",
    "current_ctc", "expected_ctc", "notice_period",
    "current_org", "current_location", "preferred_location",
    "attachment", "remarks", "record_status",
]

# ── Column aliases ─────────────────────────────────────────────────────────────
COLUMN_ALIASES: dict[str, list[str]] = {
    "general_skill": [
        "general_skill", "generalskill", "gen_skill", "genskill",
        "skill", "requirement",
    ],
    "jr_no": [
        "jr no", "jr_no", "jr no.", "jr", "jrno", "req_id", "req id",
        "requisition id", "requisition_id", "jr_number", "jr number",
        "job req", "job requisition", "job_req", "reqid",
        "requirement id", "requirement_id",
        "jr no(mention the jr number where profiles are uploaded on sf).",
        "jr_no.", "(mention the jr number where profiles are uploaded on sf).",
    ],
    "name_of_candidate": [
        "candidate name", "candidate_name", "name", "applicant name",
        "applicant_name", "full name", "full_name", "name of candidate",
        "candidate", "person name",
        "n_of_candidate", "noc", "name of cand", "name_of_cand",
        "name_of_candidate",
    ],
    "contact_number": [
        "contact number", "contact_number", "phone", "mobile", "cell",
        "phone number", "phone_number", "mobile number", "mobile_number",
        "contact no", "contact_no", "ph no", "ph_no", "contact",
        "c_number", "con_number", "con_no.", "con_num", "con no.", "cno.",
        "mobile no.", "mobile no", "contact no", "contact number",
    ],
    "email_id": [
        "email id", "email_id", "email", "e-mail", "mail id", "mail_id",
        "email address", "email_address", "e mail",
        "e_id", "e id", "emailid", "e-mail id",
    ],
    "current_org": [
        "current company", "current_company", "company", "employer",
        "current employer", "current_employer", "current org", "current_org",
        "organisation", "organization", "curr company", "present company",
        "curr_org", "current_organization", "current comp", "curr comp",
        "current / last company",
    ],
    "total_experience": [
        "total experience", "total_experience", "total exp", "total_exp",
        "experience", "exp", "total yrs", "total years",
        "tot exp", "tot_exp", "total exp (yrs)",
    ],
    "relevant_experience": [
        "relevant experience", "relevant_experience", "relevant exp",
        "relevant_exp", "rel exp", "rel_exp", "relevant yrs",
        "rel exp.", "rel experience", "rel. exp", "rel.exp",
        "rel_experience", "rel.experience", "relevant exp (yrs)",
    ],
    "current_ctc": [
        "current ctc", "current_ctc", "ctc", "current salary",
        "current_salary", "present ctc", "present_ctc", "curr ctc",
        "cctc", "currctc", "curr_ctc", "c_ctc",
        "current ctc (lakhs)", "ctcc",
    ],
    "expected_ctc": [
        "expected ctc", "expected_ctc", "expected salary", "expected_salary",
        "exp ctc", "exp_ctc", "desired ctc", "desired_ctc",
        "ectc", "e_ctc", "epected ctc", "expctc", "ect",
        "exp ctc (lakhs)", "exp ctc",
    ],
    "notice_period": [
        "notice period", "notice_period", "notice", "np", "joining time",
        "notice days", "notice_days",
        "notice period(days)",
    ],
    "current_location": [
        "current location", "current_location", "location", "city",
        "current city", "current_city", "present location",
        "current_loc", "curr location", "c loc", "c_location", "curr_location",
        "curr_loc", "current/preferred location",
    ],
    "preferred_location": [
        "preferred location", "preferred_location", "preferred city",
        "preferred_city", "pref location", "pref_location",
        "willing to relocate", "target location",
        "pref_loc", "pre location", "pref_location",
        "preferred location", "current/preferred location",
    ],
    "date": [
        "date", "submission date", "submission_date", "applied date",
        "applied_date", "profile date",
        "date of submission",
    ],
    "remarks": [
        "remarks", "comments", "comment", "note", "notes",
        "remarks/availability for interview",
    ],
}

def _normalize(s: str) -> str:
    return re.sub(r"\s+", " ", s.strip().lower())

ALIAS_MAP: dict[str, str] = {
    _normalize(alias): canonical
    for canonical, aliases in COLUMN_ALIASES.items()
    for alias in aliases
}

def _resolve_header(raw: str) -> Optional[str]:
    return ALIAS_MAP.get(_normalize(raw))

# ── Auth ──────────────────────────────────────────────────────────────────────
def get_mail_token() -> str:
    result = ConfidentialClientApplication(
        AZURE_CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{AZURE_TENANT_ID}",
        client_credential=AZURE_CLIENT_SECRET,
    ).acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Token failed: {result.get('error_description')}")
    return result["access_token"]

# ── Folder resolution ─────────────────────────────────────────────────────────
def resolve_folder_id(token: str, folder_name: str) -> Optional[str]:
    if not folder_name:
        return None
    headers = {"Authorization": f"Bearer {token}"}
    url = (
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/mailFolders"
        f"?$select=id,displayName&$top=50"
    )
    resp = requests.get(url, headers=headers, timeout=15)
    if not resp.ok:
        return None
    for folder in resp.json().get("value", []):
        if folder.get("displayName", "").strip().lower() == folder_name.strip().lower():
            return folder["id"]
        child_url = (
            f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}"
            f"/mailFolders/{folder['id']}/childFolders?$select=id,displayName&$top=50"
        )
        cr = requests.get(child_url, headers=headers, timeout=15)
        if not cr.ok:
            continue
        for child in cr.json().get("value", []):
            if child.get("displayName", "").strip().lower() == folder_name.strip().lower():
                return child["id"]
    log.warning(f"Folder {folder_name!r} not found — will use root inbox")
    return None

# ── Fetch emails (paginated, filtered by date) ────────────────────────────────
def fetch_bs_emails(token: str, folder_id: Optional[str]) -> list[dict]:
    headers = {"Authorization": f"Bearer {token}"}
    since_iso = f"{SINCE_DATE.isoformat()}T00:00:00Z"

    if folder_id:
        base = (
            f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}"
            f"/mailFolders/{folder_id}/messages"
        )
    else:
        base = f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/messages"

    url = (
        f"{base}"
        f"?$top=100"
        f"&$select=id,subject,from,toRecipients,body,receivedDateTime,hasAttachments"
        # ← removed $expand=attachments here
        f"&$filter=receivedDateTime ge {since_iso}"
        f"&$orderby=receivedDateTime asc"
    )

    all_emails: list[dict] = []
    while url:
        resp = requests.get(url, headers=headers, timeout=60)  # increased timeout
        resp.raise_for_status()
        data = resp.json()
        all_emails.extend(data.get("value", []))
        url = data.get("@odata.nextLink")

    log.info(f"Fetched {len(all_emails)} email(s) since {SINCE_DATE}.")
    return all_emails


# ── Attachment filenames (fetched on-demand, no upload) ───────────────────────
def _is_resume(filename: str) -> bool:
    return os.path.splitext(filename)[1].lower() in RESUME_EXTENSIONS

def get_attachment_names(token: str, msg: dict) -> str:
    """Fetch attachment metadata only for this message and return resume filenames."""
    if not msg.get("hasAttachments"):
        return ""
    url = (
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}"
        f"/messages/{msg['id']}/attachments?$select=name,contentType"
    )
    try:
        resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
        resp.raise_for_status()
        attachments = resp.json().get("value", [])
        names = [a.get("name", "") for a in attachments if _is_resume(a.get("name", ""))]
        return ", ".join(names)
    except Exception as exc:
        log.warning(f"Could not fetch attachments for {msg['id']}: {exc}")
        return ""
# ── Subject filter ────────────────────────────────────────────────────────────
BS_SUBJECT_RE = re.compile(
    r"^(?:\[External\]\s*:\s*)?BS\s*:",
    re.IGNORECASE,
)

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
    jr_no = None
    for pat in _JR_PATTERNS:
        m = pat.search(rest)
        if m:
            jr_no = m.group("jr")
            s, e = m.span()
            rest = re.sub(r"^[-|\s]+|[-|\s]+$", "", (rest[:s] + " " + rest[e:]).strip()).strip()
            break
    return jr_no, re.sub(r"\s{2,}", " ", rest).strip()

def parse_bs_subject(subject: str) -> Optional[tuple[str, Optional[str]]]:
    """Returns (general_skill, jr_no) for matching subjects, else None."""
    subject = subject.strip()
    if not BS_SUBJECT_RE.match(subject):
        return None
    rest = BS_SUBJECT_RE.sub("", subject).strip()
    jr_no, skill = _extract_jr_and_skill(rest)
    return skill, jr_no

# ── HTML table parser ─────────────────────────────────────────────────────────
def _strip_html(text: str) -> str:
    return re.sub(r"<[^>]+>", "", text)

def _clean_cell(text: str) -> str:
    return re.sub(r"\s+", " ", _strip_html(text)).strip()

def parse_html_table(html: str) -> list[dict]:
    rows: list[dict] = []
    for tbl in re.findall(r"<table[^>]*>(.*?)</table>", html, re.DOTALL | re.IGNORECASE):
        all_rows = re.findall(r"<tr[^>]*>(.*?)</tr>", tbl, re.DOTALL | re.IGNORECASE)
        if not all_rows:
            continue
        raw_headers = re.findall(r"<t[hd][^>]*>(.*?)</t[hd]>", all_rows[0], re.DOTALL | re.IGNORECASE)
        headers = [_clean_cell(h) for h in raw_headers]
        col_map = {i: _resolve_header(h) for i, h in enumerate(headers) if _resolve_header(h)}
        for row_html in all_rows[1:]:
            cells = re.findall(r"<t[hd][^>]*>(.*?)</t[hd]>", row_html, re.DOTALL | re.IGNORECASE)
            if not cells:
                continue
            cleaned = [_clean_cell(c) for c in cells]
            if all(v == "" for v in cleaned):
                continue
            record = {col_map[i]: v for i, v in enumerate(cleaned) if i in col_map}
            if record:
                rows.append(record)
    return rows

# ── Attachment filenames (no upload) ─────────────────────────────────────────
def _is_resume(filename: str) -> bool:
    return os.path.splitext(filename)[1].lower() in RESUME_EXTENSIONS

def get_attachment_names(msg: dict) -> str:
    """Returns a comma-separated string of resume attachment filenames."""
    attachments = msg.get("attachments") or []
    names = [
        att.get("name", "")
        for att in attachments
        if _is_resume(att.get("name", ""))
    ]
    return ", ".join(names)

# ── Helpers ───────────────────────────────────────────────────────────────────
def _extract_address(addr_obj: dict) -> str:
    try:
        return addr_obj["emailAddress"]["address"].strip().lower()
    except (KeyError, TypeError):
        return ""

def _normalize_name_from_email(email: str) -> str:
    if not email or "@" not in email:
        return ""
    local = re.sub(r"\d+", "", email.split("@", 1)[0])
    local = re.sub(r"[._\-]+", " ", local)
    return re.sub(r"\s+", " ", local).strip().title()

_DATE_FORMATS = ("%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y")

def _parse_date(raw: str) -> Optional[date]:
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(raw.strip(), fmt).date()
        except ValueError:
            continue
    return None

_REQUIRED_FIELDS = (
    "name_of_candidate", "contact_number", "email_id",
    "general_skill", "company_name", "recruiter", "email_from", "email_to",
)

def _record_status(data: dict) -> str:
    missing = [f for f in _REQUIRED_FIELDS if not data.get(f)]
    return "Fail" if missing else "Pass"

def _t(val) -> Optional[str]:
    if val is None:
        return None
    v = str(val).strip()
    return v or None

# ── Excel writer ──────────────────────────────────────────────────────────────
HEADER_FILL  = PatternFill("solid", start_color="1F4E79")
HEADER_FONT  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
CELL_FONT    = Font(name="Arial", size=10)
CENTER       = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT         = Alignment(horizontal="left",   vertical="center", wrap_text=True)

COL_WIDTHS = {
    "email_subject": 40, "recruiter": 18, "client_recruiter": 18,
    "email_from": 28, "email_to": 28, "delivery_type": 14,
    "company_name": 16, "general_skill": 22, "date": 12, "jr_no": 10,
    "name_of_candidate": 22, "contact_number": 16, "email_id": 28,
    "total_experience": 14, "relevant_experience": 16,
    "current_ctc": 14, "expected_ctc": 14, "notice_period": 14,
    "current_org": 22, "current_location": 18, "preferred_location": 18,
    "attachment": 35, "remarks": 30, "record_status": 14,
}

PRETTY_HEADERS = {
    "email_subject": "Email Subject", "recruiter": "Recruiter",
    "client_recruiter": "Client Recruiter", "email_from": "Email From",
    "email_to": "Email To", "delivery_type": "Delivery Type",
    "company_name": "Company", "general_skill": "General Skill",
    "date": "Date", "jr_no": "JR No",
    "name_of_candidate": "Candidate Name", "contact_number": "Contact Number",
    "email_id": "Email ID", "total_experience": "Total Exp",
    "relevant_experience": "Relevant Exp", "current_ctc": "Current CTC",
    "expected_ctc": "Expected CTC", "notice_period": "Notice Period",
    "current_org": "Current Org", "current_location": "Current Location",
    "preferred_location": "Preferred Location", "attachment": "Attachment",
    "remarks": "Remarks", "record_status": "Record Status",
}

def write_excel(records: list[dict], output_path: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BS Candidates"
    ws.freeze_panes = "A2"

    for col_idx, col in enumerate(EXCEL_COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx, value=PRETTY_HEADERS.get(col, col))
        cell.font      = HEADER_FONT
        cell.fill      = HEADER_FILL
        cell.alignment = CENTER
        ws.column_dimensions[
            openpyxl.utils.get_column_letter(col_idx)
        ].width = COL_WIDTHS.get(col, 16)

    ws.row_dimensions[1].height = 22

    STATUS_COLORS = {"Pass": "C6EFCE", "Fail": "FFC7CE"}

    for row_idx, rec in enumerate(records, 2):
        for col_idx, col in enumerate(EXCEL_COLUMNS, 1):
            val  = rec.get(col)
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font      = CELL_FONT
            cell.alignment = LEFT
            if col == "record_status" and val in STATUS_COLORS:
                cell.fill = PatternFill("solid", start_color=STATUS_COLORS[val])

        ws.row_dimensions[row_idx].height = 16

    ws.auto_filter.ref = ws.dimensions

    wb.save(output_path)
    log.info(f"Saved {len(records)} row(s) → {output_path}")

# ── Main ──────────────────────────────────────────────────────────────────────
def run() -> None:
    log.info("=== BS Ad-hoc Extractor starting ===")
    log.info(f"Pulling emails from: {SINCE_DATE} onwards")
    log.info(f"Folder: {INBOX_SUBFOLDER or '(root inbox)'}")

    mail_token = get_mail_token()
    log.info("Token obtained.")

    folder_id = resolve_folder_id(mail_token, INBOX_SUBFOLDER) if INBOX_SUBFOLDER else None
    emails    = fetch_bs_emails(mail_token, folder_id)

    records: list[dict] = []
    skipped = 0

    for msg in emails:
        subject = msg.get("subject", "").strip()
        parsed  = parse_bs_subject(subject)
        if parsed is None:
            skipped += 1
            continue

        general_skill, subject_jr_no = parsed

        from_addr = _extract_address(msg.get("from", {}))
        to_list   = msg.get("toRecipients", [])
        to_addr   = _extract_address(to_list[0]) if to_list else ""

        recruiter        = _normalize_name_from_email(from_addr)
        client_recruiter = _normalize_name_from_email(to_addr)
        delivery_type    = (
            "Internal"
            if VOLIBITS_DOMAIN in from_addr and VOLIBITS_DOMAIN in to_addr
            else "External"
        )

        attachment_str = get_attachment_names(mail_token, msg)

        body_html = (msg.get("body") or {}).get("content", "")
        rows      = parse_html_table(body_html)
        if not rows:
            log.warning(f"No table in email: {subject!r} — adding skeleton row.")
            rows = [{}]

        email_date: Optional[date] = None
        raw_dt = msg.get("receivedDateTime", "")
        if raw_dt:
            try:
                email_date = datetime.fromisoformat(raw_dt.replace("Z", "+00:00")).date()
            except ValueError:
                pass

        for row in rows:
            row_date = _parse_date(row["date"]) if row.get("date") else None
            if row_date is None:
                row_date = email_date or date.today()

            effective_jr_no = _t(row.get("jr_no")) or subject_jr_no
            candidate_name  = _t(row.get("name_of_candidate"))
            contact_number  = _t(row.get("contact_number"))
            email_id_val    = _t(row.get("email_id"))

            record = {
                "email_subject":       subject,
                "recruiter":           recruiter,
                "client_recruiter":    client_recruiter,
                "email_from":          from_addr,
                "email_to":            to_addr,
                "delivery_type":       delivery_type,
                "company_name":        "Birlasoft",
                "general_skill":       _t(general_skill),
                "date":                row_date,
                "jr_no":               effective_jr_no,
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
                "attachment":          _t(attachment_str),
                "remarks":             _t(row.get("remarks")),
                "record_status":       _record_status({
                    "name_of_candidate": candidate_name,
                    "contact_number":    contact_number,
                    "email_id":          email_id_val,
                    "general_skill":     _t(general_skill),
                    "company_name":      "Birlasoft",
                    "recruiter":         recruiter,
                    "email_from":        from_addr,
                    "email_to":          to_addr,
                }),
            }
            records.append(record)
            log.info(f"  + {candidate_name or '(unknown)'} | {subject!r} | {row_date}")

    log.info(f"Skipped {skipped} non-BS email(s).")
    log.info(f"Total rows to write: {len(records)}")

    if not records:
        log.warning("No matching records found. No Excel file written.")
        return

    ts          = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = f"bs_candidates_{ts}.xlsx"
    write_excel(records, output_path)
    print(f"\n✅  Done — {len(records)} row(s) written to: {output_path}\n")


if __name__ == "__main__":
    run()