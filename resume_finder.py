"""
resume_finder.py
─────────────────
Scans ALL inbox folders in the target mailbox, finds resume attachments
(PDF/DOC/DOCX), and matches them to candidate profiles in the hrvolibit table.

For each profile (optionally only those missing an attachment) the script:
  1. Enumerates every mail folder recursively.
  2. Fetches emails and checks whether the sender matches the profile's
     email_from and whether an attachment filename contains the candidate name.
  3. Uploads the matched resume to OneDrive.
  4. Updates the attachment field in the database via DB_DSN.

Required env vars (same as email_extractor.py):
  AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET
  TARGET_MAILBOX
  OD_TENANT_ID, OD_CLIENT_ID, OD_REFRESH_TOKEN, ONEDRIVE_USER
  DB_DSN

Optional:
  ONEDRIVE_FOLDER   (default: "HR Resumes")
  DB_TABLE_NAME     (default: "hrvolibit")
  ONLY_MISSING      (default: "true"  — skip profiles that already have an attachment)
  DRY_RUN           (default: "false" — when "true" uploads are skipped and DB is not updated)
  LIMIT             (default: 0 = no limit on profiles processed)
  EMAIL_PAGE_SIZE   (default: 50 — Graph API page size per folder)
"""

import os
import re
import logging
from typing import Optional

import psycopg2
from psycopg2 import sql as pgsql
import requests
from dotenv import load_dotenv
from msal import ConfidentialClientApplication, PublicClientApplication

load_dotenv()

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("resume_finder.log"),
    ],
)
log = logging.getLogger(__name__)

# ── Config ────────────────────────────────────────────────────────────────────
AZURE_TENANT_ID     = os.environ["AZURE_TENANT_ID"]
AZURE_CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
AZURE_CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]

TARGET_MAILBOX   = os.environ["TARGET_MAILBOX"]

OD_TENANT_ID     = os.environ["OD_TENANT_ID"]
OD_CLIENT_ID     = os.environ["OD_CLIENT_ID"]
OD_CLIENT_SECRET = os.environ.get("OD_CLIENT_SECRET", "")
OD_REFRESH_TOKEN = os.environ["OD_REFRESH_TOKEN"]
ONEDRIVE_USER    = os.environ["ONEDRIVE_USER"]
ONEDRIVE_FOLDER  = os.environ.get("ONEDRIVE_FOLDER", "HR Resumes")

DB_DSN       = os.environ["DB_DSN"]
DB_TABLE     = os.environ.get("DB_TABLE_NAME", "hrvolibit")

ONLY_MISSING    = os.environ.get("ONLY_MISSING", "true").strip().lower() in ("1", "true", "yes")
DRY_RUN         = os.environ.get("DRY_RUN", "false").strip().lower() in ("1", "true", "yes")
LIMIT           = int(os.environ.get("LIMIT", "0"))
EMAIL_PAGE_SIZE = int(os.environ.get("EMAIL_PAGE_SIZE", "50"))

RESUME_EXTENSIONS = {".pdf", ".doc", ".docx"}


# ── Auth ───────────────────────────────────────────────────────────────────────

def get_mail_token() -> str:
    result = ConfidentialClientApplication(
        AZURE_CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{AZURE_TENANT_ID}",
        client_credential=AZURE_CLIENT_SECRET,
    ).acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Mail token failed: {result.get('error_description')}")
    return result["access_token"]


def get_onedrive_token() -> str:
    app = PublicClientApplication(
        OD_CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{OD_TENANT_ID}",
    )
    result = app.acquire_token_by_refresh_token(
        OD_REFRESH_TOKEN,
        scopes=["https://graph.microsoft.com/.default"],
    )
    if "access_token" not in result:
        raise RuntimeError(
            f"OneDrive token refresh failed: {result.get('error_description', str(result))}"
        )
    log.info("OneDrive delegated token obtained.")
    return result["access_token"]


# ── DB helpers ────────────────────────────────────────────────────────────────

def load_profiles(cur) -> list[dict]:
    """
    Load candidate profiles from hrvolibit.
    When ONLY_MISSING=true, only rows where attachment is NULL or empty are returned.
    """
    log.info(f"Loading profiles from {DB_TABLE} …")

    missing_clause = pgsql.SQL(
        " AND (attachment IS NULL OR TRIM(attachment) = '')"
    ) if ONLY_MISSING else pgsql.SQL("")

    limit_clause = pgsql.SQL(f" LIMIT {LIMIT}") if LIMIT else pgsql.SQL("")

    query = pgsql.SQL(
        "SELECT id, name_of_candidate, contact_number, email_id, email_from, attachment"
        " FROM {table} WHERE TRUE {missing} {limit}"
    ).format(
        table=pgsql.Identifier(DB_TABLE),
        missing=missing_clause,
        limit=limit_clause,
    )

    cur.execute(query)
    cols = ["id", "name_of_candidate", "contact_number", "email_id", "email_from", "attachment"]
    rows = [dict(zip(cols, row)) for row in cur.fetchall()]
    log.info(f"Total profiles to process: {len(rows)}")
    return rows


def update_profile_attachment(cur, profile_id: int, url: str) -> bool:
    if DRY_RUN:
        log.info(f"  [DRY RUN] would set attachment for id={profile_id} → {url[:80]}")
        return True
    try:
        cur.execute(
            pgsql.SQL(
                "UPDATE {table} SET attachment = %s, modified_date = NOW() WHERE id = %s"
            ).format(table=pgsql.Identifier(DB_TABLE)),
            (url, profile_id),
        )
        return True
    except Exception as exc:
        log.error(f"  DB update failed for id={profile_id}: {exc}")
        return False


# ── Mail folder enumeration ────────────────────────────────────────────────────

def list_all_folders(token: str) -> list[dict]:
    """
    Return every mail folder in the target mailbox as a flat list.
    Recurses through child folders up to 3 levels deep.
    """
    headers = {"Authorization": f"Bearer {token}"}
    folders: list[dict] = []

    def _recurse(parent_id: Optional[str], path_prefix: str, depth: int) -> None:
        if depth > 3:
            return
        if parent_id:
            url = (
                f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}"
                f"/mailFolders/{parent_id}/childFolders"
                f"?$select=id,displayName,totalItemCount&$top=50"
            )
        else:
            url = (
                f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}"
                f"/mailFolders?$select=id,displayName,totalItemCount&$top=50"
            )

        while url:
            try:
                resp = requests.get(url, headers=headers, timeout=30)
                resp.raise_for_status()
            except Exception as exc:
                log.warning(f"Folder list error at depth={depth}: {exc}")
                break

            data = resp.json()
            for folder in data.get("value", []):
                fpath = f"{path_prefix}/{folder['displayName']}" if path_prefix else folder["displayName"]
                folders.append({
                    "id":         folder["id"],
                    "displayName": folder["displayName"],
                    "path":       fpath,
                    "totalItems": folder.get("totalItemCount", 0),
                })
                _recurse(folder["id"], fpath, depth + 1)

            url = data.get("@odata.nextLink")

    _recurse(None, "", 0)
    log.info(f"Found {len(folders)} mail folder(s) total.")
    return folders


# ── Email fetching ─────────────────────────────────────────────────────────────

def fetch_emails_from_folder(token: str, folder_id: str, folder_path: str) -> list[dict]:
    """Fetch all emails with attachments from a single folder, paginated."""
    headers = {"Authorization": f"Bearer {token}"}
    url = (
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}"
        f"/mailFolders/{folder_id}/messages"
        f"?$top={EMAIL_PAGE_SIZE}"
        f"&$select=id,subject,from,receivedDateTime,hasAttachments"
        f"&$filter=hasAttachments eq true"
    )

    emails: list[dict] = []
    page = 0
    while url:
        page += 1
        try:
            resp = requests.get(url, headers=headers, timeout=60)
            resp.raise_for_status()
        except requests.exceptions.Timeout:
            log.warning(f"Timeout on page {page} of {folder_path!r}")
            break
        except Exception as exc:
            log.warning(f"Error fetching {folder_path!r} page {page}: {exc}")
            break

        data = resp.json()
        emails.extend(data.get("value", []))
        url = data.get("@odata.nextLink")

    return emails


# ── Attachment helpers ─────────────────────────────────────────────────────────

def _is_resume(filename: str) -> bool:
    return os.path.splitext(filename)[1].lower() in RESUME_EXTENSIONS


def list_attachments(token: str, message_id: str) -> list[dict]:
    url = (
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}"
        f"/messages/{message_id}/attachments?$select=id,name,contentType"
    )
    try:
        resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
        resp.raise_for_status()
        return resp.json().get("value", [])
    except Exception as exc:
        log.warning(f"list_attachments failed for {message_id}: {exc}")
        return []


def fetch_attachment_bytes(token: str, message_id: str, att_id: str) -> Optional[bytes]:
    url = (
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}"
        f"/messages/{message_id}/attachments/{att_id}/$value"
    )
    try:
        resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=120)
        resp.raise_for_status()
        return resp.content
    except Exception as exc:
        log.error(f"fetch_attachment_bytes failed: {exc}")
        return None


# ── OneDrive upload ────────────────────────────────────────────────────────────

def _ensure_folder(od_token: str) -> None:
    if not ONEDRIVE_FOLDER:
        return
    headers = {"Authorization": f"Bearer {od_token}", "Content-Type": "application/json"}
    check = requests.get(
        f"https://graph.microsoft.com/v1.0/me/drive/root:/{ONEDRIVE_FOLDER}",
        headers=headers, timeout=30,
    )
    if check.status_code == 200:
        return
    resp = requests.post(
        "https://graph.microsoft.com/v1.0/me/drive/root/children",
        headers=headers,
        json={"name": ONEDRIVE_FOLDER, "folder": {}, "@microsoft.graph.conflictBehavior": "rename"},
        timeout=30,
    )
    if resp.status_code in (200, 201):
        log.info(f"Created OneDrive folder '{ONEDRIVE_FOLDER}'.")


def upload_to_onedrive(od_token: str, filename: str, content: bytes) -> Optional[str]:
    if DRY_RUN:
        log.info(f"  [DRY RUN] would upload {filename!r} to OneDrive")
        return f"[DRY RUN] {filename}"

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
        log.error(f"OneDrive upload failed for {filename!r}: {resp.status_code} {resp.text[:200]}")
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
                return web_url
    except Exception as exc:
        log.warning(f"createLink failed for {filename!r}: {exc}")

    return item.get("webUrl", filename)


# ── Name-matching helpers ──────────────────────────────────────────────────────

def _name_tokens(name: str) -> set[str]:
    stem = os.path.splitext(name)[0]
    # Strip bracket content like [5y_0m], [10y_0m], [Senior], etc.
    stem = re.sub(r"\[.*?\]", " ", stem)
    # Split CamelCase: insert space before an uppercase letter that follows a lowercase
    stem = re.sub(r"([a-z])([A-Z])", r"\1 \2", stem)
    # Replace common separators with spaces
    stem = re.sub(r"[_\-\.]+", " ", stem)
    # Strip digits (years, version numbers, etc.)
    stem = re.sub(r"\d+", " ", stem)
    return {t.lower() for t in stem.split() if len(t) > 1}


def _token_match_score(candidate_name: str, filename: str) -> int:
    if not candidate_name or not filename:
        return 0
    return len(_name_tokens(candidate_name) & _name_tokens(filename))


def _substr_match_score(candidate_name: str, filename: str) -> int:
    stem = os.path.splitext(filename)[0]
    stem = re.sub(r"\[.*?\]", " ", stem)
    stem = re.sub(r"([a-z])([A-Z])", r"\1 \2", stem)
    stem = re.sub(r"[_\-\.]", " ", stem)
    stem = re.sub(r"\d+", "", stem).lower().replace(" ", "")
    return sum(1 for t in _name_tokens(candidate_name) if len(t) >= 4 and t in stem)


def best_resume_for_candidate(
    candidate_name: Optional[str],
    resume_files: list[dict],
    claimed: set[str],
) -> Optional[dict]:
    """Return the best-matching unclaimed attachment for the candidate, or None."""
    name = candidate_name or ""

    # Pass 1: token overlap
    best_att, best_score = None, 0
    for att in resume_files:
        fname = att["name"]
        if fname in claimed:
            continue
        score = _token_match_score(name, fname)
        if score > best_score:
            best_score, best_att = score, att

    if best_att and best_score >= 1:
        claimed.add(best_att["name"])
        return best_att

    # Pass 2: substring fallback
    best_att, best_score = None, 0
    for att in resume_files:
        fname = att["name"]
        if fname in claimed:
            continue
        score = _substr_match_score(name, fname)
        if score > best_score:
            best_score, best_att = score, att

    if best_att and best_score >= 1:
        claimed.add(best_att["name"])
        return best_att

    return None


# ── Profile index ──────────────────────────────────────────────────────────────

def build_profile_index(profiles: list[dict]) -> dict[str, list[dict]]:
    """Group profiles by normalised sender email for O(1) lookup per email."""
    index: dict[str, list[dict]] = {}
    for p in profiles:
        sender = (p.get("email_from") or "").strip().lower()
        if sender:
            index.setdefault(sender, []).append(p)
    log.info(f"Profile index: {len(index)} unique sender(s).")
    return index


def _extract_address(addr_obj: dict) -> str:
    try:
        return addr_obj["emailAddress"]["address"].strip().lower()
    except (KeyError, TypeError):
        return ""


# ── Main ───────────────────────────────────────────────────────────────────────

def run() -> None:
    log.info("=== resume_finder starting ===")
    log.info(f"Target mailbox : {TARGET_MAILBOX}")
    log.info(f"OneDrive user  : {ONEDRIVE_USER}")
    log.info(f"OneDrive folder: {ONEDRIVE_FOLDER or '(root)'}")
    log.info(f"DB table       : {DB_TABLE}")
    log.info(f"Only missing   : {ONLY_MISSING}")
    log.info(f"Dry run        : {DRY_RUN}")
    log.info(f"Profile limit  : {LIMIT or 'unlimited'}")

    # ── DB connection ─────────────────────────────────────────────────────────
    conn = psycopg2.connect(DB_DSN)
    conn.autocommit = False
    cur  = conn.cursor()

    # ── Tokens ────────────────────────────────────────────────────────────────
    mail_token = get_mail_token()
    od_token   = get_onedrive_token()
    log.info("Tokens obtained.")

    # ── Verify OneDrive access ────────────────────────────────────────────────
    check = requests.get(
        "https://graph.microsoft.com/v1.0/me/drive",
        headers={"Authorization": f"Bearer {od_token}"},
        timeout=30,
    )
    if check.status_code == 200:
        log.info("OneDrive access confirmed.")
        if not DRY_RUN:
            _ensure_folder(od_token)
    else:
        log.warning(f"OneDrive check returned {check.status_code} — uploads may fail.")

    # ── Load profiles ─────────────────────────────────────────────────────────
    profiles = load_profiles(cur)
    if not profiles:
        log.info("No profiles to process — exiting.")
        cur.close()
        conn.close()
        return

    profile_index = build_profile_index(profiles)

    # ── Enumerate all mail folders ────────────────────────────────────────────
    all_folders = list_all_folders(mail_token)
    if not all_folders:
        log.warning("No mail folders found — check permissions.")
        cur.close()
        conn.close()
        return

    # ── Scan folders ──────────────────────────────────────────────────────────
    stats = {
        "folders_scanned":  0,
        "emails_scanned":   0,
        "resumes_found":    0,
        "profiles_updated": 0,
        "errors":           0,
    }
    matched_profile_ids: set[int] = set()

    for folder in all_folders:
        folder_id   = folder["id"]
        folder_path = folder["path"]
        total_items = folder["totalItems"]

        if total_items == 0:
            log.debug(f"Skipping empty folder: {folder_path!r}")
            continue

        log.info(f"Scanning folder: {folder_path!r} ({total_items} item(s))")
        emails = fetch_emails_from_folder(mail_token, folder_id, folder_path)
        stats["folders_scanned"] += 1
        stats["emails_scanned"]  += len(emails)

        for msg in emails:
            from_addr = _extract_address(msg.get("from", {}))
            if not from_addr:
                continue

            candidate_profiles = [
                p for p in profile_index.get(from_addr, [])
                if p["id"] not in matched_profile_ids
            ]
            if not candidate_profiles:
                continue

            subject    = msg.get("subject", "").strip()
            message_id = msg["id"]
            log.info(
                f"  Email from {from_addr!r} | {subject!r} "
                f"| {len(candidate_profiles)} profile(s)"
            )

            attachments = list_attachments(mail_token, message_id)
            resume_atts = [a for a in attachments if _is_resume(a.get("name", ""))]
            if not resume_atts:
                continue

            stats["resumes_found"] += len(resume_atts)
            claimed: set[str] = set()

            for profile in candidate_profiles:
                if profile["id"] in matched_profile_ids:
                    continue

                candidate_name = profile.get("name_of_candidate") or ""
                matched_att = best_resume_for_candidate(candidate_name, resume_atts, claimed)

                if matched_att is None:
                    log.debug(f"  No resume match for {candidate_name!r}")
                    continue

                att_id   = matched_att.get("id")
                att_name = matched_att["name"]
                log.info(f"  Matched {candidate_name!r} → {att_name!r}")

                content = fetch_attachment_bytes(mail_token, message_id, att_id) if att_id else None
                if content is None:
                    log.error(f"  Could not download {att_name!r} — skipping")
                    stats["errors"] += 1
                    continue

                od_url = upload_to_onedrive(od_token, att_name, content)
                if not od_url:
                    log.error(f"  OneDrive upload failed for {att_name!r} — skipping")
                    stats["errors"] += 1
                    continue

                ok = update_profile_attachment(cur, profile["id"], od_url)
                if ok:
                    if not DRY_RUN:
                        conn.commit()
                    matched_profile_ids.add(profile["id"])
                    stats["profiles_updated"] += 1
                    log.info(f"  ✓ Updated id={profile['id']} ({candidate_name!r}) → {od_url[:80]}")
                else:
                    conn.rollback()
                    stats["errors"] += 1

        if len(matched_profile_ids) >= len(profiles):
            log.info("All profiles matched — stopping early.")
            break

    cur.close()
    conn.close()

    unmatched = len(profiles) - len(matched_profile_ids)
    log.info("=== resume_finder complete ===")
    log.info(f"  Folders scanned  : {stats['folders_scanned']}")
    log.info(f"  Emails scanned   : {stats['emails_scanned']}")
    log.info(f"  Resumes found    : {stats['resumes_found']}")
    log.info(f"  Profiles updated : {stats['profiles_updated']}")
    log.info(f"  Unmatched        : {unmatched}")
    log.info(f"  Errors           : {stats['errors']}")

    print(
        f"\n{'[DRY RUN] ' if DRY_RUN else ''}"
        f"Done — {stats['profiles_updated']} profile(s) updated, "
        f"{unmatched} unmatched, {stats['errors']} error(s).\n"
        f"See resume_finder.log for details."
    )


if __name__ == "__main__":
    run()
