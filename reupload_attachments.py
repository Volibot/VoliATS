"""
reupload_attachments.py
-----------------------
Re-downloads resume attachments from every previously-processed email,
uploads them to the correct OneDrive folder, and updates the attachment
link in the database.

Run this once after fixing the OneDrive account / folder configuration.

Usage:
    python reupload_attachments.py            # live run
    python reupload_attachments.py --dry-run  # preview only, no DB writes
"""

import sys
import logging
import requests
import psycopg2
from psycopg2 import sql as pgsql
from typing import Optional

from email_extractor import (
    get_mail_token,
    get_onedrive_token,
    upload_all_attachments,
    match_resume_for_candidate,
    parse_html_table,
    _ensure_onedrive_folder,
    _t,
    TARGET_MAILBOX,
    ONEDRIVE_FOLDER,
    DB_DSN,
    DB_TABLE,
)

DRY_RUN = "--dry-run" in sys.argv

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("reupload.log"),
    ],
)
log = logging.getLogger(__name__)


# ─── DB helpers ────────────────────────────────────────────────────────────────

def fetch_all_processed_emails(cur) -> list[dict]:
    cur.execute("""
        SELECT message_id, from_addr, subject
        FROM hr_processed_emails
        ORDER BY processed_at ASC
    """)
    return [{"message_id": r[0], "from_addr": r[1], "subject": r[2]}
            for r in cur.fetchall()]


def find_candidate_record(
    cur,
    name: str,
    contact: str,
    email_id: str,
    from_addr: str,
) -> Optional[dict]:
    """Find the DB row for a candidate by matching identity fields + sender."""
    query = pgsql.SQL("""
        SELECT id, name_of_candidate, attachment FROM {table}
        WHERE LOWER(COALESCE(email_from, '')) = LOWER(%s)
          AND (
            (LOWER(COALESCE(name_of_candidate, '')) = LOWER(%s) AND %s <> '')
            OR (LOWER(COALESCE(contact_number,  '')) = LOWER(%s) AND %s <> '')
            OR (LOWER(COALESCE(email_id,        '')) = LOWER(%s) AND %s <> '')
          )
        LIMIT 1
    """).format(table=pgsql.Identifier(DB_TABLE))
    cur.execute(query, (
        from_addr,
        name,    name,
        contact, contact,
        email_id, email_id,
    ))
    row = cur.fetchone()
    return {"id": row[0], "name": row[1], "attachment": row[2]} if row else None


def find_single_candidate_for_email(cur, from_addr: str) -> Optional[dict]:
    """Fallback for emails where the body had no parseable table (skeleton records)."""
    query = pgsql.SQL("""
        SELECT id, name_of_candidate, attachment FROM {table}
        WHERE LOWER(COALESCE(email_from, '')) = LOWER(%s)
          AND (attachment IS NULL OR attachment = '')
        ORDER BY id DESC
        LIMIT 1
    """).format(table=pgsql.Identifier(DB_TABLE))
    cur.execute(query, (from_addr,))
    row = cur.fetchone()
    return {"id": row[0], "name": row[1], "attachment": row[2]} if row else None


def update_attachment_url(cur, record_id: int, url: str) -> None:
    if DRY_RUN:
        return
    cur.execute(
        pgsql.SQL("UPDATE {table} SET attachment = %s, modified_date = NOW() WHERE id = %s")
        .format(table=pgsql.Identifier(DB_TABLE)),
        (url, record_id),
    )


# ─── Graph helper ──────────────────────────────────────────────────────────────

def fetch_email_from_graph(mail_token: str, message_id: str) -> Optional[dict]:
    try:
        resp = requests.get(
            f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}"
            f"/messages/{message_id}"
            f"?$select=id,subject,from,toRecipients,body,hasAttachments,receivedDateTime",
            headers={"Authorization": f"Bearer {mail_token}"},
            timeout=30,
        )
        if resp.status_code == 404:
            return None
        resp.raise_for_status()
        return resp.json()
    except Exception as exc:
        log.error(f"Failed to fetch email {message_id}: {exc}")
        return None


# ─── Main ──────────────────────────────────────────────────────────────────────

def reupload_all() -> None:
    log.info("=== Attachment Re-upload %s===", "[DRY RUN] " if DRY_RUN else "")
    log.info(f"Target folder : {ONEDRIVE_FOLDER or '(root)'}")

    mail_token = get_mail_token()
    od_token   = get_onedrive_token()

    if ONEDRIVE_FOLDER and not DRY_RUN:
        _ensure_onedrive_folder(od_token, ONEDRIVE_FOLDER)

    conn = psycopg2.connect(DB_DSN)
    conn.autocommit = False
    cur  = conn.cursor()

    messages = fetch_all_processed_emails(cur)
    log.info(f"{len(messages)} processed email(s) found in tracking table.")

    total_uploaded = total_updated = total_skipped = total_missing = total_errors = 0

    for meta in messages:
        message_id = meta["message_id"]
        from_addr  = meta["from_addr"]
        subject    = meta["subject"]

        log.info(f"── {subject!r}")

        # ── Fetch email from Graph ────────────────────────────────────────────
        msg = fetch_email_from_graph(mail_token, message_id)
        if msg is None:
            log.warning(f"   Email no longer in mailbox — skipping")
            total_missing += 1
            continue

        if not msg.get("hasAttachments"):
            log.info(f"   No attachments — skipping")
            total_skipped += 1
            continue

        # ── Upload all resumes to new OneDrive location ───────────────────────
        attachment_map = upload_all_attachments(od_token, mail_token, msg)
        if not attachment_map:
            log.info(f"   No resume attachments — skipping")
            total_skipped += 1
            continue

        total_uploaded += len(attachment_map)
        log.info(f"   Uploaded {len(attachment_map)} file(s): {list(attachment_map.keys())}")

        # ── Parse candidate rows from email body ──────────────────────────────
        body_html = (msg.get("body") or {}).get("content", "")
        rows = parse_html_table(body_html)

        claimed: set[str] = set()
        matched_this_email = 0

        if rows:
            for row in rows:
                name    = _t(row.get("name_of_candidate")) or ""
                contact = _t(row.get("contact_number"))    or ""
                email   = _t(row.get("email_id"))          or ""

                if not name and not contact and not email:
                    continue

                new_url = match_resume_for_candidate(name, attachment_map, claimed)
                if not new_url:
                    log.debug(f"   No resume match for {name!r}")
                    continue

                record = find_candidate_record(cur, name, contact, email, from_addr)
                if record is None:
                    log.warning(f"   No DB record for {name!r} from {from_addr}")
                    total_errors += 1
                    continue

                if record["attachment"] == new_url:
                    log.info(f"   Already up-to-date: {name!r}")
                    total_skipped += 1
                    continue

                update_attachment_url(cur, record["id"], new_url)
                matched_this_email += 1
                total_updated += 1
                action = "[DRY RUN] would update" if DRY_RUN else "Updated"
                log.info(
                    f"   ✓ {action} record {record['id']} ({record['name']}) "
                    f"→ {new_url[:80]}{'...' if len(new_url) > 80 else ''}"
                )
        else:
            # No table rows — skeleton record; match by sender only
            new_url = next(iter(attachment_map.values()))
            record  = find_single_candidate_for_email(cur, from_addr)
            if record:
                update_attachment_url(cur, record["id"], new_url)
                matched_this_email += 1
                total_updated += 1
                action = "[DRY RUN] would update" if DRY_RUN else "Updated"
                log.info(
                    f"   ✓ {action} skeleton record {record['id']} ({record['name']}) "
                    f"→ {new_url[:80]}{'...' if len(new_url) > 80 else ''}"
                )

        if not DRY_RUN:
            conn.commit()

        if matched_this_email == 0:
            log.warning(f"   No DB records updated for this email")

    cur.close()
    conn.close()

    log.info(
        f"=== Done — emails_missing={total_missing} uploaded={total_uploaded} "
        f"updated={total_updated} skipped={total_skipped} errors={total_errors} ==="
    )


if __name__ == "__main__":
    reupload_all()
