"""
Backfill missing columns for already-processed emails
══════════════════════════════════════════════════════
Pulls every message_id from hr_processed_emails, fetches the email body
directly from the mailbox, re-parses the candidate table, and fills in
only the blank fields in hrvolibit — never overwrites existing data.

Usage:
    python backfill.py            # preview what would be updated
    python backfill.py --update   # apply updates after preview
"""

import os
import sys
import re
import logging
import requests
import psycopg2
from datetime import datetime, date
from typing import Optional
from dotenv import load_dotenv

# Reuse all parsing/auth helpers from extractor
from extractor import (
    get_mail_token,
    fetch_thread_messages,
    parse_html_table,
    ai_extract,
    _sanitize_contact,
    _sanitize_email,
    _fix_swapped_contact_email,
    _clean_cell,
    _t,
    TARGET_MAILBOX,
    DB_SCHEMA,
    AI_PROVIDER,
)

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(), logging.FileHandler("backfill.log")],
)
log = logging.getLogger(__name__)

DB_DSN   = os.environ.get("DB_DSN") or os.environ.get("DATABASE_URL")
DB_TABLE = "hrvolibit"
TRACKING_TABLE = "hr_processed_emails"

# Fields that can be backfilled from the email table
BACKFILL_FIELDS = [
    "jr_no",
    "name_of_candidate",
    "contact_number",
    "email_id",
    "total_experience",
    "relevant_experience",
    "current_ctc",
    "expected_ctc",
    "notice_period",
    "current_org",
    "current_location",
    "preferred_location",
    "general_skill",
    "remarks",
]


# ── DB helpers ────────────────────────────────────────────────────────────────

def db_connect():
    if not DB_DSN:
        log.error("DB_DSN not set in .env")
        sys.exit(1)
    return psycopg2.connect(DB_DSN)


def get_table_columns(conn) -> set[str]:
    """Return the set of column names that actually exist in DB_TABLE."""
    with conn.cursor() as cur:
        cur.execute(
            "SELECT column_name FROM information_schema.columns "
            "WHERE table_schema = %s AND table_name = %s",
            (DB_SCHEMA, DB_TABLE),
        )
        cols = {row[0] for row in cur.fetchall()}
    if not cols:
        log.error(f"Table {DB_SCHEMA}.{DB_TABLE} not found or has no columns.")
        sys.exit(1)
    log.info(f"Table {DB_TABLE} has {len(cols)} columns.")
    return cols


def fetch_processed_emails(conn) -> list[dict]:
    """Return all rows from hr_processed_emails (message_id + from_addr + subject)."""
    with conn.cursor() as cur:
        cur.execute(
            f"SELECT message_id, from_addr, subject, processed_at "
            f"FROM {TRACKING_TABLE} ORDER BY processed_at ASC"
        )
        cols = [d[0] for d in cur.description]
        rows = [dict(zip(cols, row)) for row in cur.fetchall()]
    log.info(f"Found {len(rows)} processed email(s) in {TRACKING_TABLE}.")
    return rows


def fetch_incomplete_records(conn, table_cols: set[str]) -> list[dict]:
    """Return all hrvolibit records that have at least one blank backfill field."""
    active_fields = [f for f in BACKFILL_FIELDS if f in table_cols]
    if not active_fields:
        log.warning("None of the backfill fields exist in the table.")
        return []

    email_from_col = "email_from" if "email_from" in table_cols else None
    date_col = next((c for c in ("date", "created_at", "submission_date") if c in table_cols), None)

    select_cols = ["id"] + active_fields
    for c in [email_from_col, date_col]:
        if c and c not in select_cols:
            select_cols.append(c)

    conditions = " OR ".join(
        f"({f} IS NULL OR TRIM({f}::text) = '')" for f in active_fields
    )
    order = f"ORDER BY {date_col} DESC" if date_col else "ORDER BY id DESC"
    sql = (
        f"SELECT {', '.join(select_cols)} "
        f"FROM {DB_SCHEMA}.{DB_TABLE} "
        f"WHERE {conditions} {order}"
    )
    with conn.cursor() as cur:
        cur.execute(sql)
        cols = [d[0] for d in cur.description]
        rows = [dict(zip(cols, row)) for row in cur.fetchall()]

    log.info(f"Found {len(rows)} incomplete record(s) in {DB_TABLE}.")
    return rows


def update_record(conn, record_id: int, updates: dict) -> None:
    if not updates:
        return
    set_clause = ", ".join(f"{k} = %s" for k in updates)
    values = list(updates.values()) + [record_id]
    sql = f"UPDATE {DB_SCHEMA}.{DB_TABLE} SET {set_clause} WHERE id = %s"
    with conn.cursor() as cur:
        cur.execute(sql, values)
    conn.commit()


# ── Email fetching ────────────────────────────────────────────────────────────

def fetch_message_by_id(token: str, message_id: str) -> Optional[dict]:
    """Fetch a single email directly by its Graph API message_id."""
    url = (
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}"
        f"/messages/{message_id}"
        f"?$select=id,subject,from,toRecipients,body,receivedDateTime,hasAttachments,conversationId"
    )
    try:
        resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
        if resp.status_code == 404:
            log.warning(f"  Message {message_id!r} not found (deleted/moved).")
            return None
        resp.raise_for_status()
        return resp.json()
    except Exception as exc:
        log.warning(f"  Failed to fetch message {message_id!r}: {exc}")
        return None


# ── Match extracted row to DB record ─────────────────────────────────────────

def _norm(v) -> str:
    return (v or "").strip().lower()


def best_match(db_record: dict, extracted_rows: list[dict]) -> Optional[dict]:
    """
    Find the extracted row that best matches the DB record.
    Priority: email match → phone match → name match (case-insensitive).
    """
    for row in extracted_rows:
        if _norm(row.get("email_id")) and _norm(row.get("email_id")) == _norm(db_record.get("email_id")):
            return row
    for row in extracted_rows:
        if _norm(row.get("contact_number")) and _norm(row.get("contact_number")) == _norm(db_record.get("contact_number")):
            return row
    for row in extracted_rows:
        if _norm(row.get("name_of_candidate")) and _norm(row.get("name_of_candidate")) == _norm(db_record.get("name_of_candidate")):
            return row
    return None


# ── Compute updates for one record ───────────────────────────────────────────

def compute_updates(db_record: dict, extracted: dict) -> dict:
    """
    Return {field: new_value} for fields currently blank in DB but present
    in the extracted row. Never overwrites an existing value.
    """
    updates = {}
    field_map = {
        "jr_no":               _t(extracted.get("jr_no")),
        "name_of_candidate":   _t(extracted.get("name_of_candidate")),
        "contact_number":      _sanitize_contact(_t(extracted.get("contact_number"))),
        "email_id":            _sanitize_email(_t(extracted.get("email_id"))),
        "total_experience":    _t(extracted.get("total_experience")),
        "relevant_experience": _t(extracted.get("relevant_experience")),
        "current_ctc":         _t(extracted.get("current_ctc")),
        "expected_ctc":        _t(extracted.get("expected_ctc")),
        "notice_period":       _t(extracted.get("notice_period")),
        "current_org":         _t(extracted.get("current_org")),
        "current_location":    _t(extracted.get("current_location")),
        "preferred_location":  _t(extracted.get("preferred_location")),
        "general_skill":       _t(extracted.get("general_skill")),
        "remarks":             _t(extracted.get("remarks")),
    }
    for field, new_val in field_map.items():
        current = (db_record.get(field) or "").strip() if db_record.get(field) else ""
        if not current and new_val:
            updates[field] = new_val
    return updates


# ── Print preview table ───────────────────────────────────────────────────────

def print_preview(preview_rows: list[dict]) -> None:
    if not preview_rows:
        print("\nNo updates found.\n")
        return

    print(f"\n{'='*80}")
    print(f"  BACKFILL PREVIEW — {len(preview_rows)} record(s) would be updated")
    print(f"{'='*80}\n")

    for p in preview_rows:
        print(f"  ID {p['id']} | {p['name']} | {p['subject'][:60]}")
        for field, new_val in p["updates"].items():
            print(f"    {field:<25} ← {new_val!r}")
        print()


# ── Main ──────────────────────────────────────────────────────────────────────

def run(do_update: bool = False) -> None:
    log.info("=== Backfill starting ===")

    conn = db_connect()

    table_cols   = get_table_columns(conn)
    incomplete   = fetch_incomplete_records(conn, table_cols)

    if not incomplete:
        print("Nothing to backfill.")
        conn.close()
        return

    # Build lookup: (email_from, date) → [db_records]
    # date is stored as a date object in hrvolibit; we index all dates ±1 day
    # so slight offsets between processed_at and the table's date column don't matter
    from datetime import timedelta
    by_sender_date: dict[tuple, list[dict]] = {}
    date_col = next((c for c in ("date", "created_at", "submission_date") if c in table_cols), None)
    for rec in incomplete:
        sender = (rec.get("email_from") or "").strip().lower()
        rec_date = rec.get(date_col) if date_col else None
        if not sender or not rec_date:
            continue
        d = rec_date if isinstance(rec_date, date) else date.fromisoformat(str(rec_date)[:10])
        for delta in (-1, 0, 1):
            key = (sender, d + timedelta(days=delta))
            by_sender_date.setdefault(key, []).append(rec)

    processed_emails = fetch_processed_emails(conn)

    if not processed_emails:
        print(f"No rows in {TRACKING_TABLE}.")
        conn.close()
        return

    mail_token = get_mail_token()
    log.info("Mail token obtained.")

    preview_rows: list[dict] = []
    update_plan:  list[tuple[int, dict]] = []

    # Cache extracted rows per message_id so each fetch runs once
    extracted_cache: dict[str, list[dict]] = {}

    for email_row in processed_emails:
        message_id = email_row["message_id"]
        from_addr  = (email_row.get("from_addr") or "").strip().lower()
        subject    = email_row.get("subject") or ""

        log.info(f"Processing message_id={message_id!r} from={from_addr!r}")

        if message_id not in extracted_cache:
            msg = fetch_message_by_id(mail_token, message_id)
            if not msg:
                extracted_cache[message_id] = []
                continue

            # Also fetch the full thread in case the table is in a reply
            conv_id = msg.get("conversationId", "")
            thread_msgs = fetch_thread_messages(mail_token, conv_id) if conv_id else [msg]

            all_extracted: list[dict] = []
            for thread_msg in thread_msgs:
                body_html = (thread_msg.get("body") or {}).get("content", "")
                rows, headers_ok = parse_html_table(body_html)
                if rows and headers_ok:
                    all_extracted.extend(rows)
                elif AI_PROVIDER != "none":
                    ai_rows = ai_extract(body_html)
                    all_extracted.extend(ai_rows)

            extracted_cache[message_id] = all_extracted
            log.info(f"  Extracted {len(all_extracted)} row(s) from this message/thread.")

        all_extracted = extracted_cache[message_id]
        if not all_extracted:
            log.warning(f"  No rows extracted for {message_id!r}")
            continue

        # Find incomplete hrvolibit records matching this sender + processed date
        proc_date = email_row["processed_at"].date() if hasattr(email_row.get("processed_at"), "date") else date.fromisoformat(str(email_row.get("processed_at", ""))[:10])
        candidates = by_sender_date.get((from_addr, proc_date), [])
        if not candidates:
            log.info(f"  No incomplete records for sender={from_addr!r} date={proc_date}, skipping.")
            continue
        # Deduplicate (same record may appear via ±1 day keys)
        seen_ids: set[int] = set()
        unique_candidates = []
        for c in candidates:
            if c["id"] not in seen_ids:
                seen_ids.add(c["id"])
                unique_candidates.append(c)
        candidates = unique_candidates

        for db_rec in candidates:
            matched = best_match(db_rec, all_extracted)
            if not matched:
                log.info(f"  No match for ID {db_rec['id']} ({db_rec.get('name_of_candidate')!r})")
                continue

            updates = compute_updates(db_rec, matched)
            if not updates:
                log.info(f"  ID {db_rec['id']} — nothing new to fill.")
                continue

            preview_rows.append({
                "id":      db_rec["id"],
                "name":    db_rec.get("name_of_candidate") or "(unknown)",
                "subject": subject or from_addr,
                "updates": updates,
            })
            update_plan.append((db_rec["id"], updates))

    print_preview(preview_rows)

    if not update_plan:
        conn.close()
        return

    if do_update:
        answer = input(f"Apply {len(update_plan)} update(s)? (y/n): ").strip().lower()
        if answer == "y":
            for record_id, updates in update_plan:
                update_record(conn, record_id, updates)
                log.info(f"  Updated ID {record_id}: {list(updates.keys())}")
            print(f"\nDone — {len(update_plan)} record(s) updated.")
        else:
            print("Cancelled — no changes made.")
    else:
        print("Run with --update to apply these changes.")

    conn.close()


if __name__ == "__main__":
    run(do_update="--update" in sys.argv)
