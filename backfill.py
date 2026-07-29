"""
Backfill missing columns for already-processed emails
══════════════════════════════════════════════════════
Finds records in temp_hrvolibit_archive that have blank fields,
re-fetches those emails from the mailbox, re-parses the table,
and fills in only the missing values — never overwrites existing data.

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

DB_DSN    = os.environ.get("DB_DSN") or os.environ.get("DATABASE_URL")
DB_TABLE  = "hrvolibit"

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
    log.info(f"Table {DB_TABLE} has {len(cols)} columns: {sorted(cols)}")
    return cols


def fetch_incomplete_records(conn, table_cols: set[str]) -> list[dict]:
    """Return all records that have at least one blank backfill field."""
    # Only filter/select fields that actually exist in this table
    active_fields = [f for f in BACKFILL_FIELDS if f in table_cols]
    if not active_fields:
        log.warning("None of the backfill fields exist in the table.")
        return []

    # Columns to SELECT: always include id + active backfill fields
    # Also include subject/email columns if they exist (used for email lookup)
    subject_col  = next((c for c in ("email_subject", "subject") if c in table_cols), None)
    email_from   = "email_from"   if "email_from"  in table_cols else None
    email_to     = "email_to"     if "email_to"    in table_cols else None
    date_col     = next((c for c in ("date", "created_at", "submission_date") if c in table_cols), None)

    select_cols = ["id"] + active_fields
    for c in [subject_col, email_from, email_to, date_col]:
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

    log.info(f"Found {len(rows)} incomplete record(s). Active backfill fields: {active_fields}")
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

def search_emails_by_sender_and_date(token: str, sender: str, record_date: date) -> list[dict]:
    """
    Fetch emails sent FROM `sender` received within ±1 day of `record_date`.
    Used when the table has no email_subject column.
    """
    from datetime import timedelta
    date_from = (record_date - timedelta(days=1)).isoformat() + "T00:00:00Z"
    date_to   = (record_date + timedelta(days=1)).isoformat() + "T23:59:59Z"
    sender_esc = sender.replace("'", "''")
    url = (
        f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/messages"
        f"?$filter=from/emailAddress/address eq '{sender_esc}'"
        f" and receivedDateTime ge {date_from}"
        f" and receivedDateTime le {date_to}"
        f"&$select=id,subject,from,toRecipients,body,receivedDateTime,hasAttachments,conversationId"
        f"&$top=20"
    )
    try:
        resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
        resp.raise_for_status()
        msgs = resp.json().get("value", [])
        log.info(f"  Found {len(msgs)} email(s) from {sender!r} around {record_date}")
        return msgs
    except Exception as exc:
        log.warning(f"Sender+date search failed ({sender!r}, {record_date}): {exc}")
        return []


# ── Match extracted row to DB record ─────────────────────────────────────────

def _norm(v) -> str:
    return (v or "").strip().lower()


def best_match(db_record: dict, extracted_rows: list[dict]) -> Optional[dict]:
    """
    Find the extracted row that best matches the DB record.
    Match priority:
      1. email exact match
      2. phone exact match
      3. name exact match (case-insensitive)
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
    Return a dict of {field: new_value} for fields that are currently blank
    in the DB record but have a value in the extracted row.
    Never overwrites an existing value.
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

    # Discover actual columns so we never assume a column exists
    table_cols = get_table_columns(conn)
    subject_col = next((c for c in ("email_subject", "subject") if c in table_cols), None)

    incomplete = fetch_incomplete_records(conn, table_cols)

    if not incomplete:
        print("Nothing to backfill.")
        conn.close()
        return

    mail_token = get_mail_token()
    log.info("Mail token obtained.")

    # Group by (email_from, date) — works even when there is no subject column
    GroupKey = tuple  # (sender, date_str)
    by_sender_date: dict[GroupKey, list[dict]] = {}
    skipped_no_key = 0
    for rec in incomplete:
        sender    = (rec.get("email_from") or "").strip()
        rec_date  = rec.get("date")
        if not sender or not rec_date:
            skipped_no_key += 1
            continue
        date_str = rec_date.isoformat() if hasattr(rec_date, "isoformat") else str(rec_date)[:10]
        by_sender_date.setdefault((sender, date_str), []).append(rec)

    if skipped_no_key:
        log.warning(f"Skipped {skipped_no_key} record(s) with no email_from or date.")

    preview_rows: list[dict] = []
    update_plan:  list[tuple[int, dict]] = []

    # Cache extracted rows per (sender, date) so each mailbox query runs once
    extracted_cache: dict[GroupKey, list[dict]] = {}

    for (sender, date_str), db_records in by_sender_date.items():
        log.info(f"Processing {len(db_records)} record(s) — sender={sender!r} date={date_str}")

        if (sender, date_str) not in extracted_cache:
            record_date = date.fromisoformat(date_str)
            emails = search_emails_by_sender_and_date(mail_token, sender, record_date)

            all_extracted: list[dict] = []
            seen_conv: set[str] = set()
            for msg in emails:
                conv_id = msg.get("conversationId", "")
                if conv_id and conv_id not in seen_conv:
                    seen_conv.add(conv_id)
                    thread_msgs = fetch_thread_messages(mail_token, conv_id)
                else:
                    thread_msgs = [msg]

                for thread_msg in thread_msgs:
                    body_html = (thread_msg.get("body") or {}).get("content", "")
                    rows, headers_ok = parse_html_table(body_html)
                    if rows and headers_ok:
                        all_extracted.extend(rows)
                    elif AI_PROVIDER != "none":
                        ai_rows = ai_extract(body_html)
                        all_extracted.extend(ai_rows)

            extracted_cache[(sender, date_str)] = all_extracted
            log.info(f"  Extracted {len(all_extracted)} row(s) total from these emails.")

        all_extracted = extracted_cache[(sender, date_str)]

        if not all_extracted:
            log.warning(f"  No rows extracted for sender={sender!r} date={date_str}")
            continue

        for db_rec in db_records:
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
                "subject": f"{sender} / {date_str}",
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
