"""
Preview and fix missing jr_no values in temp_hrvolibit_archive.

Matches records that have no jr_no against the `candidates` table using:
  - name (case-insensitive)
  - email OR phone
  - date

Usage:
    python fix_missing_jrno.py            # preview only
    python fix_missing_jrno.py --update   # preview then prompt before updating
"""

import os
import sys
import psycopg2
from dotenv import load_dotenv

load_dotenv()

# ── Config — adjust column names to match your `candidates` table ──────────────
SOURCE_TABLE  = "candidates"          # table that has the correct jr_no
TARGET_TABLE  = "temp_hrvolibit_archive"  # table with missing jr_no

# Column names in SOURCE_TABLE (candidates)
SRC_NAME  = "name_of_candidate"   # candidate full name
SRC_EMAIL = "email_id"            # email
SRC_PHONE = "contact_number"      # phone / mobile
SRC_DATE  = "date"                # submission date
SRC_JR    = "jr_no"              # jr number

# Column names in TARGET_TABLE (temp_hrvolibit_archive)
TGT_NAME  = "name_of_candidate"
TGT_EMAIL = "email_id"
TGT_PHONE = "contact_number"
TGT_DATE  = "date"
TGT_JR    = "jr_no"
TGT_ID    = "id"                  # primary key — change if different
# ──────────────────────────────────────────────────────────────────────────────

PREVIEW_SQL = f"""
SELECT
    t.{TGT_ID}                             AS target_id,
    t.{TGT_NAME}                           AS name,
    COALESCE(t.{TGT_EMAIL}, t.{TGT_PHONE}) AS contact,
    t.{TGT_DATE}::text                     AS date,
    t.{TGT_JR}                             AS current_jr_no,
    s.{SRC_JR}                             AS jr_no_from_candidates,
    CASE
        WHEN t.{TGT_JR} IS NULL OR TRIM(t.{TGT_JR}) = ''       THEN 'WILL UPDATE'
        WHEN TRIM(t.{TGT_JR}) = TRIM(s.{SRC_JR})               THEN 'MATCHES OK'
        ELSE                                                          'CONFLICT'
    END                                    AS status
FROM {TARGET_TABLE} t
JOIN {SOURCE_TABLE} s
  ON  s.{SRC_JR} IS NOT NULL
  AND LOWER(TRIM(t.{TGT_NAME})) = LOWER(TRIM(s.{SRC_NAME}))
  AND (
       LOWER(TRIM(t.{TGT_EMAIL})) = LOWER(TRIM(s.{SRC_EMAIL}))
    OR TRIM(t.{TGT_PHONE})        = TRIM(s.{SRC_PHONE})
  )
  AND t.{TGT_DATE} = s.{SRC_DATE}
ORDER BY status, t.{TGT_DATE} DESC, t.{TGT_NAME};
"""

UPDATE_SQL = f"""
UPDATE {TARGET_TABLE} t
SET    {TGT_JR} = s.{SRC_JR}
FROM   {SOURCE_TABLE} s
WHERE  (t.{TGT_JR} IS NULL OR TRIM(t.{TGT_JR}) = '')
  AND  s.{SRC_JR} IS NOT NULL
  AND  LOWER(TRIM(t.{TGT_NAME})) = LOWER(TRIM(s.{SRC_NAME}))
  AND  (
        LOWER(TRIM(t.{TGT_EMAIL})) = LOWER(TRIM(s.{SRC_EMAIL}))
     OR TRIM(t.{TGT_PHONE})        = TRIM(s.{SRC_PHONE})
  )
  AND  t.{TGT_DATE} = s.{SRC_DATE};
"""


def connect():
    dsn = os.environ.get("DB_DSN") or os.environ.get("DATABASE_URL")
    if not dsn:
        print("ERROR: DB_DSN not set in .env")
        sys.exit(1)
    return psycopg2.connect(dsn)


def preview(conn):
    with conn.cursor() as cur:
        cur.execute(PREVIEW_SQL)
        rows = cur.fetchall()
        cols = [d[0] for d in cur.description]

    if not rows:
        print("No missing jr_no records found that can be filled.")
        return 0

    # Print table
    col_widths = [max(len(c), max(len(str(r[i] or "")) for r in rows))
                  for i, c in enumerate(cols)]
    sep = "+-" + "-+-".join("-" * w for w in col_widths) + "-+"
    header = "| " + " | ".join(c.ljust(col_widths[i]) for i, c in enumerate(cols)) + " |"

    print(f"\nRecords that WOULD be updated ({len(rows)} row(s)):\n")
    print(sep)
    print(header)
    print(sep)
    for row in rows:
        line = "| " + " | ".join(str(v or "").ljust(col_widths[i]) for i, v in enumerate(row)) + " |"
        print(line)
    print(sep)
    print(f"\nTotal: {len(rows)} record(s) would be updated.\n")
    return len(rows)


def update(conn):
    with conn.cursor() as cur:
        cur.execute(UPDATE_SQL)
        count = cur.rowcount
    conn.commit()
    print(f"Done — {count} record(s) updated.")


def main():
    do_update = "--update" in sys.argv
    conn = connect()
    try:
        count = preview(conn)
        if count == 0:
            return
        if do_update:
            answer = input("Proceed with update? (y/n): ").strip().lower()
            if answer == "y":
                update(conn)
            else:
                print("Cancelled — no changes made.")
        else:
            print("Run with --update to apply these changes.")
    finally:
        conn.close()


if __name__ == "__main__":
    main()
