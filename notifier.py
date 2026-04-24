"""
notifier.py — HR Bot Notification Emails (Volibits)

Sends a rich HTML email to the recruiter who submitted the email after each
processed email. The notification shows:

  • Every field extracted for each candidate row
  • Green  ✓  for fields that were successfully extracted
  • Red    ✗  for fields that are missing / failed
  • "Updated" section for rows where empty fields were filled in existing record
  • A top-level summary banner (All Passed / Partial / All Failed)

Also sends an informational FYI email when the same candidate is found under
a different recruiter, notifying both recruiters of the duplicate submission.

CC recipients (team lead + manager) are loaded from the NOTIFY_CC_EMAILS secret.
Mail is sent via the mail-reading app (AZURE_* creds) which has Mail.Send.
"""

import os
import re
import requests
import logging
from datetime import datetime
from typing import Optional

log = logging.getLogger(__name__)

TARGET_MAILBOX   = os.environ["TARGET_MAILBOX"]
NOTIFY_CC_EMAILS = os.environ.get("NOTIFY_CC_EMAILS", "")
# Manager emails CC'd specifically on recruiter-level duplicate notifications
DIFF_RECRUITER_MANAGER_EMAILS = os.environ.get("DIFF_RECRUITER_MANAGER_EMAILS", "")


# ─── Human-readable field labels ────────────────────────────────────────────────
FIELD_LABELS: dict[str, str] = {
    "name_of_candidate":   "Candidate Name",
    "contact_number":      "Contact Number",
    "email_id":            "Email ID",
    "general_skill":       "General Skill",
    "company_name":        "Company Code",
    "recruiter":           "Recruiter",
    "email_from":          "Email From",
    "email_to":            "Email To",
    "jr_no":               "JR No.",
    "client_recruiter":    "Client Recruiter",
    "date":                "Date",
    "total_experience":    "Total Experience",
    "relevant_experience": "Relevant Experience",
    "current_ctc":         "Current CTC",
    "expected_ctc":        "Expected CTC",
    "notice_period":       "Notice Period",
    "current_org":         "Current Organisation",
    "current_location":    "Current Location",
    "preferred_location":  "Preferred Location",
    "delivery_type":       "Delivery Type",
    "is_duplicate":        "Duplicate Flag",
    "attachment":          "Attachment (OneDrive link)",
    "remarks":             "Remarks",
    "record_status":       "Record Status",
}

REQUIRED_FIELDS = {
    "name_of_candidate", "contact_number", "email_id",
    "general_skill", "company_name",
    "recruiter", "email_from", "email_to",
}

FIELD_FAILURE_REASONS: dict[str, str] = {
    "name_of_candidate": (
        "Candidate name was not found in the table. "
        "Ensure the column is labelled 'Candidate Name', 'Name', or similar."
    ),
    "contact_number": (
        "Contact number is required for candidate identification and duplicate checks. "
        "Add a 'Mobile', 'Phone', or 'Contact No.' column to the table."
    ),
    "email_id": (
        "Email ID is required for candidate identification and duplicate checks. "
        "Add an 'Email', 'Email ID', or 'Mail ID' column to the table."
    ),
    "general_skill": (
        "Skill/role was not extracted from the email subject. "
        "Ensure the subject follows the format:  CODE: Skill  (e.g. BS: Java Developer)."
    ),
    "company_name": (
        "Company code was not found in the subject line. "
        "Ensure the subject starts with a known code followed by a colon (e.g. BS:, RS:, HCL:)."
    ),
    "recruiter": (
        "Recruiter name could not be derived from the sender email address. "
        "The email must be sent from a valid Volibits address."
    ),
    "email_from": "Sender email address was missing from the message headers.",
    "email_to": (
        "Recipient email address was missing from the message headers. "
        "The email must have at least one To: recipient."
    ),
}

DISPLAY_ORDER = [
    "name_of_candidate", "contact_number", "email_id",
    "general_skill", "company_name", "jr_no", "client_recruiter",
    "recruiter", "date", "current_org",
    "total_experience", "relevant_experience",
    "current_ctc", "expected_ctc", "notice_period",
    "current_location", "preferred_location",
    "delivery_type", "is_duplicate", "attachment", "remarks",
    "record_status",
]

# Fields shown in conflict comparison table
CONFLICT_COMPARE_FIELDS = [
    "name_of_candidate", "contact_number", "email_id",
    "jr_no", "general_skill", "company_name",
    "total_experience", "relevant_experience",
    "current_ctc", "expected_ctc", "notice_period",
    "current_org", "current_location", "preferred_location",
    "recruiter", "client_recruiter", "attachment", "remarks",
]


# ─── CSS ────────────────────────────────────────────────────────────────────────
_CSS = """
  body { font-family:'Segoe UI',Arial,sans-serif; font-size:14px;
         color:#1a1a2e; background:#f4f6f9; margin:0; padding:0; }
  .wrap  { max-width:800px; margin:24px auto; background:#ffffff;
           border-radius:8px; overflow:hidden;
           box-shadow:0 2px 12px rgba(0,0,0,.10); }
  .hdr   { padding:24px 32px; }
  .hdr.pass    { background:#1a7a4a; }
  .hdr.fail    { background:#b71c1c; }
  .hdr.partial { background:#e65100; }
  .hdr.conflict{ background:#6a1b9a; }
  .hdr h1 { margin:0; font-size:20px; color:#ffffff; font-family:'Segoe UI',Arial,sans-serif; }
  .body  { padding:24px 32px; background:#ffffff; color:#1a1a2e; }
  .summary-bar { display:flex; gap:12px; margin-bottom:24px; flex-wrap:wrap; }
  .stat  { flex:1; min-width:110px; padding:14px 18px; border-radius:6px;
           text-align:center; }
  .stat.ins    { background:#e8f5e9; border:1px solid #a5d6a7; }
  .stat.skip   { background:#fff8e1; border:1px solid #ffe082; }
  .stat.err    { background:#ffebee; border:1px solid #ef9a9a; }
  .stat.upd    { background:#e3f2fd; border:1px solid #90caf9; }
  .stat.conf   { background:#f3e5f5; border:1px solid #ce93d8; }
  .stat .num { font-size:26px; font-weight:700; display:block; }
  .stat .lbl { font-size:12px; color:#555555; }
  .stat.ins  .num { color:#2e7d32; }
  .stat.skip .num { color:#f57f17; }
  .stat.err  .num { color:#c62828; }
  .stat.upd  .num { color:#1565c0; }
  .stat.conf .num { color:#6a1b9a; }
  .card  { border:1px solid #e0e0e0; border-radius:6px; margin-bottom:20px;
           overflow:hidden; }
  .card-hdr { display:flex; align-items:center; gap:10px;
              padding:12px 18px; font-weight:600; font-size:15px; }
  .card-hdr.pass    { background:#e8f5e9; color:#1b5e20; border-bottom:1px solid #c8e6c9; }
  .card-hdr.fail    { background:#ffebee; color:#7f0000; border-bottom:1px solid #ffcdd2; }
  .card-hdr.partial { background:#fff3e0; color:#bf360c; border-bottom:1px solid #ffe0b2; }
  .card-hdr.skip    { background:#f3f4f6; color:#555555; border-bottom:1px solid #e0e0e0; }
  .card-hdr.upd     { background:#e3f2fd; color:#0d47a1; border-bottom:1px solid #bbdefb; }
  .card-hdr.nochange{ background:#f5f5f5; color:#757575; border-bottom:1px solid #e0e0e0; }
  .card-hdr.conflict{ background:#f3e5f5; color:#4a148c; border-bottom:1px solid #e1bee7; }
  .badge { padding:2px 10px; border-radius:12px; font-size:12px; font-weight:700;
           margin-left:auto; }
  .badge.pass     { background:#2e7d32; color:#ffffff; }
  .badge.fail     { background:#c62828; color:#ffffff; }
  .badge.partial  { background:#e65100; color:#ffffff; }
  .badge.skip     { background:#78909c; color:#ffffff; }
  .badge.upd      { background:#1565c0; color:#ffffff; }
  .badge.nochange { background:#9e9e9e; color:#ffffff; }
  .badge.conflict { background:#6a1b9a; color:#ffffff; }
  .fields { padding:14px 18px; background:#ffffff; color:#1a1a2e; }
  table.ft { width:100%; border-collapse:collapse; }
  table.ft td { padding:6px 10px; vertical-align:top;
                border-bottom:1px solid #f0f0f0; font-size:13px;
                color:#1a1a2e; font-family:'Segoe UI',Arial,sans-serif; }
  table.ft td:first-child { width:38%; color:#555555; font-weight:500; }
  .val-ok   { color:#1b5e20; }
  .val-miss { color:#c62828; font-style:italic; }
  .val-upd  { color:#1565c0; font-weight:600; }
  .reason   { font-size:12px; color:#7f0000; margin-top:3px;
              padding:4px 8px; background:#fff8f8;
              border-left:3px solid #ef9a9a; border-radius:0 4px 4px 0; }
  .icon-ok   { color:#2e7d32; font-size:16px; }
  .icon-fail { color:#c62828; font-size:16px; }
  .icon-upd  { color:#1565c0; font-size:16px; }
  /* Conflict comparison table */
  table.cmp { width:100%; border-collapse:collapse; font-size:13px; }
  table.cmp th { padding:8px 12px; background:#f3e5f5; color:#4a148c;
                 text-align:left; font-weight:600; border-bottom:2px solid #ce93d8; }
  table.cmp td { padding:7px 12px; border-bottom:1px solid #f0f0f0;
                 vertical-align:top; color:#1a1a2e; }
  table.cmp tr:nth-child(even) td { background:#fafafa; }
  .cmp-empty { color:#bbbbbb; font-style:italic; }
  .cmp-diff  { background:#fff9c4 !important; }
  /* Action buttons */
  .actions { display:flex; gap:12px; margin:20px 0 8px; flex-wrap:wrap; }
  .btn { display:inline-block; padding:11px 22px; border-radius:6px;
         font-weight:700; font-size:14px; text-decoration:none;
         text-align:center; min-width:140px; }
  .btn-update { background:#1565c0; color:#fff; }
  .btn-new    { background:#2e7d32; color:#fff; }
  .btn-skip   { background:#78909c; color:#fff; }
  .urgent-banner { background:#4a148c; color:#fff; padding:14px 20px;
                   border-radius:6px; margin-bottom:20px; font-size:14px; }
  .footer { padding:16px 32px; background:#f8f9fa; font-size:11px;
            color:#9e9e9e; border-top:1px solid #e0e0e0; }
  /* Duplicate-specific card headers */
  .card-hdr.dup-recruiter { background:#ffebee; color:#7f0000; border-bottom:1px solid #ffcdd2; }
  .card-hdr.dup-contact   { background:#fff3e0; color:#bf360c; border-bottom:1px solid #ffe0b2; }
  .badge.dup-recruiter    { background:#c62828; color:#fff; }
  .badge.dup-contact      { background:#e65100; color:#fff; }
  /* Red inline dup banner */
  .dup-banner-red    { background:#ffebee; border-left:3px solid #ef5350; }
  .dup-banner-orange { background:#fff3e0; border-left:3px solid #ffa726; }
"""


def _val(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


# ─── Candidate card (standard processing report) ─────────────────────────────────
def _candidate_card_html(row_summary: dict) -> str:
    rd              = row_summary.get("record_data", {})
    missing         = set(row_summary.get("missing", []))
    outcome         = row_summary.get("outcome", "inserted")
    updated_fields  = set(row_summary.get("updated_fields", []))
    name            = _val(rd.get("name_of_candidate")) or row_summary.get("name", "(unknown)")
    dup             = _val(row_summary.get("dup_flag"))
    db_err          = row_summary.get("db_error")

    # ── Card style per outcome ───────────────────────────────────────────────
    if outcome == "skipped":
        card_cls, badge_cls, badge_txt = "skip", "skip", "Skipped — Duplicate"
    elif outcome == "no_change":
        card_cls, badge_cls, badge_txt = "nochange", "nochange", "No Change"
    elif outcome == "updated":
        card_cls, badge_cls, badge_txt = "upd", "upd", f"Updated — {len(updated_fields)} field(s) filled"
    elif outcome == "conflict":
        card_cls, badge_cls, badge_txt = "conflict", "conflict", "Conflict — Pending Review"
    elif outcome == "error":
        card_cls, badge_cls, badge_txt = "fail", "fail", "Error"
    elif dup in ("Duplicate Recruiter",) and outcome == "inserted":
        card_cls, badge_cls, badge_txt = "dup-recruiter", "dup-recruiter", "Inserted — Duplicate Recruiter"
    elif dup in ("Duplicate Contact",) and outcome == "inserted":
        card_cls, badge_cls, badge_txt = "dup-contact", "dup-contact", "Inserted — Duplicate Contact"
    elif missing:
        card_cls, badge_cls, badge_txt = "partial", "partial", "Partial — Pass with warnings"
    else:
        card_cls, badge_cls, badge_txt = "pass", "pass", "Pass"

    html = (
        f'<div class="card">'
        f'<div class="card-hdr {card_cls}">'
        f'<span>👤 {name}</span>'
        f'<span class="badge {badge_cls}">{badge_txt}</span>'
        f'</div>'
        f'<div class="fields">'
    )

    # ── Outcome-specific body ────────────────────────────────────────────────
    if outcome == "skipped":
        html += (
            '<p style="margin:0;color:#555">'
            'This candidate was <strong>not inserted</strong> because an identical record '
            '(same details, same date, same recruiter) already exists in the database '
            'and no new information was available to fill. No action needed.</p>'
        )
        html += '</div></div>'
        return html

    if outcome == "no_change":
        html += (
            '<p style="margin:0;color:#757575">'
            'An existing record was found for this candidate (same recruiter, same date). '
            'All updatable fields already have values — nothing was changed.</p>'
        )
        html += '</div></div>'
        return html

    if outcome == "updated":
        html += (
            '<p style="margin:0 0 10px;color:#0d47a1">'
            '✏️ An existing record was found and the following previously empty fields '
            f'have been filled: <strong>{", ".join(sorted(updated_fields))}</strong></p>'
        )
        html += '</div></div>'
        return html

    if outcome == "conflict":
        conflict_id       = row_summary.get("conflict_id", "")
        existing_recruiter = row_summary.get("existing_recruiter", "unknown")
        html += (
            f'<p style="margin:0;color:#4a148c">'
            f'⚠️ This candidate already exists under recruiter <strong>{existing_recruiter}</strong>. '
            f'A separate high-priority notification has been sent to both recruiters with '
            f'options to update the existing record, add as new, or skip.<br>'
            f'<small style="color:#888">Conflict ID: {conflict_id}</small></p>'
        )
        html += '</div></div>'
        return html

    if db_err:
        html += (
            f'<p style="color:#c62828"><strong>Database insert failed.</strong></p>'
            f'<p style="font-size:12px;color:#555">Error: {db_err}</p>'
            f'<p style="font-size:12px;color:#555">Please contact your system administrator.</p>'
        )
        html += '</div></div>'
        return html

    # ── Full field table (inserted / partial / error) ────────────────────────
    html += '<table class="ft">'
    for field in DISPLAY_ORDER:
        label       = FIELD_LABELS.get(field, field.replace("_", " ").title())
        raw         = rd.get(field)
        value       = _val(raw)
        is_required = field in REQUIRED_FIELDS
        is_missing  = field in missing

        if is_missing:
            reason = FIELD_FAILURE_REASONS.get(field, "This field was empty or could not be extracted.")
            required_sup = '&nbsp;<sup style="color:#c62828">required</sup>' if is_required else ""
            html += (
                f'<tr>'
                f'<td><span class="icon-fail">✗</span> {label}{required_sup}</td>'
                f'<td><span class="val-miss">Not extracted</span>'
                f'<div class="reason">⚠ {reason}</div></td>'
                f'</tr>'
            )
        else:
            display_value = value if value else '<span style="color:#aaa">—</span>'
            html += (
                f'<tr>'
                f'<td><span class="icon-ok">✓</span> {label}</td>'
                f'<td><span class="val-ok">{display_value}</span></td>'
                f'</tr>'
            )

    if dup:
        dup_meta = {
            "Duplicate":          ("🔴", "#ffebee", "#ef5350", "#7f0000",
                                   "Same phone AND email already exist in the DB. Record inserted and flagged."),
            "Duplicate Cell":     ("🟠", "#fff3e0", "#ffa726", "#5d4037",
                                   "Same phone number already exists (different email). Record inserted and flagged."),
            "Duplicate Email":    ("🟠", "#fff3e0", "#ffa726", "#5d4037",
                                   "Same email address already exists (different phone). Record inserted and flagged."),
            "Duplicate Recruiter":("🔴", "#ffebee", "#ef5350", "#7f0000",
                                   "Same candidate (phone + email + JR/skill) was already submitted by a different "
                                   "recruiter within the duplicate check window. Record inserted and flagged. "
                                   "Both recruiters and manager have been notified."),
            "Duplicate Contact":  ("🟠", "#fff3e0", "#ffa726", "#5d4037",
                                   "Same phone or email exists in the DB under a different JR/skill or recruiter "
                                   "within the duplicate check window. Record inserted and flagged."),
        }
        icon, bg, border, txt_color, note = dup_meta.get(
            dup, ("⚠️", "#fff8e1", "#ffc107", "#5d4037", dup)
        )
        html += (
            f'<tr><td colspan="2" style="padding:8px 10px;background:{bg};'
            f'border-left:3px solid {border};font-size:12px;color:{txt_color}">'
            f'{icon} <strong>Duplicate Flag: {dup}</strong> — {note}'
            f'</td></tr>'
        )

    html += '</table></div></div>'
    return html


# ─── Conflict comparison table ───────────────────────────────────────────────────
def _comparison_table_html(existing_row: dict, new_record_data: dict) -> str:
    existing_recruiter = existing_row.get("recruiter", "—")
    new_recruiter      = new_record_data.get("recruiter", "—")

    html = (
        f'<table class="cmp">'
        f'<thead><tr>'
        f'<th style="width:28%">Field</th>'
        f'<th style="width:36%">Existing Record<br><small>(by {existing_recruiter})</small></th>'
        f'<th style="width:36%">New Submission<br><small>(by {new_recruiter})</small></th>'
        f'</tr></thead><tbody>'
    )

    for field in CONFLICT_COMPARE_FIELDS:
        label    = FIELD_LABELS.get(field, field.replace("_", " ").title())
        ex_val   = _val(existing_row.get(field))
        new_val  = _val(new_record_data.get(field))
        ex_disp  = ex_val  if ex_val  else '<span class="cmp-empty">empty</span>'
        new_disp = new_val if new_val else '<span class="cmp-empty">empty</span>'

        # Highlight row if values differ (ignoring case/whitespace)
        differ = ex_val.strip().lower() != new_val.strip().lower()
        row_cls = ' class="cmp-diff"' if differ and (ex_val or new_val) else ""

        html += (
            f'<tr{row_cls}>'
            f'<td><strong>{label}</strong></td>'
            f'<td>{ex_disp}</td>'
            f'<td>{new_disp}</td>'
            f'</tr>'
        )

    html += '</tbody></table>'
    return html


# ─── Standard processing report ──────────────────────────────────────────────────
def build_email_html(
    original_subject: str,
    from_addr: str,
    rows_summary: list[dict],
    email_inserted: int,
    email_skipped: int,
    email_errors: int,
    email_updated: int = 0,
    email_conflicts: int = 0,
) -> tuple[str, str]:
    now_str = datetime.now().strftime("%d %b %Y, %I:%M %p")
    total   = email_inserted + email_skipped + email_errors + email_updated + email_conflicts

    if email_errors == 0 and email_skipped == 0 and email_conflicts == 0 and email_inserted > 0:
        overall, hdr_cls, hdr_icon = "All Passed", "pass", "✅"
    elif email_inserted == 0 and email_updated == 0 and email_skipped == 0 and email_conflicts == 0:
        overall, hdr_cls, hdr_icon = "All Failed", "fail", "❌"
    else:
        overall, hdr_cls, hdr_icon = "Partially Processed", "partial", "⚠️"

    subject_line = f"[HR Bot] {hdr_icon} {overall} — {original_subject}"

    cards_html = "".join(_candidate_card_html(r) for r in rows_summary)
    if not cards_html:
        cards_html = (
            '<div class="card"><div class="card-hdr fail">No candidate rows found</div>'
            '<div class="fields"><p style="color:#c62828">No data table was found in the '
            'email body.</p></div></div>'
        )

    # Build summary bar — only show non-zero stats to keep it clean
    stat_blocks = [
        ("ins",  email_inserted,  "Inserted"),
        ("upd",  email_updated,   "Updated"),
        ("skip", email_skipped,   "Skipped"),
        ("conf", email_conflicts, "Conflicts"),
        ("err",  email_errors,    "Errors"),
    ]
    stats_html = "".join(
        f'<div class="stat {cls}">'
        f'<span class="num">{count}</span>'
        f'<span class="lbl">{lbl}</span>'
        f'</div>'
        for cls, count, lbl in stat_blocks
    )
    # Always show Total
    stats_html += (
        f'<div class="stat" style="background:#e8eaf6;border:1px solid #9fa8da">'
        f'<span class="num" style="color:#283593">{total}</span>'
        f'<span class="lbl">Total</span>'
        f'</div>'
    )

    html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>{_CSS}</style></head>
<body><div class="wrap">
  <div class="hdr {hdr_cls}">
    <h1 style="margin:0;font-size:20px;color:#ffffff;font-family:'Segoe UI',Arial,sans-serif">{hdr_icon} HR Bot Processing Report</h1>
    <p style="margin:6px 0 0;color:#ffffff;font-size:13px;font-family:'Segoe UI',Arial,sans-serif">
      Submitted by: <strong>{from_addr}</strong> &nbsp;|&nbsp; Processed at: {now_str}
    </p>
  </div>
  <div class="body" style="padding:24px 32px;background:#ffffff;color:#1a1a2e;font-family:'Segoe UI',Arial,sans-serif">
    <div class="summary-bar">{stats_html}</div>
    {cards_html}
  </div>
  <div class="footer" style="padding:16px 32px;background:#f8f9fa;font-size:11px;color:#9e9e9e;border-top:1px solid #e0e0e0">
    This is an automated message from {TARGET_MAILBOX} — please do not reply directly.<br>
    For issues, contact your system administrator or check extractor.log.
  </div>
</div></body></html>"""

    return subject_line, html

ENABLE_NOTIFICATIONS = os.environ.get("ENABLE_NOTIFICATIONS", "true").strip().lower() == "true"

def send_notification_email(
    token: str,
    from_addr: str,
    original_subject: str,
    rows_summary: list[dict],
    email_inserted: int,
    email_skipped: int,
    email_errors: int,
    email_updated: int = 0,
    email_conflicts: int = 0,
) -> None:
    subject_line, html_body = build_email_html(
        original_subject, from_addr, rows_summary,
        email_inserted, email_skipped, email_errors,
        email_updated, email_conflicts,
    )
    if not ENABLE_NOTIFICATIONS:
        log.info("Notifications disabled (ENABLE_NOTIFICATIONS=false) — skipping.")
        return

    cc_list = []
    for addr in NOTIFY_CC_EMAILS.split(","):
        addr = addr.strip()
        if addr:
            cc_list.append({"emailAddress": {"address": addr}})

    payload = {
        "message": {
            "subject":      subject_line,
            "body":         {"contentType": "HTML", "content": html_body},
            "toRecipients": [{"emailAddress": {"address": from_addr}}],
            "ccRecipients": cc_list,
        },
        "saveToSentItems": "false",
    }

    _send_graph_mail(token, payload, from_addr)

def send_diff_recruiter_notification_email(
    token: str,
    new_recruiter_addr: str,
    existing_recruiter_addr: str,
    existing_row: dict,
    new_record_data: dict,
    original_subject: str,
    jr_no: str,
    inserted_ok: bool,
) -> None:
    """
    Send an informational FYI email when the same candidate has been submitted
    by a different recruiter.  The new record is already inserted — no action
    is needed from either recruiter.  Both recruiters are notified.
    """
    if not ENABLE_NOTIFICATIONS:
        log.info("Notifications disabled (ENABLE_NOTIFICATIONS=false) — skipping.")
        return
        candidate_name = (
        _val(new_record_data.get("name_of_candidate"))
        or _val(existing_row.get("name_of_candidate"))
        or "Unknown Candidate"
    )
    existing_recruiter = existing_row.get("recruiter", "—")
    new_recruiter      = new_record_data.get("recruiter", "—")
    insert_date        = _val(new_record_data.get("date")) or datetime.now().strftime("%d %b %Y")
    now_str            = datetime.now().strftime("%d %b %Y, %I:%M %p")

    status_line = (
        "The record has been added to the database."
        if inserted_ok
        else "The record could not be inserted — please contact your administrator."
    )

    subject_line = (
        f"[HR Bot] Candidate added by multiple recruiters — "
        f"{candidate_name} | {original_subject}"
    )

    comparison_html = _comparison_table_html(existing_row, new_record_data)

    html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>{_CSS}</style></head>
<body><div class="wrap">

  <div class="hdr" style="background:#1565c0">
    <h1 style="margin:0;font-size:20px;color:#ffffff;font-family:'Segoe UI',Arial,sans-serif">FYI: Candidate Submitted by Multiple Recruiters</h1>
    <p style="margin:6px 0 0;color:#ffffff;font-size:13px;font-family:'Segoe UI',Arial,sans-serif">Detected at: {now_str}</p>
  </div>

  <div class="body" style="padding:24px 32px;background:#ffffff;color:#1a1a2e;font-family:'Segoe UI',Arial,sans-serif">

    <div style="background:#e3f2fd;border:1px solid #90caf9;border-radius:6px;
                padding:14px 20px;margin-bottom:20px;font-size:14px;color:#0d47a1">
      <strong>FYI — no action required.</strong><br>
      The candidate <strong>{candidate_name}</strong> was submitted for
      <strong>{jr_no}</strong> on <strong>{insert_date}</strong>
      by recruiter <strong>{new_recruiter}</strong>.<br>
      An earlier record for the same candidate already exists, added by
      <strong>{existing_recruiter}</strong>.<br><br>
      {status_line}
    </div>

    <h3 style="margin:0 0 12px;color:#1565c0;font-family:'Segoe UI',Arial,sans-serif">Field Comparison</h3>
    <p style="font-size:12px;color:#888888;margin:0 0 10px;font-family:'Segoe UI',Arial,sans-serif">
      Highlighted rows have different values between the two submissions.
    </p>
    {comparison_html}

  </div>

  <div class="footer" style="padding:16px 32px;background:#f8f9fa;font-size:11px;color:#9e9e9e;border-top:1px solid #e0e0e0">
    This is an automated informational message from {TARGET_MAILBOX}.<br>
    No action is required — the record has been added automatically.
  </div>

</div></body></html>"""

    cc_list: list = []
    # CC the existing recruiter so both parties are informed
    if existing_recruiter_addr and existing_recruiter_addr != new_recruiter_addr:
        cc_list.append({"emailAddress": {"address": existing_recruiter_addr}})
    # Standard CC list (team lead etc.)
    for addr in NOTIFY_CC_EMAILS.split(","):
        addr = addr.strip()
        if addr:
            cc_list.append({"emailAddress": {"address": addr}})
    # Manager-specific CC for recruiter-level duplicate alerts
    _existing_cc_addrs = {r["emailAddress"]["address"] for r in cc_list}
    for addr in DIFF_RECRUITER_MANAGER_EMAILS.split(","):
        addr = addr.strip()
        if addr and addr not in _existing_cc_addrs:
            cc_list.append({"emailAddress": {"address": addr}})
            _existing_cc_addrs.add(addr)

    payload = {
        "message": {
            "subject":      subject_line,
            "body":         {"contentType": "HTML", "content": html},
            "toRecipients": [{"emailAddress": {"address": new_recruiter_addr}}],
            "ccRecipients": cc_list,
        },
        "saveToSentItems": "false",
    }

    _send_graph_mail(token, payload, new_recruiter_addr, label="diff-recruiter notification")


# ─── Shared Graph send helper ────────────────────────────────────────────────────
def _send_graph_mail(
    token: str,
    payload: dict,
    primary_recipient: str,
    label: str = "notification",
) -> None:
    if not ENABLE_NOTIFICATIONS:
        log.info("Notifications disabled (ENABLE_NOTIFICATIONS=false) — skipping.")
        return
    url  = f"https://graph.microsoft.com/v1.0/users/{TARGET_MAILBOX}/sendMail"
    resp = requests.post(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type":  "application/json",
        },
        json=payload,
        timeout=30,
    )
    if resp.status_code == 202:
        log.info(f"  ✉ {label} sent → {primary_recipient}")
    else:
        log.error(
            f"  ✉ {label} failed for {primary_recipient}: "
            f"{resp.status_code} — {resp.text[:300]}"
        )
