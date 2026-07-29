-- ── Preview: see all matches before updating ──────────────────────────────────
-- status = WILL UPDATE  → jr_no is blank, will be filled
-- status = MATCHES OK   → jr_no already exists and agrees (match logic is correct)
-- status = CONFLICT     → jr_no already exists but differs (review before updating)

SELECT
    t.id                                            AS target_id,
    t.name_of_candidate                             AS name,
    COALESCE(t.email_id, t.contact_number)          AS contact,
    t.date::text                                    AS date,
    t.jr_no                                         AS current_jr_no,
    s.jr_no                                         AS jr_no_from_candidates,
    CASE
        WHEN t.jr_no IS NULL OR TRIM(t.jr_no) = '' THEN 'WILL UPDATE'
        WHEN TRIM(t.jr_no) = TRIM(s.jr_no)         THEN 'MATCHES OK'
        ELSE                                             'CONFLICT'
    END                                             AS status
FROM hrvolibit t
JOIN candidates s
  ON  s.jr_no IS NOT NULL
  AND LOWER(TRIM(t.name_of_candidate)) = LOWER(TRIM(s.name))
  AND (
       LOWER(TRIM(t.email_id))  = LOWER(TRIM(s.email))
    OR TRIM(t.contact_number)   = TRIM(s.phone)
  )
  AND t.date = s.date
ORDER BY status, t.date DESC, t.name_of_candidate;


-- ── Update: run only after reviewing the preview above ────────────────────────

UPDATE hrvolibit t
SET    jr_no = s.jr_no
FROM   candidates s
WHERE  (t.jr_no IS NULL OR TRIM(t.jr_no) = '')
  AND  s.jr_no IS NOT NULL
  AND  LOWER(TRIM(t.name_of_candidate)) = LOWER(TRIM(s.name))
  AND  (
        LOWER(TRIM(t.email_id))  = LOWER(TRIM(s.email))
     OR TRIM(t.contact_number)   = TRIM(s.phone)
  )
  AND  t.date = s.date;
