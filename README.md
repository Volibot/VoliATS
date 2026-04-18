# HR Email Extractor — Volibits

Monitors `HRvolibot@volibits.com` for recruitment emails, parses candidate tables, and inserts records into the `hrvolibit` PostgreSQL table. Runs automatically every 5 minutes via GitHub Actions.

---

## Files in this repo

```
├── .github/
│   └── workflows/
│       └── email_extractor.yml   ← GitHub Actions (runs every 5 min)
├── email_extractor.py            ← Core logic
├── scheduler.py                  ← Local alternative to GitHub Actions
├── requirements.txt
├── .env.example                  ← Copy to .env for local runs
├── .gitignore
└── README.md
```

---

## One-time GitHub Setup

### 1. Push to GitHub

```bash
git init
git add .
git commit -m "Initial commit: HR email extractor"
git remote add origin https://github.com/YOUR_ORG/YOUR_REPO.git
git push -u origin main
```

### 2. Add GitHub Secrets

**Repo → Settings → Secrets and variables → Actions → New repository secret**

| Secret Name | Where to find it |
|---|---|
| `AZURE_TENANT_ID` | Azure Portal → Azure Active Directory → Overview |
| `AZURE_CLIENT_ID` | Azure Portal → App Registrations → your app → Overview |
| `AZURE_CLIENT_SECRET` | Azure Portal → your app → Certificates & secrets → Client secrets |
| `TARGET_MAILBOX` | `HRvolibot@volibits.com` |
| `DB_DSN` | `host=... dbname=... user=... password=... port=5432` |

### 3. Enable Actions

**Repo → Actions tab** → confirm enabling workflows if prompted.

To run manually: **Actions → HR Email Extractor → Run workflow**.

---

## How It Works

```
Every 5 minutes
      │
      ▼
Fetch unread emails from HRvolibot@volibits.com
      │
      ▼
Parse subject line
  BS: Java Developer   → process  (code=BS, company=Birlasoft, skill=Java Developer)
  FW: BS: Java Dev     → SKIP     (forwarded)
  Re: RS: AWS Lead     → SKIP     (reply)
  Random subject       → SKIP     (no match)
      │
      ▼
Parse HTML table in body (fuzzy header matching)
      │
      ▼
Duplicate check per candidate row
      │
      ▼
Insert into hrvolibit (always — even duplicates, even on error)
      │
      ▼
Mark email as read
```

---

## recruiter & client_recruiter logic

Both fields use the **username part (before @) of the email address**, regardless of domain:

| Email address | Stored as |
|---|---|
| `meenakshi.randhawa@volibits.com` | `meenakshi.randhawa` |
| `vishwa.chintam@birlasoft.com` | `vishwa.chintam` |
| `kirtikumar.dhruv@volibits.com` | `kirtikumar.dhruv` |

- `recruiter` = sender's username
- `client_recruiter` = first To: recipient's username

---

## Delivery Type

| From | To | delivery_type |
|---|---|---|
| @volibits.com | @volibits.com | Internal |
| Anything else | — | External |

---

## Duplicate Logic

| Condition | is_duplicate |
|---|---|
| Same phone AND email already in DB | `Duplicate` |
| Only phone matches | `Duplicate Cell` |
| Only email matches | `Duplicate Email` |
| No match | *(null)* |

Records are **always inserted** regardless of duplicate status.

---

## Column Header Aliases

The script accepts many variations of column names automatically:

| DB Column | Example accepted headers |
|---|---|
| `jr_no` | JR No, JR, Req ID, Requisition ID, JR Number, Requirement ID |
| `contact_number` | Phone, Mobile, Cell, Contact No, Ph No |
| `current_org` | Current Company, Company, Employer, Organisation |
| `total_experience` | Total Exp, Experience, Total Yrs |
| `name_of_candidate` | Candidate Name, Name, Applicant Name, Full Name |

Add more in `COLUMN_ALIASES` inside `email_extractor.py`.

---

## Error Handling

- DB insert error → partial record inserted with `record_status = 'Error - Partial Insert'` — no record is ever silently lost
- Full run log written to `extractor.log`
- On GitHub Actions failure → log uploaded as an artifact for 7 days

---

## Local Development

```bash
pip install -r requirements.txt
cp .env.example .env        # fill in real values
python email_extractor.py   # single run
python scheduler.py         # continuous polling (every 5 min)
```

---

## Adding Company Codes

Edit `COMPANY_CODES` in `email_extractor.py`:

```python
COMPANY_CODES: dict[str, str] = {
    "BS":  "Birlasoft",
    "BW":  "BeWealthy",
    "ACC": "Accenture",   # ← add new entries here
}
```

The key must match the prefix used in email subjects (e.g. `ACC: Python Developer`).
