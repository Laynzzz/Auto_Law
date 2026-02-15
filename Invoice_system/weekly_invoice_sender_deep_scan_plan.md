# Law Firm Automation — Phase 5: Weekly Invoice Sender (Deep Scan Edition)

> **Purpose:** This document is an implementation-ready plan for a code agent to build a “deep scan” weekly invoice sender that decides to email invoices **based on dates found inside the invoice content**, not on file metadata.

---

## 1) Goal

Build a Python script (`weekly_send.py`) that runs every Friday. Unlike standard file automations that rely on “Date Modified,” this script must **open the invoice .docx and read the content** to decide if it should be sent.

### Decision Rule (Core)
1. Locate the correct invoice folder for each law firm using a **keyword match** (fuzzy folder search).
2. Find the invoice file for the **current month** (e.g., `February 2026.docx`).
3. **Open & read** the `.docx` to extract service dates.
4. Email the invoice **only if** the invoice contains **any date** within the current work week **(Monday–Friday)**.

**Examples**
- **Scenario A (Send):** Today is **Fri Feb 13, 2026**. Doc contains **Feb 11, 2026** → ✅ SEND
- **Scenario B (Skip):** Today is **Fri Feb 13, 2026**. Doc contains **Feb 17, 2026** (future) → ❌ DO NOT SEND

---

## 2) Inputs & Data Model

### A) Directory Layout
The script must handle “fuzzy” folder names where the actual folder may contain extra words (e.g., “LLC”, “Group”, “NYC”).

```text
/invoice
  /ABC Law Group NYC
    /February 2026.docx
    /February 2026.pdf        <-- optional
  /Law Offices of XYZ
    /February 2026.docx
```

### B) Configuration: `config/firms.yml`
The config maps an “official name” to a folder keyword and email recipients.

```yaml
firms:
  - name: "ABC Law"
    folder_keyword: "ABC Law"     # folder name contains this
    billing_email_to: ["billing@abclaw.com"]
    cc: []
    email_template: "Attached is the invoice for coverage services provided this week."

  - name: "XYZ Associates"
    folder_keyword: "XYZ"         # matches "Law Offices of XYZ"
    billing_email_to: ["invoices@xyz.com"]
    cc: []
    email_template: "Attached is the invoice for coverage services provided this week."
```

### C) Runtime State Files
Store state for idempotency and auditing.

- `state/sent_log.jsonl`  
  One JSON record per successful send.
- `logs/weekly_send.log`  
  General logs (info/warn/error).
- `logs/skipped.log`  
  Human-friendly reasons for skipping.

**Sent log record shape (suggested):**
```json
{
  "timestamp": "2026-02-13T17:31:22-05:00",
  "firm_name": "ABC Law",
  "month_file": "February 2026.docx",
  "folder_path": "invoice/ABC Law Group NYC",
  "attachment_path": "invoice/ABC Law Group NYC/February 2026.pdf",
  "week_start": "2026-02-09",
  "week_end": "2026-02-13",
  "dates_found": ["2026-02-11", "2026-02-17"],
  "outlook_subject": "Weekly Coverage Invoice — Feb 9–Feb 13, 2026"
}
```

---

## 3) The “Deep Scan” Logic (Core Algorithm)

The script must strictly follow this decision tree **for every firm**.

### Step 1: Find the Folder (Fuzzy Keyword Match)
- Iterate directories in `invoice/`
- Find directories whose name contains `config.folder_keyword` (case-insensitive recommended)

**Constraints**
- If **0 matches** → log error and **skip firm**
- If **>1 matches** → log error and **skip firm** (avoid sending to wrong firm)

### Step 2: Find the File (Current Month Docx)
- Construct expected filename: `{Current_Full_Month} {Current_Year}.docx`  
  Example: `"February 2026.docx"`
- Check if this `.docx` exists in the firm folder.
- If a corresponding PDF exists (`February 2026.pdf`), prefer it as the **attachment**, but still use `.docx` for reading.

### Step 3: Parse Dates from the Invoice Content
- Load `.docx` using `python-docx`
- Extract text from:
  - All paragraphs (`document.paragraphs`)
  - All tables (iterate `document.tables`, then rows/cells, read `cell.text`)
- Run a regex scan for common date patterns:
  - `M/D/YY` (e.g., `2/11/26`)
  - `MM/DD/YYYY` (e.g., `02/11/2026`)
  - `Month DD, YYYY` (e.g., `February 11, 2026`)
- Convert found date strings to Python `datetime` objects.
  - Use `dateutil.parser.parse` with `fuzzy=True` where helpful.
  - Normalize to dates in local timezone (America/New_York).

**Important:** Keep results distinct (deduplicate after normalization).

### Step 4: Validate Dates Against Current Week Window
Define:
- `Week_Start` = current week Monday at `00:00:00`
- `Week_End` = current week Friday at `23:59:59`

Validation:
- If **any** parsed date `d` satisfies `Week_Start <= d <= Week_End` → ✅ VALID (queue for sending)
- Else → ❌ IGNORE (even if the file was modified today)

---

## 4) Tech Stack

- Python 3.10+
- `python-docx` — open and read `.docx` text/tables
- `python-dateutil` — `dateutil.parser.parse` for flexible parsing
- `pywin32` — Outlook automation (`win32com.client`)
- `glob`, `os`, `pathlib` — folder search + file operations
- `PyYAML` — load `firms.yml`

---

## 5) Implementation Plan

### Module A: `scan_utils.py`
Helper module to keep main logic clean.

**Functions**
1. `find_firm_folder(root_path: Path, keyword: str) -> Path`
   - Case-insensitive contains match on directory names
   - Raise `FileNotFoundError` if 0 matches
   - Raise `RuntimeError` (or custom) if >1 matches

2. `extract_dates_from_docx(file_path: Path) -> list[datetime]`
   - Read all paragraph text + table cell text
   - Regex scan for candidate date strings
   - Parse candidates into `datetime`
   - Deduplicate and return sorted list

3. `is_date_in_range(d: datetime, start: datetime, end: datetime) -> bool`
   - Inclusive comparison

4. (Optional but recommended) `get_week_window(now: datetime) -> tuple[datetime, datetime]`
   - Return Monday start and Friday end for the week containing `now`

---

### Module B: `weekly_send.py` (Main Script)

**Workflow**
1. Load `config/firms.yml`
2. Compute:
   - `now` (America/New_York)
   - `This_Monday`, `This_Friday`
   - Month filename: `"{Month} {Year}.docx"`
3. For each firm:
   1. `folder = find_firm_folder(invoice_root, firm.folder_keyword)`
   2. `docx_path = folder / month_filename`
      - If not exists → log and continue
   3. `dates = extract_dates_from_docx(docx_path)`
      - If none found → log and continue (or treat as invalid)
   4. If any date in range:
      - Check `state/sent_log.jsonl` to ensure not already sent **for this firm and this week**
      - Compose Outlook email:
        - To: `billing_email_to`
        - CC: `cc`
        - Subject: `Weekly Coverage Invoice — {Week_Range}`
        - Body: from `email_template` + brief date list + week range
      - Attach:
        - Prefer `{Month} {Year}.pdf` if exists
        - Else attach `.docx`
      - Send (unless `--dry-run`)
      - Log success to `sent_log.jsonl`
   5. Else:
      - Log to `logs/skipped.log` with reason (dates found, outside week)

**CLI Flags**
- `--dry-run`:
  - Do not open Outlook / do not send
  - Print decision report to console

- (Optional) `--invoice-root <path>`:
  - Default: `./invoice`

- (Optional) `--week-end <YYYY-MM-DD>`:
  - For backtesting (simulate a Friday)

---

### Module C: Safety Controls

1. **Dry Run Mode (`--dry-run`)**
   - Must not open Outlook or send emails
   - Must print a clear report per firm, e.g.:
     - `Scanning 'invoice/ABC Law Group NYC'...`
     - `Found dates: 2026-02-11, 2026-02-17`
     - `DECISION: SEND (because 2026-02-11 is in range 2026-02-09..2026-02-13)`

2. **Recipient allowlist**
   - Only send to emails specified in `firms.yml`
   - If To list empty → error and skip

3. **Duplicate send prevention**
   - Before sending, check `sent_log.jsonl` for matching keys:
     - firm name + week_start + week_end + month_filename
   - If found → skip and log “already sent”

4. **Ambiguous folder match protection**
   - If multiple folders match keyword, skip (do not choose one automatically)

---

## 6) Example Scenario (Test Case)

**Context**
- Today: **Friday, Feb 13, 2026**
- Week Window: **Feb 9 – Feb 13, 2026**
- File: `invoice/ABC Law Group NYC/February 2026.docx`

### Scenario A (Send)
Invoice contains:
- `Service Date: 2/11/26`
- `Service Date: 2/17/26`

Result:
- `2/11/26` is inside Feb 9–Feb 13 → ✅ Email generated and sent (once)

### Scenario B (Skip)
Invoice contains only:
- `Service Date: 2/17/26`

Result:
- No date inside Feb 9–Feb 13 → ❌ No email generated

---

## 7) Open Questions (If you want zero guessing)
1. Should the script scan **only the current month file**, or also prior month files if the week spans months?
2. If multiple dates are found, do you want to restrict to only dates near “Service Date” labels (to avoid catching unrelated dates)?
3. Email subject/body format requirements from your firm (signature, disclaimers, matter number, etc.)?
4. Do you want to **send one email per firm** with a single monthly invoice, or attach multiple files if there are multiple invoices?
