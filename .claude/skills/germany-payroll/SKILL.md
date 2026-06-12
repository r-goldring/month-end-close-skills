---
name: germany-payroll
description: >
  Process Germany monthly payroll data into a NetSuite journal entry. Wraps the
  production Python script at scripts/germany-payroll/process_germany_payroll.py.
  Use this skill when the user mentions: Germany payroll, German payroll,
  Acme B.V. payroll, Germany payroll JE, or drops a
  41601_spreadsheet_YYYYMM_01.xlsx file into Monthly Payroll/Pay Runs/Germany/.
---

# Germany Payroll JE Skill

## Overview

Processes the monthly Germany payroll spreadsheet into a NetSuite journal entry.
The Python script [`process_germany_payroll.py`](../../../scripts/germany-payroll/process_germany_payroll.py)
handles all mapping logic and writes a workbook with a `raw` tab + `JE` tab
matching the accountant's historical backup format.

**Cadence:** Germany is **monthly** — one pay run on the last day of the month.

**Subsidiary:** `Acme Holdings : Acme, Inc. : Acme Netherlands`.
There is no separate Germany subsidiary; German employees book to the Netherlands
entity.

**Script:** `scripts/germany-payroll/process_germany_payroll.py`
**Mapping files:**
- `scripts/germany-payroll/Germany Payroll Employee Mapping.csv` — Employee Name to Department. Department string drives COGS (511xxx) vs OpEx (611xxx) routing: any department starting with `COGS :` uses COGS accounts, anything else uses OpEx.
- `scripts/germany-payroll/Germany Payroll GL Mapping.csv` — raw column to GL account, split by COGS / OpEx side.

## Step 1 — the accountant drops the raw file

the accountant drops `41601_spreadsheet_YYYYMM_01.xlsx` into
`Monthly Payroll/Pay Runs/Germany/MM-YYYY/` (e.g., `04-2026`).

The raw file has a single sheet `Tabelle1`:
- Row 1: title (`Acme B.V. Spreadsheet {Month} {Year}`)
- Row 4: headers (`Co | Check Date | System No. | Cost Center | Employee Name | Net Pay | Employers Cost | Salary | Maternity Pay | ...`)
- Rows 5+: one row per employee
- Totals row + `Netpay` / `Social Security contributions` / `Total` summary rows at the bottom

## Step 2 — Ask about severance

Before running the script, ask the accountant:
> "Any severance to include in this month's Germany payroll? If yes, which employee, what amount, and what are the severance-related payroll taxes (if any)?"

If yes, collect `{employee, amount, taxes}` and pass as `severance_spec`. Severance
always codes to account `511175 COGS - Salary and Compensation : COGS - Severance`
and department `EBITDA Adjustments`. Severance-related payroll taxes post to
`511350 COGS - Salary and Compensation : COGS - Payroll Taxes` also under
`EBITDA Adjustments`.

If the accountant doesn't mention severance, proceed without.

## Step 3 — Run the script

```bash
cd scripts/germany-payroll
python process_germany_payroll.py \
  "../../Monthly Payroll/Pay Runs/Germany/MM-YYYY/41601_spreadsheet_YYYYMM_01.xlsx"
```

Default output: `Monthly Payroll/Pay Runs/Germany/MM-YYYY/YYYY-MM Germany Payroll Backup.xlsx`.

For severance runs, invoke programmatically:
```python
from process_germany_payroll import generate_je
generate_je(
    "path/to/41601_spreadsheet_YYYYMM_01.xlsx",
    "Germany Payroll Employee Mapping.csv",
    "Germany Payroll GL Mapping.csv",
    "path/to/output.xlsx",
    severance_spec={"employee": "Last, First", "amount": 10000.00, "taxes": 2200.00},
)
```

## Step 4 — Handle halts

The script refuses to proceed on unknowns. The mapping CSVs are the authoritative
knowledge base — when the script halts, the accountant updates the mapping and re-runs.

| Halt reason | Fix |
|---|---|
| `Unknown employees (not in Germany Payroll Employee Mapping.csv): ...` | Ask the accountant which department the employee belongs to (check Zenefits/BambooHR or ask). Add a row to `Germany Payroll Employee Mapping.csv`. Re-run. |
| `Unknown columns in raw file ...` | Raw file has a new column (new pay code from the provider). Ask the accountant what GL account it should map to on COGS side and OpEx side. Add a row to `Germany Payroll GL Mapping.csv`. Re-run. |
| `JE out of balance: Debits=X, Credits=Y, Imbalance=Z` | Investigate the raw file — a provider change may have broken the Net Pay + Social Security contributions totals. |

## Step 5 — Review the generated JE

Open the generated `YYYY-MM Germany Payroll Backup.xlsx`. The `JE` tab has the
same columns as the accountant's historical backups: `Date | Journal Entry Memo | Account
| Debit | Credit | Line Memo | Subsidiary | Department`. The script aggregates
lines by (Account, Department, Line Memo) to keep the JE summarized (same style
as historical).

Display a preview of the JE in chat (account, debit, credit, dept, memo) with
totals, then ask:

> **"Review the above carefully. Post to NetSuite? Type 'yes' to confirm or anything else to cancel."**

## Step 6 — Generate CSV for NetSuite UI upload (on `yes`)

**MCP-based JE posting is suspended** (see [`_shared/approval-required.md`](../_shared/approval-required.md) — server-side auto-approver was bypassing SoD). Generate a CSV alongside the Backup.xlsx instead.

1. Write `Monthly Payroll/Pay Runs/Germany/MM-YYYY/{YYYY-MM} Germany Payroll JE Import.csv` with columns: `Date, Journal Entry Memo, Account, Debit, Credit, Line Memo, Subsidiary, Department`. Date format `M/D/YYYY`. Empty cells stay empty (no zeros). Quote the subsidiary path because of the comma.
2. Tell the accountant:
   > "CSV ready at {path}. Upload via NetSuite UI: **Lists → Import Assistant → Import Type: Transactions → Record Type: Journal Entry** → upload the CSV → confirm field mapping → run import. The JE will land in the controller's Pending Approval queue."
3. Append to `audit_log.json` with `action: "GENERATE_CSV"`:
   ```json
   {
     "timestamp": "YYYY-MM-DDTHH:MM:SS",
     "skill": "germany-payroll",
     "action": "GENERATE_CSV",
     "description": "YYYY-MM Germany Payroll - N lines, EUR X balanced",
     "csv_file": "Monthly Payroll/Pay Runs/Germany/MM-YYYY/{YYYY-MM} Germany Payroll JE Import.csv",
     "posted_via": "CSV upload pending"
   }
   ```
4. After the accountant imports and tells you the JE number, append a follow-up entry with `je_number` + `netsuite_internal_id`.

**DO NOT** call `ns_createRecord` for journal entries.

## Key differences from Canada payroll

- **Cadence:** monthly (Canada is bi-weekly)
- **Subsidiary:** Netherlands (Canada uses its own subsidiary)
- **Liability split:** single `231200 Payroll Liability` for both Net Pay and taxes (Canada splits to 231200 + 231350)
- **Input:** `.xlsx` single-sheet, one row per employee (Canada is `.xls` with two sheets and block-style layout)
- **Routing:** per-employee department (Canada was per cost-center with 70/30 Infrastructure rule)
- **Gift Card column:** ignored (per policy)
- **Severance:** manually specified by the accountant; auto-codes to `EBITDA Adjustments`
## Pre-upload gut-check review (auto)

Immediately after writing the JE Import CSV, **run the gut-check before
telling the accountant to upload.** See [`gut-check/SKILL.md`](../gut-check/SKILL.md) for
the orchestration. Short version:

1. `payroll_gut_check.run_gut_check(folder_path, audit_log_path,
   suiteql_runner=None)` returns the priors-search SQL.
2. Run that SQL via `mcp__claude_ai_NetSuite__ns_runCustomSuiteQL` to find the
   prior 2 same-skill JEs (memo + subsidiary + before-current-date).
3. Fetch each prior's lines via `build_line_fetch_sql(tranid, sub)`.
4. `analyze(...)`, `format_chat_report(...)`, `write_workbook_tab(...)` to
   produce the PASS/WARN/FAIL report and append a timestamped tab to the
   backup workbook.
5. **Gate the upload instruction on FAIL**: if `result.has_fail`, tell the accountant
   "DO NOT UPLOAD. Fix the CSV first; re-run /gut-check after edits." Do NOT
   print upload instructions. Otherwise tell the accountant it's safe to upload.

Variance thresholds, sign-flip logic, and JB EBITDA / severance routing
checks live in the shared utility — no per-skill tuning needed for v1.
