---
name: poland-payroll
description: >
  Process Poland monthly payroll data into a NetSuite journal entry. Wraps the
  production Python script at scripts/poland-payroll/process_poland_payroll.py.
  Use this skill when the user mentions: Poland payroll, Polish payroll,
  Acme Poland payroll, Poland payroll JE, or drops a
  ACME CORP Payroll report MM YYYY.xlsx file into Monthly Payroll/Pay Runs/Poland/.
---

# Poland Payroll JE Skill

## Overview

Processes the monthly Polish payroll spreadsheet into a NetSuite journal entry.
The Python script [`process_poland_payroll.py`](../../../scripts/poland-payroll/process_poland_payroll.py)
handles all mapping logic and writes a workbook with a `raw` tab + `JE` tab
matching the accountant's historical backup format.

**Cadence:** Poland is **monthly** - one pay run on the last day of the month.
Payroll date comes from the raw file and the JE posts on the last day of that month.

**Subsidiary:** `Acme Holdings : Acme, Inc. : Acme Poland` (subsidiary internal ID 6). Currency is **PLN**.

**Script:** `scripts/poland-payroll/process_poland_payroll.py`
**Mapping files:**
- `scripts/poland-payroll/Poland Payroll Employee Mapping.csv` - Employee Name -> Department. Department string drives COGS (511xxx) vs OpEx (611xxx) routing: any department starting with `COGS :` uses COGS accounts, anything else uses OpEx. All current Polish employees are `COGS : Professional Services`.
- `scripts/poland-payroll/Poland Payroll GL Mapping.csv` - raw column -> GL account, split by COGS / OpEx side.

## Step 1 - the accountant drops the raw file

the accountant drops `ACME CORP Payroll report {MM} {YYYY}[ v2].xlsx` into
`Monthly Payroll/Pay Runs/Poland/MM-YYYY/` (e.g., `04-2026`).

The raw file has a single sheet `Acme Corp {MM} {YYYY}`:
- Row 1: English headers (use this row for column lookup)
- Row 2: Polish headers
- Rows 3+: one row per employee
- Sum row followed by a summary block at the bottom with labels `NET SALARY for employees`, `EE & ER Social Insurance & Health Insurance`, `TAX`, `PPK`, `Bank fee`, `TOTAL`

## Step 2 - Run the script

```bash
cd scripts/poland-payroll
python process_poland_payroll.py \
  "../../Monthly Payroll/Pay Runs/Poland/MM-YYYY/ACME CORP Payroll report MM YYYY.xlsx"
```

Default output: `Monthly Payroll/Pay Runs/Poland/MM-YYYY/YYYY-MM Poland Payroll Backup.xlsx`.

## Step 3 - Handle halts

The script refuses to proceed on unknowns. The mapping CSVs are the authoritative
knowledge base.

| Halt reason | Fix |
|---|---|
| `Unknown employees (not in Poland Payroll Employee Mapping.csv): ...` | Ask the accountant which department the new employee belongs to. Add a row to the Employee Mapping CSV. Re-run. |
| `Unknown columns in raw file ...` | Raw file has a new column (new pay code from provider). Ask the accountant what GL account it should map to on COGS and OpEx sides. Add a row to the GL Mapping CSV. Re-run. |
| `Cannot reconcile DR to CR. ... Residual did not match Y, Z, or Y+Z` | An unusual column is contributing to the provider's summary total that the script isn't tracking. Diagnose and extend the mapping. |
| `JE out of balance` | Investigate the raw file - provider may have changed the summary format. |

## Step 4 - Variable Y / Z columns (auto-detected)

Two columns - `Equivalent for unused day` (Y, column Y) and `Odprawa bez ZUS` (Z,
column Z) - are classified as `VariableExpense` and `VariableSeverance` in the GL
mapping. They're sometimes cash expenses and sometimes informational only; the
provider decides month by month.

The script reconciles: fixed-expense DR total vs. provider summary CR total. The
residual is compared against Y, Z, and Y+Z. Whichever matches determines which
columns are expensed this month:

| Residual | Script action |
|---|---|
| ~= 0 | Neither Y nor Z expensed. Both informational (or zero). |
| = Y total | Vacation payout IS cash -> route to Bonus (511150/611150). |
| = Z total | Severance IS cash -> route to Severance (511175 + EBITDA Adjustments). |
| = Y+Z | Both cash -> route both accordingly. |
| other | Halt - unknown item in summary total. |

Informative messages are printed so you can see what the script decided.

## Step 5 - Ignored columns

- **Gift card** (column AA) - always ignored per policy (matches Germany / Netherlands). Script prints a note if non-zero.
- **Medical Care, Fitness, Relocation** - booked separately via vendor bills; NOT in the payroll JE.

## Step 6 - Review the generated JE

Open the generated `YYYY-MM Poland Payroll Backup.xlsx`. The `JE` tab has the
same columns as the accountant's historical backups: `Date | Journal Entry Memo | Account
| Debit | Credit | Line Memo | Subsidiary | Department`. Lines are aggregated
by (Account, Department, Line Memo) - matches historical style.

Display a preview of the JE in chat (account, debit, credit, dept, memo) with
totals, then ask:

> **"Review the above carefully. Post to NetSuite? Type `yes` to confirm or anything else to cancel."**

## Step 7 - Generate CSV for NetSuite UI upload (on `yes`)

**MCP-based JE posting is suspended** (see [`_shared/approval-required.md`](../_shared/approval-required.md)). Generate a CSV alongside the Backup.xlsx instead.

1. Write `Monthly Payroll/Pay Runs/Poland/MM-YYYY/{YYYY-MM} Poland Payroll JE Import.csv` with columns: `Date, Journal Entry Memo, Currency, Account, Debit, Credit, Line Memo, Subsidiary, Department`. Currency `PLN` populated on every row. Date format `M/D/YYYY`. Empty cells stay empty (no zeros). Quote the subsidiary path because of the comma.
2. Tell the accountant:
   > "CSV ready at {path}. Upload via NetSuite UI: **Lists → Import Assistant → Import Type: Transactions → Record Type: Journal Entry** → upload the CSV → confirm field mapping → run import. The JE will land in the controller's Pending Approval queue."
3. Append to `audit_log.json` with `action: "GENERATE_CSV"`. After the accountant imports and reports the JE number, append a follow-up entry with `je_number` and `netsuite_internal_id`.

**DO NOT** call `ns_createRecord` for journal entries.

## Key differences from Germany / Netherlands payroll skills

- **Liability split:** Poland splits into THREE accounts (`231200 Payroll Liability` + `231350 Payroll Tax Liability` + `231250 401K payable`) - unlike Germany (single 231200) and Netherlands (different pattern).
- **Currency:** PLN (has its own subsidiary, unlike Germany which routes to Netherlands).
- **PPK (Pracowniczy Plan Kapitalowy):** Polish employee retirement plan; treated like 401K Match (511250).
- **Variable Y/Z logic:** Poland has inconsistent monthly handling of vacation payout and severance; the script reconciles against summary totals to auto-classify.
- **Bilingual raw headers:** English on row 1, Polish on row 2 - script keys off the English row.

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
