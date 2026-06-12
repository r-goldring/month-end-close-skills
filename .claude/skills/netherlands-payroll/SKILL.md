---
name: netherlands-payroll
description: >
  Process Netherlands monthly payroll data into a balanced NetSuite journal entry.
  Wraps the production Python scripts at scripts/netherlands-payroll/.
  Use this skill when the user mentions: Netherlands payroll, NL payroll, Dutch
  payroll, BV payroll, Acme BV payroll, or drops a
  CompanyEmployeeWageComponents*.xlsx file into
  Monthly Payroll/Pay Runs/Netherlands/MM-YYYY/.
---

# Netherlands Payroll JE Skill

## Overview

Processes the monthly Netherlands payroll export from the Dutch payroll system
into a balanced NetSuite journal entry. Mirrors the Canada/US/Germany payroll
skill pattern: preflight check for new employees and component codes, then
the main script that aggregates the raw component-per-employee data into a
16-row JE across the two Product (R&D) and Professional Services (COGS)
departments.

**Cadence:** Netherlands is **monthly** — one pay run on the last day of the
month (e.g., `03-2026` folder for March 2026).

**Subsidiary:** `Acme Holdings : Acme, Inc. : Acme Netherlands`.

**Currency:** EUR throughout — no USD conversion in the JE.

**Scripts:**
- [`scripts/netherlands-payroll/check_mappings.py`](../../../scripts/netherlands-payroll/check_mappings.py) — preflight
- [`scripts/netherlands-payroll/process_netherlands_payroll.py`](../../../scripts/netherlands-payroll/process_netherlands_payroll.py) — JE generator

**Mapping files** (authoritative knowledge base — update here when new employees or codes appear):
- `scripts/netherlands-payroll/Netherlands Payroll Employee Mapping.csv` — Employee ID to Department
- `scripts/netherlands-payroll/Netherlands Payroll Component Mapping.csv` — component name to OpEx account, COGS account, liability account

## Step 1 — the accountant drops the raw file

the accountant drops the Dutch payroll export into:
```
Monthly Payroll/Pay Runs/Netherlands/MM-YYYY/
```
(e.g., `04-2026` for April 2026). The raw file is named:
- `CompanyEmployeeWageComponentsPerPeriod_Acme Corp_BV*.xlsx` (regular recurring months with P1..P12 columns), OR
- `CompanyEmployeeWageComponentsPerRunPeriod_Acme Corp_BV*.xlsx` (bonus / multi-run months with `R1/P1`, `R2/P2` columns, used for MIP months)

Single sheet `Page_1` with columns:
- `Nr` (Employee ID, e.g., 2002, 2008)
- `Naam` (employee name)
- `Component Nr.` (Dutch payroll code, e.g., 1000, 6539, 8800)
- `Component naam` (component name, e.g., "Gross salary", "Wage tax (table)")
- Period columns (P1..P12 or R?/P?) with EUR amounts
- `Cumulative` — running total

## Step 2 — REQUIRED: Run the preflight

Before running the main script, always run the mapping preflight. It scans the
raw file for Employee IDs and Component Names and flags any that aren't mapped
yet, so they can be added to the knowledge base (mapping CSVs) instead of being
silently dropped from the JE.

```bash
cd scripts/netherlands-payroll
python check_mappings.py "../../Monthly Payroll/Pay Runs/Netherlands/MM-YYYY/"
```

Exit 0 = all mapped, safe to proceed. Exit 1 = new items found.

### If the preflight surfaces new items

| Finding | Fix |
|---------|-----|
| New Employee ID (e.g., new hire) | Ask the accountant which department — Product (R&D) or Professional Services (COGS). Add a row to `Netherlands Payroll Employee Mapping.csv` with `Employee ID, Employee Name, Department, Subsidiary`. Re-run preflight. |
| New Component Name | Ask the accountant how this Dutch payroll line should be routed: to which OpEx and COGS expense account, and to which liability account (if any). Check existing similar rows in `Netherlands Payroll Component Mapping.csv` for the pattern. Add a row. Re-run preflight. |
| Departed employee (in mapping but not in raw) | No action needed; preflight only flags items in the raw that aren't mapped, not the reverse. |

## Step 3 — Run the main script

Only when preflight is clean (or the accountant has explicitly acknowledged any unmapped items):

```bash
cd scripts/netherlands-payroll
python process_netherlands_payroll.py "../../Monthly Payroll/Pay Runs/Netherlands/MM-YYYY/"
```

Script behavior:
- Auto-detects the raw file in the folder (any `CompanyEmployeeWageComponents*.xlsx` that isn't a backup)
- Parses year + month from the folder name (`03-2026` produces 2026-03)
- Finds the current-month period column (prefers `P{month}`, falls back to `R?/P{month}` for bonus-run files)
- Looks up department via employee mapping; looks up GL accounts via component mapping
- Aggregates expense lines by (department, account); aggregates liability lines by account
- Writes the output workbook to `Monthly Payroll/Pay Runs/Netherlands/MM-YYYY/YYYY-MM Netherlands Payroll Backup.xlsx` with three sheets:
  - `Sheet1` — aggregation by (Department × Component)
  - `Page_1` — raw records with Department added
  - `JE` — the 16-row balanced journal entry with hardcoded debit/credit values

## Step 4 — Handle halts and warnings

| Condition | What it means | Fix |
|-----------|---------------|-----|
| Unmapped components warning | Preflight was skipped or a new code appeared mid-processing | Stop, run preflight, update mapping CSV, re-run |
| `Pension Wg present but Pension Wn missing` warning | Raw file (typically a bonus-run export) is missing the employee pension contribution rows | Open the output JE and manually add the Pension Wn amounts to the `231205 Netherlands Pension Payable` credit line. Ask the accountant for the correct figures (he has them from the full payroll report). |
| `STATUS: OUT OF BALANCE` | Totals don't match to the penny | Review debits vs credits; the most common cause is missing Pension Wn on a bonus run. Manual adjustment required before posting. |

## Step 5 — Review the JE

Open `YYYY-MM Netherlands Payroll Backup.xlsx` and inspect the `JE` tab.
Standard 16-row layout (rows 4-19):

| # | Account | Side | Department |
|---|---------|------|------------|
| 1 | 611100 Salaries and Wages | Debit | R&D : Product |
| 2 | 511100 COGS Salaries and Wages | Debit | COGS : Professional Services |
| 3 | 231206 Holiday Pay Accrual (placeholder 0) | Debit | Product |
| 4 | 231206 Holiday Pay Accrual (placeholder 0) | Debit | Professional Services |
| 5 | 611400 Other Benefits (placeholder 0) | Debit | Product |
| 6 | 511300 COGS Other Benefits (placeholder 0) | Debit | Professional Services |
| 7 | 611350 401k Match | Debit | Product |
| 8 | 511250 COGS 401k Match | Debit | Professional Services |
| 9 | 611150 Bonus | Debit | Product |
| 10 | 511150 COGS Bonus | Debit | Professional Services |
| 11 | 611450 Payroll Taxes | Debit | Product |
| 12 | 511350 COGS Payroll Taxes | Debit | Professional Services |
| 13 | 231200 Payroll Liability | Credit | (aggregate) |
| 14 | 231205 Netherlands Pension Payable | Credit | (aggregate) |
| 15 | 231350 Payroll Tax Liability | Credit | (aggregate) |
| 16 | 231206 Netherlands Holiday Pay Accrual | Credit | (aggregate) |

Display a preview of the JE in chat (account, debit, credit, department, memo)
with totals, then ask the accountant:

> **"Review the above carefully. Post to NetSuite? Type 'yes' to confirm or anything else to cancel."**

## Step 6 — Generate CSV for NetSuite UI upload (on `yes`)

**MCP-based JE posting is suspended** (see [`_shared/approval-required.md`](../_shared/approval-required.md)). Generate a CSV alongside the Backup.xlsx instead.

1. Write `Monthly Payroll/Pay Runs/Netherlands/MM-YYYY/{YYYY-MM} Netherlands Payroll JE Import.csv` with columns: `Date, Journal Entry Memo, Currency, Account, Debit, Credit, Line Memo, Subsidiary, Department`. Currency `EUR` populated on every row. Date format `M/D/YYYY`. Empty cells stay empty (no zeros). Quote the subsidiary path because of the comma.
2. Tell the accountant:
   > "CSV ready at {path}. Upload via NetSuite UI: **Lists → Import Assistant → Import Type: Transactions → Record Type: Journal Entry** → upload the CSV → confirm field mapping → run import. The JE will land in the controller's Pending Approval queue."
3. Append to `audit_log.json` with `action: "GENERATE_CSV"`. After the accountant imports and reports the JE number, append a follow-up entry with `je_number` and `netsuite_internal_id`.

**DO NOT** call `ns_createRecord` for journal entries.

## Component routing reference

From the January / February / March 2026 historical JEs, each Dutch component routes as:

| Component | Expense side | Liability side | Notes |
|-----------|--------------|----------------|-------|
| Gross salary | 611100 / 511100 Salaries | - | Base salary |
| Reservation holiday allowance | 611100 / 511100 Salaries (added to salary debit) | 231206 Holiday Pay Accrual credit | Booked to BOTH sides in the same amount |
| Incentive Bonus(special rate) | 611150 / 511150 Bonus | - | MIP / bonus runs |
| Pension premium Wg | 611350 / 511250 401k Match | 231205 Pension Payable | Hits both |
| Pension premium Wn | - | 231205 Pension Payable (absolute) | Liability only, raw is negative |
| Aof / WGA / Unemployment / WKO / ZVW / ZW premiums | 611450 / 511350 Payroll Taxes (summed) | 231350 Payroll Tax Liability (summed) | 6 different tax codes all feed same expense + liability |
| Wage tax (table) / Wage tax (spec. rate) | - | 231350 Payroll Tax Liability (absolute) | Liability only, raw is negative |
| Total net / Per bankaccount | - | 231200 Payroll Liability | Net pay owed to employees |

## Key differences from Canada payroll

- **Cadence:** monthly (Canada is bi-weekly, 15th + last day)
- **Folder convention:** `MM-YYYY` (e.g., `03-2026`); Canada uses `MM.DD.YYYY`
- **Input format:** Dutch payroll `CompanyEmployeeWageComponents*.xlsx` with one row per (employee, component); Canada uses ADP `.xls` with two sheets and block-per-cost-center layout
- **Routing key:** Employee ID to Department (2002 = Professional Services, 2008 = Product); Canada uses Cost Center codes
- **Currency:** EUR (Canada is CAD); output amounts unconverted
- **No 70/30 Infrastructure split:** flat per-employee department assignment
- **Liability structure:** 4 separate liability accounts (231200 net pay, 231205 pension, 231206 holiday accrual, 231350 tax); Canada uses 3 (231200, 231250 401k, 231350)

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
