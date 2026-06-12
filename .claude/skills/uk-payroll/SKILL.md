---
name: uk-payroll
description: >
  Process UK monthly payroll into a NetSuite journal entry, including HMRC P32
  reconciliation. Wraps the production Python script at
  scripts/uk-payroll/process_uk_payroll.py. Use this skill when the user
  mentions: UK payroll, Acme UK Ltd payroll, UK payroll JE, PAYE, NI,
  UK Pension Fund Provider, P32, or drops a PREReport-XLSX file into Monthly Payroll/Pay Runs/UK/.
---

# UK Payroll JE Skill

## Overview

Processes the monthly UK payroll spreadsheet plus an HMRC P32 Summary PDF into a
NetSuite journal entry. The Python script
[`process_uk_payroll.py`](../../../scripts/uk-payroll/process_uk_payroll.py)
handles all bucket logic, dual COGS/OpEx routing across departments, P32 PDF
reconciliation (SMP Unrecovered + Apprenticeship Levy), and writes a backup
workbook with a `raw` + `JE` tab.

**Cadence:** monthly, last day of month.
**Subsidiary:** `Acme Holdings : Acme, Inc. : Acme UK Ltd` (NetSuite ID 8).
**Currency:** GBP.

**Script:** `scripts/uk-payroll/process_uk_payroll.py`
**Mapping files:**
- `scripts/uk-payroll/UK Payroll Employee Mapping.csv` - Employee Name -> Department.
- `scripts/uk-payroll/UK Payroll GL Mapping.csv` - raw column -> COGS/OpEx GL accounts + role classification.
- **Dependency:** `pdfplumber` (for P32 PDF parsing). Already installed; if missing on a fresh machine: `pip install pdfplumber`.

## Step 1 - the accountant drops two files

the accountant drops both files into `Monthly Payroll/Pay Runs/UK/YYYY-MM/`:
1. **Raw provider file:** `PREReport-XLSX*.xlsx` (sheet name `PreReport`)
2. **HMRC P32 Summary PDF:** `P32Summary_YYYY-MM.pdf`

**IMPORTANT:** the filename's `MM` is the **UK tax month** (1=April, 12=March), NOT the calendar month. The script reads the suffix and finds the matching `TOTAL for Month {N}` row inside the PDF.

| Calendar month | UK tax month | Filename suffix |
|---|---|---|
| Jan 2026 | 10 | `P32Summary_2025-10.pdf` |
| Feb 2026 | 11 | `P32Summary_2025-11.pdf` |
| Mar 2026 | 12 | `P32Summary_2025-12.pdf` |
| Apr 2026 | 1  | `P32Summary_2026-01.pdf` |
| May 2026 | 2  | `P32Summary_2026-02.pdf` |
| ... | ... | ... |

If the P32 PDF is missing, the script halts with a clear message.

## Step 2 - Run the script

```bash
cd scripts/uk-payroll
python process_uk_payroll.py \
  "../../Monthly Payroll/Pay Runs/UK/YYYY-MM/PREReport-XLSX (NN).xlsx"
```

The PDF is auto-discovered in the same folder. Default output: `Monthly Payroll/Pay Runs/UK/YYYY-MM/YYYY-MM UK Payroll Backup.xlsx`.

## Step 3 - Per-employee bucket formulas (Feb/Mar pattern - definitive)

| Bucket | Raw columns (sum) | GL Account (COGS) | GL Account (OpEx) |
|---|---|---|---|
| **Salary** | `Salary (£)` + `Company Maternity Pay (£)` | 511100 COGS - Salaries and Wages | 611100 Salaries and Wages |
| **Severance** | `Holiday Pay (£)` + `Lieu of Notice (£)` + `Termination Payment` + `Termination Under £30k` + `Additional Pay (£)` | 511175 COGS - Severance | 611250 Severance |
| **Bonus** | `Bonus (£)` + `Incentives (£)` | 511150 COGS - Bonus | 611150 Bonus |
| **Commission** | `Commission (£)` | 511200 COGS - Commission | 611200 Commission |
| **401K Match** | `PenEr` | 511250 COGS - 401k Match | 611350 401k Match |
| **Payroll Taxes** | `ErNI` + `Class1ANics` | 511350 COGS - Payroll Taxes | 611450 Payroll Taxes |
| **Other Benefits** | `Car Allowance (£)` + `Mobile Allowance (£)` + `Gift Over £50` | 511300 COGS - Other Benefits | 611400 Other Benefits |
| **Health Benefits** (CR) | `Health Insurance (£)` (sign-flipped: provider stores as negative deduction) | 511200 COGS - Health Benefits | 611300 Health Benefits |

Routing: department string starting with `COGS :` -> COGS account; otherwise OpEx. Severance always routes to `EBITDA Adjustments` department, regardless of employee's normal department.

## Step 4 - HMRC P32 PDF reconciliation

The P32 PDF has 24 numeric columns per "TOTAL for Month N" line. The script extracts these key fields:

| Field | Meaning | Position |
|---|---|---|
| PAYE Tax | Income tax withheld from employees | col 1 |
| SMP Recovered | What HMRC reimburses for statutory maternity pay | col 6 |
| App Levy | Apprenticeship Levy (employer payroll tax) | col 19 (third from end) |
| Total Due | Net amount HMRC will draw | last col |

The script uses these to:

1. **HMRC Liability credit** = PDF Total Due. Posted to `231200 Payroll Liability` with memo `YYYY-MM UK Payroll - HMRC Liability`.
2. **SMP Unrecovered DR** = `provider_SMP - PDF_SMP_Recovered`. If positive, posted to the SMP recipient's department (511100/611100 Salaries and Wages). If multiple SMP recipients, allocated pro-rata.
3. **Apprenticeship Levy DR** = PDF App Levy. Posted to `611450 Payroll Taxes` (or `511350` for COGS-side recipient) under the SMP recipient's department. If no SMP recipient this month, defaults to `Sales & Marketing : Revenue : New Business` (Natalie Little's department - per the accountant's historical convention).

## Step 5 - Liability credits

Three lines, all credit `231200 Payroll Liability`:

| Memo | Source |
|---|---|
| `YYYY-MM UK Payroll Liability` | sum(`NetPay` col AE) |
| `YYYY-MM UK Payroll - Pension Liability (EE+ER)` | sum(`TotalPens` col AJ) |
| `YYYY-MM UK Payroll - HMRC Liability` | PDF `Total Due` (already nets SMP Recovered + App Levy) |

## Step 6 - Handle halts

The script refuses to proceed on unknowns. The mapping CSVs are the authoritative knowledge base.

| Halt reason | Fix |
|---|---|
| `Unknown employees (not in UK Payroll Employee Mapping.csv): ...` | Ask the accountant which department the new employee belongs to. Add a row to the Employee Mapping CSV. Re-run. |
| `Unknown columns in raw file ...` | Provider added a new pay code column. Ask the accountant how to map it (COGS + OpEx). Add a row to the GL Mapping CSV. Re-run. |
| `JE out of balance: Debits=X, Credits=Y, Imbalance=Z` | Investigate; a new code may have appeared in the provider summary that the script isn't tracking. |
| `No P32Summary_*.pdf found in {folder}` | Drop the HMRC P32 PDF into the same folder as the raw payroll. |
| `Could not find 'TOTAL for Month {N}' line in {pdf}` | The PDF's tax month doesn't match the filename suffix. Verify the filename. |

## Step 7 - Review the generated JE

Open the output `YYYY-MM UK Payroll Backup.xlsx`. The `JE` tab matches the historical backup column structure (`Date | Journal Entry Memo | Account | Debit | Credit | Line Memo | Subsidiary | Department`), aggregated by (Account, Dept, Memo). Lines with zero amounts are skipped.

Display the JE preview in chat (group by category, show DR/CR with totals), include diagnostic info from the script (SMP Unrecovered, App Levy, HMRC Total Due), then ask:

> **"Review the above carefully. Post to NetSuite? Type `yes` to confirm or anything else to cancel."**

## Step 8 - Generate CSV for NetSuite UI upload (on `yes`)

**MCP-based JE posting is suspended** (see [`_shared/approval-required.md`](../_shared/approval-required.md)). Generate a CSV alongside the Backup.xlsx instead.

1. Write `Monthly Payroll/Pay Runs/UK/YYYY-MM/{YYYY-MM} UK Payroll JE Import.csv` with columns: `Date, Journal Entry Memo, Currency, Account, Debit, Credit, Line Memo, Subsidiary, Department`. Currency `GBP` populated on every row. Date format `M/D/YYYY`. Empty cells stay empty (no zeros). Quote the subsidiary path because of the comma.
2. Tell the accountant:
   > "CSV ready at {path}. Upload via NetSuite UI: **Lists → Import Assistant → Import Type: Transactions → Record Type: Journal Entry** → upload the CSV → confirm field mapping → run import. The JE will land in the controller's Pending Approval queue."
3. Append to `audit_log.json` with `action: "GENERATE_CSV"`. After the accountant imports and reports the JE number, append a follow-up entry with `je_number` and `netsuite_internal_id`.

**DO NOT** call `ns_createRecord` for journal entries.

## Key differences from other payroll skills

- **Largest:** 32 employees, 9 active departments, dual COGS/OpEx routing, ~45 line JE.
- **P32 PDF:** unique to UK; auto-parsed for HMRC reconciliation.
- **SMP Unrecovered + App Levy:** computed from PDF, posted as separate DR lines (different from the accountant's historical Jan which combined them).
- **Health Insurance is a CREDIT:** EE-paid premium deducted from employee net; reduces company expense via CR to Health Benefits account. Provider stores it as a negative number; script flips the sign.
- **611250 OpEx Severance:** UK is the first skill to use the OpEx severance account. May need COA repull.
- **Tax-month filename mapping:** `P32Summary_YYYY-MM.pdf` where MM is UK tax month (Apr=1, Mar=12).

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
