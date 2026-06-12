---
name: uruguay-payroll
description: >
  Process Uruguay monthly payroll (1-3 payruns per month) into NetSuite JEs.
  Wraps the production Python script at scripts/uruguay-payroll/process_uruguay_payroll.py.
  Use this skill when the user mentions: Uruguay payroll, UY payroll, UY Payroll Provider, Aguinaldo,
  egreso, off-cycle, or drops SGN_892_*.xls* files into Monthly Payroll/Pay Runs/Uruguay/YYYY-MM/.
---

# Uruguay Payroll JE Skill

## Overview

Processes Uruguay monthly payroll into 2-6 NetSuite JEs per month. Each raw file (one per payrun) generates two balanced JEs: **Block 1 (Main Payroll)** and **Block 2 (Aguinaldo Accrual)**. Egreso (severance) payruns produce only Block 1 with severance routing.

**Subsidiary:** `Acme Holdings : Acme, Inc. : Acme Uruguay` (NetSuite ID 13). **Currency:** UYU.

**Script:** `scripts/uruguay-payroll/process_uruguay_payroll.py`
**Mapping files:**
- `scripts/uruguay-payroll/Uruguay Payroll Employee Mapping.csv` - Employee ID -> Department + COGS Split %
- `scripts/uruguay-payroll/Uruguay Payroll GL Mapping.csv` - Pay code -> bucket / GL account

**Dependency:** none beyond stdlib + pandas + openpyxl (already installed).

## Step 1 - the accountant drops raw files

the accountant drops 1-N `SGN_892_*.xls*` raw files into `Monthly Payroll/Pay Runs/Uruguay/YYYY-MM/`. Files include:
- `SGN_892_Mensual_*` or `SGN_892_{MES}_*` (e.g., MARZO, ABRIL): regular monthly run
- `SGN_892_*Pando*` / `SGN_892_*Espinoza*`: off-cycle for a specific employee
- `SGN_892_LiquidacionExtraRun*`: extra-run (typically new hire)
- `SGN_892_Egreso*`: severance/termination

Optional reference files (NOT used by script): `XX. PRCX UR Funding Request *.xlsx`, `*Detalle de Provisiones por Funcionario*.xlsx`.

## Step 2 - Run the script

```bash
cd scripts/uruguay-payroll
python process_uruguay_payroll.py "../../Monthly Payroll/Pay Runs/Uruguay/YYYY-MM"
```

Default Infrastructure COGS split is **30%** (April 2026 onward). For regression testing pre-April months, pass `--infra-cogs-pct=20`.

Default output: `Monthly Payroll/Pay Runs/Uruguay/YYYY-MM/YYYY-MM Uruguay Payroll Backup.xlsx` containing N raw tabs + 2*N (or 2*N-1 for egreso months) JE tabs.

## Step 3 - Per-payrun JE structure

Each payrun produces TWO JEs (except egreso = ONE):

### Block 1 - Main Payroll JE

| DR Bucket | Source (pay code) | GL (COGS) | GL (OpEx) |
|---|---|---|---|
| Salaries & Wages | Total Gross Pay (2000) − Medical Coverage (67) − Special Bonus (50) | 511100 | 611100 |
| Special Bonus (e.g., MIP/BONO) | Code 50 | 511150 | 611150 |
| Payroll Taxes | ER BPS (7001) + ER FRL+FG (7003) + ER FONASA (7007) + Egreso BPS Aguinaldo (7002) + Egreso FRL+FG Aguinaldo (7004) | 511350 | 611450 |

CR side:
- 231200 Payroll Liability = sum of Net Pay (code 2030)
- 231350 Payroll Tax Liability = balance (DR total - Net Pay total)

### Block 2 - Aguinaldo Accrual JE

| DR Bucket | Source (pay code) | GL (COGS) | GL (OpEx) |
|---|---|---|---|
| Aguinaldo Bonus | MONTHLY BONUS PROVISION (6300) | 511150 | 611150 |
| Aguinaldo PR Tax | LICENSE CS PROVISION (6303) + SOCIAL SECURITY BONUS PROVISION (6304) | 511350 | 611450 |

CR side:
- 231207 Accrued PTO = sum of Aguinaldo Bonus DR
- 231171 Accrued Bonus Payroll Tax Liability = sum of Aguinaldo PR Tax DR

### Egreso (severance) routing

When the raw filename contains "Egreso", that payrun's Block 1 expense lines route to:
- Salaries -> `511175 COGS - Severance` / `611250 Severance` (instead of 511100/611100)
- Payroll Taxes -> 511350 / 611450 (unchanged)
- Department -> `EBITDA Adjustments` (overrides employee's normal dept)
- COGS Split % stays the same per employee

Egreso payruns do NOT generate Block 2 (no aguinaldo accrual for terminated employees).

## Step 4 - COGS / OpEx routing

Each employee in `Uruguay Payroll Employee Mapping.csv` has a `COGS Split %`:
- **Engineering, Product, IT employees:** 0% (100% to OpEx 611xxx)
- **Infrastructure / DevOps employees (cc=3, cc=4):** 30% COGS / 70% OpEx (Apr 2026+; was 20/80 pre-Apr)

The split applies to all four expense buckets (Salaries, Payroll Taxes, Aguinaldo Bonus, Aguinaldo PR Tax).

## Step 5 - Handle halts

| Halt reason | Fix |
|---|---|
| `Unknown employees in {file}: ...` | Add row to `Uruguay Payroll Employee Mapping.csv` with their Employee ID, Name, Department, COGS Split %. Re-run. |
| `Unknown pay codes in {file}: [N, ...]` | Provider added a new pay code. Add row to `Uruguay Payroll GL Mapping.csv` with Role and (if applicable) GL accounts. Re-run. |
| `Block N imbalanced` | A new pay code may be misclassified. Check the raw and the mapping; check for missing summary rows. |
| `[FLAG] Special bonus (code 50/MIP)` | Not a halt - a runtime warning. Confirm with the accountant whether the special bonus line is correctly classified (e.g., MIP for prior year). |

## Step 6 - Review the generated JEs

Open the generated `YYYY-MM Uruguay Payroll Backup.xlsx`. Each raw payrun has its own raw tab (`raw_*`) and 1-2 JE tabs (`JE01_...`, `JE02_...`, etc.).

Display previews of all JEs in chat (group by payrun, show Block 1 + Block 2 totals, dept breakdowns). Then ask:

> **"Review the above carefully. Post all N JEs to NetSuite? Type `yes` to confirm or anything else to cancel."**

## Step 7 - Generate CSV for NetSuite UI upload (on `yes`)

**MCP-based JE posting is suspended** (see [`_shared/approval-required.md`](../_shared/approval-required.md)). For each JE that would have been posted, generate a separate CSV file alongside the Backup.xlsx. With 1-3 payruns per month and 1-2 JEs per payrun, expect 2-5 CSVs per month.

For EACH JE:
1. Write `Monthly Payroll/Pay Runs/Uruguay/YYYY-MM/{JE Memo} JE Import.csv` (e.g., `2026-04 Uruguay Payroll JE Import.csv`, `2026-04 Uruguay Payroll - Aguinaldo Accrual JE Import.csv`, `2026-04 Uruguay Payroll - Egreso (GabrielZerbino) JE Import.csv`).
2. Columns: `Date, Journal Entry Memo, Currency, Account, Debit, Credit, Line Memo, Subsidiary, Department`. Currency `UYU` populated on every row. Date format `M/D/YYYY`. Empty cells stay empty (no zeros). Quote the subsidiary path.

Tell the accountant:
> "{N} CSV files ready in {folder}. Upload EACH one separately via NetSuite UI: **Lists → Import Assistant → Import Type: Transactions → Record Type: Journal Entry** → upload the CSV → confirm field mapping → run import. Each JE will land in the controller's Pending Approval queue."

Append to `audit_log.json` with `action: "GENERATE_CSV"` listing all CSV file paths. After the accountant imports and reports the JE numbers, append a follow-up entry mapping each `je_number` to its `netsuite_internal_id`.

**DO NOT** call `ns_createRecord` for journal entries.

## Key differences from other payroll skills

- **Multi-payrun:** 1-3 raw files per month -> 2-5 JEs (each balanced). Most skills are 1 file -> 1-2 JEs.
- **Two-block JE structure:** every regular/extra-run/offcycle payrun produces 2 separate JEs (Main + Aguinaldo Accrual). Egreso produces 1.
- **Aguinaldo accrual:** monthly 1/12 of base salary accrued to `231207 Accrued PTO` (paid in June + December - reversal handled manually, NOT by skill).
- **Infrastructure COGS split:** 30/70 (Apr+) - configurable via `--infra-cogs-pct` CLI flag for regression.
- **Egreso routing:** severance accounts (511175/611250) + EBITDA Adjustments department.
- **Special bonus (MIP):** auto-detected via pay code 50; emitted as separate Bonus line and excluded from Salaries DR.
- **Multiple raw file formats:** mixed `.xlsx` / `.xls` from provider. Bilingual headers (English row 11 in some, Spanish row 11 in others). Pay codes (row 12) are the canonical column key.

## Out-of-scope for v1

- Aguinaldo reversal in June + December (manual JE)
- License Provision / Vacation Salary Provision posting (tracked in Detalle de Provisiones audit file, not posted by skill)
- Funding Request file parsing (wire transfer reference only)
- FX conversion (JE posts in UYU directly)

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
