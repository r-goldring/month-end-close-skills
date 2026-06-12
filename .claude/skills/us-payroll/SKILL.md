---
name: us-payroll
description: >
  Process US bi-weekly ADP payroll into a balanced NetSuite journal entry and
  post it directly via the NetSuite MCP. Use this skill when the user mentions:
  US payroll, ADP payroll, payroll JE, payroll journal entry, payroll import,
  payroll mapping, or drops a raw ADP export (GXXXXXX000060.xls/xlsx) into
  Monthly Payroll/Pay Runs/US/{MM.DD.YYYY}/.
---

# US Payroll JE Skill

## Overview

Processes the US ADP payroll export into a NetSuite journal entry. The Python
script [`payroll_mapper.py`](../../../scripts/us-payroll/payroll_mapper.py)
handles all extraction, mapping, and aggregation; this skill previews the JE,
generates a CSV for upload via the NetSuite UI Import Assistant (per the
2026-05-01 SoD policy in [`_shared/approval-required.md`](../_shared/approval-required.md)),
and appends the audit log. the accountant reports the resulting `JE#####` tranid back
after his manual import; the skill records that tranid in the audit log.

**Cadence:** Bi-weekly — two pay runs per month (typically the 15th and the
last day of the month). Each pay run gets its own folder and its own JE.

**Subsidiary:** `Acme Holdings : Acme, Inc.` (consolidated US).

**Script:** `scripts/us-payroll/payroll_mapper.py`
**Preflight:** `scripts/us-payroll/check_mappings.py`
**Mapping files (authoritative knowledge base):**
- `scripts/us-payroll/US Payroll Department Mapping File.csv` — cost center to department
- `scripts/us-payroll/Payroll_Mapping_with_GL_Accounts (final).csv` — payroll code to GL account (Opex and COGS sides)

## Step 1 — the accountant drops the raw file

the accountant drops the ADP export into `Monthly Payroll/Pay Runs/US/{MM.DD.YYYY}/`
where `{MM.DD.YYYY}` is the pay date (e.g., `04.15.2026`, `04.30.2026`).

The raw file has the ADP job-number naming pattern (`GXXXXXX000060.xls` or
`.xlsx`). It contains six-to-seven sheets named `Sheet1`, `Sheet2`, etc., each
with the ADP employee block layout. The mapper auto-discovers the file in the
folder as long as only the raw file is present (anything with `Backup` in its
name, `Review_Report_*`, or `~$*` temp files is ignored).

## Step 2 — Preflight mapping check (advisory, not blocking)

Always run the preflight before the mapper:

```bash
cd scripts/us-payroll
python check_mappings.py "../../Monthly Payroll/Pay Runs/US/{MM.DD.YYYY}/"
```

Exit 0 = every cost center and payroll code in the raw file is mapped.
Exit 1 = new cost centers or codes were found; surface the report to the accountant.

**US-specific behavior:** Unlike Germany/Netherlands, exit 1 is NOT a hard halt
for US. the accountant sometimes cannot tell where a brand-new payroll code should be
coded until he sees which side of the JE is unbalanced. If the preflight reports
unmapped items, show the accountant the report and ask him to pick:

1. **Add mappings and re-run preflight.** Open the appropriate CSV, add a row,
   save, re-run `check_mappings.py`. Repeat until clean. (Default recommendation.)
2. **Proceed to the mapper anyway.** Any unmapped codes will emit rows with
   `Account = UNMAPPED_ACCOUNT` in the JE. Do NOT post an unbalanced JE —
   instead, use the debit/credit delta in the preview to triangulate where each
   new code's missing side should go, then go back to option 1.
3. **Abort** this pay run.

Refuse to post a JE that contains any `UNMAPPED_ACCOUNT` rows. Always route
through option 1 before posting.

## Step 3 — Run the mapper

```bash
cd scripts/us-payroll
python payroll_mapper.py "../../Monthly Payroll/Pay Runs/US/{MM.DD.YYYY}/"
```

The mapper:
- Auto-discovers the raw file in the folder.
- Parses the pay date from the folder name (`MM.DD.YYYY`).
- Applies all routing rules (see "Key mapper behaviors" below).
- Writes `{MM.DD.YYYY} US Payroll Backup.xlsx` into the pay-run folder. The
  backup contains every raw ADP sheet copied verbatim plus one new tab
  `{MM.DD.YYYY}_US_Payroll_JE (pivot)` — that pivot tab is the JE the accountant
  reviews, mirroring the format he has used historically.
- Prints summary: row count, total debits, total credits, balance status,
  and any unmapped cost centers or payroll codes.

For in-skill use, import `generate_je(folder_path)` from the script — it
returns the JE rows, totals, output path, and unmapped lists directly.

Pass `--verify-suffix` (or `verify=True`) to write
`... US Payroll Backup (verify).xlsx` instead of overwriting an existing
backup. Use this when dry-running a pay run whose JE has already been posted.

## Step 4 — Preview the pivot and PAUSE

Render the `{MM.DD.YYYY}_US_Payroll_JE (pivot)` tab in chat, grouped by
department then by account within each department (liabilities at the bottom
with blank Department). Include at minimum:

```
US PAYROLL PIVOT — {MM.DD.YYYY}
Total lines:   XXX
Total debits:  $X,XXX,XXX.XX
Total credits: $X,XXX,XXX.XX
Balanced:      YES  (or NO — delta $X,XXX.XX)

Top 5 debit accounts:
  611100 Salaries:  $XXX,XXX.XX
  [...]

Liability (credit) lines:
  231200 Payroll Liability:     $-X,XXX,XXX.XX
  231201 HSA Payable:           $-X,XXX.XX
  231202 FSA Payable:           $-X,XXX.XX
  231250 401K Payable:          $-XXX,XXX.XX
  231350 Payroll Tax Liability: $-XXX,XXX.XX

Unmapped cost centers: 0 (or list)
Unmapped codes:        0 (or list)
```

If the JE is unbalanced or contains `UNMAPPED_ACCOUNT` rows, do not proceed —
go back to Step 2.

**This is a PAUSE point.** The pivot JE is roughly 90% of the final JE that
posts to NetSuite. the accountant's assistant controller manually adds a `Reclasses` tab
to the backup workbook each pay run; the actual posting is the pivot plus
those reclass adjustments. Do NOT post yet.

Tell the accountant something like:

> Pivot ready and balanced ($X,XXX,XXX.XX, N lines). Backup written to
> `{MM.DD.YYYY} US Payroll Backup.xlsx`. Pausing for the assistant controller
> to add the `Reclasses` tab. When she's done, ask me to "apply reclasses and
> post for {MM.DD.YYYY}".

## Step 5 — Apply reclasses (Phase 2)

Triggered when the accountant says something like *"apply reclasses and post for
{MM.DD.YYYY}"*. The `Reclasses` tab in the backup workbook drives this step.

### Reclasses tab format

Sections are identified by a non-empty cell in column A:

- **Severance** (`"Severance"` in column A) — header row
  `Department | Name | Severance | Taxes` (cols A–D), then one row per
  severance employee with their dept short-name, name, severance amount, and
  employer-tax amount. Subtotal rows (no dept/name, only amounts) are ignored.
  An optional Infrastructure continuation block in cols F–I may show a
  pre-computed COGS/OpEx split — the skill ignores it and computes the split
  itself using the current 30/70 rule.
- **Borland to EBITDA** (substring `"borland"` + `"ebitda"` in column A) — JB
  is the only board-member reclass and is hard-coded. Subsequent rows:
  `[B] Component label | [C] Amount`. Valid component labels (case-insensitive
  substring match): `Salary`, `Medical Benefits`, `401K Match`,
  `Other Benefits`, `Taxes`.

**Any other section header in column A halts the skill.** Echo the section
verbatim and ask the accountant for instructions. Charles-Newman-style one-off dept
transfers are NOT supported and must be handled manually.

### Run the apply step

```python
from payroll_mapper import apply_reclasses
result = apply_reclasses("Monthly Payroll/Pay Runs/US/{MM.DD.YYYY}/")
```

`apply_reclasses` does the following:

1. Reads the existing pivot tab from the backup workbook.
2. Reads the `Reclasses` tab and dumps its raw contents back to chat (so any
   format drift from the assistant controller is caught early).
3. Parses the two known section types; halts on any unknown section.
4. Builds reclass entries:
   - **Severance tax reclass** — for each severance employee, reduce home-dept
     `611450 OpEx Payroll Taxes` (or `511350 COGS Payroll Taxes` for COGS
     depts) by the tax amount and add the same amount to `EBITDA Adjustments`
     `611250 OpEx Severance` (or `511175 COGS Severance`). Infrastructure
     splits 30% to COGS / 70% to OpEx on both source and destination sides.
     Line memo: `"US PAYROLL - PR Tax Severance ({Home Dept full name})"`,
     e.g., `(Customer Success Mgmt)`, `(Professional Services)`,
     `(Infrastructure - COGS)`, `(Infrastructure - Opex)`. 231350 Payroll
     Tax Liability is NOT touched — only the debit side moves between expense
     accounts.
   - **JB board member reclass** — for each component, reduce
     `General & Administrative : GA`'s row on the matching account by the
     amount and add an equal positive entry under `EBITDA Adjustments` with
     line memo suffixed `"(JB)"` — e.g., `"US PAYROLL - Salary (JB)"`,
     `"US PAYROLL - PR Tax (JB)"`. **Sanity check** (non-blocking): if any JB
     component amount exceeds the GA pivot row's amount on that account, halt
     and ask the accountant before applying — prevents typos creating negative GA
     balances.
5. Concatenates reclass entries with pivot rows, re-runs the
   (Department, Account, Memo, Subsidiary) groupby, drops zero-rounded rows.
6. Writes a `JE with Reclasses` tab to the backup workbook (alongside the
   pivot — pivot stays untouched so the accountant can verify code mappings separately).
7. Returns `{status, totals, final_rows, reclasses, raw_reclasses_dump, ...}`.

## Step 6 — Preview the JE with Reclasses and confirm

Render the JE with Reclasses tab in chat with:

```
US PAYROLL JE WITH RECLASSES — {MM.DD.YYYY}
Pivot rows: XXX → Final rows: XXX (+N reclass entries)
Total debits:  $X,XXX,XXX.XX
Total credits: $X,XXX,XXX.XX
Balanced:      YES

Reclasses applied:
  Severance tax: N entries across M employees (depts: ...)
  Borland (JB):  N components, total $X,XXX.XX moved GA → EBITDA Adjustments
  Unknown sections: 0

Account-level deltas vs pivot:
  611450 Payroll Taxes:  $-X,XXX.XX (severance tax moved to severance accts)
  611250 Severance:      $+X,XXX.XX
  611100 Salaries (GA):  $-9,425.00 (JB)
  611100 Salaries (EBITDA): $+9,425.00 (JB)
  [...]
```

Then prompt verbatim:

> **"Review the above carefully. Post to NetSuite? Type 'yes' to confirm or anything else to cancel."**

Only the exact string `yes` proceeds to posting. Refuse to post if any
unknown sections, JB warnings, or unresolved severance employees remain.

## Step 7 — Generate CSV for NetSuite UI upload

**MCP-based JE posting is suspended** (see [`_shared/approval-required.md`](../_shared/approval-required.md)). Generate a CSV alongside the Backup.xlsx.

1. Build the JE rows from the **JE with Reclasses** sheet (NOT the pivot). One row per line: positive = debit, negative = credit (absolute value into the Credit column). Liability lines have no department.
2. Write `Monthly Payroll/Pay Runs/US/{MM.DD.YYYY}/{MM.DD.YYYY} US Payroll JE Import.csv` with columns: `Date, Journal Entry Memo, Account, Debit, Credit, Line Memo, Subsidiary, Department`. Date `M/D/YYYY`. Empty cells stay empty (no zeros). Quote subsidiary path.
3. Append to `audit_log.json` with `action: "GENERATE_CSV"`, `description`, source files, and the reclasses_applied counts. (After the accountant imports and reports the JE number, append a follow-up entry with `je_number` and `netsuite_internal_id`.)

**DO NOT** call `ns_createRecord` for journal entries.

## Step 8 — Pre-upload gut-check review (auto)

Immediately after writing the CSV in Step 7, **run the gut-check** before
telling the accountant to upload. See [`gut-check/SKILL.md`](../gut-check/SKILL.md) for
the full orchestration. Short version:

1. Call `payroll_gut_check.run_gut_check(folder_path, audit_log_path,
   suiteql_runner=None)` to get the priors-search SQL.
2. Run that SQL via `mcp__claude_ai_NetSuite__ns_runCustomSuiteQL` to find the
   prior 2 same-skill JEs in NetSuite (memo + subsidiary + before-current-date).
3. For each prior, run `payroll_gut_check.build_line_fetch_sql(tranid, sub)`
   via SuiteQL to pull its lines.
4. Call `payroll_gut_check.analyze(...)` + `format_chat_report(...)` and
   `write_workbook_tab(...)` to produce the PASS/WARN/FAIL report and append
   a `Gut-Check {YYYY-MM-DD HH_MM}` tab to the backup workbook.
5. **Gate the upload instruction on FAIL severity:**
   - If `result.has_fail`: tell the accountant **"DO NOT UPLOAD. Fix the CSV (or
     Reclasses tab) per the FAIL findings, then re-run /gut-check after
     edits."** — do NOT print the upload instructions.
   - Else: tell the accountant **"Safe to upload. Lists -> Import Assistant ->
     Transactions -> Journal Entry -> upload CSV. WARNs are advisory."**

The report covers: aggregate variance per account family (Salaries, Payroll
Tax, Health/Other Benefits, 5 liability accounts), per-(dept, account)
variance for every line, canonical sign violations on gross-expense accounts,
prior-2 sign flips on net accounts (Health/Other Benefits), severance
routing to EBITDA Adjustments, and JB EBITDA reclass drift vs prior 2.

## Step 10 — Report to the accountant

Confirm the post in chat using the `JE#####` tranid (not the internal ID):

> Posted `JE#####` — {MM.DD.YYYY} US Payroll, N lines, $X,XXX,XXX.XX balanced.

## Key mapper behaviors

- **Infrastructure 30/70 COGS/OpEx split.** Any payroll line routed to the
  Infrastructure department is split: 30% to COGS (511xxx) and 70% to OpEx
  (611xxx). Changed from 80/20 on 2026-04-22; the script enforces 70% OpEx.
- **Severance.** Code `SEVR` (and any GL named "Severance") overrides the
  posting department to `EBITDA Adjustments` and routes to `511175 COGS
  Severance` or `611250 OpEx Severance` based on the employee's home-dept
  COGS/OpEx nature. The Line Memo is suffixed with the home department's
  short name in parentheses (e.g., `US PAYROLL - Severance (Professional Services)`)
  so the aggregation produces one row per home-dept × account, preserving
  visibility into who the severance was for. Severance-related employer
  payroll taxes flow into regular Payroll Taxes (611450 / 511350), not into
  the Severance account.

  **Infrastructure-origin severance** (since 2026-05-06): the principal also
  splits 30% COGS / 70% OpEx across `511175` and `611250`, mirroring the
  regular Infrastructure 30/70 split. Memos suffix `" - Opex"` and
  `" - COGS"` (e.g., `US PAYROLL - Severance (Infrastructure - Opex)` /
  `(Infrastructure - COGS)`), matching the existing PR-Tax-Severance pattern.
  This makes Infrastructure severance internally consistent with its
  hybrid-cost-center treatment in normal months. Pre-2026-05-06 pay runs
  (JE#####, JE#####, JE#####) booked Infra severance principal 100% to
  611250 OpEx — those are not retroactively adjusted.
- **Health Benefits / Other Benefits sign inversion.** Employee-Deds amounts
  mapped to these accounts are inverted (positive withholdings become
  negative credits; negative reversals become positive debits). The rule is
  scoped to the Employee Deds category only — Earnings-category codes
  routed to these accounts (`HSAER`, `MEDCR`) stay as positive debits since
  they represent employer contributions.
- **Other Benefits — employer contribution codes.** `HSAER` (Earnings-side),
  `MEDCR`, `MEDS`, `MPS` all route to 611400 / 511300 Other Benefits as
  positive debits. The HSAER Employee-Deds-side credits 231201 HSA Payable
  separately.
- **Payroll Tax Liability (231350) uses ADP footer totals, not per-code sum.**
  The mapper dynamically locates the ADP `Report Totals:` row on the last
  sheet and reads the Employee Taxes and Employer Tax Exp subtotals. The
  credit on 231350 = `-(ADP Employee Taxes + ADP Employer Tax Exp − NYPFLEE
  − NYSDIEE)`. NYPFLEE and NYSDIEE are Employee-Tax codes that ADP folds
  into both totals, so we back them out to avoid double-counting. This also
  means individual employee-tax codes (USFIT, USMEDEE, state SITs, local
  taxes like CODENEE / OH1585) are not routed per-code.
- **Payroll Tax Expense (611450 / 511350) per department.** Sum of the
  employee's Employer Tax Exp codes per cost-center, with NYPFLEE and
  NYSDIEE for that cost-center backed out. Employer-tax codes route
  uniformly to 611450 / 511350 by department without per-code GL mapping —
  new or rarely-seen employer tax codes (CODENER, NMWCER, etc.) are
  handled automatically.
- **401K Payable consolidation.** Employer 401K match expense (611350 /
  511250) is consolidated as a credit to 231250 401K Payable, alongside the
  employee 401K deductions already mapped to 231250.
- **231200 Payroll Liability** = `-(ADP Total Net + WBNK1)`. The mapping
  for `WBNK1` (Type: Liability) routes it directly to 231200, so the
  groupby naturally produces the combined credit.
- **Ignored codes.** `DPIID`, `DPIIM`, `DPIIV`, `GTL`, `GIFT` are never
  extracted.
- **Employer Ded Exp filter.** Of the Employer Ded Exp category, only
  `401KM` is extracted; all other codes in that category are intentionally
  dropped.

## Pre-flight / JE sanity checklist

Before typing `yes`:
- Debits equal credits (preview prints `Balanced: YES`).
- No `UNMAPPED_DEPT` or `UNMAPPED_ACCOUNT` entries remain.
- Severance (if present) is in `EBITDA` with `511175` or `611250`.
- Infrastructure lines appear on both 511xxx and 611xxx accounts
  (approximately 30% / 70% split of the total).
- All five liability accounts are negative (credits): 231200, 231201, 231202,
  231250, 231350.

## Subsidiary

`Acme Holdings : Acme, Inc.`
