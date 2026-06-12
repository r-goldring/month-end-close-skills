---
name: bonus-accrual
description: >
  Build the monthly bonus accrual JE from FP&A's bonus accrual workbook
  (typically `20YY-MM_Bonus Accrual_vAccounting_vF.xlsx`, sent monthly by
  the FP&A Lead / FP&A). Produces one NetSuite Import Assistant CSV covering all
  5 subsidiaries (US, Canada, Netherlands, UK, Uruguay). Each entity is its
  own balanced JE (`USB2605`, `CANB2605`, `NLB2605`, `UKB2605`, `URYB2605`)
  that reverses 12/31 of the same year, so the full-year accrual builds up
  in 231170/231171 until December when actuals are booked. Use this skill
  when the user mentions: bonus accrual, bonus JE, monthly bonus, accrual
  CSV, bonus spreadsheet, the FP&A Lead bonus file, FP&A bonus file, 231170,
  231171, or drops a `*Bonus Accrual*.xlsx` file into
  `Monthly Bonus Accrual/`.
---

# Bonus Accrual Skill

## What this skill does

FP&A (the FP&A Lead) sends one Excel workbook each month with the bonus accrual
target for every department across every Acme Corp subsidiary. This skill
reads that workbook and produces a single NetSuite Import Assistant CSV
containing one balanced JE per subsidiary, each reversing on 12/31.

**Annual cycle:**
- **Months 1-11**: book monthly accrual per FP&A's file. Each JE reverses
  12/31 automatically.
- **December**: book actuals at year-end. Prior monthly accruals reverse
  out on 12/31 (NetSuite handles via the reversal date); actuals replace
  them in the 231170 / 231171 balance.
- **March (next year)**: bonuses pay out. Cash hits 231170 / 231171,
  clearing the year-end actuals.

## References

- Subsidiary paths, currencies: `../_shared/subsidiary-constants.md`
- NetSuite verification queries: `../_shared/netsuite-queries.md`
- CSV-only-upload policy: `../_shared/approval-required.md`
- Check-and-balance pattern: `../_shared/check-and-balance.md`
- ID lookup: `../_shared/id-lookup-guide.md`

## Per-skill files

- Script: `scripts/bonus-accrual/build_bonus_je.py`
- Source files (gitignored): `Monthly Bonus Accrual/{YYYY-MM} {Month Name} {Year}/`

## CSV-only policy

This skill writes a CSV that the accountant imports via NetSuite UI (Setup -> Import/
Export -> Import CSV Records -> Journal Entries). Each entity's External ID
becomes its own JE in the controller's Pending Approval queue. The skill
never calls `ns_createRecord` for these JEs (per CLAUDE.md rule #1).

---

## Workflow

### Phase 1 - Confirm inputs

1. **Period**: ask the user, or infer from context (e.g. "May 2026" -> `2605`).
2. **Source file**: default location is
   `Monthly Bonus Accrual/{YYYY-MM} {Month Name} {Year}/`. The script
   auto-discovers one `*Bonus Accrual*.xlsx` file there, preferring a
   filename containing `vF` (final) if multiple are present.
3. **Tab**: ask which one to use unless the user has already told you.
   - `Bonus - with CSM` (default; total = row 41) — books CSM bonus as
     part of the same JE.
   - `Bonus - excluding CSM` (total = row 40) — non-CSM only.
4. **Sanity-check the workbook structure**: open it with openpyxl
   (data_only=True) and verify the chosen tab exists and the Total row
   (41 with-CSM, 40 without-CSM) actually contains totals. If the
   structure has changed (rows shifted, new dept inserted), STOP and ask
   the accountant before running the build.

### Phase 2 - Verify prior month is in NetSuite

Before building the current month, confirm last month's accrual posted.
Run this SuiteQL with the prior period's external IDs (e.g. for May 2026,
check April: `USB2604`, `CANB2604`, `NLB2604`, `UKB2604`, `URYB2604`):

```sql
SELECT t.externalid, t.tranid, t.trandate, t.reversaldate, t.memo,
       a.acctnumber, tl.amount, BUILTIN.DF(tl.subsidiary) AS subsidiary,
       c.symbol AS currency
FROM transactionline tl
JOIN transaction t ON t.id = tl.transaction
JOIN account a ON a.id = tl.account
JOIN currency c ON c.id = t.currency
WHERE t.externalid IN ('USB{PREV}','CANB{PREV}','NLB{PREV}','UKB{PREV}','URYB{PREV}')
  AND a.acctnumber IN ('231170','231171')
ORDER BY t.externalid
```

Confirm:
- All 5 entities are present (or fewer if an entity skipped a month).
- Reversal date = `12/31/{year}` on each.
- 231170 + 231171 amounts match what we expected from last month's build.

If any entity is missing or has the wrong reversal date, flag it before
proceeding. The CSV upload can still happen, but the accountant needs to know
something is off in NS.

### Phase 3 - Run the build script

Folder mode (preferred — matches the drop-and-run workflow):

```bash
python scripts/bonus-accrual/build_bonus_je.py \
  --folder "Monthly Bonus Accrual/2026-05 May 2026" \
  --tab with-csm
```

Explicit mode (re-run, ad-hoc, non-standard path):

```bash
python scripts/bonus-accrual/build_bonus_je.py \
  --file "/path/to/2026-05_Bonus Accrual_vAccounting_vF.xlsx" \
  --period 2605 \
  --tab with-csm \
  --output "/path/to/output.csv"
```

The script:
- Auto-discovers the workbook in the folder (`*Bonus Accrual*.xlsx`,
  preferring `vF`).
- Infers period from the folder name (`2026-05 ...` -> `2605`).
- Writes the CSV alongside the input: `{YYYY-MM} Bonus Accrual JE Import.csv`.
- Omits zero-amount rows by default for a clean upload. Pass `--keep-zeros`
  to retain them (useful for row-by-row reconciliation against FP&A's
  workbook).
- Verifies each entity balances to zero in its own currency.
- Prints a per-entity summary (currency, bonus total, PT total).
- With `--json`, emits a machine-readable summary on stdout for audit-log
  integration.

### Phase 4 - Gut-check vs prior months

Before presenting, run a quick sanity check against the prior 2 months.

For each entity, compute:
- This month's bonus total vs. prior month's
- This month's bonus total vs. 2 months ago

Flag any entity where the absolute % change is >25% for the accountant to confirm.
Common causes that ARE OK:
- Headcount change (new hires / departures)
- Customer Success over-accrual reversal (negative amounts)
- CSM allocation change

Causes that are NOT OK:
- Workbook structure changed and rows shifted (verify dept mapping by
  re-reading rows 6-38 with their labels)
- Currency column index changed (verify cols 17-22 still match)

### Phase 5 - Present and confirm

Show the accountant a per-entity summary table:

| External ID | Subsidiary | Currency | Lines | Bonus | PT |
|---|---|---|---|---|---|
| USB2605 | Acme, Inc. | USD | XX | $XXX | $XXX |
| ... | ... | ... | ... | ... | ... |

Ask for confirmation before announcing the file is ready (per
`../_shared/check-and-balance.md`).

### Phase 6 - Save and audit

If saving to a different location than the script default, do that.
Otherwise, the CSV is already at:
`Monthly Bonus Accrual/{YYYY-MM} {Month Name} {Year}/{YYYY-MM} Bonus Accrual JE Import.csv`.

Tell the accountant:

> "CSV ready at `{path}`. Upload via NetSuite UI: Setup -> Import/Export
> -> Import CSV Records -> Journal Entries. Each of the 5 External IDs
> (USB{YYMM}, CANB{YYMM}, NLB{YYMM}, UKB{YYMM}, URYB{YYMM}) will create a
> separate JE in the controller's Pending Approval queue."

Append a `GENERATE_CSV` entry to `audit_log.json` with the period, output
path, and per-entity totals. After the accountant provides JE numbers post-import,
append a follow-up `POST_JE` entry with each `je_number` and
`netsuite_internal_id`.

---

## JE shape (per entity)

For each entity in the workbook with non-zero activity:

| Account | Direction | Amount | Department |
|---------|-----------|--------|------------|
| 511150 COGS Bonus | DR | per COGS dept | dept-specific |
| 611150 OpEx Bonus | DR | per OpEx dept | dept-specific |
| 231170 Accrued Bonus | CR | sum of all bonus DRs | blank |
| 511350 COGS Payroll Tax | DR | bonus_per_dept * pt_rate | dept-specific |
| 611450 OpEx Payroll Tax | DR | bonus_per_dept * pt_rate | dept-specific |
| 231171 PT Liability | CR | sum of all PT DRs | blank |

**Payroll tax rates by entity:**
- US: 9.00%
- Canada: 10.00%
- UK: 13.80%
- Netherlands: 0% (no PT lines emitted)
- Uruguay: 12.625%

**JE Date**: last day of period (e.g. `5/31/2026`).
**Reversal Date**: `12/31/{year}` on every line.
**External ID**: `{PREFIX}{YYMM}` — `USB2605`, `CANB2605`, etc.

---

## CSV format (NetSuite Style B)

Headers, in order:

```
External ID, Journal Entry Memo, Line Memo, Date, Reversal Date,
Subsidiary, Department, Account, Currency, Debit
```

- Positive `Debit` = DR; negative `Debit` = CR.
- Credit lines (231170, 231171) have blank Department.
- All lines for one External ID share the same Date, Reversal Date,
  Subsidiary, Currency, and Memo.

---

## Workbook layout (reference)

Both tabs share the same column structure. Column indices are 0-based.

| Col | Header |
|-----|--------|
| 0   | Department |
| 1   | GL account description |
| 2   | % allocation |
| 3-8 | YTD Balance (CAD, NL, PL, UK, US, URY) |
| 10-15 | YTD Accrued (same order) |
| **17** | **Current Month - CAD** |
| **18** | **Current Month - NL** |
| 19  | Current Month - Poland (ignored — always zero) |
| **20** | **Current Month - UK** |
| **21** | **Current Month - US** |
| **22** | **Current Month - URY** |

Row layout — `Bonus - with CSM` tab:

| Rows | Content |
|------|---------|
| 6-10 | COGS depts (511150) |
| 13-25 | OpEx depts (611150) |
| **26** | **Customer Success Mgmt** |
| 27-38 | Remaining OpEx depts |
| 41 | Total row |

Row layout — `Bonus - excluding CSM` tab: same as above but no row 26;
rows 27-38 shift up by one (becoming 26-37). Total row is 40 instead of
41. The script handles the offset automatically with `--tab without-csm`.

---

## Department to NS path mapping

COGS rows (511150 / 511350):

| Spreadsheet dept | NS Department path |
|---|---|
| Infrastructure | Engineering : Infrastructure |
| Professional Services | COGS : Professional Services |
| Consulting | COGS : Consulting |
| SRE | COGS : SRE |
| Managed Survey QA | COGS : Managed Survey QA |

OpEx rows (611150 / 611450):

| Spreadsheet dept | NS Department path |
|---|---|
| Infrastructure (OpEx) | Engineering : Infrastructure |
| Marketing | Sales & Marketing : Marketing |
| Revenue | Sales & Marketing : Revenue |
| SW Quality Engineering | Research & Development : Technology : SW Quality Engineering |
| New Business | Sales & Marketing : Revenue : New Business |
| Product | Research & Development : Product |
| General Mgmt and Growth | General & Administrative : General Mgmt and Growth |
| People | General & Administrative : People |
| Engineering | Research & Development : Technology : Engineering |
| Solutions Consulting | Sales & Marketing : Revenue : Solutions Consulting |
| Customer Enablement Ops | Sales & Marketing : Customer Enablement Ops |
| Customer Success | Sales & Marketing : Customer Success |
| Customer Success Mgmt | Sales & Marketing : Revenue : Customer Success Mgmt *(with-CSM tab only)* |
| Data Science Analytics | Research & Development : Technology : Data Science Analytics |
| Info Security Privacy | Research & Development : Technology : Info Security Privacy |
| Workforce Transformation | Sales & Marketing : Marketing : Workforce Transformation |
| Business Operations | Sales & Marketing : Revenue : Business Operations |
| Strategic Expansion | Sales & Marketing : Revenue : Strategic Expansion |
| Information Technology | General & Administrative : Information Technology |
| Sales Development | Sales & Marketing : Marketing : Sales Development |
| Revenue Enablement | Sales & Marketing : Revenue Enablement |
| GA | General & Administrative : GA |
| Legal | General & Administrative : Legal |
| Technology | Research & Development : Technology |
| EBITDA Adj. | EBITDA Adjustments |

The mapping lives in `scripts/bonus-accrual/build_bonus_je.py` under
`DEPT_ROWS_WITH_CSM`. If FP&A adds a new dept row, edit the script.

---

## Edge cases & notes

- **Negative amounts are valid** — e.g. Customer Success over-accrual
  reversals. Leave as-is; they become credits in the DR column.
- **Netherlands has no payroll tax** (rate = 0%). No 231171 line is
  emitted for NLB{YYMM}.
- **Poland column** (index 19) is always zero in the workbook and is
  ignored by the script. Poland bonus accrual, if any, is handled
  outside this skill.
- **Zero-amount rows are omitted by default** so the CSV is shorter and
  easier to review before upload. Use `--keep-zeros` if the accountant wants every
  dept row present for row-by-row reconciliation against FP&A's workbook
  (which is how the historical ~30 months of production files were
  built).
- **Rounding**: each line rounds to 2 decimal places with `ROUND_HALF_UP`.
  Cent-level discrepancies between the sum-of-rounded-parts and the
  spreadsheet's unrounded total are normal and within tolerance.
- **Filename variations from FP&A**: the script accepts any
  `*Bonus Accrual*.xlsx`. If there are multiple files, it prefers the one
  containing `vF` (final), otherwise the newest by mtime.
- **Workbook structure changes** (rows inserted, depts renamed): if the
  Total row no longer holds the column total, the dept mapping in
  `DEPT_ROWS_WITH_CSM` needs to be updated. The script will not detect
  this automatically — the accountant should spot-check the per-entity totals
  against the workbook's Total row before uploading.

---

## Year-end transition (December)

December is special. The skill builds a regular accrual in early December,
but FP&A also produces a year-end actuals workbook. The actuals JE is
**not** a reversing accrual — it's the real expense. Treatment:

1. Book the monthly Dec accrual as normal (reverses 12/31).
2. Book the year-end actuals as a separate JE. This is typically done
   manually outside this skill. If FP&A asks the skill to build it,
   confirm the date is 12/31, the reversal date is blank, and the
   memo names "year-end actuals" explicitly.

The accrual JE reverses on 12/31 (its built-in reversal date), and the
actuals JE replaces it. The net effect on 231170 / 231171 is the year's
true bonus expense, which clears when bonuses pay out in March.

---

## Customer-Payment / Other non-bonus JEs

Out of scope. This skill only builds bonus accruals using FP&A's
workbook. Year-end true-ups, prior-period corrections, and bonus payouts
are separate JEs handled outside this skill.
