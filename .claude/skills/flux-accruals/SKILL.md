---
name: flux-accruals
description: >
  Budget-driven monthly accrual + reclass identification. Uses FP&A's Vendor Budget
  v*.xlsx as the source-of-truth for "what each vendor SHOULD post each month and
  where," compares to NetSuite actuals + BillFlow open bills, and writes a single
  Accrual_Reclass_Candidates_{YYYY-MM}.xlsx for the accountant to review in Excel. Approved
  rows produce up to 3 JE Import CSVs: Accruals, Dept Reclass, Software Reclass.
  All JEs flow through NetSuite UI Import Assistant -> Pending Approval (no MCP
  posting per CLAUDE.md rule 1). Use this skill when the user mentions: accruals,
  identify accruals, missing vendors, month-end accruals, accrue for vendor, dept
  reclass, software reclass, vendor in wrong department, or asks to check what
  needs an accrual before close.
---

# Flux Accruals & Reclass Skill

## What this skill does

Replaces the previous clunky "zero-month above threshold" identification with a
budget-driven scan. For each closing month:

1. Loads FP&A's Vendor Budget v*.xlsx (221 vendors, monthly budgets, expected
   department, expected GL account).
2. Pulls NetSuite actuals for the closing month + 3 trailing months across the
   four flux-detail reports (537, 540, 542, 721).
3. Optionally merges BillFlow open bills if the accountant dropped a CSV under
   `Monthly Flux Analysis/{YYYY}/{YYYY-MM}/BillFlow Export/`.
4. Computes one of four signals per budget vendor:
     - `missing_budgeted` (budget > 0, actual = 0)
     - `partial_below_baseline` (actual < 50% budget, gap > category threshold)
     - `over_budget_review` (actual > 120% budget; informational)
     - `unbudgeted_vendor` (actual posted but no budget row; informational)
   BillFlow matches override with `billcom_confirmed` (HIGH confidence).
5. Computes dept-drift rows (vendor's actual dept != budget dept; non-software
   accounts only).
6. Computes software reclass rows (671xxx + 511425 activity where actual dept
   != budget dept). Primary use case: monthly amortization landing in the
   default dept. Source-agnostic - works for vendor bills AND amortization JEs.
7. Writes `Accrual_Reclass_Candidates_{YYYY-MM}.xlsx` with 6 working tabs +
   3 audit tabs. Column A is a Y/N/EDIT dropdown.
8. After the accountant marks column A, emits up to 3 JE Import CSVs.

## References

- Vendor Budget schema: `references/vendor-budget-schema.md`
- Signal algorithm: `references/signal-rules.md`
- Accrual thresholds: `references/accrual-thresholds.md`
- Accrual account map: `references/accrual-account-map.md`
- NetSuite report IDs and SuiteQL: `../_shared/netsuite-queries.md`
- Subsidiary constants: `../_shared/subsidiary-constants.md`
- ID lookup: `../_shared/id-lookup-guide.md`
- CSV-only-upload policy: `../_shared/approval-required.md`

## Per-skill files

- Scripts: `scripts/flux-accruals/`
  - `load_vendor_budget.py` - FP&A file loader
  - `parse_reports.py` - parse the 4 ns_runReport JSONs
  - `load_billcom_export.py` - BillFlow CSV + NS validation helper
  - `candidate_workbook.py` - openpyxl writer for the review workbook
  - `build_candidates.py` - main orchestrator (CLI)
  - `refresh_pivot_template.py` - copy master pivot template into month folder + populate data tabs
  - `generate_je_csvs.py` - emit JE Import CSVs from approved rows
- FP&A budget: `Monthly Flux Analysis/Vendor Budget v*.xlsx`
- Master pivot template (never modified): `Monthly Flux Analysis/Flux Pivot Template.xlsx`
- Per-month workspace: `Monthly Flux Analysis/{YYYY}/{YYYY-MM}/`
  - `_cache/{537,540,542,721}.json` - NS report payloads
  - `_cache/billcom_already_in_ns.json` - VendBills already exported
  - `BillFlow Export/TransactionListByExportStatus*.csv` - the accountant drops here
  - `Accrual_Reclass_Candidates_{YYYY-MM}.xlsx` - review workbook (built)
  - `Flux Pivot Template {YYYY-MM}.xlsx` - per-month pivot copy (built)

---

## Step 1 - Determine period

Ask the accountant for the closing month if not stated. Convert to YYYY-MM. Derive prior 2
months for trailing context: M, M-1, M-2.

```
M    = "2026-04"   (Apr 2026)
M-1  = "2026-03"
M-2  = "2026-02"
```

Set up the month workspace:

```
Monthly Flux Analysis/2026/2026-04/
  _cache/                          # NS report JSONs go here
  BillFlow Export/                 # the accountant drops the CSV here (optional)
```

If `_cache/` doesn't exist, create it.

---

## Step 2 - Pull NetSuite reports (one call per report, 3-month range)

For each of the four reports, call `ns_runReport` with a single date range
covering M-2 through M. Save the raw JSON to `_cache/{report_id}.json`.

| Report ID | Category   |
|-----------|------------|
| 537       | COGS       |
| 540       | Contractors|
| 542       | ProfFees   |
| 721       | Software   |

```
ns_runReport(
    reportId: 537,
    subsidiaryId: -2,
    startDate: "2026-01-01",     # first day of M-3
    endDate:   "2026-04-30"      # last day of M
)
```

Always use `subsidiaryId: -2` (consolidated USD per CLAUDE.md rule 5).

After each call, save the response JSON exactly as returned:

```python
import json
with open("Monthly Flux Analysis/2026/2026-04/_cache/537.json", "w") as f:
    json.dump(response, f)
```

The Python parser handles the `reportData` structure; do not pre-process.

**4 calls total.** This is intentionally one-call-per-report instead of one-per-month
- much more efficient than the old skill's 12 calls.

---

## Step 3 - BillFlow validation (optional but standard)

Check whether the accountant dropped a BillFlow export at:
```
Monthly Flux Analysis/{YYYY}/{YYYY-MM}/BillFlow Export/TransactionListByExportStatus*.csv
```

If present, validate the bills against NetSuite to filter out bills that DID
already export. Run this SuiteQL once for the closing period:

```sql
SELECT t.id, t.tranid, TO_CHAR(t.trandate, 'YYYY-MM-DD') AS trandate,
       v.companyname AS companyname, t.foreignamount
FROM transaction t
LEFT JOIN vendor v ON t.entity = v.id
WHERE t.type = 'VendBill'
  AND t.trandate >= TO_DATE('{M_first_day}', 'YYYY-MM-DD')
  AND t.trandate <= TO_DATE('{M_last_day}', 'YYYY-MM-DD')
  AND t.voided = 'F'
```

Save the result to `_cache/billcom_already_in_ns.json` as a list of
`{tranid, companyname, trandate}` objects:

```python
import json
ns_bills = ns_runCustomSuiteQL(query=...)["rows"]
with open(".../billcom_already_in_ns.json", "w") as f:
    json.dump(ns_bills, f)
```

If no BillFlow CSV exists, skip this step entirely.

---

## Step 4 - Build the candidate workbook + refresh the pivot template

Run both scripts. Working directory must be the repo root.

```bash
python scripts/flux-accruals/build_candidates.py 2026-04
python scripts/flux-accruals/refresh_pivot_template.py 2026-04
```

`refresh_pivot_template.py` copies the master `Flux Pivot Template.xlsx`
into the month folder as `Flux Pivot Template {YYYY-MM}.xlsx`, populates the
hidden data tabs from the cache, and updates the pivot source ranges. the accountant
opens the per-month copy in Excel and clicks **Data > Refresh All** to rebuild
the pivot tables themselves (openpyxl can't execute the calculation; only the
source-range update). The master is never modified.

The script:
- Loads the latest `Vendor Budget v*.xlsx` from `Monthly Flux Analysis/`.
- Reads all 4 cache JSONs and parses transactions.
- Reads the BillFlow CSV (if present) and validates against
  `billcom_already_in_ns.json`.
- Computes signals, dept-drift, software-reclass.
- Writes `Monthly Flux Analysis/2026/2026-04/Accrual_Reclass_Candidates_2026-04.xlsx`.
- Prints a summary to stdout.

The workbook has 9 tabs:

| Tab | Purpose |
|-----|---------|
| Summary | Counts, totals, instructions |
| Accruals_Software | Software accrual candidates |
| Accruals_Contractors | Contractor accrual candidates |
| Accruals_ProfFees | Professional fees accrual candidates |
| Accruals_COGS | COGS accrual candidates |
| Reclass_Dept_Drift | Non-software vendors in wrong dept |
| Reclass_Software | Software-account activity in wrong dept (incl. amortizations) |
| Unmatched_Actuals | Vendors in actuals but not in FP&A budget (FP&A follow-up) |
| Billcom_AlreadyInNS | BillFlow rows already exported (filter validation) |
| Billcom_Unmatched | BillFlow vendors with no fuzzy match in budget |

---

## Step 5 - Hand off to the accountant for review

Send the accountant a message in this exact shape (substitute the real path and counts):

```
Candidate workbook built:
  Monthly Flux Analysis/2026/2026-04/Accrual_Reclass_Candidates_2026-04.xlsx

Summary:
  Software:     {n} candidates,  ${total:,.2f}
  Contractors:  {n} candidates,  ${total:,.2f}
  ProfFees:     {n} candidates,  ${total:,.2f}
  COGS:         {n} candidates,  ${total:,.2f}
  Reclass_Dept_Drift:  {n} candidates
  Reclass_Software:    {n} candidates

Open the workbook. On each Accruals/Reclass tab, mark column A:
  Y     = approve as suggested
  N     = reject
  EDIT  = override; set 'edited_amount' (or 'edited_target_dept' for reclass)

Save when done. Reply 'build' and I'll emit the JE Import CSVs.

(Optional) If you also want to refresh the Flux Pivot Template for visual
verification, run /flux-analysis-workbook in a separate session - it uses the
same NetSuite reports.
```

---

## Step 6 - Emit JE Import CSVs

After the accountant replies that he's marked the workbook:

```bash
python scripts/flux-accruals/generate_je_csvs.py 2026-04
```

The script:
- Reads only Y/EDIT rows from each tab.
- Emits up to 3 CSVs to `Monthly Flux Analysis/2026/2026-04/`:
  - `2026-04 Accruals JE Import.csv`
  - `2026-04 Dept Reclass JE Import.csv`
  - `2026-04 Software Reclass JE Import.csv`
- Appends GENERATE_CSV entries to `audit_log.json` per CLAUDE.md rule 8.

Tell the accountant:

> "CSVs ready. Upload via NetSuite UI:
>  Lists -> Import Assistant -> Transactions -> Journal Entry
> Each JE will land in the controller's Pending Approval queue.
> Once posted, send me the JE numbers and I'll log them."

After he reports the JE numbers, append a follow-up entry to `audit_log.json`
for each, with both `je_number` (tranid) and `netsuite_internal_id`. Per
CLAUDE.md rule 8, ALWAYS query for the internal ID:

```sql
SELECT id FROM transaction WHERE tranid = 'JE20XXX'
```

---

## Step 7 - Verification (one-time, when first running on a new month)

Before handing off the workbook, sanity-check:

1. Did all 4 ns_runReport calls succeed? Verify each `_cache/{id}.json`
   contains a `reportData` dict with non-empty rows.
2. Does the `unmatched_actuals` count look reasonable (< 30 vendors)? A high
   count suggests the FP&A file is stale - tell the accountant.
3. Does `Billcom_AlreadyInNS` have 0 or near-0 rows? If many, the accountant's export
   filter is misbehaving.
4. Spot-check one large `missing_budgeted` candidate: does the vendor truly
   have $0 in NS for the period? Confirm via:

```sql
SELECT SUM(tl.debitforeignamount - tl.creditforeignamount) AS total
FROM transaction t
JOIN transactionline tl ON t.id = tl.transaction AND tl.mainline = 'F'
JOIN account a ON tl.account = a.id
LEFT JOIN vendor v ON t.entity = v.id
WHERE a.acctnumber = '{account_number}'
  AND v.id = {vendor_id}
  AND t.trandate >= TO_DATE('{M_first_day}', 'YYYY-MM-DD')
  AND t.trandate <= TO_DATE('{M_last_day}', 'YYYY-MM-DD')
  AND t.voided = 'F'
```

---

## What NOT to do

- Do NOT call `ns_createRecord` for `journalentry` or
  `advintercompanyjournalentry`. CSV upload only per CLAUDE.md rule 1.
- Do NOT skip the `subsidiaryId: -2` filter on report calls; it gives
  consolidated USD across all subs.
- Do NOT re-pull reports if the cache already has fresh JSONs for the period.
  If the accountant asks to re-run, delete `_cache/` first or pass a `--force` flag
  (not yet implemented; for now, manual cache clear).
- Do NOT modify the Vendor Budget file. FP&A owns it; the accountant replaces the file
  when a new version arrives.
- Do NOT touch `flux-analysis-workbook` from this skill. They are
  complementary; run them separately.
