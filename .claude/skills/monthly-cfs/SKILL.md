---
name: monthly-cfs
description: >
  Build the 7 monthly Cash Flow Statement workbooks for FP&A (US, BV, CAD,
  Poland, UK, UY + Consolidated with FX embedded) in Monthly CFS/{YYYY-MM}/.
  Copies the prior month's workbooks, regenerates the Income Statement and
  Balance Sheet tabs in NetSuite-export-identical format from SuiteQL data,
  rolls the CFS formulas, re-points the monthly-shifting IS references,
  derives manual cells by zeroing the check figures (highlighted yellow),
  drops in period-end FX rates, repoints the consolidated workbook's external
  links, and runs a full validation suite. Use this skill when the user
  mentions: CFS, cash flow statement, monthly CFS, CFS workbooks, FP&A cash
  flow, consolidated CFS, FX embedded, or asks to build/refresh the cash flow
  statements after close.
---

# Monthly CFS — Cash Flow Statement workbooks for FP&A

## What this skill produces

For target month `{YM}` (e.g. `2026-05`), in `Monthly CFS/{YM}/`:

| File | Contents |
|------|----------|
| `1. {YM} CFS Consolidated {M.D.YYYY} (FX embedded).xlsx` | Consolidated tab + 6 per-subsidiary USD tabs (external links to the files below) |
| `2. {YM} CFS BV {M.D.YYYY}.xlsx` | USD CFS, EUR CFS, Income Statement, Balance Sheet |
| `3. {YM} CFS CAD …` / `4. … Poland …` / `5. … UK …` / `7. … UY …` | same structure (local currency + USD CFS tabs) |
| `6. {YM} CFS US {M.D.YYYY}.xlsx` | single USD CFS tab + IS + BS |
| `_build_report_{YM}.json` | every derived manual cell, flag, and check |
| `CHECKS {YM}.md` | validation results |

Prep date in the filename = run date. the accountant reviews, then copies the files to
`G:\Shared drives\Finance Team\CFS\{YYYY}\{YM}\` (the consolidated workbook's
external links point there).

## Workflow

### Phase 1 — Pull (Claude + MCP, read-only)

Refresh `Monthly CFS/_data/` (this folder is gitignored — it holds raw GL
balances). Run the SuiteQL below, write each result straight to its JSON file
(`{ENT}_balances.json`, `coa.json`, `fx.json`), and verify with the zero-sum
check — every entity's full ledger must net to 0.00 at each month-end:

1. **Balances** — one `ns_runCustomSuiteQL` per entity (subsidiary IDs:
   US 2, CAD 3, BV 4, PL 6, UK 8, UY 13). MUST use `transactionaccountingline`
   (`SUM(NVL(tal.debit,0)-NVL(tal.credit,0))`), NOT transactionline foreign
   amounts (wrong for BS: mixes transaction currencies; off even for USD).
   Cumulative CASE columns at each month-end from prior-FY December through
   the target month:

   ```sql
   SELECT a.id, a.acctnumber, a.accttype,
     SUM(CASE WHEN ap.enddate <= TO_DATE('{prior FY end}','YYYY-MM-DD')
         THEN NVL(tal.debit,0)-NVL(tal.credit,0) ELSE 0 END) AS m_base,
     -- ... one column per month-end ...
     SUM(NVL(tal.debit,0)-NVL(tal.credit,0)) AS m_target
   FROM transaction t
   JOIN transactionaccountingline tal ON t.id = tal.transaction
   JOIN transactionline tl ON tal.transaction = tl.transaction
        AND tal.transactionline = tl.id
   JOIN account a ON tal.account = a.id
   JOIN accountingperiod ap ON t.postingperiod = ap.id
   WHERE tl.subsidiary = {sid} AND tal.posting = 'T'
     AND ap.enddate <= TO_DATE('{target month end}','YYYY-MM-DD')
   GROUP BY a.id, a.acctnumber, a.accttype
   ORDER BY a.acctnumber NULLS LAST, a.id
   ```

   Write to `{ENT}_balances.json` (keys: id, acctnumber, accttype,
   m2512/m2601/... per `tab_builders.MONTH_KEYS` — extend MONTH_KEYS when a
   new month is added). IS activity = adjacent-column differences; RE/NI are
   computed from the IS-account columns. Update `MONTH_KEYS` in
   `tab_builders.py` for each new month.

2. **FX** — `consolidatedexchangerate.currentrate`, fromsubsidiary in the 5
   foreign subs, tosubsidiary `Acme, Inc.` → `fx.json`. These match the
   workbook rates exactly (verified to 5+ decimals).

3. **COA** — full account list with parents → `coa.json` (refresh when new
   accounts appear).

### Phase 2 — Build subsidiaries

```
python scripts/monthly-cfs/build_subsidiary_cfs.py {YM}
```

Per entity: copies the prior-month workbook, rewrites IS/BS tabs, renames CFS
tabs, re-points IS references by account label (IS row positions shift every
month), updates period dates and note-row text, writes FX rates (cell
placement auto-detected), then derives manual cells:
- specific rules first (`cfs_config.MANUAL_CELL_RULES` — US accrued interest
  S14/S28 from BS 261277, term-loan payment W40 from Δ261270, interest-income
  note B54 from IS 711000),
- then per-column plugs that zero each column's Check Figure. The
  "Effect of Exchange Rate on Cash" plug (row 47) is **rounded to whole
  dollars** to match the accountant's manual style — this leaves a sub-dollar check
  residual exactly like his files; every other plug stays exact (check = 0).
  Every touched cell is yellow-filled and listed in the build report.

### Phase 3 — Build consolidated

```
python scripts/monthly-cfs/build_consolidated_cfs.py {YM}
```

Zip-level external-link repoint (shared-drive paths, month + filename swap,
source sheet names), tab renames, formula month rewrite, FX rates, manual
cells mirrored from the subsidiary workbooks. Tabs that carry pasted values
instead of live links (the accountant's pattern varies by month) are value-refreshed
from the subsidiary workbooks and flagged.

### Phase 4 — Validate

```
python scripts/monthly-cfs/validate_cfs.py {YM}
```

Writes `CHECKS {YM}.md`. Report PASS/WARN/FAIL in chat with manual-cell and
new-account flags up top. All checks passed on the 2026-04 historical rebuild
(51/51) and the May 2026 pilot (51 pass + the expected leasehold-write-off WARN).

Includes a **gross fixed-asset decrease (disposal / write-off) WARN**: a
one-off drop in a gross 141xxx account (excluding accumulated-depreciation
contras, and suppressing steady monthly amortization like ROU assets by
comparing this month's drop to last month's). When it fires, the capex plug
silently absorbed the disposal into investing — confirm the disposal-vs-capex
split (the disposal-loss formula F12 only captures the accumulated-depreciation
side unless the gross movement is added by hand that month, e.g. the May 2026
US leasehold write-off). Do NOT hardcode the gross term into F12 permanently —
it would mis-classify leasehold *additions*; the flag-and-judge approach keeps
the formula stationary and matches the accountant's files in every quiet month.

### Phase 5 — Compare (on demand)

```
python scripts/monthly-cfs/compare_workbooks.py {YM} {ENTITY}
```

Diffs the generated workbook against the accountant's manual one (auto-pairs by prep
date), annotating known-variance rows.

## Known UI-display variances (do NOT chase these)

The generated IS/BS tabs are **current-GL truth**. the accountant's NetSuite UI exports
apply display-layer adjustments that exist in no API or table (verified
empirically during development):

1. **IC AR/AP gross-up** — `121900`/`211900` UI values differ from GL by an
   equal-and-offsetting amount (open-item gross-up + dynamic revaluation of
   open IC balances). Net IC ties to GL; both sides shift together; the CFS
   is unaffected (both land in the Intercompany row).
2. **Undeposited-funds netting (US)** — UI shows `111100` as 0 and nets its
   balance (XXX,XXX.XX, constant since before 2026) into the `241xxx`
   deferred-revenue rows. Validation warns if the GL balance ever moves.
3. **Dynamic revaluation noise** — `Unrealized Gain/Loss`, `Net Income`,
   `Retained Earnings` can differ slightly from a UI export (the report
   revalues open foreign-currency AR/AP at display time; GL holds posted
   FxReval transactions instead).
4. **Equity presentation** — US `300220` vs `Retained Earnings` carry a
   constant X,XXX.XX offset of unknown origin (nets to zero inside equity,
   no transaction-table source, stable since Mar 2026).

A stale comparison caveat: workbooks exported weeks ago will not match
current GL where JEs were edited after the export (verify with
`t.lastmodifieddate > export date`). Only same-day comparisons are exact.

## Hard-won implementation facts

- IS export row set = accounts with nonzero activity in either displayed
  month; row positions move EVERY month, so CFS references into the IS tab
  (D&A rows etc.) are re-pointed by label, never carried.
- Whether an account renders as parent-header + own-posting row is a property
  of the Acme report definition (`_data/observed_parents.json`), not the COA.
- Two report layout families: US-style (US, CAD: Income/COS/Expense/
  Net Ordinary Income) and UK-style (BV, PL, UK, UY: Sales/Purchases/
  Overheads/Operating Profit). UK-style "Other Expenses" displays
  credit-positive and is ADDED in the net formula; missing sections become
  literal `0` placeholders in computed formulas (NetSuite behavior).
- BS tabs are template-driven from the prior month (row set is stable); new
  accounts are inserted in hierarchy order, yellow-flagged, and all formulas
  below shift. Columns D+ (the accountant's side calculations, e.g. the US fixed-asset
  delta column) are preserved untouched.
- `Retained Earnings` row = -(IS-account balance at prior FY end + RE-account
  311100); `Net Income` row = -(FY-to-date IS activity). FY = calendar year.
- Tab names drift month to month in the historical files; everything is
  resolved fuzzily (`tab_builders.find_tab`) and renamed to canonical.
- **External links survive the openpyxl round-trip via `external_links.finalize`,
  run as the LAST step of every build.** openpyxl rewrites `xl/externalLinks/*`
  on save (drops the XML declaration + mc/x14 namespaces Excel authored), which
  makes Excel pop a harmless "Repaired Records: external formula reference"
  prompt on open. `finalize` overwrites those parts with the pristine template
  bytes (subsidiary files: the dead 2023 legacy links restored verbatim;
  consolidated: same restore plus the month/filename repoint of the 6
  subsidiary links). Do NOT re-save with openpyxl after `finalize`.
- No NetSuite writes anywhere; no audit_log entries needed.

## Excel smoke test (the accountant, after each build)

Open each subsidiary workbook (formulas recalc, no #REF!), then the
consolidated workbook and accept "Update Links" (sources must be in the
shared-drive month folder, or temporarily next to the file locally).
