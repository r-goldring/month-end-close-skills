# Accrual & Reclass Signal Rules

How `build_candidates.py` decides which budget vendors and which actuals are
worth the accountant's review. Each candidate row carries one signal so it's clear *why*
it surfaced.

## Accrual signals (one per budget vendor, in-scope category only)

For each `BudgetRow` in the FP&A file with `category in {Software, Contractors,
ProfFees, COGS}`, compute:

```
budget_amt  = budget_row.budget_for(month)
actual_amt  = sum of all NS actuals for this vendor in this period and category
gap         = budget_amt - actual_amt
threshold   = THRESHOLDS[category]   # see accrual-thresholds.md
```

Then assign signal:

| Signal | Trigger | Suggested Amount | Reason in workbook |
|---|---|---|---|
| `missing_budgeted` | budget > 0 AND actual < $1 | `budget_amt` | vendor has a budget but didn't post |
| `partial_below_baseline` | actual < 50% of budget AND gap > threshold | `gap` | vendor posted but materially short |
| `over_budget_review` | actual > 120% of budget AND (actual - budget) > threshold | `0` (informational only) | flag for variance review, NOT an accrual |
| `unbudgeted_vendor` | actual posted, no budget row | `0` (informational only) | budget file is missing the vendor |

Vendors whose actuals are within 50%-120% of budget get **no row** (they're fine).

## BillFlow override (highest precedence)

If BillFlow export contains an open bill (after NS validation) for a budget
vendor, the candidate row is overridden:
- `signal` -> `billcom_confirmed`
- `confidence` -> `HIGH`
- `suggested_amount` -> sum of matching BillFlow bills (replaces the
  budget-vs-actual estimate)
- `notes` includes invoice numbers

This applies even if no other accrual signal would have fired (e.g., vendor
has actuals matching budget but ALSO has an open bill — we still want to
accrue the open bill since it represents work performed but not yet billed).

Match key: normalized vendor name (exact, then SequenceMatcher >= 0.85).

## Confidence levels

| Confidence | When |
|---|---|
| HIGH | `billcom_confirmed`, OR `missing_budgeted` with budget >= 5x threshold |
| MED | `missing_budgeted` and `partial_below_baseline` (default) |
| LOW | not currently used (reserved for future signals) |

## Dept-drift rules (Reclass_Dept_Drift tab)

Bucket all current-month transaction lines by `(vendor, account, actual_dept)`.
Skip software accounts (671100, 671000, 511425) - those go to Reclass_Software.

For each bucket where the vendor exists in the budget file AND
`actual_dept != budget_row.department`, write one row with:
- `actual_dept`
- `expected_dept` (from budget)
- `actual_amount` (sum of the bucket)
- `reason` ("FP&A budget dept = X; posted to Y")

Sort descending by absolute amount.

## Software reclass rules (Reclass_Software tab)

Same logic as dept-drift but ONLY for software accounts. Captures both:
1. **Primary case (most common):** monthly software amortization JEs landing
   in the wrong dept. Source-agnostic - bills, amortization JEs, etc.
2. **Secondary case:** vendor's budget account is non-software but actual
   posted to a software account, or vice versa. Flagged via an extra
   "(also account drift)" note.

## Out-of-scope categories

Vendors with `category == "Other"` (Marketing / Rent / T&E / Office / Salary)
are NOT processed by accrual signals. They still appear in dept-drift if their
actuals land in the wrong dept (since dept-drift scans ALL non-software
actuals, not just in-scope categories).

## Sorting

Within each accrual tab, candidates are sorted by `suggested_amount` desc,
then vendor name asc. Reclass tabs sort by `abs(actual_amount)` desc.
