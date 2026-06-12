# Accrual Thresholds by Category

These thresholds determine when budget-vs-actual gaps warrant a candidate row in
the workbook. Rules driven by `signal-rules.md`.

**Missing-budgeted (highest signal):** vendor has a non-zero monthly budget AND
zero actuals in the period. No threshold gate — every missing-budget vendor in
an in-scope category surfaces.

**Partial-below-baseline:** vendor posted some actuals, but they're below 50% of
budget AND the gap (`budget - actual`) exceeds the category threshold. The
threshold gates partial-shortfall noise; missing-budget always wins.

**Over-budget review (informational):** actuals > 120% of budget AND
`actual - budget` exceeds the threshold. Surfaces for variance review, NOT for
accrual.

| Category | NetSuite Report | GL Accounts | Threshold | Notes |
|----------|----------------|-------------|-----------|-------|
| Software | 721 | 671100, 671000, 511425 | $500 | Lowered from $750 on 2026-05-12 per the accountant |
| Contractors | 540 | 511370, 611700 | $100 | Includes PEO-managed contractors |
| Professional Fees | 542 | 651100, 651150 | $500 | Legal + other professional services |
| COGS (General) | 537 | 511xxx | $500 | Default for other COGS accounts |

## To adjust a threshold

Edit this file and change the threshold value. The flux-accruals skill reads this file
to determine which vendors to flag. Changes take effect on the next run.

## Rationale

- **Software ($500)**: the CFO and the Director scrutinize software closely. Most subscriptions
  are $500+ monthly. Lowered from $750 to catch mid-tier annual renewals (e.g., SEMRUSH
  $4.6K, Temporal $1.9K) that hit Brex once a year and would otherwise miss the gate.
- **Contractors ($100)**: Even small contractor amounts matter for COGS accuracy.
  PEO providers like PEO Provider often have small adjustments.
- **Professional Fees ($500)**: Out-of-period invoices are common (Acme UK External Auditor); accrual
  helps match the expense to the right period.
- **COGS ($500)**: Default; large hosting providers (IT Hardware Partner ~$350-370K) rarely
  have zero months so this usually catches only smaller items.
