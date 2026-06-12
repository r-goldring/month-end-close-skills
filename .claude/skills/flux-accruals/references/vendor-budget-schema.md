# FP&A Vendor Budget — Schema and Refresh Policy

The FP&A team produces a vendor-level budget file used by `flux-accruals` as the
source-of-truth for "what each vendor SHOULD post each month and where."

## File location

```
Monthly Flux Analysis/Vendor Budget v{M.D.YYYY}.xlsx
```

The version date in the filename reflects when FP&A last published. The loader
auto-discovers the latest by mtime — the accountant can drop a new version next to the old
one and the next run picks it up.

## Sheet

Single sheet `Export Budget (5)`. ~221 rows as of v4.20.2026.

## Columns (26)

| Column | Type | Used by | Notes |
|---|---|---|---|
| Vendor | str | — | short code |
| Vendor ID | int | **primary join key** | NetSuite internal ID |
| Vendor Name | str | display + name fallback | |
| Functional Department | str | — | high-level (S&M / R&D / G&A) |
| Department | str | **dept-drift, dept reclass** | leaf cost center |
| Subsidiary L3 | str | — | usually "Acme, Inc." |
| Account Parent | str | category mapping | Software / Professional Services / COGS / Contractors / Marketing / Rent / T&E / Office / Salary |
| Account | str | account number + name | "671100 - Software Subscriptions" — parsed by `_parse_account_field()` |
| Department Grouping | str | — | |
| Currency | str | — | always USD in current file |
| Location | str | — | |
| Contract End Date | date or NaT | candidate notes | nullable |
| Comment | str | candidate notes | mid-year change flags |
| Jan / Feb / ... / Dec | float | **monthly budget** | per-month USD; constant unless `Comment` notes a change |

## Category mapping

`load_vendor_budget.classify_category()` buckets each row to one of:

- `Software`
- `Contractors`
- `ProfFees`
- `COGS`
- `Other` (out of scope for accruals — Marketing / Rent / T&E / Office / Salary)

Bucketing prefers `Account Parent` first, falls back to the parsed account
number range. Software / Contractors / Professional Fees / COGS are the four
in-scope categories per CLAUDE.md.

## Refresh policy

- FP&A owns the file. Do NOT modify it from any skill.
- When a new version arrives (e.g., `Vendor Budget v5.5.2026.xlsx`), the accountant drops
  it next to the existing one. Old versions can stay; the loader picks the
  most recent by mtime.
- If the accountant needs a vendor-specific override that FP&A hasn't published yet,
  the right path is for the accountant to ask FP&A to add the row; do NOT maintain a
  side-car overrides file in this repo.

## Vendor name normalization

`_normalize_name()` strips punctuation, collapses whitespace, lowercases, and
trims common corp suffixes (Inc / LLC / Ltd / Corp / etc.). Used for both
exact-match (NS report vendor name -> budget) and fuzzy-match (BillFlow vendor
name -> budget) joins.

If a vendor is in actuals but the normalized name doesn't match any budget
row, it surfaces in the workbook's `Unmatched_Actuals` tab. The fix is for
FP&A to add the vendor; not for the skill to fuzzy-match aggressively.
