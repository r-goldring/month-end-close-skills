# Accrual Detection Rules — Pivot-Driven Pattern Library

This file codifies how the accrual builder (`build_accrual_suggestions.py`) decides which vendors to suggest accruing for in the current month. Logic is grounded in two months of observed the accountant + the VP decisions (Mar + Apr 2026 closes) plus 4 months of meeting transcripts.

## Inputs to the detector

For each vendor in the 4 detail-report data sheets (COGS / Contractors / ProfFees / Software):
1. **Per-period totals**: Feb (M-2), Mar (M-1), Apr (M).
2. **Last month's accrual amount** (cross-referenced from BOTH `{YYYY-{MM-1}} Accruals JE Import.csv` AND NetSuite's posted JE with externalid `{YYYY-{MM-1}}-ACCRUALS`).
3. **Memo keywords** in transaction lines (quarterly / annual / renewal / etc.).

## Rule set (first match wins)

### Rule A — Always-Accrue Vendors (highest confidence)

Hardcoded list of vendors that ALWAYS get accrued every month regardless of pattern:

| Vendor | Account | Typical amount | Notes |
|---|---|---|---|
| IT Hardware Partner Networked Solutions Group, LLC | 511400 | $350–377K | Bills in arrears; accrue at the AP Specialist's draft estimate |
| DNS Provider A | 511400 | $X,XXX.XX | Consistent monthly |
| Health Premium Billing Intermediary | 511200 / 611300 | $50–88K | Always present, amount varies |
| D and O Broker | 651250 | $X,XXX.XX | Constant amortization |

Confidence: **High**. Suggestion = last month's actual amount (or last accrual amount if available).

### Rule B — Last-Month-Accrual Recurrence

If a vendor appears in last month's accrual JE (either CSV or NS-posted version), re-suggest the same amount this month with high confidence.

**Cross-reference both sources**:
1. Read `{YYYY-{MM-1}}/JE Imports/{YYYY-{MM-1}} Accruals JE Import.csv`
2. Query NetSuite for the actual posted JE: `SELECT account, sum(netamount), name FROM transactionline WHERE transaction IN (SELECT id FROM transaction WHERE externalid = '{YYYY-{MM-1}}-ACCRUALS')`
3. If CSV amount ≠ NS posted amount, trust NS (the accountant may have adjusted in NS after generating the CSV). Flag the diff in the suggestion note.

Confidence: **High**.

### Rule C — Recurring Missing-Current

Trigger: vendor in M-2 AND M-1 with avg > threshold, M = $0 (or < 5% of avg).

| Account family | Threshold |
|---|---|
| Software (671xxx, 511425) | $500 |
| Contractors (511370, 611700) | $100 |
| Professional Fees (651100, 651101) | $500 |
| COGS General (other 511xxx) | $500 |

Suggested accrual: average of (M-2, M-1).

Confidence: **High** if 2-of-2 prior months matched within 20% of each other; **Medium** otherwise.

### Rule D — Step-Down From Prior

Trigger: vendor in all 3 months but M is < 30% of avg(M-2, M-1) AND gap > threshold.

Suggested accrual: avg(M-2, M-1) − M.

Confidence: **Medium**. Often signals a billing-cycle issue (bill due but not yet received) vs. an actual rate change. Flag both possibilities.

### Rule E — Quarterly / Annual Cadence

Trigger: vendor's memo contains `quarterly` / `Q1` / `Q2` / `annual` / `renewal` AND last hit was 3+ months ago.

Suggested accrual: prior amount, prorated by the months since (so 1/3 of quarterly amount if it's been 1 month since the last hit, etc.).

Confidence: **Low** without confirming the cadence. Always note the cadence and ask for confirmation.

Known quarterly vendors: RetireBridge 401k, A-Census, Tax Firm A FY tax provision.

### Rule F — Brex Card Categorical Vendors (DO NOT suggest)

Skip rows whose Entity (Line): Name is in the Brex categorical denylist (`Lodging`, `Airfare`, `Meals`, `Transportation`, `Taxi`, `Hotel`, `Uber`). These are NetSuite report quirks, not real vendor activity — filtered out by `parse_gl_detail` before the suggester even sees them.

### Rule G — Always-Skip Vendors

Hardcoded list of vendors that NEVER get accrued (per the accountant's policy):

| Vendor | Why |
|---|---|
| Former FSA Administrator | Terminated 2025; refunds roll to 2025 per the VP |
| AcquiredCo lease | Lease expired Feb 2026 |
| Customer Conference / Customer Conference ticket revenue | Offsets event expense, not accrued separately |

### Rule H — Always-Verify Vendors (suggest with Low confidence + note)

Vendors that have judgment calls each month:

| Vendor | Why |
|---|---|
| Anthropic / Claude | Mixed Brex charges + annual licenses; the VP's allocation decision pending |
| OpenAI | $950/seat Brex pattern; verify annual vs monthly |
| Microsoft (Brex one-time licenses) | Quarterly/annual purchases; cardholder confirmation needed |
| SEMRUSH (annual renewals) | $X,XXX.XX hit once a year; convert to prepaid not accrual |
| Legacy Hosting Stack-related Brex charges | the Assistant Controller owns reclass to COGS, not an accrual |

## Confidence levels and what they mean

| Level | Meaning | Action |
|---|---|---|
| **High** | Clear recurring pattern + last-month-accrual match | Auto-write to candidate JSON with `approved: false`; the accountant only needs to confirm. |
| **Medium** | 2-of-3 month pattern or plausible cadence; some judgment | Write to JSON; surface in chat summary for the accountant to glance at. |
| **Low** | Irregular history, one-time-looking, or always-verify list | Write to JSON with explicit "the accountant should verify" tag. Don't write into the workbook static table. |

## Output schema

The suggester writes to `Monthly Flux Analysis/{YYYY}/{YYYY-MM}/State/accrual_candidates.json`:

```json
{
  "version": 1,
  "generated_at": "2026-05-12T...",
  "candidates": [
    {
      "vendor": "IT Hardware Partner Networked Solutions Group, LLC",
      "account": "511400",
      "department": "Engineering : Infrastructure",
      "amount": 377091.00,
      "confidence": "high",
      "rule": "always_accrue",
      "reason": "Recurring monthly hosting accrual (bills in arrears). Last month CSV: $377K; NS posted: $377K (match).",
      "approved": false,
      "notes": ""
    },
    {
      "vendor": "Pension Audit Firm LLP",
      "account": "651100",
      "department": "General & Administrative : GA",
      "amount": 8500.00,
      "confidence": "medium",
      "rule": "quarterly_cadence",
      "reason": "401k audit; last hit Apr 2026 $8.5K. Quarterly cadence. Suggest accrual.",
      "approved": false,
      "notes": "Confirm engagement timing with the VP."
    }
  ]
}
```

Also writes a "Accrual Suggestion" annotation to columns **M / N / O** of the pivot template's static F-L table for each vendor:
- M: suggested amount (number)
- N: confidence (text: "High", "Medium", "Low")
- O: reason (short text — 1 sentence max)

## What the suggester does NOT do

- **Does not auto-write the JE CSV.** That's `build_accrual_csv.py`, which only runs after the accountant approves entries in the JSON or workbook.
- **Does not use BillFlow data as primary signal.** Per the accountant's Apr 2026 feedback, BillFlow creates too many false positives. Pivot template historical activity is the trusted signal.
- **Does not modify the accountant-approved entries.** If `accrual_candidates.json` already has `edited_by: ryan` for a vendor, the suggester preserves the accountant's amount/approval/notes and appends its current-month suggestion to a `history` field for transparency.
- **Does not use FP&A budget as a primary signal.** The existing `flux-accruals` skill handles budget-vs-actual; that's secondary and not part of this pivot-driven flow.

## Tuning the rules

To add a new always-accrue vendor or change a threshold:
1. Edit this file's rule sections.
2. Update `scripts/flux-analysis-workbook/build_accrual_suggestions.py` constants if the logic isn't config-driven.
3. Test on the prior month's data before running on current month.
