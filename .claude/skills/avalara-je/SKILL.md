---
name: avalara-je
description: >
  Generate the monthly Avalara sales tax journal entry CSV for NetSuite import.
  Wraps the production Python script at scripts/avalara/generate_avalara_je.py.
  Use this skill when the user mentions: Avalara, sales tax JE, tax liability, AvaTax,
  Avalara journal entry, sales tax payable, or provides a TaxLiabilityWorksheet file.
---

# Avalara Sales Tax JE Skill

## Overview

Processes the monthly Avalara tax liability report into a NetSuite import CSV. The Python
script `generate_avalara_je.py` handles all state mapping and JE structure.

**Script location:** `scripts/avalara/generate_avalara_je.py`
**Mapping file:** `scripts/avalara/State Department of Taxation Name.csv`
- Maps state abbreviations (CA, NY, TX, etc.) to full vendor names (e.g., "California Department of Tax and Fee Administration")
- Covers 35 states + DC

## Step 1 — Get the input file

Ask the accountant for the Avalara tax liability file. Two file types may be needed:
- `TaxLiabilityWorksheetSummaryReturnDetail.xlsx` — primary input with state amounts
- `TransactionImpactResults.xls` — supplementary for prior period adjustments

Typical location: drop into `Monthly Avalara/{YYYY-MM}/` (create the folder if needed — this is the parallel to `Monthly Payroll/Pay Runs/`).

## Step 2 — Run the script

```bash
cd scripts/avalara
python generate_avalara_je.py
```

The script:
- Reads the Avalara report, finds "State" and "Amount Due To Avalara" columns dynamically
- Extracts reporting period from the filename (YYYYMM format)
- Calculates JE date as last day of next month
- Generates a 3-part JE structure per state (see below)

Output: `Avalara_NetSuite_JE_{Month}_{Year}.csv`

## Step 3 — JE structure per state

For each state with tax due:
1. **Cash line** (Credit): `111070 Cash - Chase Checking x0001` — total amount due
2. **Sales Tax Expense line** (Credit): `651210 Sales Tax Expense` — prior period discount + rounding adjustments (if any)
3. **Sales Tax Payable line** (Debit): `235100 Sales Taxes Payable - Avatax` — net amount (Amount Due minus adjustments)

All lines use the state's vendor name in the `Name` field.

## Step 4 — Review the output

Check:
- All 35 states + DC covered (warn if any states appear in the report but not in the mapping file)
- Cash line total = sum of all state amounts
- JE date = last day of next month from the reporting period
- Memo format: `{MonthAbbr} {Year}` (e.g., "Dec 2025" for the November period)

## Step 5 — Check-and-balance and upload

Display summary:
```
AVALARA JE — {Reporting Period}
States: XX
Total cash: $X,XXX,XXX.XX
JE date: {Date}
JE memo: {Memo}

Unmapped states: none / [list any]
```

Confirm with the accountant, then either:
- Upload the CSV via NetSuite JE import
- Or post via MCP with `../_shared/check-and-balance.md` confirmation **plus
  `../_shared/approval-required.md`** — set `"approved": false` on the `journalentry`
  payload so it routes to the controller's Pending Approval queue rather than auto-posting.

## Key accounts

| Account | Description |
|---------|-------------|
| `111070 Cash - Chase Checking x0001` | Operating cash account (credit — cash goes out) |
| `651210 Sales Tax Expense` | For prior period discounts/rounding (credit — reduces expense) |
| `235100 Sales Taxes Payable - Avatax` | Sales tax payable (debit — reduces liability) |

## Subsidiary

`Acme Holdings : Acme, Inc.`
