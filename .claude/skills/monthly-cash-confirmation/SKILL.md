---
name: monthly-cash-confirmation
description: >
  Confirm FP&A's monthly 1T cash inflows/outflows + debt service requests. Walks
  every line item in the inbound email from the Director or the FP&A Lead, traces
  it through the BillFlow payments_out CSV, the month-end banking workbook, and
  NetSuite SuiteQL, then reports CONFIRMED / VARIANCE / NOT CONFIRMED in chat.
  Skill does NOT draft the email reply - the accountant writes the response himself.
  Use this skill when the user mentions: monthly cash confirmation, cash
  confirmation, 1T cash, cash inflows, cash outflows, confirm cash, FP&A cash
  request, sublease confirmation, debt service confirmation, the Director cash,
  the FP&A Lead cash, or drops a "{Month}'{YY} Cash Confirmations.pdf" into
  Monthly Cash Confirmations/.
---

# Monthly Cash Confirmation Skill

## Overview

Every month FP&A emails the accountant asking him to confirm a handful of 1T cash items
(sublease payments/income, severance, retention bonuses, debt service, other 1T
legal/consulting). the accountant currently traces each one by hand. This skill automates
the trace and reports the result in chat.

**Hard rule: no hallucinating.** If a line cannot be confidently confirmed
against the source files, report it as `NOT CONFIRMED - manual review needed`
rather than guessing. Accuracy is the whole point.

**Tolerance**: $1 (FP&A's amounts are rounded to whole dollars; anything within
$1 of the FP&A figure counts as a match).

**References:**
- SuiteQL patterns: `../_shared/netsuite-queries.md`
- Subsidiary constants: `../_shared/subsidiary-constants.md`
- ID lookup: `../_shared/id-lookup-guide.md`
- Vendor alias dictionary: `./vendor_aliases.yaml`

---

## Inputs (the accountant drops these into `Monthly Cash Confirmations/`)

| File | Required | Notes |
|---|---|---|
| `{Mon}'{YY} Cash Confirmations.pdf` | Yes | Gmail print-to-PDF of inbound email |
| `payments_out_{YYYY-MM-DD}.csv` | Yes | BillFlow export of outgoing payments |
| `{Mon}-{YY} Banking Sheet*.xlsx` (Jun 2026 onwards) or `{Mon}-{YY} Banking Transactions*.xlsx` (through May 2026) | Yes | Reused from `bank-statement-posting`. If absent locally, fetch from Drive (Step 1) |

The BillFlow CSV columns are: `Confirmation Number, Vendor, Process Date,
Payment Status, Payment Method, Payment Amount, Arrival Date, Invoice Number,
Paid From, Expires On, Vendor Credit, Currency`.

The banking XLSX has one tab per bank account (16 accounts: Wells Fargo x0003, Chase
x0001, Chase x0002, foreign banks) plus Summary/Access/Bill Summary/ME Summary.
Transaction rows have Date, Description, Amount, CR/DR, Type, TAG, WEEK, Posted.

---

## Step 1 - Identify the target month and locate inputs

Ask the user (or infer from their message) which month to confirm. Normalize
to `(Mon, YY, YYYY-MM)`, e.g. `("Mar", "26", "2026-03")`.

Resolve file paths:
```python
folder = "Monthly Cash Confirmations"
pdf_path = f"{folder}/{Mon}'{YY} Cash Confirmations.pdf"
banking_glob = f"{folder}/{Mon}-{YY} Banking Transactions*.xlsx"
billcom_glob = f"{folder}/payments_out_*.csv"
```

Use `Glob` to confirm each exists. If any is missing, tell the accountant what's missing
and stop.

---

## Step 2 - Parse the email PDF

Use the `Read` tool on the PDF. The PDF text is reliably extractable.

**Parse rules:**
- Only parse the **first inbound email** (the one from the Director / the FP&A Lead
  the FP&A Lead). Ignore the quoted thread of the accountant's prior responses.
- Recognize these section headers (case-insensitive):
  - `Sublease Payments`
  - `Sublease Income`
  - `Severance`
  - `1T/Retention Bonuses`
  - `Debt Service`
  - `Other 1T`
- Under each section, capture bulleted/indented line items in the form
  `{name}: ${amount}` optionally followed by a parenthetical note.
- Some lines have no `$` (e.g., `Interest Income: Can you please confirm?`) -
  set amount to `null` and capture the note text.
- Some lines have a parenthetical breakdown (e.g., `$X,XXX.XX ($150K bonus + $9K
  payroll tax)`) - capture the top-line $ as the amount; put the breakdown in
  notes.

Build a list of items:
```python
items = [
  {"section": "Sublease Payments", "name": "Landlord B", "amount": 26415, "note": None},
  {"section": "Sublease Income", "name": "Premier Plans", "amount": 11692, "note": "slightly higher than prior months of $X,XXX.XX"},
  {"section": "Debt Service", "name": "Interest Income", "amount": None, "note": "Can you please confirm?"},
  ...
]
```

Show the accountant the parsed item list so he can sanity-check before lookups run.

---

## Step 3 - Route each item to a lookup strategy

For each item, determine the lookup source based on section + name:

| Section | Item | Source |
|---|---|---|
| Sublease Payments | any | BillFlow CSV |
| Sublease Income | any | Banking XLSX (Wells Fargo + Chase, search Description + Name + Transaction Detail) |
| Other 1T | NL Social Insurance Agency | NetSuite vendor 69232: check records to ING-EUR. Original currency EUR; multiply by month's EUR/USD rate (~1.16 in Mar 2026) for the USD-equivalent FP&A is asking about. |
| Severance | (one total) | NetSuite SuiteQL (accts 511175 + 611250) |
| 1T/Retention Bonuses | (one total) | NetSuite SuiteQL (accts 511150 + 611150, excl accruals) |
| Debt Service | Term Loan Principal | Banking XLSX Chase Operating (tab `chase_op`) |
| Debt Service | Interest | NetSuite SuiteQL (acct 711150) |
| Debt Service | Interest Income | NetSuite SuiteQL (acct 711000) |
| Other 1T | JB Fees / JB Travel Fees | NetSuite SuiteQL (memo search) |
| Other 1T | Legacy Hosting Stack Acquisition | Banking XLSX Chase Operating (tab `chase_op`, large outflow) |
| Other 1T | Legacy Hosting Stack employee Travel Fees | NetSuite SuiteQL (T&E accts filtered to EBITDA Adj dept 108) |
| Other 1T | Acme Labs Contractors | NetSuite SuiteQL: vendor bills + line memos containing "Acme Lab"; aggregate PEO Provider (BillFlow) + AP Consulting (NS-only bills) |
| Other 1T | Medical Carrier Aical / Dental | BillFlow CSV (Marked as paid) |
| Other 1T | (everything else - legal/consulting/contractors) | BillFlow CSV |

---

## Step 4 - Execute BillFlow lookups

Load `payments_out_*.csv` with pandas:
```python
import pandas as pd
df = pd.read_csv(billcom_path)
df['Process Date'] = pd.to_datetime(df['Process Date'])
df['Payment Amount'] = pd.to_numeric(df['Payment Amount'])
month_df = df[(df['Process Date'] >= period_start) & (df['Process Date'] <= period_end)]
```

For each BillFlow-routed item:
1. Look up `name` in `vendor_aliases.yaml` to get the list of vendor patterns.
   If `name` is not in the dict, fall back to using `name` itself as a single
   substring pattern.
2. Filter `month_df` to rows where `Vendor` (case-insensitive substring) matches
   any of the patterns.
3. Sum `Payment Amount`. Note: if Currency != USD, treat as USD-equivalent-not-
   available and flag for manual (BillFlow export gives original-currency only).
4. Compare against the FP&A amount with $1 tolerance.

Record sources used (Confirmation Number(s) and Process Date(s)) for the report.

---

## Step 5 - Execute banking XLSX lookups

Load the banking workbook with openpyxl:
```python
from openpyxl import load_workbook
wb = load_workbook(banking_xlsx, data_only=True)
```

### Format detection (REQUIRED — workbook layout changed mid-2026)

The banking workbook layout was overhauled around June 2026 by the accountant's boss,
which `bank-statement-posting` now produces in the new shape. Both old and
new files may need to be read, so detect format at the start and normalize.

**Old format (through May 2026, file titled `{Mon}-{YY} Banking Transactions.xlsx`):**
- Header rows 1-3 carry month carry-forward totals; column headers live on **row 4**
- Account-tab names: `1.Wells Fargo x0003`, `2. Chase x0001`, `3. Chase x0002`, `4. ING-EUR`...
- Columns: `Date | Transaction Description | Amount (unsigned) | CR/DR | Transaction Detail | Type | TAG | WEEK | Assignee | Posted | JE # | Notes`
- Direction encoded in separate `CR/DR` column

**New format (June 2026 onwards, file titled `{Mon}-{YY} Banking Sheet.xlsx`):**
- Column headers live on **row 1** (no carry-forward header block)
- Account-tab names: `1. Chase Op (US)`, `2. Bank of America MMA (US)`, `3. Wells Fargo`, `4. ING BV (EUR)`...
- Columns: `Date | Description | Amount (SIGNED) | Tag | Owner | Week | Posted? | Ref #`
- Direction encoded by amount sign: `Amount < 0` = outflow, `Amount > 0` = inflow
- Description is the merged ACH detail field (no separate Transaction Detail column)

Detect with a one-shot probe on any account tab:

```python
def detect_format(wb):
    """Returns 'new' or 'old' based on row-1 headers."""
    for sn in wb.sheetnames:
        if sn.lower() in ('summary','summary (usd)','summary (lc)','access','bill summary','me summary'):
            continue
        ws = wb[sn]
        r1 = [ws.cell(1, c).value for c in range(1, 10)]
        # New format: row 1 has "Date" + "Description" + "Amount" + "Tag"
        if any(v == 'Date' for v in r1) and any(v == 'Tag' for v in r1):
            return 'new'
        # Old format: row 1 has dates/totals; "Date" appears on row 4
        return 'old'
    raise RuntimeError("could not detect banking workbook format")
```

### Tab name resolver (works for both formats)

| Logical account | Old tab name | New tab name |
|---|---|---|
| Chase Operating x0001 | `2. Chase x0001` | `1. Chase Op (US)` |
| Bank of America MMA x0002 | `3. Chase x0002` | `2. Bank of America MMA (US)` |
| Wells Fargo x0003 (lockbox) | `1.Wells Fargo x0003` | `3. Wells Fargo` |
| ING BV EUR | `12. ING BV EUR` | `4. ING BV (EUR)` |
| ING BV USD | `14. ING BV USD` | `5. ING BV (USD)` |
| ING BV GBP | `13. ING BV GBP` | `6. ING BV (GBP)` |
| TD CAD | `10. TD-CAD` | `7. TD` |
| Barclays UK | `11. Barclays-UK Ltd` | `8. Barclays UK` |
| ING EUR | `4. ING-EUR` | `14. ING (EUR)` |
| ING GBP | `5. ING-GBP` | `15. ING (GBP)` |
| ING USD | `7. ING-USD` | `13. ING (USD)` |
| Bank of America PLN | `8. Bank of America PLN x0006` | `10. Bank of America (PLN)` |
| Bank of America USD | `9. Bank of America - USD x0007` | `9. Bank of America (USD)` |
| Banco Santander UYU | `15. Banco Santander UYU` | `11. Banco Santander (UYU)` |
| Banco Santander USD | `16. Banco Santander USD` | `12. Banco Santander (USD)` |

```python
TAB_ALIASES = {
    'chase_op':    ['1. Chase Op (US)',  '2. Chase x0001'],
    'chase_mma':   ['2. Bank of America MMA (US)', '3. Chase x0002'],
    'wellsfargo':        ['3. Wells Fargo',           '1.Wells Fargo x0003'],
    'ing_bv_eur':['4. ING BV (EUR)', '12. ING BV EUR'],
    'ing_bv_usd':['5. ING BV (USD)', '14. ING BV USD'],
    'ing_bv_gbp':['6. ING BV (GBP)', '13. ING BV GBP'],
    'td':        ['7. TD',           '10. TD-CAD'],
    'barclays_uk':    ['8. Barclays UK',       '11. Barclays-UK Ltd'],
    'bofa_usd':   ['9. Bank of America (USD)',    '9. Bank of America - USD x0007'],
    'bofa_pln':   ['10. Bank of America (PLN)',   '8. Bank of America PLN x0006'],
    'santander_uy_uyu':['11. Banco Santander (UYU)','15. Banco Santander UYU'],
    'santander_uy_usd':['12. Banco Santander (USD)','16. Banco Santander USD'],
    'ing_usd':   ['13. ING (USD)',   '7. ING-USD'],
    'ing_eur':   ['14. ING (EUR)',   '4. ING-EUR'],
    'ing_gbp':   ['15. ING (GBP)',   '5. ING-GBP'],
}
def get_tab(wb, key):
    for name in TAB_ALIASES[key]:
        if name in wb.sheetnames:
            return wb[name]
    return None
```

### Row reader (returns a normalized record regardless of format)

```python
def read_rows(ws, fmt):
    """Yields dicts with normalized fields: date, description, amount (signed),
    tag, posted_ref. Works for both old and new formats."""
    if fmt == 'new':
        header_row = 1
        col_date, col_desc, col_amt, col_tag, col_ref = 1, 2, 3, 4, 8
    else:  # old
        # Find row where col A == 'Date' or 'AsOfDate'
        header_row = next(r for r in range(1, 8)
                          if str(ws.cell(r,1).value or '').lower() in ('date','asofdate','statement date'))
        # Old format positions vary by tab; resolve by header text
        headers = {ws.cell(header_row, c).value: c for c in range(1, ws.max_column+1)}
        col_date = headers.get('Date') or headers.get('AsOfDate')
        # Old Chase has 'Transaction Description' (col 2) + 'Transaction Detail' (col 5);
        # we concatenate both for the searchable description
        col_desc_main = headers.get('Transaction Description') or headers.get('Description') or 2
        col_desc_detail = headers.get('Transaction Detail') or headers.get('Detail')
        col_amt = headers.get('Amount') or headers.get('Debit') or 3
        col_crdr = headers.get('CR/DR')
        col_tag = headers.get('TAG') or headers.get('Tag')
        col_ref = headers.get('Posted') or headers.get('POSTED') or headers.get('JE #')
    for r in range(header_row+1, ws.max_row+1):
        if fmt == 'new':
            d = ws.cell(r, col_date).value
            amt = ws.cell(r, col_amt).value
            if d is None and amt is None: continue
            yield {
                'date': d,
                'description': ws.cell(r, col_desc).value or '',
                'amount': float(amt) if amt is not None else 0.0,
                'tag': ws.cell(r, col_tag).value,
                'posted_ref': ws.cell(r, col_ref).value,
            }
        else:
            d = ws.cell(r, col_date).value
            if d is None: continue
            desc_main = ws.cell(r, col_desc_main).value or ''
            desc_detail = ws.cell(r, col_desc_detail).value if col_desc_detail else ''
            description = f"{desc_main} {desc_detail or ''}".strip()
            amt_raw = ws.cell(r, col_amt).value
            try: amt = float(amt_raw) if amt_raw is not None else 0.0
            except: continue
            # Apply sign from CR/DR if present (old format)
            if col_crdr and ws.cell(r, col_crdr).value == 'DR':
                amt = -abs(amt)
            elif col_crdr and ws.cell(r, col_crdr).value == 'CR':
                amt = abs(amt)
            yield {
                'date': d, 'description': description, 'amount': amt,
                'tag': ws.cell(r, col_tag).value if col_tag else None,
                'posted_ref': ws.cell(r, col_ref).value if col_ref else None,
            }
```

After this, all lookups use the normalized record schema:
- `r['amount'] > 0` = inflow (was CR/DR=='CR' in old format)
- `r['amount'] < 0` = outflow (was CR/DR=='DR' in old format)
- `r['description']` = single searchable field (merged for old format)
- `r['posted_ref']` = PYMT#/JE#/DEP#/TRN#/ACH_Debit ref

### Lookups (using normalized records)

**Sublease Income**: Iterate `get_tab(wb, 'wellsfargo')` and `get_tab(wb, 'chase_op')`.
Filter `r['amount'] > 0` and `target_month`. Match the tenant by alias
against `r['description']` (it now contains the full ACH metadata including
the `NAME:ORIG:...` segment that holds tenant names like `GA OPERATING CASH
DISB`). Sum amounts; cite `r['posted_ref']` (PYMT#/DEP#). March 2026 verified
(old format): Premier Plans $X,XXX.XX on Chase Operating, DEP878.

**Term Loan Principal**: Iterate `get_tab(wb, 'chase_op')`. Filter
`r['amount'] < 0` AND `r['description']` contains `TERM LOAN`, `PRINCIPAL`,
or `PRINCIPAL AND AGENCY FEE`. The end-of-month ACH typically shows as
`MISCELLANEOUS DEBIT / GENERAL TRANSFER / ACME CORP PRINCIPAL AND AGENCY FEE`.
March 2026 verified: $X,XXX.XX on 2026-03-31 = FP&A $X,XXX.XX. Note: agency
fee (~$X,XXX.XX) is bundled into this ACH but classifies to acct 711150 (Interest
Expense - Term Loans), which creates a small variance against FP&A's
"Interest: $0" line — flag, don't reconcile.

**Legacy Hosting Stack Acquisition**: Iterate `get_tab(wb, 'chase_op')`. Filter
`r['amount'] < 0` and `r['description']` contains `LYCEUM`. If absent, mark
`MANUAL - large M&A wire not in source files`.

**Interest Payment (quarterly, large)**: Iterate `get_tab(wb, 'chase_op')`.
Filter `r['amount'] < 0` and `r['description']` contains `ACME CORP INTEREST`
or matches the FP&A amount exactly. The quarterly interest ACH shows as
`MISCELLANEOUS DEBIT / GENERAL TRANSFER / ACME CORP INTEREST {Q-end date}`.
April 2026 verified: $X,XXX.XX on 2026-04-02 for the 3/31 interest =
FP&A exact match.

---

## Step 6 - Execute NetSuite SuiteQL lookups

Use `ns_runCustomSuiteQL` per CLAUDE.md rule 1 (read-only NetSuite operations
are unaffected by the JE-CSV-only policy).

All consolidated-basis queries: `subsidiary = -2` per CLAUDE.md rule 5.

**Important: do NOT add `tl.mainline = 'F'` to any of these aggregate queries.**
That filter excludes payroll JE lines (which are where severance/bonus/JB
activity lives) and returned 0 rows in March 2026 testing. Use `tl.account`,
`posting = 'T'`, `voided = 'F'`, and trust the account scope to keep results
correct.

### Severance total
```sql
SELECT SUM(NVL(tl.debitforeignamount,0) - NVL(tl.creditforeignamount,0)) AS total
FROM transaction t
JOIN transactionline tl ON t.id = tl.transaction
JOIN account a          ON tl.account = a.id
WHERE a.acctnumber IN ('511175', '611250')
  AND t.trandate BETWEEN TO_DATE('{period_start}','YYYY-MM-DD')
                     AND TO_DATE('{period_end}','YYYY-MM-DD')
  AND t.voided = 'F'
  AND t.posting = 'T'
```
March 2026 verified: 511175 $X,XXX.XX + 611250 $X,XXX.XX = $X,XXX.XX vs FP&A $X,XXX.XX. **Match.**

### 1T / Retention Bonuses (memo-targeted, NOT account-wide)
Account-wide totals on 511150 + 611150 are dominated by regular bonus accruals
($1M+ in March 2026 alone) and miss the point. The cash-impacting retention
bonuses are tagged in line memos with `Legacy Hosting Stack Bonus`, `Retention`, `1T Bonus`,
or `Signing Bonus`. Search memos directly:

```sql
SELECT a.acctnumber, a.fullname,
       SUM(NVL(tl.debitforeignamount,0) - NVL(tl.creditforeignamount,0)) AS total
FROM transaction t
JOIN transactionline tl ON t.id = tl.transaction
JOIN account a          ON tl.account = a.id
WHERE a.acctnumber IN ('511150','611150','611450','231171')
  AND t.trandate BETWEEN TO_DATE('{period_start}','YYYY-MM-DD')
                     AND TO_DATE('{period_end}','YYYY-MM-DD')
  AND t.voided = 'F'
  AND t.posting = 'T'
  AND (LOWER(NVL(tl.memo,'')) LIKE '%lyceum bonus%'
       OR LOWER(NVL(tl.memo,'')) LIKE '%retention%'
       OR LOWER(NVL(tl.memo,'')) LIKE '%1t bonus%'
       OR LOWER(NVL(tl.memo,'')) LIKE '%signing bonus%')
GROUP BY a.acctnumber, a.fullname
```

The bonus portion lives on 611150 / 511150; the payroll tax portion on 611450
(or 231171 if accrued). Both should be summed and compared to FP&A's
`$X bonus + $Y payroll tax` breakdown.

March 2026 verified: JE##### (3/15 US Payroll) carried 3 Legacy Hosting Stack Bonus lines
= $150K bonus + $13.5K payroll tax = $163.5K vs FP&A $159K ($150K + $9K). The
bonus piece matches exactly; FP&A's payroll tax estimate was $4.5K low.

If variance > $1k vs FP&A figure, mark `PUNT TO the Assistant Controller - he has visibility into
1T vs retention split`.

### Interest Expense - Term Loans
```sql
SELECT SUM(tl.debitforeignamount - tl.creditforeignamount) AS total
FROM transaction t
JOIN transactionline tl ON t.id = tl.transaction AND tl.mainline = 'F'
JOIN account a          ON tl.account = a.id
WHERE a.acctnumber = '711150'
  AND tl.subsidiary = -2
  AND t.trandate BETWEEN TO_DATE('{period_start}','YYYY-MM-DD')
                     AND TO_DATE('{period_end}','YYYY-MM-DD')
  AND t.voided = 'F'
  AND t.posting = 'T'
```

### Interest Income
```sql
-- 711000 is an Other Income account; natural balance is CR (negative debit-credit).
-- Flip sign so the reported number is positive ("how much interest we earned").
SELECT -SUM(tl.debitforeignamount - tl.creditforeignamount) AS total
FROM transaction t
JOIN transactionline tl ON t.id = tl.transaction AND tl.mainline = 'F'
JOIN account a          ON tl.account = a.id
WHERE a.acctnumber = '711000'
  AND tl.subsidiary = -2
  AND t.trandate BETWEEN TO_DATE('{period_start}','YYYY-MM-DD')
                     AND TO_DATE('{period_end}','YYYY-MM-DD')
  AND t.voided = 'F'
  AND t.posting = 'T'
```

**Always compute Interest Income even if FP&A omitted the amount.**

### Legacy Hosting Stack employee Travel Fees (EBITDA Adj dept, T&E accounts)
**Legacy Hosting Stack employee travel is NOT tagged "Legacy Hosting Stack" in memos.** It is tagged via
the `EBITDA Adjustments` department (id `108`) on T&E P&L accounts. Per the accountant
(2026-05-19): "Go into the T&E PL accounts and filter for Mar EBITDA Adj dept,
that should get you to narrow down on the lyceum travel."

T&E accounts (id list from COA): 660000, 661000, 661100 (Airfare), 661150
(Transportation), 661200 (Lodging), 661250 (Meals), 661260 (Sales Support),
662000 (Client Travel Non-billable), 621400 (Conference Travel & Registration).

```sql
SELECT a.acctnumber, a.fullname,
       SUM(NVL(tl.debitforeignamount,0) - NVL(tl.creditforeignamount,0)) AS total
FROM transaction t
JOIN transactionline tl ON t.id = tl.transaction
JOIN account a          ON tl.account = a.id
WHERE tl.department = 108
  AND a.acctnumber IN ('660000','661000','661100','661150','661200',
                       '661250','661260','662000','621400')
  AND t.trandate BETWEEN TO_DATE('{period_start}','YYYY-MM-DD')
                     AND TO_DATE('{period_end}','YYYY-MM-DD')
  AND t.voided = 'F'
  AND t.posting = 'T'
GROUP BY a.acctnumber, a.fullname
```
March 2026 verified: $482.35 on JE##### (Brex card reimbursements for ELT
offsite at TCV) = FP&A $482. **Match.**

### JB Fees / JB Travel Fees (memo search)
```sql
SELECT t.tranid, t.trandate, a.acctnumber, a.fullname,
       tl.memo, (tl.debitforeignamount - tl.creditforeignamount) AS amount
FROM transaction t
JOIN transactionline tl ON t.id = tl.transaction AND tl.mainline = 'F'
JOIN account a          ON tl.account = a.id
WHERE tl.subsidiary = -2
  AND t.trandate BETWEEN TO_DATE('{period_start}','YYYY-MM-DD')
                     AND TO_DATE('{period_end}','YYYY-MM-DD')
  AND t.voided = 'F'
  AND t.posting = 'T'
  AND (REGEXP_LIKE(tl.memo, '(^|[^A-Za-z])JB([^A-Za-z]|$)', 'i')
       OR LOWER(tl.memo) LIKE '%john borland%')
ORDER BY t.trandate
```

Sum the amounts. Separate JB Travel Fees if the FP&A email distinguishes them
(filter to travel-related accounts: 651xxx Travel, T&E accounts).

---

## Step 6.5 - April-2026 lessons learned

### BillFlow CSV — header quirks
The export is UTF-8 with BOM, and header casing varies between exports
(`"Process date"` vs `"Process Date"`). Always read with
`encoding='utf-8-sig'` and normalize columns to Title Case:
```python
df = pd.read_csv(path, encoding='utf-8-sig')
df.columns = [c.strip().title() for c in df.columns]
```

### Banking workbook — Google Sheets fallback
If the `{Mon}-{YY} Banking Transactions*.xlsx` isn't in
`Monthly Cash Confirmations/`, the source lives as a Google Sheet titled
`{Mon}-{YY} Banking Transactions` on the shared Drive. Fetch via
`mcp__claude_ai_Google_Drive__search_files` + `download_file_content`
with `exportMimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'`.
The download returns a JSON wrapper with the XLSX bytes base64-encoded
(saved to a temp file if too large to inline). Decode and write to
`Monthly Cash Confirmations/{Mon}-{YY} Banking Transactions.xlsx`.

### Quarterly debt payments
Term Loan Principal and Interest Payment are usually quarterly
(Mar/Jun/Sep/Dec), not monthly. The quarterly interest pay on Chase x0001
shows up as `MISCELLANEOUS DEBIT / GENERAL TRANSFER / ACME CORP INTEREST
{Q-end date}` (April 2026: $X,XXX.XX on 4/2/2026 for the 3/31 interest).
If FP&A doesn't list either line in a given month, don't try to find them
in NetSuite — they didn't move cash.

### BillFlow vendor totals can be too broad
FP&A sometimes asks about a SPECIFIC subset of a vendor's invoices, not the
vendor's full month total. April 2026: Legal Firm B total = $X,XXX.XX
across 3 invoices, but FP&A's $X,XXX.XX = the two 4/3 invoices only. When
the vendor total varies > $1 from FP&A, drop to invoice-level: report
which invoices match the FP&A figure and which are extra. The right
match is at the invoice level.

### Multi-currency severance (Uruguay UYU + Netherlands EUR + US USD)

`transactionline.debitforeignamount` / `creditforeignamount` are in the LINE
currency, NOT in the consolidated USD reporting currency. For Acme Corp US
sub-owned lines this IS USD (so March 2026 came out clean). For Netherlands
(EUR), Uruguay (UYU), and other foreign-sub-owned severance lines, summing
foreign amounts mixes currencies and produces nonsense totals.

The safe pattern: group activity by subsidiary + line currency, then convert
each non-USD bucket using the period's consolidated FX rate before summing.

```sql
SELECT sub.name, sub.currency,
       a.acctnumber,
       SUM(NVL(tl.debitforeignamount,0) - NVL(tl.creditforeignamount,0)) AS amt
FROM transaction t
JOIN transactionline tl ON t.id = tl.transaction
JOIN account a          ON tl.account = a.id
LEFT JOIN subsidiary sub ON tl.subsidiary = sub.id
WHERE ...
GROUP BY sub.name, sub.currency, a.acctnumber
```

April 2026 verified pattern: FP&A's $X,XXX.XX (ex Parminder's bonus) =
US severance ($X,XXX.XX) minus "Bonus Severance (Infrastructure)" lines
($X,XXX.XX = Parminder's package, both 611250 + 511175 components) PLUS
Uruguay egreso (UYU 163,567.97 ≈ $X,XXX.XX USD at ~40 UYU/USD). Germany
severance (€7,609 ≈ $X,XXX.XX USD) was NOT included in FP&A's number.

### AP Consulting LLC = Acme Labs Contractors (BillFlow)
Apr 2026: AP Consulting bills posted in March (Bill 001 $X,XXX.XX on 3/20,
Bill 002 $X,XXX.XX on 3/31) were both paid in April via BillFlow
(P26041601-5474524 on 4/17 for Bill 001, P26040901-4070844 on 4/10 for
Bill 002), totaling $X,XXX.XX = FP&A exact match. Bill 003 (posted 4/30,
$X,XXX.XX) paid in May 2026, not April. Lesson: FP&A's Acme Labs Contractors
number is the CASH OUTFLOW that month, which may be 1-2 months lagged
from the bill post date. Use BillFlow payments by Process Date, NOT
NetSuite bill date.

### Severance ex-Parminder ex-Uruguay-Egreso (April 2026 pattern)
FP&A may quote a "regular" US/Germany severance number excluding specific
named individuals' separation packages (e.g. "$X,XXX.XX excluding Parminder's
bonus"). The NetSuite total on 511175 + 611250 will be larger because it
includes:
- US payroll severance lines (regular)
- "Bonus Severance ({Department})" lines (likely the named exclusion)
- Uruguay JE titled `Egreso ({EmployeeName})` (separate severance package)
- Germany payroll severance lines (if any)
When variance > $1k, flag with the full line breakdown so the accountant can ask FP&A
what bucket of severance their number represents.

## Step 7 - Build the chat report

Output a markdown report with these sections (omit any section that has no
items):

```
## {Month} {YYYY} Cash Confirmations — Auto-Confirmation Report

### CONFIRMED (within $1 of FP&A figure)
| Section | Item | FP&A | Found | Source |
|---|---|---:|---:|---|
| Sublease Pay | Landlord B | $X,XXX.XX | $X,XXX.XX | BillFlow P26031701-9405268, 2026-03-13 |
| Other 1T | Law Firm D | $X,XXX.XX | $X,XXX.XX | BillFlow P26030501-7105043, 2026-03-06 |
| ... | | | | |

### VARIANCE > $1 (the accountant reviews)
| Item | FP&A | Found | Delta | Notes |
|---|---:|---:|---:|---|
| 1T/Retention Bonuses | $X,XXX.XX | $X,XXX.XX | -$X,XXX.XX | PUNT TO the Assistant Controller — 1T vs retention split unclear |

### NOT CONFIRMED — manual review needed
- **Legacy Hosting Stack Acquisition** ($X): no matching wire found in Chase x0001
- **{name}** (${amount}): no matching source rows

### FP&A OPEN QUESTIONS (FP&A asked but didn't give a $)
- **Interest Income**: not provided by FP&A — computed $X,XXX.XX from acct 711000
- **Term Loan Principal**: FP&A asked "thought it was $225K" — actual: $X,XXX.XX (matches FP&A's figure of $X,XXX.XX)

### Summary
- {N} confirmed
- {N} variance
- {N} not confirmed
- {N} FP&A open questions
```

Ascii-only. No emoji. No em-dashes (use `-` or `—` is fine via dash, but spell
"to"/"from"/"vs" per CLAUDE.md rule 7 for any text that might end up in a CSV
or JE memo).

After the report, tell the accountant:
> "Report complete. Review variances and not-confirmed items above. I have not
> drafted a reply email — you'll write the response to FP&A yourself."

---

## Out of scope (do NOT do)

- Drafting the email reply to FP&A. the accountant writes that himself.
- Any NetSuite writes (`ns_createRecord`, `ns_updateRecord`). This is read-only.
- Auto-pulling the BillFlow CSV from BillFlow. the accountant exports manually.
- Updating the debt schedule.
- Adding entries to `audit_log.json` (no NetSuite writes happen, so no audit).

---

## After running

If the accountant tells you new vendor name patterns or aliases that should be added,
update `./vendor_aliases.yaml`. The file is expected to grow over time.
