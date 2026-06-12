---
name: bank-statement-posting
description: >
  Process daily bank statement CSV drops from any of the 16 Acme Corp bank accounts
  (US: Chase x0001, Bank of America MMA x0002, Wells Fargo x0003; foreign: ING BV EUR/USD/GBP, TD,
  Barclays UK, Bank of America USD/PLN, Banco Santander UYU/USD, ING USD/EUR/GBP, ING Savings EUR).
  Per-account dispatch via accounts.yaml. Classifies every row, matches customer
  payments to NetSuite open invoices, posts customer payments / checks / deposits
  via MCP, generates JE Import CSVs for transfers and bank fees (per the
  CSV-only JE policy), and produces a multi-tab paste-back XLSX for the current
  month's banking sheet. Always use this skill when the user mentions: Chase, Wells Fargo,
  TD, Barclays, Bank of America, ING BV, ING, customer payments, lockbox deposits, bank fees,
  weekly cash, daily transactions, banking transactions CSV, foreign bank, posting
  payments, cash receipts, paste-back, or PYMT numbers.
---

# Bank Statement Posting Skill

This skill processes daily bank-statement CSV drops from any Acme Corp bank account,
classifies every row, posts what can be posted via MCP (customer payments, checks,
deposits), generates JE Import CSVs for what cannot (bank-transfer JEs, Adv IC JEs),
and produces a paste-back XLSX matching the current month's banking sheet layout.

**Four phases — never skip ahead. Phase 3 always requires the accountant's explicit approval.**

1. **Parse & Classify** — identify customer payment rows; deduplicate against NetSuite
2. **Match** — link each payment to a NetSuite customer and open invoice
3. **Review** — produce an Excel confirmation file; wait for approval
4. **Post & Write-back** — create NetSuite Customer Payment records; record PYMT numbers

---

## Per-account dispatch (read accounts.yaml at start of every run)

This skill is parameterized per bank account via `accounts.yaml` (sibling of this file).
At the start of every run:

1. Read `accounts.yaml`. It defines all 16 Acme Corp bank accounts with their `gl_account_id`,
   `subsidiary_id`, `currency`, `sheet_tab`, `csv_format`, `filename_glob`, and
   `bank_account_string` (the actual account number that appears in the CSV's account column).
2. Identify which account each input CSV belongs to:
   - First try matching the CSV's account-number column against `bank_account_string`.
   - Fall back to the filename matching `filename_glob`.
3. If the matched account has `status: stub`, STOP and tell the accountant:
   > "I don't have a parser for {account_key} yet. Drop a historical sample CSV at
   > `Weekly Cash Activities/Historical Exampels/{account_key}/sample.csv` and I'll build one."
4. Otherwise dispatch Phase 1 parsing to the function for that `csv_format`
   (`chase_standard`, `wellsfargo_activity_detail`, future foreign formats).
5. All hardcoded constants below get pulled from `accounts.yaml` at runtime, not memorized.

## NetSuite customerpayment posting — REQUIRED FIELDS (changed 2026-05-01)

NetSuite started enforcing the Payment Method field on `customerpayment` POSTs around
2026-05-01. Both fields below are now MANDATORY on every customer payment payload:

- **`customForm`** — `{"id": "70"}` (Standard Customer Payment)
  - Do NOT use form 146 (Acme Customer Payment) — it requires a SuiteScript-only field.
- **`paymentOption`** — pick by source row type:

  | Source row pattern | paymentOption.id | refName |
  |---|---|---|
  | LOCKBOX TRANSACTION (Chase) | `2` | Check |
  | PREAUTHORIZED ACH CREDIT (Chase) | `8` | ACH/EFT |
  | PRE-AUTHORIZED ACH DEBIT (Chase) | `8` | ACH/EFT |
  | INCOMING/OUTGOING MONEY TRANSFER (wire, Chase) | `9` | Wire |
  | Wells Fargo `ACH Credits` | `8` | ACH/EFT |
  | Wells Fargo `Money Transfer CR/DB - Wire` | `9` | Wire |

If you omit `paymentOption`, the API rejects with HTTP 400 `"Please enter value(s) for: Payment Method."`

---

## Reference files (read as needed)

- `references/originator_lookup.json` — Pre-built lookup table: 341 originator names
  extracted from 15 months of historical Chase transactions, each matched to a NetSuite
  customer name and internal ID with a confidence score. **Read this at the start of
  Phase 2** — it is your primary name-matching resource and will save many NetSuite calls.

---

## Phase 0: Duplicate Check — Query Already-Posted Payments

**Before classifying anything**, query NetSuite for customer payments already posted
for the date range covered by the CSV file. This prevents double-posting when the same
transaction appears in multiple Chase pull files.

Determine the earliest and latest Post Date in the CSV, then run:

```sql
SELECT t.id, t.tranid, t.trandate, t.amount
FROM transaction t
WHERE t.type = 'CustPymt'
  AND t.subsidiary = 1
  AND t.trandate >= TO_DATE('<earliest_date>', 'MM/DD/YYYY')
  AND t.trandate <= TO_DATE('<latest_date>', 'MM/DD/YYYY')
ORDER BY t.trandate
```

Build a deduplication set of `(trandate, amount)` pairs from the results.
In Phase 1, mark any CSV row whose `(Post Date, Amount)` matches an already-posted
payment as **[ALREADY POSTED — SKIP]** — do not include it in the review file.

Also deduplicate across the two CSV files if the accountant drops multiple pull files for the
same period (e.g., the 4/15 and 4/16 pulls both contain 4/13–4/15 transactions).
Read all CSV files in the 2026-04 folder and deduplicate rows by `(Post Date, Transaction Description, Amount)` before processing.

---

## Phase 1: Parse & Classify

Read the raw Chase CSV. Columns:
`Account Number`, `Account Name`, `Currency`, `Post Date`, `Status`,
`Transaction Description`, `Amount`, `Bank Reference`, `Customer Reference`, `Transaction Detail`

**Status field:**
- `F` = Final (settled) — process
- `I` = Intraday (not yet final) — **process by default** (the accountant confirmed 2026-04-17: intraday
  amounts almost always settle correctly, so include them in every run unless the accountant says to skip)

**Note on intraday amounts:** If an intraday payment is posted and the final settled amount
differs, NetSuite will need a manual correction. In practice this is rare. Do not skip intraday
customer payments unless the accountant explicitly requests it.

**The three customer payment transaction types to identify:**

| Transaction Description | Always a customer payment? |
|---|---|
| `LOCKBOX DEPOSIT` | Yes — always |
| `PREAUTHORIZED ACH CREDIT` | Usually — verify via originator name |
| `INCOMING MONEY TRANSFER` | Sometimes — verify via originator name |

**Skip these — they are not customer payments:**
- Any row with a negative `Amount` (debits)
- ACH credits where the originator matches a known non-customer pattern:
  FlexBenefits, BancorpSv/BANCORPSV, ChaseSETTLE, RetireBridge, PEO Provider Tech, BREX,
  Medical Carrier A, TOB - RENT, Zellerman, ASD- (ACH settlement), BILL.COM PAYABLES
- `TRANSFER` / `ACCOUNT TRANSFER` rows
- `MISCELLANEOUS DEBIT`, `ACCOUNT ANALYSIS FEE`, `FOREIGN EXCHANGE` rows
- **INCOMING MONEY TRANSFER where the originator is "ACME INC"** — these are
  internal transfers from Acme Corp's own Wells Fargo account to Chase (confirmed pattern:
  originator detail contains "ACME INC" and "Wells Fargo BANK"). Skip all of these.

**Extracting the originator name from `Transaction Detail`:**

For ACH credits, find text after `NAME:ORIG:` — the pattern is:
`NAME:ORIG:<COMPANY NAME> INF:` or `NAME:ORIG:<COMPANY NAME> RCVR`

Example: `NAME:ORIG:GM-GMC DIVISION VENDOR PMT INF:ORIG ID:...` → originator is `GM-GMC DIVISION`

Strip common trailing suffixes before looking up:
`VENDOR PMT`, `AP PAYMENT`, `PAYMENTS`, `PYMNT`, `EDIPYMT`, `PMD PAYMENT`,
`EDIPYMTS`, `SALES`, `CONTRACTS`, `CORP PYMNT`, `EFT`, `PAYABLES`, `EDI PYMNTS`,
`EDI PAYMNT`, `DISB`

For wire transfers, look for `ORIGINATOR:` in the detail field, or look for the
remitter name in the `INWARD REMITTANCE` block.
For lockbox deposits, there is no originator — match by amount only.

After classifying, summarize before proceeding:
> "Found X candidate customer payments (F: final, I: intraday): Y lockbox, Z ACH credits, W wires.
> Skipped N non-customer credits (including M already-posted duplicates). Proceeding to match..."

---

## Phase 2: Match to NetSuite

### Step 2a — Load the historical lookup table

Read `references/originator_lookup.json`. This file maps cleaned originator names
(uppercase) to `{netsuite_id, netsuite_name, confidence, frequency}`.

For each candidate ACH credit / wire transfer:
1. Clean the originator name: uppercase, strip suffixes (see Phase 1)
2. Look up the cleaned name in the JSON
3. If found with confidence ≥ 80 → use that NetSuite customer ID and name
4. If found with confidence 60-79 → use as a candidate but mark ⚠️ for review
5. If not found or confidence < 60 → mark as "No customer match — investigate"

**⚠️ CRITICAL: Customer numbers are NOT internal entity IDs.**
When the accountant says a customer is "703 TriHealth" or "customer 703", the number (703) is the
NetSuite customer number prefix — it is NOT the internal entity ID used in API calls.
Example: "703 TriHealth, Inc." has internal entity ID 38175 (confirmed 2026-04-17).
**Always resolve the internal entity ID by searching the entity table by name:**
```sql
SELECT id, altname, entityid FROM entity
WHERE type = 'customer' AND altname LIKE '%TriHealth%' AND isinactive = 'F'
```
Never assume a customer-facing number like "703" is the `id` to pass to `ns_createRecord`.

### Step 2b — Query NetSuite for open invoices

Once you have customer IDs, fetch open invoices in **one query** — not per-payment:

```sql
SELECT
  t.id,
  t.tranid,
  t.entity,
  e.altname        AS customer_name,
  t.foreigntotal   AS amount_remaining,
  t.trandate,
  t.duedate
FROM transaction t
JOIN entity e ON t.entity = e.id
WHERE t.type    = 'CustInvc'
  AND t.status  = 'A'
  AND t.subsidiary = 1
ORDER BY e.altname, t.trandate
```

Note: use `t.status = 'A'` for open invoices — `statusRef` is not exposed in SuiteQL search.

The Chase 5839 GL account ID is `773` (confirmed). No need to re-query for it.

### Step 2c — Match each payment to a specific invoice

**CRITICAL: Only report invoice matches that come from the real NetSuite query above.
Never guess, fabricate, or invent invoice numbers. If a real query match is not found,
say so explicitly. A wrong invoice number is worse than no match.**

**LOCKBOX DEPOSIT:**
- Match by exact `amount_remaining` on the open invoices list (to the cent)
- One exact match → ✅ High confidence
- Multiple matches → ⚠️ Flag all candidates; the accountant will pick
- No match → ⚠️ Flag: "No exact AR match — investigate (need lockbox remittance from Chase portal)"

**ACH CREDIT / WIRE TRANSFER (customer identified from lookup):**
- Search that customer's open invoices for an exact amount match
- One match → ✅ High confidence
- Multiple open invoices for that customer → ⚠️ Flag: show all open invoices, the accountant selects
- No exact single-invoice match → **check for split payment before flagging** (see below)

**Split payment detection (check before flagging as unmatched):**
When no single open invoice for a customer exactly matches the payment amount, check whether
the payment is a split across 2–3 invoices:
1. Sort the customer's open invoices by amount
2. Check all 2-invoice combinations: does invoice_A + invoice_B = payment amount (to the cent)?
3. If yes → ✅ Flag as split payment, list both invoices, mark ✅ Ready
4. Check 3-invoice combinations only if no 2-invoice match found
5. If no combination matches → ⚠️ Flag: show all open invoices for that customer

**Split payment posting (Phase 4b):**
Post as a single `customerpayment` record with multiple items in the `apply` sublist:
```json
"apply": {
  "items": [
    {"doc": {"id": "<invoice1_id>"}, "apply": true, "amount": <amount1>},
    {"doc": {"id": "<invoice2_id>"}, "apply": true, "amount": <amount2>}
  ]
}
```
Confirmed working: AnchorCo $X,XXX.XX = INV-US-##### ($X,XXX.XX) + INV-US-##### ($X,XXX.XX)
→ posted as PYMT#### on 2026-04-17.

**Edge case — originator is both a customer AND vendor in NetSuite:**
```sql
SELECT id FROM entity WHERE type = 'vendor' AND altname LIKE '%[name]%' AND isinactive = 'F'
```
If vendor match found, add flag: "⚠️ Also exists as vendor — confirm this is a customer payment"

---

## Phase 3: Review — Excel Output

**Do not post anything yet.** Generate an Excel review file and wait for the accountant's explicit approval.

Use openpyxl to create a file named `chase-review-[YYYYMMDD].xlsx` with two sheets:

### Sheet 1: "Review & Approve"

**Columns:**
| Col | Header | Content |
|-----|--------|---------|
| A | # | Row number |
| B | Date | Post Date from CSV |
| C | Type | Lockbox / ACH CR / Wire |
| D | Amount | $ amount |
| E | Matched Customer | NetSuite customer name (or "NO MATCH") |
| F | Invoice (tranid) | Real invoice number from NetSuite query (or blank if none found) |
| G | Invoice Amount | Amount remaining on matched invoice |
| H | NS ID | NetSuite customer internal ID |
| I | Status | ✅ Ready / ⚠️ Review / ❗ Investigate |
| J | Flag / Notes | Reason for review flag |
| K | **APPROVE?** | **Leave blank — the accountant fills in Y / N / invoice number** |

**Color coding:**
- ✅ Ready rows: light green background
- ⚠️ Review rows: light yellow
- ❗ Investigate rows: light orange/red

**Key rule on Column F:** If you could not find a real open invoice from NetSuite that
matches this payment, leave it blank or write "Not found in NS". Never write a made-up
invoice number. the accountant will see that a match wasn't found and handle it manually.

### Sheet 2: "Skipped Transactions"

List all transactions that were classified as non-customer payments with the reason they
were skipped (debit, known non-customer originator, already posted, etc.). This is for
the accountant's audit trail.

Save the Excel file to the workspace folder and share the link.

Then ask:
> "Please open the Excel file, review the 'Review & Approve' tab, and fill in Column K:
> - **Y** = post this payment as shown
> - **N** = skip / investigate
> - Write the correct invoice number if the wrong one was matched
>
> Reply 'done' when you've filled it in and I'll process your approvals."

---

## Phase 4: Post to NetSuite & Write Back

### 4a — Read the accountant's approvals from the Excel file

Re-read the Excel file. Process every row where Column K = "Y" or contains an invoice
number override.

### 4b — Post each approved payment

For each approved row, create a `customerpayment` record using `ns_createRecord`.

**Approval rule:** Customer payments applied to one or more specific open invoices are
**exempt** from the `approved: false` requirement that applies to JEs (see
`_shared/approval-required.md`). They auto-approve here because they're routine AR
application — the accountant's senior-accountant role is appropriate to clear them. Do NOT include
the `approved` field in the payload.

**CRITICAL — use this exact payload structure** (confirmed working 2026-05-05; `arAcct` requirement added 2026-05-28):

```json
{
  "customForm": {"id": "70"},
  "paymentOption": {"id": "<2 | 8 | 9 - see Payment Method table at top of file>"},
  "customer": {"id": "<NetSuite customer internal ID as string>"},
  "tranDate": "<YYYY-MM-DD>",
  "account": {"id": "<accounts[active].gl_account_id>"},
  "arAcct": {"id": "237"},
  "subsidiary": {"id": "<accounts[active].subsidiary_id>"},
  "payment": <amount as number, exact cents>,
  "autoApply": false,
  "apply": {
    "items": [
      {
        "doc": {"id": "<invoice internal ID as string>"},
        "apply": true,
        "amount": <payment amount>
      }
    ]
  }
}
```

**`arAcct` MUST be `237`** (`121100 Accounts Receivable : Accounts Receivable`) for every sub-2 USD customer payment.

**CRITICAL FIELD-NAME CASE BUG (root cause identified 2026-05-28)** — the NS REST field name is **`arAcct`** (camelCase). NS REST is case-sensitive on JSON property names. If you write `"aracct"` (all-lowercase), NS silently drops the field and the customForm 70 default cascades — and in this NS instance that default is misconfigured to `id: 112` (`211000 Accounts Payable`), NOT `237` (AR). Result: the CR leg posts to AP, invoices stay open, applied amount = 0.
- Always type `"arAcct"` exactly as shown — camelCase, capital A in Acct.
- The NS metadata confirms this: `ns_getRecordTypeMetadata("customerpayment")` returns the property as `arAcct`.
- Underlying NS-config bug (out of scope for this skill): customForm 70's "Default A/R Account" is set to 211000 instead of 121100. A NS admin should fix the form so the default cascade lands on AR — until then, every API customerpayment MUST send `arAcct` explicitly with correct case.

**Verification (run after every customer payment POST batch):**
```sql
SELECT tl.transaction, BUILTIN.DF(tl.expenseaccount) AS gl_account, tl.foreignamount
FROM transactionline tl WHERE tl.transaction IN (<new payment ids>) ORDER BY tl.transaction, tl.id
```
Expected: line 0 = bank cash GL (debit positive), line 1 = `121100 Accounts Receivable` (credit negative). If you see `211000 Accounts Payable`, fix immediately via:
```
ns_updateRecord(recordType="customerpayment", recordId=<id>, data='{"arAcct":{"id":"237"}, "apply":{"items":[{"doc":<inv_id>, "apply":true, "amount":<amt>}]}}')
```
- Caveat: even with `arAcct` set correctly on create, the `apply` sublist is often not honored on initial POST (NetSuite returns a static-sublist error on certain payloads). Confirm `unapplied == 0` and invoice `foreignamountunpaid == 0` on each payment; if not, run `ns_updateRecord` to set `apply.items` and re-verify.

**For split payments (one payment applied to multiple invoices):**
```json
"apply": {
  "items": [
    {"doc": {"id": "<invoice1_id>"}, "apply": true, "amount": <amount1>},
    {"doc": {"id": "<invoice2_id>"}, "apply": true, "amount": <amount2>}
  ]
}
```

**Both `customForm: 70` AND `paymentOption` are MANDATORY.** Omitting either causes
the API to reject with HTTP 400 `"Please enter value(s) for: Payment Method."` See the
"NetSuite customerpayment posting — REQUIRED FIELDS" section near the top of this file
for the lockbox→2 / ACH→8 / wire→9 mapping.

**The `account` and `subsidiary` IDs come from `accounts.yaml`** based on which bank
account this row belongs to (Chase x0001 → 773/2, Chase x0002 → 774/2, Wells Fargo x0003 → 679/2,
foreign accounts → values in their accounts.yaml entry).

#### 4b-foreign — Currency-aware customer payments (ING BV, Bank of America PLN, etc.)

Foreign-currency customer payments require **`currency`** in the payload, in addition
to the standard fields. `subsidiary` and `account` must match the foreign sub's GL.

```json
{
  "customForm": {"id": "70"},
  "paymentOption": {"id": "<2|8|9>"},
  "customer": {"id": "<NS customer ID>"},
  "tranDate": "<YYYY-MM-DD>",
  "account": {"id": "<accounts[active].gl_account_id>"},
  "arAcct": {"id": "237"},
  "subsidiary": {"id": "<accounts[active].subsidiary_id>"},
  "currency": {"id": "<accounts[active].currency_id>"},
  "payment": <amount in foreign currency>,
  "autoApply": false,
  "apply": {"items": [{"doc": {"id": "<invoice id>"}, "apply": true, "amount": <amt>}]}
}
```

`arAcct` stays `237` even on foreign-sub payments — `121100 Accounts Receivable` is the consolidated control account; do not switch to a sub-specific AR sub-account unless explicitly told.

**Currency internal IDs** (verified 2026-05-19 against NS `currency` table):
| Symbol | ID |
|---|---|
| USD | 1 |
| GBP | 2 |
| CAD | 3 |
| EUR | 4 |
| PLN | 5 |

**Matched-currency case** (transaction currency = invoice currency = sub base currency):
NS auto-fills the exchange rate from the daily rate table. No `exchangeRate` field needed.

**Mixed-currency edge case** (rare on ING BV): a transfer credited as EUR but the
bank's Detail Information shows `AMT CHF 4,300.00 / ORIGINAL AMT CHF 4,300.00` because
the wire originated in CHF and Bank of America/Chase converted. The customer payment is still
posted in EUR (the receiving account currency); the customer's open invoice should
also be in EUR. If the open invoice is in a third currency, post in the invoice's
currency and let NS compute the FX rate. Do not invent an `exchangeRate` value from
the bank narrative — let NS pull the daily rate.

You can post multiple payments in parallel (3-4 at a time) — NetSuite handles
concurrent creates without issue.

After each `ns_createRecord` succeeds, retrieve the PYMT document number:
```
ns_getRecord(recordType="customerpayment", recordId=<returned ID>, fields="tranId,payment,tranDate")
```
The `tranId` field is the `PYMT####` number. Record it.

### 4c — PYMT numbers feed into the paste-back XLSX

PYMT numbers are written back via the multi-tab paste-back XLSX produced in Phase 5
(see new layout below). Each posted customer payment ends up on its bank account's tab
with `Posted=*` and `Ref #=PYMT####`. There is NO separate writeback CSV step —
the XLSX is the single artifact the accountant pastes into the Banking Sheet.

### 4d — Final summary

```
✅ Posted X payments totaling $X,XXX,XXX.XX
⚠️  Y payments skipped / held for investigation

Posted:
  PYMT#### | $X,XXX.XX | AnchorCo Inc.  | INV-US-#####
  PYMT#### | $X,XXX.XX | University Hospital System     | INV-US-#####
  ...

Held:
  $X,XXX.XX  | Lockbox | No exact AR match
  $X,XXX.XX  | MediaCo Inc. | Multiple open invoices — needs selection

📋 Write-back CSV: [link]
```

---

## Phase 5: Paste-back XLSX — multi-tab, all 16 accounts

After Phase 4 is complete (or at any point the accountant drops a new raw CSV), produce a single
XLSX with one tab per bank account from `accounts.yaml`. Each tab matches the May-26
Banking Sheet transaction-tab column layout exactly. Tab order matches the sheet (1-16).

**This phase replaces the manual copy-paste grind.** The goal: the accountant drops a CSV (or three),
Claude posts what can be posted (customer payments / checks / deposits via MCP), and hands
back one XLSX with each tab ready to copy-paste into the corresponding tab of the live
Banking Sheet.

---

### Step 5a — Read the current Banking Sheet from Google Drive

Use the Google Drive MCP tool (`read_file_content`) to read the current month's Banking
Sheet. Sheet ID rotates monthly. **Use month rotation logic** rather than a hardcoded ID.

Compute the current month from `tranDate` (not today's clock — end-of-month files belong
to the prior month's sheet). Format: `{MonShort}-{YY} Banking Sheet` (e.g.,
`May-26 Banking Sheet`). Search Google Drive for that title; pick most recent.

**Known IDs (cache; refresh on title-mismatch):**
- `Apr-26 Banking Transactions` — `1oBU3pFRAUBqcGDNrkJ1fMTSjah04fSl0QaIsnWyYpCs`
- `May-26 Banking Sheet` — `1S9JvwUM12UmsXYgfSMrDfsXKNdWw8ejsa5KNvzslA0o`

Tabs to reference (per `accounts.yaml.sheet_tab` for each account). The May-26 sheet has
16 transaction tabs at the bottom of the workbook in the order shown in `accounts.yaml`.

For dedup: read the active tab(s) and extract already-entered rows as
`(Date, Amount, Description)` tuples. Combined with NetSuite Phase 0 dedup, this is the
"already entered" set.

---

### Step 5b — Compare raw CSV against already-entered set

Take all rows from the raw Chase CSV(s) for the month. For each row:
- If `(Post Date, Amount, Transaction Description)` is in the already-entered set → **skip** (already in sheet)
- If not → **gap row** that needs to be classified and added

Build the list of gap rows.

---

### Step 5c — Classify all gap rows

For every gap row, assign the following fields based on the classification logic below:

| Field | How to determine |
|-------|-----------------|
| **Type** | See classification table below |
| **TAG** | See classification table below |
| **WEEK** | Derive from Post Date: WK1=days 1–7, WK2=days 8–14, WK3=days 15–21, WK4=days 22–28, WK5=days 29–31 |
| **Assignee** | the accountant for customer payments & Chase-owned items; the AP Specialist for benefits/FSA/HSA; the Assistant Controller for wires/transfers |
| **Posted** | `*` if already posted to NetSuite; blank if not yet posted |
| **JE #** | PYMT#### for customer payments just posted; JE number if known; blank otherwise |
| **Notes** | Customer name for customer payments; payee/description for others |

**Classification table** (tag values per the new tag dictionary, effective 2026-05-15):

| Transaction Description | Originator / Pattern | Type | TAG |
|---|---|---|---|
| PREAUTHORIZED ACH CREDIT (customer) | Verified customer from `originator_lookup.json` | Customer | `N/A - Customer` |
| LOCKBOX DEPOSIT | — | Customer | `N/A - Customer` |
| INCOMING MONEY TRANSFER (customer) | Third-party customer originator (not Acme Corp) | Customer | `N/A - Customer` |
| INCOMING MONEY TRANSFER (internal) | Originator = ACME INC / Wells Fargo BANK | Transfer | `N/A - TRN` |
| ACCOUNT TRANSFER (Chase x0001 ↔ x0002) | RMKS TO/FROM 081003* or 082070* | Transfer | `N/A - TRN` |
| FOREIGN EXCHANGE DEBIT | Cross-sub cash funding wire (US→Canada via Adv IC JE) | Transfer | `N/A - TRN` |
| PREAUTHORIZED ACH CREDIT (non-customer credit) | Hanover, Former FSA Administrator COBRA, vendor overpayment refunds | MISC | `N/A - MISC` |
| LOCKBOX DEPOSIT (unmatched / customer refund credit) | No AR match — investigate | MISC | `N/A - MISC` |
| PREAUTHORIZED ACH DEBIT | RetireBridge / RET PLAN | 401K | `BEN` |
| PREAUTHORIZED ACH DEBIT | FlexBenefits / BancorpSv / WAGEWORKS | Benefits | `BEN` |
| PREAUTHORIZED ACH DEBIT | Medical Carrier A, Regional Medical Plan, HSA Custodian, CNA | Benefits | `BEN` |
| PREAUTHORIZED ACH DEBIT | State DOR / FTB / IRS / CRA | TAX | `TAX` |
| PREAUTHORIZED ACH DEBIT | BILL.COM PAYABLES | Bills | `BILL` |
| PREAUTHORIZED ACH DEBIT | BREX INC. BREX REIMB (small individual amounts) | EXP | `EXP` |
| PREAUTHORIZED ACH DEBIT | BREX INC. (bulk monthly statement settlement) | BREX | `BREX` |
| PREAUTHORIZED ACH DEBIT | PEO Provider Tech, Verizon, Springfield/TOB rent, CNA, vendor | Bills | `DOM` |
| OUTGOING MONEY TRANSFER | Workhorse HR payroll wires (DD/Tax) | Payroll | `PAYUS` |
| PRE-AUTHORIZED ACH DEBIT | KRONOS SAASHR (Workhorse HR garnishment) | Payroll | `PAYUS` |
| OUTGOING MONEY TRANSFER | Final-pay wire to terminated employee | Payroll | `PAYUS` |
| OUTGOING MONEY TRANSFER | International AP vendor wire | Bills | `INT` |
| OUTGOING MONEY TRANSFER | Customer refund wire (Acme Corp → overpaying customer) | Customer | `CUST` |
| ACCOUNT ANALYSIS FEE / MISCELLANEOUS FEES | Bank fees | Bank Fees | `DOM` |
| MISCELLANEOUS FEES | Chase LC fees | LOCINT | `LOCINT` |
| Term Loan principal ACH | Quarterly TL principal | DEBT | `LOAN` |
| LOC principal paydown | LOC principal | DEBT | `LOC` |
| Term Loan interest ACH | Quarterly TL interest | DEBT | `LOANINT` |
| Other ACH debits not matching above | Unknown — flag for review | UNKNOWN | (mark `MISC` and flag) |

For any row that can't be confidently classified, mark Type as `UNKNOWN` and Notes as
`"Needs manual classification"` so the accountant can handle it.

**ING BV classifier (`ing_bv_eur` / `ing_bv_gbp` / `ing_bv_usd`, sub 4 NL B.V.):**

Match on `Transaction Type Name` + `Detail Information` from the tab-separated ING BV CSV.
Originator lookup file: [references/ing_bv_lookup.json](references/ing_bv_lookup.json).

| Transaction Type Name | Detail pattern | Type | TAG | Owner | Posting |
|---|---|---|---|---|---|
| `Charges and Other Expenses` | `ACCOUNT MAINTENANCE FEE FOR <MMM><YYYY>` | Bank Fees | `INT` | the accountant | Phase 7h auto-Check `ACH_Debit_MM.DD.YY_ChaseBV` |
| `Charges and Other Expenses` | Paired with `INWARD REMITTANCE` (same Bank Reference) | Bank Fees | `INT` | the accountant | Phase 7h auto-Check `ACH_Debit_MM.DD.YY_ChaseBV` — memo "Bank fees for accepting customer payment PYMT####" |
| `Charges and Other Expenses` | Paired with `OUTWARD REMITTANCE` (wire fee) | Bank Fees | `INT` | the accountant | Phase 7h auto-Check `ACH_Debit_MM.DD.YY_ChaseBV<CCY>` |
| `Transfer` | `INWARD REMITTANCE` / `SEPA INCOMING` + originator matches lookup | Customer | `N/A - Customer` | the accountant | Phase 4 customerpayment in foreign currency (4b-foreign) |
| `Transfer` | `INWARD REMITTANCE` / `SEPA INCOMING` + originator = `ACME INC` / `TRANSFER OF ING` / similar | Transfer | `N/A - TRN` | the accountant | classify-only (cross-account internal funding; cross-sub leg handled by `adv-interco-je`) |
| `Transfer` | `OUTWARD REMITTANCE` + `UK Payroll Processor CLIENT ACCOUNT` (GBP) | Payroll | `PAYUK` | the accountant | Phase 7g auto-Check `ACH_Debit_MM.DD.YY_PAYUK{N}` |
| `Transfer` | `OUTWARD DOMESTIC REMITTANCE` (EUR, recurring beneficiary on payroll whitelist) | Payroll | `PAYBV` | the accountant | Phase 7g auto-Check `ACH_Debit_MM.DD.YY_PAYBV{N}` — NL payroll provider TBD, currently classify-only until whitelist seeded |
| `Transfer` | `OUTWARD REMITTANCE` + Dutch tax authority (Belastingdienst) | Tax | `TAX` | the accountant | Phase 7h auto-Check `ACH_Debit_MM.DD.YY_BVCIT` (DR Dutch VAT/CIT GL — verify with the Assistant Controller's prior history) |
| `Transfer` | `OUTWARD REMITTANCE` other | Bills | `INT` | the accountant | classify-only (the AP Specialist books AP) |
| `Miscellaneous` | `INITIAL FEE` (one-off, June 2025) | Bank Fees | `INT` | the accountant | classify-only — historical |

**Bank of America PLN classifier (`bofa_pln`, GL 560, sub 6 Poland, PLN):**

Match on `Description` (and `Beneficiary/ Remitter` if non-blank) from the Bank of America PLN XLSX.
Lookup file: [references/bofa_pln_lookup.json](references/bofa_pln_lookup.json).

| Description / Beneficiary | Type | TAG | Owner | Posting |
|---|---|---|---|---|
| `OPLATA MIESIECZNA ZA OBSLUGE KARTY DEBETOWEJ` | Bank Fees | `INT` | the AP Specialist | Phase 7h auto-Check `ACH_Debit_MM.DD.YY_PLN` |
| `LOCAL PAYMENTS CHARGES`, `CHARGE`, `ACCOUNTS MAINTENANCE CHARGE`, `INTERNAL DEBIT` | Bank Fees | `INT` | the AP Specialist | Phase 7h auto-Check (same monthly bundle) |
| Beneficiary = `PL Social Insurance` (ZUS) | Payroll | `PAYPN` | the AP Specialist | Phase 7g auto-Check (vendor 1111116) |
| Beneficiary = `PL National Tax Admin` (KAS) | Payroll | `PAYPN` | the AP Specialist | Phase 7g auto-Check (vendor 1111117) |
| Beneficiary = `PL Regional Tax Office` (regional tax office) | Tax | `TAX` | the AP Specialist | classify-only — the AP Specialist books manually |
| Beneficiary = `PL Pension Fund (PPK)` (PPK pension) | Payroll | `PAYPN` | the AP Specialist | classify-only until vendor created in NS, then Phase 7g auto-Check |
| Beneficiary = `PL Health Benefits Provider` | Bills | `INT` | the AP Specialist | classify-only (the AP Specialist Bill Pay) |
| Beneficiary = `PL Payroll Advisory` | Bills | `INT` | the AP Specialist | classify-only |
| Beneficiary = `PL Telecom` | Bills | `INT` | the AP Specialist | classify-only |
| Beneficiary = recognized Polish employee (Employee_PL_001, Employee_PL_002, etc.) | Payroll | `PAYPN` | the AP Specialist | Phase 7g auto-Check IF employee NS ID resolved AND on whitelist; else classify-only |
| Beneficiary = unrecognized | UNKNOWN | `MISC` | the AP Specialist | classify-only — the accountant/the AP Specialist to review |

**TD Canada classifier (account `td_canada`, GL 231, subsidiary 3, CAD):**

Match on the concatenation `Description 1 | Description 2` from the TD CSV. Full
detail in [.claude/skills/bank-statement-posting/references/td_canada_lookup.json](.claude/skills/bank-statement-posting/references/td_canada_lookup.json).

| Description 1/2 pattern | Type | TAG | Owner | Posting |
|---|---|---|---|---|
| `Direct Deposits (PDS) service total / GRADS8159700000` | Payroll | PAYCA | the accountant | classify-only (canada-payroll skill posts JE) |
| `Funds transfer credit / TT ACME CORP IN` | Transfer | N/A | the accountant | classify-only (US side `adv-interco-je` covers both legs) |
| `Funds transfer fee / TT ACME CORP IN` | Bank Fees | INT | the AP Specialist | classify-only |
| `Investment / CANLIFE-GPAS` | 401k | BEN | the AP Specialist | classify-only |
| `Misc Payment / HSA CAD CAN` | Bills | BEN | the AP Specialist | classify-only |
| `COMM CASH MGMT / AIRWALLEX` | Expense | EXP | the AP Specialist | classify-only (Brex card pmts) |
| `COMM BILL PAYMENT / AIRWALLEX` | Expense | EXP | the AP Specialist | classify-only (Brex card pmts) |
| `Business PAD / SCB Ltd.` | Benefits | BEN | the AP Specialist | classify-only (Medical Premium Broker medical) |
| `Web payment / WIRE####` | Bills | INT | the AP Specialist | classify-only |
| `Bill Payment / KINDERA LIVING` | Customer | N/A | the accountant | **POST customerpayment** (CAD, account 231, sub 3) |
| `Funds transfer credit / TB ROYAL BK OF` | Transfer | N/A | the accountant | review (rare) |
| `Item returned NSF` | MISC | N/A | the accountant | classify-only |
| `Monthly fee` / `Activity fee` / `NSF item fee` | Bank Fees | INT | the AP Specialist | classify-only (could post as Check; ask the accountant) |

> **Foreign-account tag rule (per the accountant 2026-05-27):** TD is a foreign bank account, so bills / invoices / payments / bank fees are tagged **`INT`**, never `DOM`. `DOM` is reserved for **US accounts only** (Chase x0001/x0002, Wells Fargo). The same rule applies to every foreign account (ING BV, Barclays UK, Bank of America PLN, etc.) — they never use `DOM`.
| `Bill payment - 8369 / CRA TX ISTAL` | Tax | TAX | the accountant | review (CRA installment) |

Cross-account dedup: `Funds transfer credit / TT ACME CORP IN` on the Canada side
pairs with the OUTGOING WIRE on Chase x0001 the same day. The US side runs through
`adv-interco-je` (Adv IC JE pattern, e.g. JE#####) which posts BOTH the US cash leg
AND the Canada cash leg of the JE in CAD. So the bank-statement-posting skill MUST
NOT post the Canada inflow as a separate JE — it's classify-only on the Canada tab.

**Assignee rules:**
- `the accountant` — customer payments, and any Chase credits/debits the accountant owns
- `the AP Specialist` — anything related to benefits, FSA, HSA (FlexBenefits, BancorpSv, RetireBridge, Medical Carrier A, WAGEWORKS)
- `the Assistant Controller` — wire transfers, tax payments, loan payments

---

### Step 5d — Generate the paste-back XLSX (multi-tab, all 16 accounts)

Run the canonical script `scripts/generate_paste_back.py` (sibling of this SKILL.md). It
reads `accounts.yaml`, parses every CSV in the current month folder via the matching
`csv_format` parser, applies the classifier from Step 5c, and writes one tab per account.

**Output XLSX columns (matches May-26 Banking Sheet exactly):**

```
Date | Description | Amount | Tag | Owner | Week | Posted? | Ref #
```

**Strict field rules — DO NOT deviate:**

| Field | Rule |
|---|---|
| Date | `M/D/YYYY` (no leading zeros) |
| Description | **Customer Conference VERBATIM** copy of the bank's `Transaction Detail` (Chase) / `Reference Detail` (Wells Fargo) / equivalent field for foreign banks. NEVER edit, summarize, truncate, or annotate. The point is forensic traceability — the accountant must be able to grep this column and find the original bank-statement row. |
| Amount | Signed number. Negative for outflows, positive for inflows. No currency symbols. |
| Tag | See full tag dictionary below — bare tags for outflows + categorized inflows. Reverted policy 2026-05-15: suffixed `N/A - Customer`, `N/A - TRN`, `N/A - MISC` ARE now the configured values for the three inflow types. Never plain `N/A` or plain `MISC` going forward. |
| Owner | `the accountant` for customer payments, transfers, payroll, tax, debt service, FlexBenefits/Bancorp (FSA/HSA), and the routine TD patterns ownership already flipped over. `the AP Specialist` for remaining benefits/BillFlow/Brex-adjacent items still in her queue. |
| Week | `WK1`-`WK5` format, **business-day weeks anchored on Mondays** (effective 2026-05-15). Algorithm: count Mondays strictly after day 1 of the month through the date — `WK{count + 1}`. May 2026 example: 5/1 (Fri) = WK1, 5/4-5/8 = WK2, 5/11-5/15 = WK3, 5/18-5/22 = WK4, 5/25-5/29 = WK5. Handles month-starts-on-Monday correctly (June 6/1 alone = WK1, 6/8 = WK2). Never `WEEK 1`. |

#### Tag dictionary (effective 2026-05-15)

The Tag column on every paste-back row maps to one of these. Same dictionary applies to **all 16 bank accounts** going forward.

**Outflows (negative amount):**

| Tag | Use case | Source-row signal |
|---|---|---|
| `BILL` | AP - BillFlow sweeps | `BILL.COM PAYABLES` originator on Chase; bulk multi-vendor ACH on Wells Fargo |
| `DOM` | AP - Domestic. **US accounts ONLY (Chase x0001/x0002, Wells Fargo).** One-off domestic vendor wires (Marsh McLennan, CHRO Assembly Series, Verizon, CNA, Springfield rent, TOB rent, etc.); US-side bank fees. NEVER use DOM on a foreign account. | One-off domestic vendor wires; US bank fees |
| `INT` | AP - International. **Every foreign account uses this for bills/invoices/payments/bank fees** (TD, ING BV, Barclays UK, Bank of America PLN, etc.) — i.e. what would be DOM on a US account is INT on a foreign account. Also non-US vendor wires from US accounts. Per the accountant 2026-05-27. | Any AP/fee/payment on a foreign bank account; international wire originators |
| `EXP` | Expense reports (Brex Reimb individual employee reimbursements via ACH) | `NAME:ORIG:BREX INC. BREX REIMB` detail; small individual amounts |
| `TAX` | Tax payments to state/federal authorities | State Dept of Revenue, FTB, IRS, CRA; Wells Fargo tax checks |
| `BREX` | Brex credit card statement payments (bulk monthly settlement) | `NAME:ORIG:BREX INC.` detail with larger bulk amount (>$200K range) |
| `CUST` | Customer refunds (Acme Corp paying back a customer who overpaid) | Outgoing wire/ACH to a NetSuite customer entity |
| `BEN` | Benefits — FSA, HSA, medical/dental insurance, 401k contributions | FlexBenefits, BancorpSv FlexBenefits-DBI, Cigna, Regional Medical Plan, HSA Custodian, RetireBridge |
| `PAYUS` | US payroll wires | Workhorse HR DD/Tax/Garnishment wires (3 per pay date); final-pay employee wires to US bank accounts |
| `PAYCA` | Canada payroll | TD `Direct Deposits (PDS) service total / GRADS8159700000` (ADP Canada) |
| `PAYDE` | Germany payroll | Germany payroll wires |
| `PAYBV` | Netherlands (Acme B.V.) payroll | NL payroll wires from ING BV EUR — `OUTWARD DOMESTIC REMITTANCE` to NL payroll provider |
| `PAYPN` | Poland payroll | Poland payroll wires |
| `PAYUK` | UK payroll | UK payroll wires |
| `PAYUY` | Uruguay payroll | UY payroll wires |
| `LOAN` | Chase Term Loan principal payment | Quarterly principal hits |
| `LOC` | Chase LOC principal payment | LOC principal paydowns |
| `LOANINT` | Chase Term Loan interest | Quarterly TL interest (`ACH_Debit_*_LOANINT`) |
| `LOCINT` | Chase LOC interest + LC fees | LC charges, unused commitment fees |

**Inflows (positive amount):**

| Tag | Use case | Source-row signal |
|---|---|---|
| `N/A - Customer` | Customer payment applied to AR invoice | LOCKBOX DEPOSIT (Chase); PREAUTHORIZED ACH CREDIT from a verified-customer originator (lookup hit); INCOMING MONEY TRANSFER from third-party customer (not "ACME INC"); Wells Fargo ACH Credits from customer; TD Bill Payment / KINDERA LIVING |
| `N/A - MISC` | Miscellaneous credit — refunds, NSF returns, mystery credits | Hanover/insurance refunds, Former FSA Administrator COBRA reimbursements, vendor overpayment returns, FlexBenefits FSA Settlement Credits, self-refunds, unmatched lockbox credits, NSF returns |
| `N/A - TRN` | Internal transfer — between Acme Corp's own accounts (intra- or cross-sub) | Chase `TRANSFER` rows between x0001/x0002; INCOMING MONEY TRANSFER from `ACME INC`; `FOREIGN EXCHANGE DEBIT` (cross-sub FX wire); Wells Fargo `Money Transfer DB - Wire` (sweep out); TD `Funds transfer credit / TT ACME CORP IN`. Same tag for both directions — sign of Amount determines in/out. |
| Posted? | `*` if posted to NetSuite this session OR found by Phase 0 dedup. Blank if classify-only (the AP Specialist) or pending CSV upload. |
| Ref # | `PYMT####` for customer payments; `JE####` for transfers; `ACH_Debit_MM.DD.YY_<bank>` for bank-fee Checks; blank for classify-only or pending CSV upload. |

**Output file:** `Weekly Cash Activities/{YYYY-MM}/banking-transactions-paste-{YYYYMMDD}.xlsx`

**Tab order matches `accounts.yaml`:** every account gets a tab, even if the account had
no rows this run (the empty tab is fine — the accountant still pastes nothing into the corresponding
sheet tab).

Sort within each tab: **Date ASC (oldest first)**. The Google Sheet tabs grow in
chronological order — pasting our oldest-first XLSX content directly underneath the
existing rows is a clean append, no manual reordering. Within a date, preserve source
CSV row order (stable sort).

Then tell the accountant:
> "Paste-back XLSX ready at `{path}`. Each tab maps 1:1 to a tab in the {Month}-{YY}
> Banking Sheet. Customer payments / checks already posted are marked Posted=`*` with
> PYMT/Check tranid in Ref #. Pending JE CSV imports show blank Ref # — fill in after
> upload."

---

### Step 5e — Month transition checklist

At the start of a new month, before running the skill:
1. Compute target month from `tranDate`. Search Google Drive for `{MonShort}-{YY} Banking Sheet`.
2. If not found, fall back to last known month's sheet ID (Step 5a) and warn the accountant.
3. CSV folder path is `Weekly Cash Activities/{YYYY-MM}/` — auto-derived; no skill edit needed.

---

---

## Phase 6: Transfer Posting

Bank-to-bank transfers that move cash between Acme Corp accounts. These appear in the
banking sheet as Type=Transfer.

### 6a — Same-subsidiary transfers (journalentry) — CSV UPLOAD ONLY

Use when moving cash between two accounts that are both under Acme, Inc. (e.g.,
Wells Fargo x0003 → Chase x0001, or Chase x0001 → Chase x0002).

**🚨 MCP-based JE posting is suspended (effective 2026-05-01). DO NOT use `ns_createRecord`.** Generate a CSV instead. See `_shared/approval-required.md` for the policy.

**CSV file (one per run, all same-subsidiary JEs combined):**
`Weekly Cash Activities/{YYYY-MM}/JE Imports/{YYYY-MM-DD} Cash Transfers JE Import.csv`

**Consolidation rule (effective 2026-05-14, per the accountant):** every same-subsidiary JE from
this run (Wells Fargo sweeps, x0001↔x0002 transfers, any other plain journalentry) goes into
**ONE CSV**. NetSuite Import Assistant groups lines into separate JEs by **External ID**.
This way the accountant uploads one workbook instead of one per JE.

Columns: `External ID, Date, Journal Entry Memo, Account, Debit, Credit, Line Memo, Subsidiary, Department`. Date format `M/D/YYYY`. Subsidiary path quoted because of comma. Liability lines have no department.

**🚨 CSV Account column MUST use full NS hierarchy, NOT leaf-only (rule added 2026-05-29 after upload failure).** The Import Assistant errors with `Invalid account reference key {leaf} for subsidiary N` if you use just the leaf name. Use `{acctnumber} {fullname}` where fullname is the **complete** path from NS.
- ❌ Wrong: `111070 Chase Checking x0001`
- ✅ Right: `111070 Cash and Cash Equivalents : Chase Checking x0001`
- ❌ Wrong: `211900 Intercompany Payables`
- ✅ Right: `211900 Accounts Payable : Accounts Payable : Intercompany Payables`

**Authoritative lookup:** [`references/ns_account_paths.json`](references/ns_account_paths.json) — every account used in CSV uploads with its canonical `csv_format` string. If an account isn't there, pull `acctnumber` + `fullname` via `SELECT id, acctnumber, fullname FROM account WHERE id = N` and add it to the file. Same rule applies to plain Journal Entry CSVs AND Adv IC JE CSVs.

**External ID naming:** `{MonShort}{DD}JE-{short-key}` (max 30 chars). Examples:
- `May14JE-x0002Tx0001` — x0002 to x0001 same-sub transfer
- `May14JE-Wells Fargo5839Sweep` — Wells Fargo sweep into Chase x0001
- `May14JE-Sched` — Springfield deposit reclass
Keep the date part = the date the JEs are uploaded (so a single file uploaded on 5/14
gets all `May14JE-*` IDs even if the underlying tranDates span 5/13-5/14).

**Adv IC JE stays separate** — Adv IC JEs use the `advintercompanyjournalentry` record
type and a DIFFERENT Import Assistant ("Adv IC JE Import"), so they MUST be in their own
CSV with the Adv IC JE column layout (`External ID, Header Memo, Header Subsidiary, Date,
Account, Journal Entry Currency, Debit, Credit, Exchange Rate, Memo, NAME, VENDOR,
Department, DUE TO/FROM SUBSIDIARY, Eliminate, Line Subsidiary`). File naming:
`{YYYY-MM-DD} {Direction} Adv IC JE.csv` (e.g. `2026-05-12 Canada FX Cash Transfer Adv IC JE.csv`).

**Adv IC JE Department value must be the full path** `General & Administrative : GA`
(not the abbreviated `General & Administrative`, which is not a valid dept). the accountant flagged
this on 2026-05-14 after the FX wire CSV had to be edited before upload. Apply to every
line in the Adv IC JE CSV (all four lines of a US-to-foreign cash funding pattern).

Tell the accountant:
> "CSV ready at {path}. Upload via NetSuite UI: **Lists → Import Assistant → Import Type: Transactions → Record Type: Journal Entry** → upload the CSV → confirm field mapping → run import. The JE will land in the controller's Pending Approval queue."

Append to `audit_log.json` with `action: "GENERATE_CSV"`. After the accountant reports the JE number post-import, append a follow-up entry with `je_number` and `netsuite_internal_id`.

Common account IDs (still useful if the accountant asks for context):
- Chase Checking x0001 (111070): `773`
- Chase x0002 MMDA: look up via SuiteQL on account table if needed
- Wells Fargo x0003: look up via SuiteQL on account table if needed

### 6b — Cross-subsidiary FX cash transfers — DELEGATE TO `adv-interco-je`

Use when moving cash from a US account to a foreign subsidiary account (e.g., Chase x0001
→ TD-CAD Canada via a FOREIGN EXCHANGE DEBIT on Chase paired with a Funds Transfer Credit
on the foreign side). These require a 4-line `advintercompanyjournalentry`.

**Do not duplicate the format here.** The authoritative spec lives in
[`.claude/skills/adv-interco-je/SKILL.md`](../adv-interco-je/SKILL.md). That skill defines
the column layout, the single-currency rule, the FX-rate placement rule, file naming,
and worked examples (including JE##### for US→Canada). This skill (`bank-statement-posting`)
only **identifies** the FX wire row and **emits a CSV in the adv-interco-je format**;
all rules below are summaries of that spec — when there's any conflict, adv-interco-je
wins.

**Identifying an FX wire row on Chase x0001:**
- `Transaction Description = FOREIGN EXCHANGE DEBIT`
- `Transaction Detail` contains `FOREIGN EXCHANGE/ FOREX/ AMT {foreign_ccy} {foreign_amt} EX {rate}`
- The other leg of the wire shows up on the foreign account (e.g., TD Canada as
  `Funds transfer credit / TT ACME CORP IN`) usually on the same day.

**What the skill produces:**
- CSV name: `{YYYY-MM-DD} {Direction} Cash Transfer Adv IC JE.csv` (e.g.
  `2026-05-12 Canada FX Cash Transfer Adv IC JE.csv`)
- Lives in `Weekly Cash Activities/{YYYY-MM}/JE Imports/` separate from the same-sub
  consolidated JE Import CSV (Adv IC uses a different NS Import Assistant).
- External ID: `{MonShort}{YY}USADVJE-{TargetCCy}FX` (e.g. `May26USADVJE-CADFX`).

**Summary of the rules (full spec in `adv-interco-je/SKILL.md`):**

| Rule | Value |
|---|---|
| Header Subsidiary | Source sub (US for a US→CA wire) |
| Journal Entry Currency | Source transaction currency (USD for our outbound FX wire from Chase) |
| Amount on every line | USD wire amount (e.g. $X,XXX.XX) — NOT the foreign currency amount |
| Exchange Rate (source-side lines) | `1` or blank |
| Exchange Rate (foreign-side lines) | The bank-displayed FX rate (e.g. `1.35154` for $X,XXX.XX USD → CAD 500,000) |
| Department | `General & Administrative : GA` (full path, every line) |
| NAME (on 121900 / 211900) | IC-{counterparty short name} per `_shared/subsidiary-constants.md` |
| Eliminate | `Yes` on 121900 and 211900 lines, `No` on cash lines |

**🚨 Single-currency rule.** All four lines carry the JE currency amount, not the foreign
currency amount. NS reads back the translated foreign amount via the Exchange Rate column.
This is the same rule `adv-interco-je` enforces; the JE##### worked example confirms it.
The local April CSV `Monthly Intercompany JEs/2026-04 April 2026/Canada/2026-04 Canada FX
Cash Transfer Adv IC JE.csv` is a stale draft that violates the rule — do NOT mirror it.
The actually-posted JE##### follows the rule.

**4-line structure (US → Canada, USD wire = $X,XXX.XX USD, FX rate 1.35154):**

| Line | Account | Line Sub | DR/CR | JE-currency Amount | Rate | NAME | Eliminate |
|------|---------|----------|-------|-------------------|------|------|-----------|
| 1 | 111070 Chase x0001 | US | Credit | 369,948.36 USD | 1 | — | No |
| 2 | 121900 IC Receivables | US | Debit | 369,948.36 USD | 1 | IC-Acme Canada | Yes |
| 3 | 111011 TD/Barclays Canada | Canada | Debit | 369,948.36 USD | **1.35154** | — | No |
| 4 | 211900 IC Payables | Canada | Credit | 369,948.36 USD | **1.35154** | IC-Acme, Inc. | Yes |

After the accountant uploads, append the JE number to `audit_log.json` per the standard
GENERATE_CSV → POST_JE follow-up pattern.

---

## Phase 7: Payroll Posting

**Approval rule for Phase 7:** `check` and `deposit` records are **exempt** from the
`approved: false` requirement that applies to JEs (see `_shared/approval-required.md`).
These record types do not carry an approval field in the accountant's NetSuite UI — they post and
clear immediately. Do NOT add `approved` to any Phase 7 payload below.

### 7a — Bi-monthly Workhorse HR payroll (PAYUS)

Acme Corp runs payroll via Workhorse HR twice a month. **Pay dates are always the 15th and the
last day of the month** (e.g., April 15 and April 30). Wires go out the business day before
the pay date — use the **pay date** (not the wire date) in the memo.

Three check records per pay cycle:

| Check | CONTRA AC | tranId pattern | DR account |
|-------|-----------|---------------|------------|
| DD wire | 4399256023 | `ACH_Debit_[M.DD.YY]_Workhorse HR_DD` | 231200 Payroll Liability (acct 249) |
| Tax wire | 4833794589 | `ACH_Debit_[M.DD.YY]_Workhorse HR_Tax` | 231350 Payroll Tax Liability (acct 252) |
| Garnishment ACH ($X,XXX.XX fixed) | — | `ACH_Debit_[M.DD.YY]_Workhorse HR` | 231200 Payroll Liability (acct 249) |

**Memo format:** Use the pay date month abbreviation, e.g., `4.15.26` for April 15 payroll.

**Record type:** `check`, form `49` (Standard Check)
**Entity:** Workhorse HR (id: `1111113`)
**Account (bank):** Chase x0001 (id: `773`)
**Subsidiary:** `2`

**Expense sublist offset account:** matches the DR account above (Payroll Liability or Tax Liability)

**How to derive amounts:** Look up prior month's corresponding checks in NetSuite
(search `transaction` table for tranId LIKE `%Workhorse HR%` in the prior period). Mirror amounts
unless the accountant provides updated figures.

### 7b — Final pay for terminated employees

Occasionally (roughly once a month, not every month), terminated employees receive final
pay via ACH. These appear in the banking sheet as individual ACH debits with the employee
name in the transaction detail.

**Pattern:** Mirror the format of historical final pay checks (e.g., `ACH_Debit_3.16.2026_R.Wormald`).

| Field | Value |
|-------|-------|
| Record type | `check`, form `49` |
| Entity | Employee's NS entity ID (look up via SuiteQL: `SELECT id, altname FROM entity WHERE altname LIKE '%[Last Name]%'`) |
| Account (bank) | Chase x0001 (id: `773`) |
| tranId | `ACH_Debit_[M.D.YYYY]_[F].[Last]` |
| Memo | `[F]. [Last] Final Pay` |
| Expense DR account | 231200 Payroll Liability (acct 249) |
| Department | G&A:GA (dept id: `2`) |

**Known employee entity IDs:**
- Employee_008: `35382`
- Employee_009: `1785196`
- Employee_007: `10963`

### 7c — RetireBridge 401k

Monthly 401k contribution wire. Pay date is embedded in the RCVR ID field of the
transaction detail: `233528 MMDDYYYY`.

| Field | Value |
|-------|-------|
| Record type | `check`, form `49` |
| Entity | RetireBridge (id: `29`) |
| Account (bank) | Chase x0001 (id: `773`) |
| tranId | `ACH_Debit_MM.DD.YYYY_401k` (use pay date from RCVR ID) |
| Memo | `401k contribution for [pay date]` |
| Expense DR account | 231250 401k Payable (acct 250) |

### 7d — Springfield sublease deposit

Monthly sublease payment received from Premier Plans Inc. for the Springfield, IL
office. Appears as `GA Operating` or similar in transaction detail.

| Field | Value |
|-------|-------|
| Record type | `deposit`, form `1` |
| Entity | Premier Plans Inc. (id: `1111116`) |
| tranId | Auto-assigned by NetSuite (DEP###) — do not set manually |
| Other line account | 261100 (acct id: `260`) |
| Department | EBITDA Adjustments (dept id: `108`) |
| Memo | `RENT - Springfield Office` |

**Sublist:** Use the `other` sublist (not `payment`) for deposit offset entries.

**Amount:** Confirm with the accountant — amount may vary slightly month to month. Historical: ~$X,XXX.XX.

### 7e — FlexBenefits FSA/HSA Check Posting (Chase x0001)

FlexBenefits and BancorpSv FlexBenefits-DBI rows on Chase x0001 are taken over from the AP Specialist starting
2026-05-05. Routing is **deterministic** from the bank's `Transaction Detail` field, so the
skill auto-posts these as Check records via MCP on every run. No paste-back classify-only
for these patterns going forward.

**Vendor:** FSA HSA Administrator. — id `1111115`. Subsidiary `2` (Acme, Inc.). No department
on any line. CR account is always `773` (Chase x0001, GL 111070).

**Routing dispatch (match on substring of Chase `Transaction Detail`):**

| Bank Description contains | tranid suffix | DR account | Header memo | Line memo |
|---|---|---|---|---|
| `FlexBenefits HEALTH INC PLAN FUND` (LARGER amount of the day, ~$8K-$13K, on payroll dates 15th / EOM) | `-EE_HSA` | 557 (`231201 HSA Payable`) | `FlexBenefits - HSA - Employee Payroll Deduction` | `FlexBenefits - HSA - Employee Payroll Deduction` |
| `FlexBenefits HEALTH INC PLAN FUND` (SMALLER amount, ~$2K-$4K, same payroll dates) | `-ER_HSA` | 557 (`231201 HSA Payable`) | `FlexBenefits - HSA - Employer Payroll Deduction` | `FlexBenefits - HSA - EmployeR Payroll Deduction` (preserve historical typo "EmployeR") |
| `FlexBenefits HEALTH INC CLAIM FUND` | `_FSA` | 558 (`231202 FSA Payable`) | `Medical FSA Carryover` | `Medical FSA Carryover - MM/DD/YYYY` |
| `BANCORPSV ... FlexBenefits HEALTH DBI ... SETTLE PURCHASE` | `_FSA` | 558 (`231202 FSA Payable`) | `Medical FSA Carryover` | `Medical FSA Carryover - MM/DD/YYYY` |

**Tranid format:**
- First Check of the day for that suffix: `ACH Debit YYYY.MM.DD_FSA` (or `_HSA`).
- Multiple FSA Checks same day: append `-1`, `-2`, ... before the suffix: `ACH Debit YYYY.MM.DD-1_FSA`, `-2_FSA`, etc.
- HSA pair always: `ACH Debit YYYY.MM.DD-EE_HSA` and `ACH Debit YYYY.MM.DD-ER_HSA`.

**EE vs ER disambiguation rule:**
- Two PLAN FUND rows on the same day → larger = EE, smaller = ER.
- Sanity check ratio: EE typically 3-4x ER. If they're within 30% of each other (rare),
  STOP and surface to the accountant with the two amounts; don't guess.

**Payload template (mirrors `ACH_Debit_05.04.26_Chase` pattern, no `customForm`):**
```json
{
  "tranDate": "2026-05-01",
  "subsidiary": {"id": "2"},
  "account": {"id": "773"},
  "entity": {"id": "1111115"},
  "tranId": "ACH Debit 2026.05.01-EE_HSA",
  "memo": "FlexBenefits - HSA - Employee Payroll Deduction",
  "expense": {
    "items": [
      {
        "account": {"id": "557"},
        "amount": 11922.66,
        "memo": "FlexBenefits - HSA - Employee Payroll Deduction"
      }
    ]
  }
}
```

For FSA, use `account.id = "558"` and replace memos with `Medical FSA Carryover` /
`Medical FSA Carryover - MM/DD/YYYY`.

**Approval:** Check records are exempt from the JE auto-approver issue per
`_shared/approval-required.md` — do NOT include `approved: false`.

**Phase 0 dedup applies:** before posting any FlexBenefits Check, query NS for the same
`(account=773, trandate, amount)` to skip rows the AP Specialist has already posted historically.

**FlexBenefits FSA/HSA refund credits (BANCORPSV ... DBI-99995-SETTLE CREDIT, FlexBenefits HEALTH ... CLAIM RETURN):**

Inverse of an FSA/HSA Check — funds returning from FlexBenefits (declined claim, incomplete
verification, employee never spent the amount). Post as a **Deposit** mirroring DEP893
(NS id 1839730, the canonical HSA refund pattern).

Routing dispatch on positive-amount (CR) FlexBenefits rows:

| Bank Description contains | Offset (CR) account | memo |
|---|---|---|
| `BANCORPSV ... DBI-99995-SETTLE CREDIT` | 558 (231202 FSA Payable) | `FSA Settlement Credit Returned Funds` |
| `FlexBenefits HEALTH ... CLAIM FUND` (positive amount) | 558 (231202 FSA Payable) | `FSA Claim Returned Funds` |
| `FlexBenefits HEALTH ... PLAN FUND` (positive amount) | 557 (231201 HSA Payable) | `HSA Incomplete Verification Returned Funds` |

**Payload template (mirrors DEP893):**
```json
{
  "customForm": {"id": "1"},
  "tranDate": "2026-04-29",
  "subsidiary": {"id": "2"},
  "account": {"id": "773"},
  "other": {"items": [
    {
      "account": {"id": "558"},
      "entity": {"id": "1111115"},
      "amount": 464.00,
      "memo": "FSA Settlement Credit Returned Funds"
    }
  ]}
}
```

NetSuite auto-assigns `tranid` as `DEP###`; do not set manually. No department on either
line. Confirmed working: DEP898 = $464 / 4/29/2026 / 1844030 (FSA refund).

### 7f — Write-back after payroll posting

After posting payroll, FlexBenefits, and deposit records, update the gap CSV and banking tracking sheet:
- Mark each row Posted=`*`
- Fill in the JE # column with the tranId (for checks) or DEP### (for deposits)

### 7g — Foreign-currency payroll bank-side legs (PAYUK / PAYNL / PAYPN / PAYUY)

**Scope:** post the bank-side leg of foreign payroll wires as Check records. Mirrors
the US-side pattern (Phase 7a Workhorse HR Tax/DD/Garnishment + 7c RetireBridge 401k) but for
foreign accounts. The detail side (CR Payroll Liability / DR Payroll Expense) is
posted by the country payroll skills (`uk-payroll`, `netherlands-payroll`,
`poland-payroll`, `uruguay-payroll`); this skill only books the cash-out leg.

**Whitelist required.** Auto-post is gated on `references/payroll_whitelist.yaml`.
The whitelist enumerates exact (account_key, beneficiary substring) pairs allowed
to auto-Check. Any wire that doesn't match the whitelist stays paste-back-only and
pings the accountant. See whitelist file for current entries.

**Standard payload:**

```json
{
  "customForm": {"id": "1"},
  "account": {"id": "<accounts[active].gl_account_id>"},
  "subsidiary": {"id": "<accounts[active].subsidiary_id>"},
  "currency": {"id": "<accounts[active].currency_id>"},
  "entity": {"id": "<beneficiary NS id from lookup>"},
  "tranDate": "<YYYY-MM-DD>",
  "tranid": "ACH_Debit_MM.DD.YY_PAY{XX}{N?}",
  "memo": "<condensed bank narrative>",
  "expense": {"items": [{
    "account": {"id": "249"},
    "amount": <bank amount in foreign CCY>,
    "department": {"id": "2"},
    "memo": "Payroll liability - {Month YYYY} {country} payroll"
  }]}
}
```

**Key fields per country:**

| Country | Tag | account (bank GL) | subsidiary | currency | tranid suffix | Common vendor |
|---|---|---|---|---|---|---|
| UK | `PAYUK` | 783 (ING BV GBP) | 4 | 2 (GBP) | `_PAYUK{N}` | UK Payroll Processor (1111118) |
| Netherlands | `PAYNL` | 782 (ING BV EUR) | 4 | 4 (EUR) | `_PAYNL{N}` | TBD (resolve from the Assistant Controller's history) |
| Poland | `PAYPN` | 560 (Bank of America PLN) | 6 | 5 (PLN) | per-beneficiary (vendor-specific) | ZUS (1111116), KAS (1111117), employees (per lookup) |
| Uruguay | `PAYUY` | 773 (Chase x0001, USD source) | 2 | 1 (USD) | `_PAYUY` | UY Payroll Processor (1111119) |

Account 249 = `231200 Payroll Liability` — shared across all subs (verified 2026-05-19).

**Numbering convention:** when 2+ payroll wires hit in the same month-day, suffix
with `{N}` (e.g., `ACH_Debit_4.28.26_PAYUK3` if it's the 3rd UK wire of April-26).
Look at the prior period's tranid sequence to pick `N`.

**Approval:** like other Checks, no `approved` field needed — Checks have no approval
workflow in NetSuite (per `_shared/approval-required.md`).

**Memo style (concise — per the accountant 2026-05-26):** Header memos should be short and
contextual. Vendor name is on the entity field; date is on tranDate. The memo just
needs the "what" and the period. Avoid mechanical descriptions of which skill
posts which side, and avoid verbose bank-narrative quotes.

Good examples (use these):
- `"May 2026 Uruguay payroll wire - UY Payroll Processor for UY Payroll Provider SAS"`
- `"April 2026 UK payroll - UK Payroll Processor"`
- `"Q2 2026 Poland ZUS payroll tax"`
- `"2025 TX franchise tax extension"`
- `"NJ state tax - Acme Buyer Inc"`

Avoid (too verbose):
- ~~`"UY Payroll Processor for UY Payroll Provider SAS - May 2026 Uruguay payroll wire (BPS and IRPF) - bank-side leg only; uruguay-payroll skill posts payroll detail JE"`~~
- ~~`"Remittance of TX franchise tax 2025 - WEBFILE TAX PYMT (ACD-1357754) - matches 2025 TX Extension Form 05-164/05-165 prepared by Tax Firm A LLP, $X,XXX.XX extension payment"`~~

The bank reference / detailed context still lives in the audit log entry —
not the NS memo field.

### 7h — Foreign-account bank-fee Checks (mirrors 7e FlexBenefits pattern)

For every Bank Fees row on a foreign account (ING BV maint fees, ING BV FX
conversion fees paired with inbound remittances, ING BV outbound wire fees,
Bank of America PLN monthly fee bundle), post a Check with:

```json
{
  "customForm": {"id": "1"},
  "account": {"id": "<accounts[active].gl_account_id>"},
  "subsidiary": {"id": "<accounts[active].subsidiary_id>"},
  "currency": {"id": "<accounts[active].currency_id>"},
  "entity": {"id": "1111114"},                  // Chase Bank, Ltd. (verified 2026-05-19)
  "tranDate": "<YYYY-MM-DD>",
  "tranid": "ACH_Debit_MM.DD.YY_ChaseBV<CCY?>",  // _ChaseBV (EUR/USD), _ChaseBVGBP, _PLN
  "memo": "<verbatim bank narrative OR 'Bank fees for accepting customer payment PYMT####'>",
  "expense": {"items": [
    {"account": {"id": "715"}, "amount": <fee amount>, "memo": "Bank Fees"},
    {"account": {"id": "223"}, "amount": 0,            "memo": "VAT"}
  ]}
}
```

Account 715 = `651160 Bank Fees`. Account 223 = `121400 VAT on Purchases` — the
zero-amount VAT line is required for Dutch B.V. (sub 4) VAT reporting. the Assistant Controller's
prior Checks (verified Feb-Mar 2026) all include it.

**tranid pattern by source account:**
- ING BV EUR / USD: `ACH_Debit_MM.DD.YY_ChaseBV`
- ING BV GBP: `ACH_Debit_MM.DD.YY_ChaseBVGBP` (disambiguates from EUR/USD)
- Bank of America PLN: `ACH_Debit_MM.DD.YY_PLN`

When the fee is paired with a customer inbound, the memo should reference the
resulting PYMT# from Phase 4b-foreign (e.g., `"Bank fees for accepting customer
payment PYMT####"`). When standalone (monthly maint), memo is the bank narrative
or `"{Month YYYY} - bank fees"`.

---

## Phase 8: Debt / Interest Payment Posting

Monthly and quarterly interest and fee payments to Chase Bank, Ltd. These appear in the
banking sheet as Type=DEBT with TAG=LOANINT or LOCINT. Post as `check` records mirroring
the prior period's pattern.

**Entity:** Chase Bank, Ltd. (id: `1111114`)
**Record type:** `check`, form `49` (Standard Check)
**Account (bank):** Chase x0001 (id: `773`)
**Subsidiary:** `2` (Acme, Inc.)
**Department:** G&A:GA (id: `2`)

### Patterns by TAG

| TAG | Account | Account ID | tranId Pattern | Memo |
|-----|---------|------------|---------------|------|
| LOANINT | 711150 Interest Expense - Term Loans | `675` | `ACH_Debit_MM.DD.YYYY_LOANINT` | `Remittance of Q[N]'YY TL interest.` |
| LOCINT | 711155 Interest Expense - Line of Credit | `716` | `ACH_Debit_MM.DD.YY_Chase` | See note below |

**LOCINT memo variants:**
- LC maturity/fee charges: `LC Charges and Fees`
- Unused commitment fee (mid-month): `Q[N]-YY LC and Unused Commitment Fees`

**tranId date format:**
- LOANINT: 4-digit year (e.g., `ACH_Debit_04.02.2026_LOANINT`)
- LOCINT: 2-digit year (e.g., `ACH_Debit_04.08.26_Chase`, `ACH_Debit_04.15.26_Chase`)

**Multiple LC charges on the same date:** combine into a single check with multiple expense
lines (one per LC number), each line with the same `LC Charges and Fees` memo. Use a single
tranId for the combined check.

### How to identify and post

1. Filter gap CSV rows where Type=DEBT (TAG=LOANINT or LOCINT)
2. Look up prior period's corresponding checks in NetSuite:
   ```sql
   SELECT t.id, t.tranid, t.trandate, t.memo, t.total
   FROM transaction t
   WHERE t.entity = {NS_INTERNAL_ID}
     AND t.trandate >= TO_DATE('YYYY-MM-01', 'YYYY-MM-DD')
     AND t.trandate <= TO_DATE('YYYY-MM-30', 'YYYY-MM-DD')
     AND t.recordtype = 'check'
   ORDER BY t.trandate
   ```
3. Mirror the amount, account, memo, and tranId pattern — update the date/quarter reference
4. Post using `ns_createRecord` with `check` record type

### April 2026 posted checks (reference)

| tranId | Date | Amount | Account | Memo |
|--------|------|--------|---------|------|
| ACH_Debit_04.02.2026_LOANINT | 4/2/2026 | $X,XXX.XX | 711150 | Remittance of Q1'26 TL interest. |
| ACH_Debit_04.08.26_Chase | 4/8/2026 | $447.71 | 711155 | LC Charges and Fees (2 lines: $369.58 LC S542145 + $78.13 LC S542068) |
| ACH_Debit_04.15.26_Chase | 4/15/2026 | $X,XXX.XX | 711155 | Q1-26 LC and Unused Commitment Fees |

---

## Standing Rules

- **Never post without explicit approval.** Always pause at Phase 3.
- **Never fabricate invoice numbers.** If a real NetSuite match isn't found, say so.
- **Amount precision.** Match and post to the cent.
- **Include intraday (Status=I) by default.** the accountant confirmed intraday amounts almost always
  settle correctly. Only skip intraday transactions if the accountant explicitly requests it.
- **Subsidiary.** Chase 5839 = Acme, Inc. (id: "2").
- **Scope.** Process the accountant's transactions — customer payments, transfers, debt, payroll,
  misc, and (as of 2026-05-05) **FlexBenefits and BancorpSv FlexBenefits-DBI rows on Chase x0001**
  (Phase 7e auto-posts these as FSA/HSA Checks). Continue to skip Bills/DOM rows that
  the AP Specialist books via BillFlow (Brex reimbursements, Airwallex, Medical Premium Broker, etc.).
- **Payments always applied.** Never post unapplied. Every payment must link to a specific invoice.
- **Date format for NetSuite REST API.** Use ISO format: `YYYY-MM-DD` (not MM/DD/YYYY).
- **Deduplicate every run.** Always run Phase 0 before Phase 1 to avoid double-posting.
