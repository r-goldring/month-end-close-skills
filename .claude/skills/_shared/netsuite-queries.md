# NetSuite Query Reference

Common SuiteQL patterns used across all skills. Reference these rather than writing new queries from scratch.

## MCP tools available
- `ns_runCustomSuiteQL` — run a SuiteQL query, returns JSON rows
- `ns_runReport` — run a saved NetSuite report by ID, returns tabular data

## GL Detail query (used by recon-fill)

Returns all transaction lines for a specific GL account + subsidiary + date range.
Matches the "Custom Acme Balance Sheet Detail" report format.

```sql
SELECT
    t.recordtype                                        AS type,
    t.trandate                                          AS date,
    t.tranid                                            AS document_number,
    CASE
        WHEN v.companyname IS NOT NULL THEN v.companyname
        WHEN e.firstname IS NOT NULL THEN e.firstname || ' ' || COALESCE(e.lastname, '')
        ELSE e.entityid
    END                                                 AS name,
    (tl.debitforeignamount - tl.creditforeignamount)   AS amount,
    tl.memo,
    NULL                                                AS description,
    loc.name                                            AS location
FROM transaction t
JOIN transactionline tl    ON t.id = tl.transaction AND tl.mainline = 'F'
JOIN account a             ON tl.account = a.id
JOIN subsidiary sub        ON tl.subsidiary = sub.id
LEFT JOIN entity e         ON t.entity = e.id
LEFT JOIN vendor v         ON t.entity = v.id
LEFT JOIN location loc     ON tl.location = loc.id
WHERE a.acctnumber = '{ACCOUNT_NUMBER}'
  AND sub.fullname = '{SUBSIDIARY_FULLNAME}'
  AND t.trandate >= TO_DATE('{YYYY-MM-01}', 'YYYY-MM-DD')
  AND t.trandate <= TO_DATE('{YYYY-MM-LAST}', 'YYYY-MM-DD')
  AND t.voided = 'F'
ORDER BY t.trandate, t.tranid, tl.linesequencenumber
```

## Beginning balance query (for recon-fill running balance)

Returns the account balance as of the day before the period start.

```sql
SELECT SUM(tl.debitforeignamount - tl.creditforeignamount) AS beginning_balance
FROM transaction t
JOIN transactionline tl ON t.id = tl.transaction AND tl.mainline = 'F'
JOIN account a          ON tl.account = a.id
JOIN subsidiary sub     ON tl.subsidiary = sub.id
WHERE a.acctnumber = '{ACCOUNT_NUMBER}'
  AND sub.fullname = '{SUBSIDIARY_FULLNAME}'
  AND t.trandate < TO_DATE('{YYYY-MM-01}', 'YYYY-MM-DD')
  AND t.voided = 'F'
```

## IC transaction query (used by adv-interco-je) — parameterized by sub + accounts

Pulls all lines hitting IC expense accounts in a given month on a given subsidiary. Used both for the forward US-side scan and the foreign-side reverse / F-to-F scans.

The 681xxx series is target-named (681100=NL, 681110=CA, 681130=US, 681140=UK, 681150=UY). For each scan, exclude the source sub's own IC account (e.g., a CA scan excludes 681110 since "CA owes CA" is meaningless).

| Scan | `{SUBSIDIARY_FULLNAME}` | `{ACCOUNT_LIST}` |
|------|------------------------|------------------|
| US-side (forward) | `Acme Holdings : Acme, Inc.` | `'681100','681110','681140','681150'` |
| CA-side (reverse) | `Acme Holdings : Acme, Inc. : Acme Canada` | `'681100','681130','681140','681150'` |
| NL-side (F-to-F) | `Acme Holdings : Acme, Inc. : Acme Netherlands` | `'681110','681130','681140','681150'` |
| UK-side (F-to-F) | `Acme Holdings : Acme, Inc. : Acme UK Ltd` | `'681100','681110','681130','681150'` |
| UY-side (F-to-F) | `Acme Holdings : Acme, Inc. : Acme Uruguay` | `'681100','681110','681130','681140'` |

```sql
SELECT
    t.id                                                AS transaction_id,
    t.tranid,
    t.trandate,
    t.recordtype,
    t.memo                                              AS transaction_memo,
    t.exchangerate,
    cur.symbol                                          AS currency_code,
    tl.id                                               AS line_id,
    tl.linesequencenumber,
    a.acctnumber,
    a.fullname                                          AS account_name,
    tl.memo                                             AS line_memo,
    tl.debitforeignamount,
    tl.creditforeignamount,
    sub.fullname                                        AS line_subsidiary,
    d.fullname                                          AS department,
    v.companyname                                       AS vendor_name
FROM transaction t
JOIN transactionline tl  ON t.id = tl.transaction
JOIN account a           ON tl.account = a.id
LEFT JOIN subsidiary sub ON tl.subsidiary = sub.id
LEFT JOIN department d   ON tl.department = d.id
LEFT JOIN vendor v       ON t.entity = v.id
LEFT JOIN currency cur   ON t.currency = cur.id
WHERE a.acctnumber IN ({ACCOUNT_LIST})
  AND t.trandate >= TO_DATE('{YYYY-MM-01}', 'YYYY-MM-DD')
  AND t.trandate <= TO_DATE('{YYYY-MM-LAST}', 'YYYY-MM-DD')
  AND sub.fullname = '{SUBSIDIARY_FULLNAME}'
  AND t.voided = 'F'
ORDER BY a.acctnumber, t.trandate, t.tranid, tl.linesequencenumber
```

## Foreign sub P&L scan (used by adv-interco-je Phase 1B fallback)

When a saved P&L report ID is unavailable for a foreign sub, fall back to this SuiteQL. Returns all P&L (Expense, COGS, Other Expense) line activity for a sub in a month, grouped by vendor + account + memo so cross-sub items can be flagged.

```sql
SELECT
    t.id                                                AS transaction_id,
    t.tranid,
    t.trandate,
    t.recordtype,
    cur.symbol                                          AS currency_code,
    a.acctnumber,
    a.fullname                                          AS account_name,
    a.accttype,
    tl.memo                                             AS line_memo,
    tl.debitforeignamount,
    tl.creditforeignamount,
    d.fullname                                          AS department,
    v.companyname                                       AS vendor_name
FROM transaction t
JOIN transactionline tl  ON t.id = tl.transaction
JOIN account a           ON tl.account = a.id
LEFT JOIN subsidiary sub ON tl.subsidiary = sub.id
LEFT JOIN department d   ON tl.department = d.id
LEFT JOIN vendor v       ON t.entity = v.id
LEFT JOIN currency cur   ON t.currency = cur.id
WHERE a.accttype IN ('Expense', 'COGS', 'OthExpense')
  AND t.trandate >= TO_DATE('{YYYY-MM-01}', 'YYYY-MM-DD')
  AND t.trandate <= TO_DATE('{YYYY-MM-LAST}', 'YYYY-MM-DD')
  AND sub.fullname = '{SUBSIDIARY_FULLNAME}'
  AND t.voided = 'F'
ORDER BY v.companyname, a.acctnumber, t.trandate
```

## Adv IC JE line detail by tranid (used by adv-interco-je for self-validation)

Returns the full multi-line structure of any historical Adv IC JE so the skill can diff a newly-built CSV against a known-good production example. Useful before the first import of a new direction (e.g., before the first April CA → US import, diff the structure against JE#####).

```sql
SELECT
    t.id, t.tranid, t.trandate, t.memo AS header_memo, t.externalid,
    t.exchangerate, cur.symbol AS tran_currency,
    a.acctnumber, a.fullname AS account_name,
    tl.memo AS line_memo,
    tl.debitforeignamount, tl.creditforeignamount,
    sub.fullname AS line_sub,
    tl.linesequencenumber, tl.eliminate
FROM transaction t
JOIN transactionline tl  ON t.id = tl.transaction
JOIN account a           ON tl.account = a.id
LEFT JOIN subsidiary sub ON tl.subsidiary = sub.id
LEFT JOIN currency cur   ON t.currency = cur.id
WHERE t.tranid = '{JE_NUMBER}'
  AND t.voided = 'F'
ORDER BY tl.linesequencenumber
```

## Transaction lookup by ID (used by je-error-fix)

```sql
SELECT
    t.tranid,
    t.trandate,
    t.recordtype,
    t.memo                                              AS header_memo,
    a.acctnumber,
    a.fullname                                          AS account_name,
    tl.memo                                             AS line_memo,
    tl.debitforeignamount                               AS debit,
    tl.creditforeignamount                              AS credit,
    sub.fullname                                        AS subsidiary,
    d.fullname                                          AS department,
    tl.id                                               AS line_id,
    tl.linesequencenumber
FROM transaction t
JOIN transactionline tl  ON t.id = tl.transaction AND tl.mainline = 'F'
JOIN account a           ON tl.account = a.id
LEFT JOIN subsidiary sub ON tl.subsidiary = sub.id
LEFT JOIN department d   ON tl.department = d.id
WHERE t.tranid = '{JE_NUMBER}'
  AND t.voided = 'F'
ORDER BY tl.linesequencenumber
```

## Gut-check JE-line fetch (with subsidiary filter)

Used by the `gut-check` skill to pull JE lines for a posted prior payroll run.
Differs from the `je-error-fix` lookup by also filtering on subsidiary
(prevents accidental cross-sub matches when the same tranid exists in another
NS company file). Returns native-currency debit/credit so foreign-payroll
comparisons stay in EUR/GBP/PLN/UYU/CAD.

```sql
SELECT
    t.id                                                AS internal_id,
    t.tranid,
    TO_CHAR(t.trandate, 'YYYY-MM-DD')                   AS trandate,
    t.memo                                              AS header_memo,
    t.voided, t.posting, t.reversaldate,
    a.acctnumber,
    a.fullname                                          AS account_name,
    tl.memo                                             AS line_memo,
    tl.debitforeignamount                               AS debit,
    tl.creditforeignamount                              AS credit,
    sub.fullname                                        AS subsidiary,
    d.fullname                                          AS department,
    tl.linesequencenumber
FROM transaction t
JOIN transactionline tl  ON t.id = tl.transaction AND tl.mainline = 'F'
JOIN account a           ON tl.account = a.id
LEFT JOIN subsidiary sub ON tl.subsidiary = sub.id
LEFT JOIN department d   ON tl.department = d.id
WHERE t.tranid = '{JE_NUMBER}'
  AND sub.fullname = '{SUBSIDIARY_FULLNAME}'
  AND t.voided = 'F'
ORDER BY tl.linesequencenumber
```

## Validate prior JE candidates (used by gut-check)

Returns a row only for tranids that are valid baselines: not voided, posting,
no reversal date. Caller drops candidates absent from the result.

```sql
SELECT
    t.tranid,
    t.id                                AS internal_id,
    t.voided,
    t.posting,
    t.reversaldate,
    TO_CHAR(t.trandate, 'YYYY-MM-DD')   AS trandate,
    t.memo
FROM transaction t
WHERE t.tranid IN ('JE#####', 'JE#####', 'JE#####')   -- caller fills with up to ~8 candidates
  AND t.voided = 'F'
  AND t.posting = 'T'
  AND t.reversaldate IS NULL
ORDER BY t.trandate DESC
```

## Flux detail reports (used by flux-analysis-workbook and flux-accruals)

Use `ns_runReport` with these report IDs. Always use `subsidiaryId: -2` for consolidated USD.

| Report ID | Content |
|-----------|---------|
| 537 | COGS detail by vendor |
| 540 | Contractor payroll detail |
| 542 | Professional Fees detail |
| 721 | Software subscriptions detail |

Parameters:
```json
{
  "reportId": 537,
  "subsidiaryId": -2,
  "startDate": "YYYY-MM-01",
  "endDate": "YYYY-MM-DD"
}
```

## Date helper: last day of month (Python)

```python
import calendar, datetime

def last_day(year, month):
    return datetime.date(year, month, calendar.monthrange(year, month)[1])

# Example: last_day(2026, 3) → datetime.date(2026, 3, 31)
```
