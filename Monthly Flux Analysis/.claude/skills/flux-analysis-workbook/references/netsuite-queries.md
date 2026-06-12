# NetSuite Queries Reference — Flux Analysis Skill

**Last Updated**: April 7, 2026
**Validated Against**: Feb 2026 (penny-for-penny match) and Mar 2026 (penny-for-penny match)

---

## Overview

This reference documents the exact NetSuite report configurations and parsing rules for the 4 detail tabs (COGS, Contractors, Professional Fees, Software) in the flux analysis workbook.

**Primary approach**: `ns_runReport` with confirmed report IDs
**Fallback**: `ns_runCustomSuiteQL` (SuiteQL queries documented below as backup)

---

## Critical: Subsidiary Configuration

**ALWAYS use subsidiary ID -2 — Acme, Inc. (Consolidated).**

The flux workbook is prepared on a consolidated USD basis. Validated April 2026 after testing all subsidiary options:
- Sub **-2** (Acme, Inc. Consolidated): Returns all subs FX-converted to USD ✅ USE THIS
- Sub -1 (Holdings Consolidated): `ns_runReport` returns empty account shell — no transaction detail ❌
- Sub 2 (US only): Misses international entries — e.g., ProfFees Mar 2026: $X,XXX.XX vs $X,XXX.XX with sub -2 ❌

| Subsidiary | ID | Notes |
|---|---|---|
| Acme Holdings (Consolidated) | **1** | ALWAYS use this for flux analysis |
| Acme, Inc. | 2 | US-only — do NOT use |
| Acme Corp UK | 3+ | Included automatically via consolidated |
| Acme Netherlands | 3+ | Included automatically via consolidated |
| Acme Uruguay | 3+ | Included automatically via consolidated |

---

## Report 1: COGS Detail (Report ID 537)

### Report Info
- **Title**: Income Statement Detail COGS - BW
- **Report ID**: 537
- **Type**: Income Statement Detail
- **has_subsidiary_filter**: true

### MCP Call
```
ns_runReport(
  reportId: 537,
  dateFrom: "YYYY-MM-DD",     # First day of M-2
  dateTo: "YYYY-MM-DD",       # Last day of M
  subsidiaryId: 1              # Consolidated
)
```

### Date Range Example (March 2026 close)
```
dateFrom: "2026-01-01"
dateTo: "2026-03-31"
```

### Output Columns (detailLineValues)
13 columns per detail row:
```
Type, Date, Document Number, Name, Entity (Line), Clr, Split, Amount,
Memo, Message, Account (Line): Name (GL-style), Department: Name,
Accounting Period: Name
```

### Parsing Rules
1. Filter: `isDetailLine == true` AND `Type` is not null
2. Account filter: Only rows where Financial Row contains one of: `511400, 511425, 511450, 511510, 511520, 511550, 511600`
   - This EXCLUDES salary/benefits accounts (511100, 511150, 511175, 511200, 511250, 511300, 511350, 511370) that also appear in this report
3. Financial Row extraction: Split `Account (Line): Name (GL-style)` on `\x01` (or `:` after cleaning), take last segment
4. **NEGATE amounts**: The IS report shows expenses as negative. Multiply Amount by -1.
5. Final Name: `COALESCE(Entity (Line), Name)` — skip "- No Entity -"
6. Clean `\x01` → `:` in account paths
7. Parse dates: `"Mon, 01 Dec 2025 00:00:00 GMT"` → datetime

### Target Columns in Template (IncomeStatementDetailCOGS, Row 7 headers)
```
A: Financial Row       H: Clr
B: Type                I: Split
C: Date                J: Amount (NEGATED)
D: Document Number     K: Memo
E: Name                L: Message
F: Entity (Line)       M: Account (Line): Name (GL-style)
G: Final Name          N: Department: Name
                       O: Accounting Period: Name
```

### COGS Sub-Account Reference
| Account | Description |
|---|---|
| 511400 | COGS - Hosting |
| 511425 | COGS - Software Subscriptions |
| 511450 | COGS - Translation |
| 511510 | COGS - Client Billable |
| 511520 | COGS - Client Non Billable |
| 511550 | COGS - Paper Survey |
| 511600 | COGS - Other |

---

## Report 2: Contractors Detail (Report ID 540)

### Report Info
- **Title**: General Ledger detail Contractor Payroll - BW
- **Report ID**: 540
- **Type**: GL Detail
- **has_subsidiary_filter**: true

### MCP Call
```
ns_runReport(
  reportId: 540,
  dateFrom: "YYYY-MM-DD",
  dateTo: "YYYY-MM-DD",
  subsidiaryId: -2
)
```

### Output Columns (detailLineValues)
13 columns per detail row:
```
Account, Type, Date, Document Number, Name, Entity (Line): Name,
"" (unnamed = net amount), Memo, Message, Department: Name,
Subsidiary: Name, Accounting Period: Name, Balance
```

### Parsing Rules
1. Filter: `isDetailLine == true` AND `Type` is not null
2. Account filter: Only rows where Account contains `511370` or `611700`
3. **Do NOT negate** — GL Detail amounts are already in correct sign convention
4. Split unnamed `""` column into Debit/Credit: positive → Debit, negative → abs → Credit
5. Final Name: `COALESCE(Entity (Line): Name, Name)`

### Target Columns in Template (GeneralLedgerdetailContra, Row 7 headers)
```
A: Account             H: Debit
B: Type                I: Credit
C: Date                J: Net (raw unnamed value)
D: Document Number     K: Memo
E: Name                L: Message
F: Entity (Line): Name M: Department: Name
G: Final Name          N: Subsidiary: Name
                       O: Accounting Period: Name
```

### Target GL Accounts
| Account | Description |
|---|---|
| 511370 | COGS - Contractor Payroll (under 511001) |
| 611700 | Contractor Payroll (under 611000) |

---

## Report 3: Professional Fees Detail (Report ID 542)

### Report Info
- **Title**: General Ledger detail Professional Fees - BW
- **Report ID**: 542
- **Type**: GL Detail
- **has_subsidiary_filter**: true

### MCP Call
```
ns_runReport(
  reportId: 542,
  dateFrom: "YYYY-MM-DD",
  dateTo: "YYYY-MM-DD",
  subsidiaryId: -2
)
```

### Output & Parsing
Same structure and rules as Contractors (Report {NS_REPORT_ID}), except:
- **Account filter**: Only rows where Account contains `651100` or `651101`
- **Target template sheet**: GeneralLedgerdetailProfes

### Target GL Accounts
| Account | Description |
|---|---|
| 651100 | Professional Fees |
| 651101 | Professional Fees (sub) |

---

## Report 4: Software Detail (Report ID 721)

### Report Info
- **Title**: General Ledger detail Software Expense - BW
- **Report ID**: 721
- **Type**: GL Detail
- **has_subsidiary_filter**: true

### MCP Call
```
ns_runReport(
  reportId: 721,
  dateFrom: "YYYY-MM-DD",
  dateTo: "YYYY-MM-DD",
  subsidiaryId: -2
)
```

### Output & Parsing
Same structure and rules as Contractors (Report {NS_REPORT_ID}), except:
- **Account filter**: Only rows where Account contains `671000` or `671100`
- **Target template sheet**: GeneralLedgerdetailSoftwa

### Target GL Accounts
| Account | Description |
|---|---|
| 671000 | Software |
| 671100 | Software Subscriptions |

---

## Report Output Format (All 4 Reports)

All reports return this JSON structure:
```json
[{
  "type": "text",
  "text": "{\"reportData\":{\"0\":{...},\"1\":{...},...},\"currency\":\"$\",\"title\":\"...\"}"
}]
```

The `reportData` object has numbered string keys ("0", "1", "2", ...) where each value is a row object:

### Detail Row (transaction data — what we parse)
```json
{
  "isDetailLine": true,
  "detailLineValues": [
    {"Type": "Journal"},
    {"Date": "Mon, 01 Dec 2025 00:00:00 GMT"},
    {"Document Number": "JE#####"},
    {"Name": null},
    {"Entity (Line)": "IT Hardware Partner Networked Solutions Group, LLC"},
    ...
  ]
}
```

### Section Row (headers/totals — skip, but track for COGS context)
```json
{
  "isDetailLine": false,
  "value": "511400 - COGS - Hosting",
  "summaryLineValues": [...]
}
```

### Key Parsing Details
- `detailLineValues` is an **array of single-key objects**, not a flat dict
- The GL Detail reports have an **unnamed column** (`""` key) that contains the net amount
- Dates are in format: `"Mon, 01 Dec 2025 00:00:00 GMT"`
- Account hierarchy uses `\x01` (SOH character) as separator, e.g., `"511000 - Cost of Goods Sold\x01511400 - COGS - Hosting"`
- `"NaN"` appears in summary row Amount fields — ignore these

---

## Tie-Back Verification

After populating the template, always verify current-month totals:

```python
# For each tab, sum the current period's amounts from:
# 1. The template data sheet (after writing)
# 2. The raw report data (re-parsing with period filter)
# Both must match to the penny.
```

### Validated Results

**February 2026** (initial validation):
```
COGS - IT Hardware Partner Dec 2025:  Template $X,XXX.XX = NetSuite $X,XXX.XX  ✓
COGS - LogPlatform all months:  Template $X,XXX.XX  = NetSuite $X,XXX.XX   ✓
Contractors - HR Consulting Partner Dec: Template $X,XXX.XX  = NetSuite $X,XXX.XX   ✓
Software - SecuritySaaS:      Template $X,XXX.XX   = NetSuite $X,XXX.XX    ✓
```

**March 2026** (end-to-end validation):
```
COGS:              Template $X,XXX.XX    = NetSuite $X,XXX.XX    ✓
Contractors:       Template $X,XXX.XX   = NetSuite $X,XXX.XX   ✓
Professional Fees: Template $X,XXX.XX    = NetSuite $X,XXX.XX    ✓
Software:          Template ($X,XXX.XX) = NetSuite ($X,XXX.XX) ✓
```

---

## SuiteQL Fallback Queries

If `ns_runReport` is unavailable, use `ns_runCustomSuiteQL` with the `transactionline` table (NOT `transactionaccountingline` — the latter has poor vendor resolution).

### Key Join Pattern for Vendor Resolution
```sql
FROM transactionline tl
JOIN transaction t ON t.id = tl.transaction
JOIN account a ON a.id = tl.account
JOIN accountingperiod ap ON ap.id = tl.postingperiod
LEFT JOIN entity le ON tl.entity = le.id        -- LINE-level entity (critical)
LEFT JOIN entity te ON t.entity = te.id          -- Transaction-level entity
LEFT JOIN vendor tv ON t.entity = tv.id          -- Vendor record
LEFT JOIN department d ON d.id = tl.department
LEFT JOIN subsidiary s ON s.id = tl.subsidiary
```

**IMPORTANT**: Use `tl.entity` (line-level) for vendor resolution, NOT just `t.entity` (transaction-level). JEs and Brex card charges store the vendor at the line level.

### Vendor Name Resolution
```sql
COALESCE(le.entityid, te.entityid, tv.companyname) AS "Final Name"
```

### NULL Handling for Amounts
```sql
SUM(COALESCE(tl.debit, 0) - COALESCE(tl.credit, 0)) AS "Net"
```
(Without COALESCE, NULL debit or credit causes the entire expression to return NULL.)

### Example: COGS Detail via SuiteQL
```sql
SELECT
  a.acctnumber || ' - ' || a.displayname AS "Financial Row",
  t.type AS "Type",
  t.trandate AS "Date",
  t.tranid AS "Document Number",
  te.entityid AS "Name",
  le.entityid AS "Entity (Line)",
  COALESCE(le.entityid, te.entityid) AS "Final Name",
  COALESCE(tl.debit, 0) - COALESCE(tl.credit, 0) AS "Amount",
  tl.memo AS "Memo",
  a.acctnumber AS "Account Number",
  d.name AS "Department",
  ap.periodname AS "Accounting Period"
FROM transactionline tl
JOIN transaction t ON t.id = tl.transaction
JOIN account a ON a.id = tl.account
JOIN accountingperiod ap ON ap.id = tl.postingperiod
LEFT JOIN entity le ON tl.entity = le.id
LEFT JOIN entity te ON t.entity = te.id
LEFT JOIN department d ON d.id = tl.department
WHERE a.acctnumber IN ('511400','511425','511450','511510','511520','511550','511600')
  AND t.subsidiary = 1
  AND ap.periodname IN ('Jan 2026','Feb 2026','Mar 2026')
  AND t.voided = 'F'
ORDER BY a.acctnumber, t.trandate
```

---

## MCP Tools Reference

| Tool | Purpose |
|---|---|
| `ns_runReport` | Primary tool — run saved NetSuite financial reports |
| `ns_runCustomSuiteQL` | Fallback — run custom SQL queries against NetSuite |
| `ns_listAllReports` | Discover available reports, check has_subsidiary_filter |
| `ns_getSubsidiaries` | Get valid subsidiary IDs |
| `ns_getSuiteQLMetadata` | Explore database schema for SuiteQL queries |

---

## Document Control

**Last Updated**: April 7, 2026
**Validated By**: Alex Reed (Controller)
**NetSuite Environment**: Acme Holdings
**Fiscal Calendar**: US Calendar (Jan-Dec)
