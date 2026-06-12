# Acme Corp Close — Claude Code Operating Instructions

> **Note:** This is the public, illustrative version of an in-production accounting-automation repo. "Acme Corp" is a stand-in for the real company. You will need to replace the example IDs, paths, vendor names, and bank accounts with your own — see [README.md](README.md#customize-for-your-company) for the walkthrough.

## Who uses this repo

A senior accountant at a multi-entity SaaS company on NetSuite uses these skills to automate the monthly close. The patterns assume:

- **ERP:** NetSuite (with the NetSuite MCP for read operations).
- **Multi-entity:** US parent with foreign subsidiaries (the examples show US / Canada / Netherlands / UK setup; adapt for yours).
- **Payroll:** ADP for US and Canada (bi-weekly); local providers for other countries (monthly).
- **AP automation tool** for vendor bills, **corporate card platform** for spend, **separate reconciliation tool** for month-end recs.

## Skills available

| Skill | Trigger phrases | Status |
|-------|----------------|--------|
| `flux-analysis-workbook` | flux analysis, month-end close workbook, flux workbook, variance analysis | Production |
| `adv-interco-je` | intercompany JE, IC JE, reclass to NL/UK/UY/CA, 681xxx accounts | Production |
| `bank-statement-posting` | bank statement posting, customer payments, lockbox deposits, bank fees, daily transactions | Production |
| `monthly-cash-confirmation` | monthly cash confirmation, confirm cash, FP&A cash request, sublease confirmation, debt service | Production |
| `monthly-cfs` | CFS, cash flow statement, monthly CFS, consolidated CFS, FX embedded, FP&A cash flow workbooks | Production |
| `bonus-accrual` | bonus accrual, bonus JE, monthly bonus, FP&A bonus file, 231170/231171 | Production |
| `flux-accruals` | accruals, identify accruals, month-end accruals, dept reclass, software reclass | Production |
| `je-error-fix` | fix JE error, wrong department, period error, reclass JE, correcting entry | Production |
| `us-payroll` | US payroll, ADP payroll, payroll JE | Production |
| `canada-payroll` | Canada payroll, Canadian payroll, RRSP/CRRSP | Production |
| `germany-payroll` | Germany payroll, German payroll | Production |
| `netherlands-payroll` | Netherlands payroll, NL payroll, Dutch payroll | Production |
| `poland-payroll` | Poland payroll, Polish payroll | Production |
| `uk-payroll` | UK payroll, PAYE/NI/P32 | Production |
| `uruguay-payroll` | Uruguay payroll, UY payroll, Aguinaldo | Production |
| `avalara-je` | Avalara, sales tax JE | Production wrapper |
| `gut-check` | gut-check, pre-upload review, validate JE | Production |

## Core operating rules — read before any action

### 1. JE-style transactions: CSV upload only (separation-of-duties policy)

**Skills MUST NOT post `journalentry` or `advintercompanyjournalentry` records via the NetSuite MCP.** The NetSuite REST integration role typically has approval permissions a senior accountant should not — that violates separation of duties. All JE-style entries must be generated as CSVs and uploaded through the NetSuite UI (Lists → Import Assistant) so they land in the controller's Pending Approval queue.

Every JE-producing skill must:
1. Build the JE data in memory.
2. Display a formatted preview (account, dept, debit/credit, memo, totals).
3. Generate a `{YYYY-MM} {Description} JE Import.csv` next to the existing Backup workbook.
4. Tell the user: **"CSV ready at {path}. Upload via NetSuite UI: Lists → Import Assistant → Transactions → Journal Entry. JE will land in the controller's Pending Approval queue."**
5. Append a `GENERATE_CSV` entry to `audit_log.json`. After the user provides the resulting JE number post-import, append a follow-up with `je_number` and `netsuite_internal_id`.

**Exempt records (continue using `ns_createRecord`)** — these have no approval workflow:
- `customerpayment` (routine AR application)
- `check` (payroll wires, retirement contributions, terminated-employee final pay)
- `deposit` (sublease deposits, misc)

Read-only NetSuite operations (SuiteQL queries, ns_runReport) are unaffected.

### 2. Use local reference files for ID lookups — not NetSuite MCP

For vendor IDs, customer IDs, account IDs, department IDs, subsidiary IDs: read from local `.xls`/`.csv` exports of your NetSuite reference data. Re-export quarterly. Only fall back to a NetSuite MCP call if the entity isn't in the local file (e.g., added recently).

Why: API calls are slow, eat into rate limits, and create needless audit-log noise. The reference files are authoritative for 99% of lookups.

See [examples/sample-coa.csv](examples/sample-coa.csv), [examples/sample-vendors.csv](examples/sample-vendors.csv) etc. for the schemas the skills expect.

### 3. Skills cross-reference each other via _shared/

Common logic lives in `.claude/skills/_shared/`. Never duplicate:
- SuiteQL query patterns → `_shared/netsuite-queries.md`
- Subsidiary full-name paths and IDs → `_shared/subsidiary-constants.md`
- Check-and-balance pattern → `_shared/check-and-balance.md`
- ID lookup steps → `_shared/id-lookup-guide.md`
- Approval / CSV-upload policy → `_shared/approval-required.md`

### 4. Subsidiary `-2` for consolidated reporting

When running NetSuite reports for consolidated USD data, use subsidiary ID `-2`. For entity-specific queries, use the subsidiary's own ID from your local subsidiaries reference file.

### 5. Payroll Python scripts are production-ready — do not rewrite

The existing Python scripts under `scripts/{country}-payroll/` are mature and tested:
- `scripts/us-payroll/payroll_mapper.py`
- `scripts/canada-payroll/process_canada_payroll.py`
- `scripts/germany-payroll/process_germany_payroll.py`
- `scripts/netherlands-payroll/process_netherlands_payroll.py`
- `scripts/poland-payroll/process_poland_payroll.py`
- `scripts/uk-payroll/process_uk_payroll.py`
- `scripts/uruguay-payroll/process_uruguay_payroll.py`

Reference and invoke them. Do not rewrite their logic.

### 6. ASCII only in CSV fields and JE memos

All CSV fields and NetSuite JE memos must use plain ASCII text. No Unicode arrows (→), em-dashes (—), or special characters. Spell out directional words: "to", "from", "vs", etc.

### 7. Audit log

Every NetSuite write is appended to `audit_log.json` at the repo root. Format:
```json
{
  "timestamp": "YYYY-MM-DDTHH:MM:SS",
  "skill": "skill-name",
  "action": "POST_JE | GENERATE_CSV | POST_CUSTOMER_PAYMENT",
  "description": "what the entry covers",
  "je_number": "JE#####",
  "netsuite_internal_id": "1234567"
}
```
**Always record both `je_number` (tranid, e.g. `JE#####`) and `netsuite_internal_id` (numeric internal ID).** When reporting to the user in chat, use the `je_number` — most accountants cannot natively search internal IDs in the NetSuite UI.

### 8. Approval workflow — `approved: false` on every API-posted JE

Every record posted via `ns_createRecord` that creates a journal entry MUST include `"approved": false` so it routes to the controller's Pending Approval queue rather than auto-posting. Applies to: `journalentry` and `advintercompanyjournalentry`.

**Exempt — do NOT include `approved`:**
- `customerpayment` applied to specific open invoices
- `check` records
- `deposit` records

Why: a senior accountant should not have unilateral approval rights over manual JEs. Setting `approved: false` enforces the control regardless of role permissions. See `.claude/skills/_shared/approval-required.md` for the full rule.

## Payroll cadence

- **US and Canada:** bi-weekly — two pay runs per month (typically 15th + last day).
- **Germany, Netherlands, Poland, UK, Uruguay:** monthly — single pay run on the last day.

Folder conventions:
- Bi-weekly: `Monthly Payroll/Pay Runs/{Country}/{MM.DD.YYYY}/`
- Monthly:   `Monthly Payroll/Pay Runs/{Country}/{MM-YYYY}/`

## Accrual thresholds (flux-accruals skill)

| Account group | GL range | Threshold (illustrative) |
|---|---|---|
| Software | 671xxx, 511425 | $500 |
| Contractors / Payroll | 511370, 611700 | $100 |
| Professional Fees | 651100, 651150 | $500 |
| COGS General | 511xxx | $500 |

These are example thresholds. Tune for your company's materiality.
