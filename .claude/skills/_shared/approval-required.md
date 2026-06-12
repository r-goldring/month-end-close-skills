# NetSuite JE Posting Policy — CSV Upload ONLY (effective 2026-05-01)

## Rule (HARD STOP)

**Skills MUST NOT post journal entries via `mcp__claude_ai_NetSuite__ns_createRecord`.** All JE-style transactions go through **CSV upload via the NetSuite UI** so they enter the `Pending Approval` queue under the controller's review.

This applies to every skill that produces a JE:
- `journalentry` — every flavor (payroll, accruals, reclasses, error corrections)
- `advintercompanyjournalentry`

## Why we changed (the SoD failure that prompted this)

On 2026-05-01 we discovered every skill-posted JE was being silently auto-approved by a NetSuite-side process (`lastmodifiedby = -4` with no audit-trail entry for the status change), even though the script set `approved: false` correctly. Examples that were SoD-violated: JE##### (Germany), JE##### (UK), JE#####-JE##### (Uruguay), JE##### (NL).

Root cause is server-side (a workflow / scheduled SuiteScript / saved-form override on the "Acme Corp Journal Entry" custom form auto-flips Pending → Approved within ~10-15 minutes of creation). The script side cannot prevent it. The integration also authenticates as the accountant personally (createdby = 135114), so auditors see "Created by the accountant AND approved" — a textbook SoD failure.

Until NetSuite admin / IT identifies and disables the auto-approver and configures a separate non-approving integration user, **MCP-based JE posting is suspended**. CSV upload routes through the proper UI workflow.

## Exempt records (continue using MCP create)

These do NOT carry an approval field in the NS UI and post normally — keep using `ns_createRecord` with NO `approved` field:

- **`customerpayment`** applied to specific open invoices (routine AR application).
- **`check`** records — Workhorse HR payroll wires, RetireBridge 401k, terminated-employee final pay, etc. (Phase 7 of `bank-statement-posting`).
- **`deposit`** records — Springfield sublease, misc deposits (Phase 7 of `bank-statement-posting`).

These are exempt because the JE-style approval workflow doesn't apply to them and they have no auto-approve issue.

## CSV upload workflow (the new pattern for all JE skills)

Every JE-producing skill must:

1. **Generate a CSV** alongside the existing Backup.xlsx in the pay-run / month folder.
   - Filename pattern: `{YYYY-MM} {Skill Description} JE Import.csv` (e.g., `2026-04 Uruguay Payroll JE Import.csv`, `2026-04 Uruguay Payroll - Aguinaldo Accrual JE Import.csv`).
   - One CSV per NetSuite JE that would have been posted.
   - Standard columns: `Date, Journal Entry Memo, Account, Debit, Credit, Line Memo, Subsidiary, Department`
   - Foreign-currency JEs (UY, etc.) add a `Currency` column after `Journal Entry Memo`.
   - Date format: `M/D/YYYY` (e.g., `4/30/2026`) to match NetSuite's CSV import expectations.
   - Empty cells (no Debit on a credit row, or no Department on a liability line) stay empty (no zeros).
   - Quote any field containing a comma (e.g., the subsidiary path with `Acme, Inc.`).

2. **Display the JE preview in chat** as before (categories, totals, dept breakdown).

3. **Tell the accountant to upload** the CSV files via NetSuite UI:
   > "CSV file(s) ready at: {path}.\n\nUpload via NetSuite: Lists → Import Assistant → Import Type: Transactions → Record Type: Journal Entry → upload the CSV → confirm field mapping → run the import. Each imported JE will land in the controller's Pending Approval queue."

4. **DO NOT call `ns_createRecord` for the JE.** Do not include any post step that would invoke the MCP for journal entries.

5. **Append to `audit_log.json`** with `action: "GENERATE_CSV"` (instead of `POST_JE`) and `posted_via: "CSV upload pending"`. Once the accountant tells you the JE numbers after his manual import, append a follow-up entry with the actual JE numbers.

## Reference CSV examples

`COA, Depts, Vendors, Customers, NetSuite/Example Journal Entry CSVs/` contains many historical examples the accountant has used:
- `2026-01 Germany Payroll JE Import.csv`
- `2026-01 Uruguay Payroll JE Import.csv`
- `2024-08 Bonus Accrual JE Import.csv` (flux-accruals format)
- `Adv JE UK Feb-26 (Brex).csv` (advanced intercompany format)

Match the column structure and date/quoting conventions from the closest example.

## Cross-reference (skills affected)

CSV-only (do NOT use ns_createRecord for JEs):
- `germany-payroll`, `netherlands-payroll`, `poland-payroll`, `uk-payroll`, `uruguay-payroll`
- `us-payroll`, `canada-payroll`, `avalara-je`
- `flux-accruals` (accrual JEs)
- `je-error-fix` (correcting JEs)
- `adv-interco-je` (already CSV-based)
- `bank-statement-posting` — JE phases (Phase 6a regular JE, Phase 6b Adv IC JE) only

MCP create still allowed for these record types:
- `customerpayment` — `bank-statement-posting` Phase 5
- `check` — `bank-statement-posting` Phase 7
- `deposit` — `bank-statement-posting` Phase 7

## When the SoD issue is fixed

When IT/admin identifies and disables the auto-approver (and ideally creates a separate non-approving integration user), this policy will be revisited. Until then, **CSV upload is mandatory for every JE the skills produce**.
