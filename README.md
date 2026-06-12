# Claude Code Skills for Month-End Close on NetSuite

Production-grade [Claude Code](https://docs.claude.com/en/docs/claude-code/overview) skills for the monthly financial close at a multi-entity SaaS company on NetSuite. Built and used in production by a senior accountant — these are the same patterns that close the books every month, repackaged with fictional sample data so you can clone, run, and adapt.

> **Building toward a fully agentic close.** This repo is the foundation; the long-term goal is an end-to-end agent that orchestrates the entire month-end process. Watch for the `agentic-close` skill in a future release.

## Who this is for

- **Senior / staff accountants** at SaaS companies running NetSuite who want to automate the repetitive parts of close: payroll JEs, accruals, intercompany reclasses, bank reconciliation, flux analysis.
- **Accountants in any vertical** who have entities in the US, Canada, Germany, Netherlands, Poland, UK, or Uruguay — the country-specific payroll skills work as-is regardless of industry.
- **Accountants on any ERP** who want to learn the *patterns* (separation-of-duties via CSV upload, local ID-lookup cache, check-and-balance, preflight-before-mapper, structured audit log) — see [Apply these patterns to your stack](#apply-these-patterns-to-your-stack) below.

## What's in the box

### Close-cycle skills

| Skill | What it does |
|---|---|
| [`flux-analysis-workbook`](.claude/skills/flux-analysis-workbook/) | Builds the monthly Flux Analysis Workbook with MoM variance and pivot tables for COGS / Contractors / Professional Fees / Software. Pulls live GL data from NetSuite. |
| [`flux-accruals`](.claude/skills/flux-accruals/) | Budget-driven monthly accrual + reclass identification. Compares FP&A's Vendor Budget to NetSuite actuals + open AP bills, produces a candidate workbook for review, generates JE Import CSVs for approved rows. |
| [`adv-interco-je`](.claude/skills/adv-interco-je/) | Advanced Intercompany Journal Entry generator. Reclasses 681xxx expenses booked under the wrong subsidiary to the right one. Supports any (source, target) pair. |
| [`bank-statement-posting`](.claude/skills/bank-statement-posting/) | Daily bank-statement processor. Classifies every row, matches customer payments to NetSuite open invoices, posts customer payments / checks / deposits, generates JE Import CSVs for transfers and bank fees, produces a paste-back XLSX. |
| [`monthly-cash-confirmation`](.claude/skills/monthly-cash-confirmation/) | Traces each line of FP&A's monthly cash-confirmation request (subleases, debt service, large one-time payments) from the AP tool and banking activity to NetSuite, and reports CONFIRMED / VARIANCE / NOT CONFIRMED per item. Map your own counterparties in `vendor_aliases.example.yaml`. |
| [`monthly-cfs`](.claude/skills/monthly-cfs/) | Builds the monthly Cash Flow Statement workbooks for FP&A (one per subsidiary + a consolidated, FX-embedded book). Regenerates the Income Statement and Balance Sheet tabs from SuiteQL in NetSuite-export format, rolls the indirect-method CFS formulas, re-points month-shifting references, embeds period-end FX rates, repoints the consolidated workbook's external links, and runs a full balance/tie-out validation suite. |
| [`bonus-accrual`](.claude/skills/bonus-accrual/) | Builds the monthly bonus-accrual JE from FP&A's bonus workbook — one balanced, auto-reversing entry per subsidiary that accrues into 231170/231171 through the year. |
| [`je-error-fix`](.claude/skills/je-error-fix/) | Fixes the three most common JE errors: wrong department, period error, GL reclass. Builds a correcting JE. |
| [`avalara-je`](.claude/skills/avalara-je/) | Generates the monthly Avalara sales-tax JE CSV from the AvaTax tax-liability worksheet. |
| [`gut-check`](.claude/skills/gut-check/) | Pre-upload review of generated payroll JE CSVs against the prior 2 same-skill JEs. Catches sign flips, severance dept misroutes, and abnormal variances before upload. |

### Payroll skills (country-specific)

| Skill | Country | Cadence | Highlights |
|---|---|---|---|
| [`us-payroll`](.claude/skills/us-payroll/) | US | Bi-weekly | ADP export → balanced JE; preflight mapping check; department reclasses; gut-check. |
| [`canada-payroll`](.claude/skills/canada-payroll/) | Canada | Bi-weekly | ADP export → balanced JE; RRSP / CRRSP handling; preflight mapping. |
| [`germany-payroll`](.claude/skills/germany-payroll/) | Germany | Monthly | German payroll exports → balanced JE. |
| [`netherlands-payroll`](.claude/skills/netherlands-payroll/) | Netherlands | Monthly | Dutch payroll components → balanced JE; pension Wn handling. |
| [`poland-payroll`](.claude/skills/poland-payroll/) | Poland | Monthly | Polish payroll → balanced JE. |
| [`uk-payroll`](.claude/skills/uk-payroll/) | UK | Monthly | UK payroll including HMRC P32 reconciliation, PAYE, NI, pension. |
| [`uruguay-payroll`](.claude/skills/uruguay-payroll/) | Uruguay | 1-3/mo | Uruguay payroll including Aguinaldo (13th-month bonus) and egreso (termination) calculations. |

## Quick start

### 1. Install Claude Code

If you don't have it: [docs.claude.com/en/docs/claude-code/quickstart](https://docs.claude.com/en/docs/claude-code/quickstart). Works on Mac, Windows, Linux. Available as CLI, desktop app (Mac/Windows), web app, and VS Code / JetBrains extensions.

### 2. Clone this repo into your accounting working directory

```bash
git clone https://github.com/r-goldring/month-end-close-skills.git
cd month-end-close-skills
```

### 3. Open in Claude Code

```bash
claude
```

Claude Code auto-discovers the skills in `.claude/skills/` and the operating instructions in `CLAUDE.md`.

### 4. Try it with the example data

```
> /us-payroll
```

The skill will look in `Monthly Payroll/Pay Runs/US/` for a raw ADP file. Drop [examples/sample-adp-export.csv](examples/sample-adp-export.csv) (renamed to a `G******000060.xlsx` pattern) into a `Monthly Payroll/Pay Runs/US/MM.DD.YYYY/` folder and try it end-to-end.

### 5. Customize for your company

Every skill ships with placeholder identifiers (`Acme Corp`, `Acme Holdings : Acme, Inc.`, bank accounts at `Chase x0001` etc.). Replace these with your real values:

1. **NetSuite reference exports.** Export your Chart of Accounts, Vendors, Customers, Departments, and Subsidiaries from NetSuite as `.xls` or `.csv`. Save under `COA, Depts, Vendors, Customers, NetSuite/`. The skills read these locally to avoid hammering the NetSuite API for ID lookups. See `.claude/skills/_shared/id-lookup-guide.md` for the schema each file should have.

2. **Subsidiary constants.** Edit `.claude/skills/_shared/subsidiary-constants.md`. Replace the `Acme Holdings : Acme, Inc. : Acme {Country}` paths with your own NetSuite subsidiary full-name strings.

3. **Bank accounts.** Copy `.claude/skills/bank-statement-posting/accounts.yaml.example` to `accounts.yaml` and fill in your real bank account GL IDs, subsidiary IDs, and `bank_account_string` patterns.

4. **Department / payroll-code mappings.** For US and Canada payroll, edit `scripts/us-payroll/department-mapping.example.csv` and `scripts/us-payroll/code-mapping.example.csv` (rename without `.example`). These map your ADP cost centers to NetSuite department IDs and ADP earnings/deduction codes to NetSuite GL accounts.

5. **NetSuite MCP.** Configure the [NetSuite MCP server](https://www.netsuite.com/portal/developers/resources/mcp.shtml) in Claude Code so the skills can run SuiteQL queries and post records. See your NetSuite admin for credentials.

## Apply these patterns to your stack

Not on NetSuite? Not in SaaS? The *patterns* in this repo transfer to any close workflow, even if the runnable code doesn't. The ideas worth stealing:

### 1. CSV upload over API write for separation-of-duties
For any record type that requires controller approval, build the data in the script but **don't** post via API — instead generate a CSV the user uploads through the ERP's UI Import Assistant. The record lands in the approval queue. No way to bypass the workflow even if the API role technically can.

### 2. Local reference-file cache for ID lookups
Most ERPs are slow to query for ID resolution. Export your COA, vendors, customers, and departments once a quarter; read locally. Only fall back to API for misses. Speeds skills up 10-100x and reduces API load.

### 3. Preflight before mapper
Before any payroll-style transformation that maps source codes (cost centers, earnings codes) to GL/dept, run a preflight that just *checks* every code maps. If anything's unmapped, surface it to the user with the exact code and suggested mapping — don't post an unbalanced JE.

### 4. Check-and-balance before write
Every JE-producing skill builds the data, then displays a structured preview (account, dept, debit, credit, memo, totals, balance check) and **pauses for user confirmation** before writing. If debits ≠ credits, refuse to proceed.

### 5. Structured audit log
Every record write gets an entry in `audit_log.json` with timestamp, skill, action, description, record number, and internal ID. When something breaks two weeks later, you have a record. When the auditor asks "who posted this?", you have a record. When Claude needs context for next month's close, it has a record.

### 6. Skills cross-reference shared knowledge
Common patterns (SuiteQL snippets, subsidiary IDs, approval rules) live in `.claude/skills/_shared/` and every skill links to them. No duplication, no drift, single source of truth.

## What's deliberately not here

To stay focused on SaaS month-end close, the following are **out of scope** and not covered:

- **Inventory accounting** (perpetual / periodic, costing methods, cycle counts).
- **E-commerce COGS from fulfillment** (3PL invoicing, freight allocation).
- **Project-based revenue recognition** (ASC 606 percent-completion, WIP).
- **Multi-state US sales tax at scale** (only basic Avalara sync is here).
- **Consolidation eliminations** beyond the simple IC reclass case.

PRs adding these are welcome but aren't on the immediate roadmap.

## Architecture

```
.claude/
├── skills/
│   ├── _shared/              # Cross-skill knowledge (subsidiary constants, SuiteQL patterns, approval policy)
│   ├── adv-interco-je/       # Each skill is a directory with SKILL.md
│   ├── bank-statement-posting/
│   ├── flux-accruals/
│   ├── flux-analysis-workbook/
│   ├── je-error-fix/
│   ├── {country}-payroll/
│   ├── gut-check/
│   └── avalara-je/
scripts/
├── _shared/                  # Shared Python utilities (JE CSV writer, skill config loader)
├── {country}-payroll/        # Production Python mappers
├── flux-accruals/
├── flux-analysis-workbook/
└── avalara/
examples/                     # Fictional Acme Corp sample data — runnable end-to-end
CLAUDE.md                     # Operating instructions Claude Code auto-loads
README.md                     # This file
LICENSE                       # MIT
```

## Contributing

Contributions welcome — especially:
- New country payroll skills (e.g., India, Brazil, France, Singapore).
- Other ERPs (Sage Intacct, Workday Financials, Oracle NetSuite alternatives).
- Improvements to the agentic close orchestrator (coming soon).
- Bug fixes to existing skills.

Open an issue first to discuss scope.

## License

[MIT](LICENSE). Use, modify, redistribute, sell as part of a commercial product — just keep the copyright notice.

## Disclaimer

These skills are illustrative, built around fictional "Acme Corp" data. Production use at your company requires you to fill in your own NetSuite IDs, COA, vendors, customers, subsidiary paths, and payroll mappings. The author makes no warranty about fitness for your specific environment — *always* run skills in a test/sandbox NetSuite instance first, and *always* review generated CSVs before uploading.
