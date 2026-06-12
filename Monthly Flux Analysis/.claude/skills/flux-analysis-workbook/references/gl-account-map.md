# GL Account Mapping Reference - Flux Analysis

**Created:** 2026-04-06
**Source Files:**
- netsuite_skill_summary.json (Chart of Accounts)
- feb2026_template_spec.json (Flux Analysis Template)

---

## 1. Income Statement Accounts

| Account # | Account Name | Type | Detail Tab | Typical Comment Pattern |
|-----------|--------------|------|-----------|------------------------|
| 411000 | Revenue | Header | — | — |
| 411100 | Subscription - Software | Detail | Revenue | "SaaS revenue" |
| 411150 | Professional Services | Detail | Revenue | "Professional services revenue" |
| 411160 | Subscription - Services | Detail | Revenue | "Services subscription revenue" |
| 411200 | Other | Detail | Revenue | "Other revenue" |
| 411300 | Revenue Reserves | Detail | Revenue | "Revenue reserve adjustment" |
| **Total - 411000** | **Total Revenue** | **Subtotal** | **—** | **—** |
| 415000 | Intercompany Revenue | Detail | — | "Intercompany transfer" |
| **Total - Income** | **Total Income** | **Subtotal** | **—** | **—** |
| 511000 | Cost of Goods Sold | Header | — | — |
| 511400 | COGS - Hosting | Detail | COGS | "Hosting/infrastructure costs" |
| 511425 | COGS - Software Subscriptions | Detail | COGS | "Third-party software subscriptions" |
| 511450 | COGS - Translation | Detail | COGS | "Translation services" |
| 511510 | COGS - Client Billable | Detail | COGS | "Billable client expenses" |
| 511520 | COGS - Client Non Billable | Detail | COGS | "Non-billable client expenses" |
| 511550 | COGS - Paper Survey | Detail | COGS | "Paper survey costs" |
| 511600 | COGS - Other | Detail | COGS | "Other COGS expenses" |
| **Total - 511000** | **Total COGS (exc. Salary)** | **Subtotal** | **—** | **—** |
| 511001 | COGS - Salary and Compensation | Header | Contractors | "COGS salary/contractor costs" |
| 511100 | COGS - Salaries and Wages | Detail | Contractors | "COGS direct salaries" |
| 511150 | COGS - Bonus | Detail | — | "COGS bonus accrual" |
| 511175 | COGS - Severance | Detail | — | "COGS severance" |
| 511200 | COGS - Health Benefits | Detail | — | "COGS health insurance" |
| 511250 | COGS - 401k Match | Detail | — | "COGS 401k match" |
| 511300 | COGS - Other Benefits | Detail | — | "COGS benefits" |
| 511350 | COGS - Payroll Taxes | Detail | — | "COGS payroll taxes" |
| 511370 | COGS - Contractor Payroll | Detail | Contractors | "COGS contractor payments" |

---

## 2. Balance Sheet Accounts

| Account # | Account Name | Type | Detail Tab | Typical Comment Pattern |
|-----------|--------------|------|-----------|------------------------|
| **ASSETS** | | Header | — | — |
| **Current Assets** | | Header | — | — |
| **Bank** | | Header | — | — |
| 111000 | Cash and Cash Equivalents | Header | — | — |
| 111001 | Wells Fargo CEO Checking x6035 | Detail | — | "Checking deposit" |
| 111005 | Wells Fargo MM | Detail | — | "Money market account" |
| 111007 | ING x0005 - EUR | Detail | — | "EUR bank account" |
| 111008 | ING - USD | Detail | — | "USD bank account" |
| 111009 | ING - GBP | Detail | — | "GBP bank account" |
| 111010 | ING - EUR | Detail | — | "EUR bank account" |
| 111011 | TD/Barclays - Canada | Detail | — | "CAD bank account" |
| 111012 | Bank of America - Poland PLN | Detail | — | "PLN bank account" |
| 111013 | Bank of America - Poland USD | Detail | — | "USD bank account (Poland)" |
| 111015 | Barclays - USD | Detail | — | "USD bank account" |
| 111016 | Petty Cash | Detail | — | "Petty cash" |
| 111017 | SVB x9054 | Detail | — | "Bank deposit" |
| 111018 | Chase x0395 | Detail | — | "Bank deposit" |
| 111019 | Chase x2079 | Detail | — | "Bank deposit" |
| 111020 | Chase x7012 | Detail | — | "Bank deposit" |
| 111021 | Chase x9687 | Detail | — | "Bank deposit" |

---

## 3. Detail Tab GL Filters

### COGS Tab
**GL Accounts included in COGS detail tab:**
- `511400` - COGS - Hosting
- `511425` - COGS - Software Subscriptions
- `511450` - COGS - Translation
- `511510` - COGS - Client Billable
- `511520` - COGS - Client Non Billable (shown in template)
- `511550` - COGS - Paper Survey
- `511600` - COGS - Other

### Contractors Tab
**GL Accounts included in Contractors detail tab:**
- `511001` - COGS - Salary and Compensation (includes 511370)
- `511370` - COGS - Contractor Payroll (sub-account of 511001)
- `611000` - Salary and Compensation (parent account)
- `611700` - Contractor Payroll (sub-account of 611000)

### Professional Fees Tab
**GL Accounts included in Professional Fees detail tab:**
- `651000` - Professional Fees (parent account)
- `651100` - Professional Fees (detail line in flux)
- Includes vendors: UY Legal Counsel, External Audit Firm, External Auditor, legal services, etc.

### Software Tab
**GL Accounts included in Software detail tab:**
- `671000` - Software (parent account)
- `671100` - Software Subscriptions (detail line in flux)
- Includes vendors: SecuritySaaS, Adobe, Anthropic, Atlassian, BillFlow, External Audit Firm, Salesforce, etc.

### Revenue Tab
**GL Accounts included in Revenue detail tab:**
- `411000` - Revenue (parent account)
- `411100` - Subscription - Software
- `411150` - Professional Services
- `411160` - Subscription - Services
- `411200` - Other

---

## 4. Cross-Reference Lookup Table

**Quick reference: "If I see a variance on account X, which detail tab has more info?"**

| Account | Account Name | Check Detail Tab | Notes |
|---------|--------------|------------------|-------|
| 411100 | Subscription - Software | **Revenue** | SaaS revenue breakdown by customer |
| 411150 | Professional Services | **Revenue** | Services revenue by project/customer |
| 411160 | Subscription - Services | **Revenue** | Services subscriptions by customer |
| 411200 | Other Revenue | **Revenue** | Other revenue sources |
| 411300 | Revenue Reserves | **Revenue** | Reserve adjustments |
| 415000 | Intercompany Revenue | — | Intercompany transactions (no detail tab) |
| 511400 | COGS - Hosting | **COGS** | Hosting/infrastructure vendors |
| 511425 | COGS - Software Subscriptions | **COGS** | Third-party software by vendor |
| 511450 | COGS - Translation | **COGS** | Translation vendor details |
| 511510 | COGS - Client Billable | **COGS** | Billable client expenses by type |
| 511520 | COGS - Client Non Billable | **COGS** | Non-billable client costs |
| 511550 | COGS - Paper Survey | **COGS** | Paper survey vendor details |
| 511600 | COGS - Other | **COGS** | Other COGS expenses |
| 511001 | COGS - Salary & Comp | **Contractors** | COGS contractor payroll |
| 511100 | COGS - Salaries | **Contractors** | COGS salary details (limited detail tab data) |
| 511150 | COGS - Bonus | — | No detail tab in flux analysis |
| 511175 | COGS - Severance | — | No detail tab in flux analysis |
| 511200 | COGS - Health Benefits | — | No detail tab in flux analysis |
| 511250 | COGS - 401k Match | — | No detail tab in flux analysis |
| 511300 | COGS - Other Benefits | — | No detail tab in flux analysis |
| 511350 | COGS - Payroll Taxes | — | No detail tab in flux analysis |
| 511370 | COGS - Contractor Payroll | **Contractors** | Contractor vendor/resource breakdown |
| 611000 | Salary and Compensation | **Contractors** | OpEx contractor payroll |
| 611100 | Salaries and Wages | — | No detail tab for OpEx salaries |
| 651000 | Professional Fees | **Professional Fees** | External advisor/service vendors |
| 651100 | Professional Fees | **Professional Fees** | Accounting, legal, consulting services |
| 651150 | Legal Fees | **Professional Fees** | Legal vendor breakdown |
| 651160 | Bank Fees | — | No detail tab in flux analysis |
| 671000 | Software | **Software** | Software subscriptions by vendor |
| 671100 | Software Subscriptions | **Software** | SaaS/tool vendor breakdown |

---

## Notes

### Hierarchy Structure
- **Headers** (no values): 411000, 511000, 511001, 111000, etc. - organizational categories only
- **Details**: Individual accounts with transaction activity
- **Subtotals**: "Total - {Account}" - calculated sums of detail lines
- **Parent/Child**: Some accounts (e.g., 511370) are children of parent accounts (511001)

### Detail Tab Coverage
- **Revenue**: Full coverage of all revenue accounts (411100-411300)
- **COGS**: Covers non-salary COGS (511400-511600) + contractor payroll (511370)
- **Contractors**: COGS and OpEx contractor payroll (511370 and 611700)
- **Professional Fees**: Account 651000 family (651100, 651150, etc.)
- **Software**: Account 671000 family (671100 and subaccounts)

### Flux Analysis Context
The February 2026 template shows:
- Income Statement: 5 revenue detail accounts + 7 COGS detail accounts + salary/benefits header
- Balance Sheet: Bank accounts (111000-111021) and other asset accounts
- Detail tabs driven by vendor/transaction activity (e.g., Software tab lists 30+ software vendors)
