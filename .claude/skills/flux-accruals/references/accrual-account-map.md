# Accrual Account Mapping

Maps expense accounts to their corresponding accrued liability accounts for month-end accrual JEs.

## Standard mapping

| Expense Account | Expense Description | Accrued Liability Account | Notes |
|----------------|--------------------|-----------------------------|-------|
| 671100 | Software Subscriptions | 231100 Accrued Liabilities | General accrual |
| 671000 | Software | 231100 Accrued Liabilities | |
| 511425 | Software - COGS | 231100 Accrued Liabilities | |
| 511370 | Contractor Payroll - COGS | 231100 Accrued Liabilities | |
| 611700 | Contractor Payroll - OpEx | 231100 Accrued Liabilities | |
| 651100 | Professional Fees | 231100 Accrued Liabilities | |
| 651150 | Legal Fees | 231100 Accrued Liabilities | |
| 511xxx (other COGS) | Various COGS | 231100 Accrued Liabilities | |

## JE memo format for accruals

```
{Vendor Name} - {Mon-YY} accrual
```
Examples:
- `Salesforce - Apr-26 accrual`
- `PEO Provider - Apr-26 accrual`
- `Acme UK External Auditor - Apr-26 accrual`

## JE header memo

```
{Mon-YY} Month-End Accruals
```
Example: `Apr-26 Month-End Accruals`

## Subsidiary for accrual JEs

- Post to `Acme Holdings : Acme, Inc.` (US) for most accruals
- If the vendor is clearly a subsidiary-specific vendor (e.g., KWPS BV for Netherlands),
  confirm with the accountant before posting to a different subsidiary
