# ID Lookup Guide

Before calling any NetSuite MCP tool that requires an internal ID, look up the ID from the local reference files. This is faster and avoids API calls.

## Reference files location

```
COA, Depts, Vendors, Customers, NetSuite/Updated Files 4.17.26 (with InternalIDs)/
├── ChartofAccounts812.xls   — GL accounts
├── Customers262.xls         — Customers
├── Departments608.xls       — Departments
├── VendorsResults678.xls    — Vendors
├── Subsidiaries467.xls      — Subsidiaries
└── Items266.xls             — Items
```

## How to look up a GL account internal ID

1. Read `ChartofAccounts812.xls` using pandas (`.xls` requires `engine='xlrd'`)
2. Filter where the account number column matches (e.g., `'121200'`)
3. Return the internal ID column value

```python
import pandas as pd
df = pd.read_excel(
    r"COA, Depts, Vendors, Customers, NetSuite/Updated Files 4.17.26 (with InternalIDs)/ChartofAccounts812.xls",
    engine='xlrd'
)
# Print column names to find the right ones first
print(df.columns.tolist())
# Then filter: df[df['Account Number'] == '121200']['Internal ID'].iloc[0]
```

## How to look up a vendor internal ID

```python
df = pd.read_excel("..VendorsResults678.xls", engine='xlrd')
# Find by company name (case-insensitive partial match)
match = df[df['Company Name'].str.lower().str.contains('presidio', na=False)]
# Return match['Internal ID'].iloc[0]
```

## How to look up a department internal ID

```python
df = pd.read_excel("..Departments608.xls", engine='xlrd')
# Filter by department name
match = df[df['Name'].str.lower().str.contains('legal', na=False)]
```

## Fallback to NetSuite MCP

If the entity is not found in the local file, fall back to SuiteQL:
```sql
SELECT id, companyname FROM vendor WHERE LOWER(companyname) LIKE '%presidio%'
```

Always warn the user when using MCP fallback: "Vendor not found in local reference file (may have been added after 4/17/26) — queried NetSuite directly."

## When to refresh the reference files

If you encounter multiple MCP fallbacks in a session, suggest to the accountant: "Several IDs were not found locally. Run `/refresh-reference-files` or re-download the reference files from NetSuite."
