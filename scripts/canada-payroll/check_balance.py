import pandas as pd

output_csv = r"c:\Users\Accountant\Documents\Finance's Requests\Antigravity\Canada Payroll Automation\Final Output Report.csv"

df = pd.read_csv(output_csv)

# Clean the Debit column
df['Debit'] = df.iloc[:, -1] # It's the last column ' Debit '

def parse_amt(val):
    if pd.isna(val): return 0.0
    val = str(val).strip()
    if not val: return 0.0
    val = val.replace('"', '').replace(',', '')
    if val.startswith('(') and val.endswith(')'):
        return -float(val[1:-1])
    return float(val)

df['Amt'] = df['Debit'].apply(parse_amt)

total_debits = df[df['Amt'] > 0]['Amt'].sum()
total_credits = df[df['Amt'] < 0]['Amt'].sum()

print(f"Total Debits (Expenses): {total_debits:.2f}")
print(f"Total Credits (Liabilities + Contra-Expenses): {total_credits:.2f}")
print(f"Net Check: {total_debits + total_credits:.2f}")

# Look at specifically Liabilities
liabs = df[df['Account'].str.contains('Liability|payable', case=False, na=False)]
print("\nLiabilities:")
for idx, row in liabs.iterrows():
    print(f"{row['Account']}: {row['Amt']}")

print("\nIf Payroll Tax Liability is a plug:")
non_tax_liab = liabs[~liabs['Account'].str.contains('Tax Liability', case=False, na=False)]['Amt'].sum()
print("Sum of other liabilities:", non_tax_liab)

# Sum of everything EXCEPT Payroll Tax Liability
non_tax_sum = df[~df['Account'].str.contains('Tax Liability', case=False, na=False)]['Amt'].sum()
print("Sum of all accounts except Tax Liability:", non_tax_sum)
expected_tax_liab = -non_tax_sum
print("Expected Tax Liability to balance to 0:", expected_tax_liab)

