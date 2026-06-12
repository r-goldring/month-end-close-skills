import pandas as pd
import sys

def parse_num(val):
    if pd.isna(val): return 0.0
    val = str(val).strip().replace(',', '').replace('"', '').replace(' ', '')
    if not val: return 0.0
    if val.startswith('(') and val.endswith(')'):
        return -float(val[1:-1])
    try:
        return float(val)
    except:
        return 0.0

f1 = r"c:\Users\Accountant\Documents\Finance's Requests\Antigravity\Canada Payroll Automation\Final Output Report.csv"
f2 = r"c:\Users\Accountant\Documents\Finance's Requests\Antigravity\Canada Payroll Automation\Generated_Output_Report.csv"

df_target = pd.read_csv(f1)
df_gen = pd.read_csv(f2)

print(f"Target lines: {len(df_target)}, Generated lines: {len(df_gen)}")

for c in df_target.columns:
    if c not in df_gen.columns:
        print(f"Missing column {c}")

df_target['Amt'] = df_target.iloc[:, -1].apply(parse_num)
df_gen['Amt'] = df_gen.iloc[:, -1].apply(parse_num)

# Group by Account, Department to compare
tgt_grouped = df_target.groupby(['Department', 'Account'])['Amt'].sum().reset_index()
gen_grouped = df_gen.groupby(['Department', 'Account'])['Amt'].sum().reset_index()

merged = pd.merge(tgt_grouped, gen_grouped, on=['Department', 'Account'], how='outer', suffixes=('_target', '_gen'))
merged['Diff'] = merged['Amt_target'].fillna(0) - merged['Amt_gen'].fillna(0)
mismatches = merged[merged['Diff'].abs() > 0.01]

if len(mismatches) > 0:
    print("\n--- Mismatches Found ---")
    print(mismatches.to_string())
else:
    print("\nMATCH: The Generated CSV amounts match the Target Output exactly!")

# Total debits / credits
print(f"\nTarget Net: {df_target['Amt'].sum():.2f}")
print(f"Generated Net: {df_gen['Amt'].sum():.2f}")
