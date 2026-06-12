import pandas as pd
df = pd.read_csv(r"c:\Users\Accountant\Documents\Finance's Requests\Antigravity\Canada Payroll Automation\Generated_Output_Report.csv")

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

df['Amt'] = df.iloc[:, -1].apply(parse_num)

salaries = df[df['Account'].str.contains('Salaries and Wages', case=False, na=False)]['Amt'].sum()
k401 = df[(df['Account'].str.contains('401k Match', case=False, na=False)) & ~(df['Account'].str.contains('payable', case=False, na=False))]['Amt'].sum()
taxes = df[(df['Account'].str.contains('Payroll Taxes', case=False, na=False)) & ~(df['Account'].str.contains('Liability', case=False, na=False))]['Amt'].sum()
commissions = df[df['Account'].str.contains('Commission', case=False, na=False)]['Amt'].sum()
bonus = df[df['Account'].str.contains('Bonus', case=False, na=False)]['Amt'].sum()
other_ben = df[df['Account'].str.contains('Other Benefits', case=False, na=False)]['Amt'].sum()

net_pay = df[df['Account'].str.contains('Payroll Liability', case=False, na=False)]['Amt'].sum()
k401_liab = df[df['Account'].str.contains('401K payable', case=False, na=False)]['Amt'].sum()
tax_liab = df[df['Account'].str.contains('Payroll Tax Liability', case=False, na=False)]['Amt'].sum()

print(f"Salaries: {salaries:.2f} (Target: 280365.76)")
print(f"401k Match: {k401:.2f} (Target: 4384.90)")
print(f"Payroll Tax Exp: {taxes:.2f} (Target: 18269.96)")
print(f"Commissions: {commissions:.2f} (Target: 0)")
print(f"Bonus: {bonus:.2f} (Target: 0)")
print(f"Other Benefits: {other_ben:.2f} (Target: -2234.04)")
print(f"Payroll Liability: {net_pay:.2f} (Target: -183365.28)")
print(f"401k Liability: {k401_liab:.2f} (Target: -9842.54)")
print(f"Payroll Tax Liab: {tax_liab:.2f} (Target: -107578.76)")
print(f"Total Debits: {salaries + k401 + taxes + commissions + bonus + other_ben:,.2f}")
print(f"Total Credits: {net_pay + k401_liab + tax_liab:,.2f}")
