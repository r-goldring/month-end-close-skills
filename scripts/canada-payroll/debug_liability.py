import pandas as pd

def parse_amt(val):
    if pd.isna(val): return 0.0
    val = str(val).strip().replace(',', '').replace('"', '').replace('$', '')
    if not val: return 0.0
    if val.startswith('(') and val.endswith(')'):
        return -float(val[1:-1])
    try:
        return float(val)
    except:
        return 0.0

raw_xls = r"c:\Users\Accountant\Documents\Finance's Requests\Antigravity\Canada Payroll Automation\G9D2KT000060 (raw file).xls"
df2 = pd.read_excel(raw_xls, sheet_name='Sheet2', header=None)

ee_taxes = 0.0
er_taxes = 0.0
col_ee_amount = -1
col_er_amount = -1

for row_idx in range(min(20, len(df2))):
    row_vals = df2.iloc[row_idx].values
    for c in range(len(row_vals)):
        val = str(row_vals[c]).strip()
        if "Employee Taxes" in val:
            print(f"Found Employee Taxes at row {row_idx}, col {c}")
            for r2 in range(row_idx+1, min(row_idx+5, len(df2))):
                for c2 in range(c, min(c+10, len(row_vals))): # Note I changed c+1 to c
                    if "Amount" in str(df2.iat[r2, c2]).strip():
                        col_ee_amount = c2
                        print(f"  -> Found Amount at {r2}, {c2}")
                        break
                if col_ee_amount != -1: break
        if "Employer Tax Exp" in val:
            print(f"Found Employer Tax Exp at row {row_idx}, col {c}")
            for r2 in range(row_idx+1, min(row_idx+5, len(df2))):
                for c2 in range(c, min(c+10, len(row_vals))):
                    if "Amount" in str(df2.iat[r2, c2]).strip():
                        col_er_amount = c2
                        print(f"  -> Found Amount at {r2}, {c2}")
                        break
                if col_er_amount != -1: break

count_rt = 0
for row_idx in range(len(df2)):
    if "Report Totals:" in " ".join([str(v) for v in df2.iloc[row_idx].values if pd.notna(v)]):
        count_rt += 1
        print(f"Found Report Totals #{count_rt} at {row_idx}")
        if count_rt == 2:
            # Need to search rows below for "Number of Pays" as the anchor
            # to accommodate if it's row_idx+1 or row_idx+2
            sum_row = -1
            for offset in range(1, 4):
                joined = " ".join([str(v) for v in df2.iloc[row_idx + offset].values if pd.notna(v)])
                if "Number of Pays" in joined or "531,374.22" in joined: # heuristic
                    sum_row = row_idx + offset
                    break
            if sum_row != -1:
                next_row = df2.iloc[sum_row].values
                print("Sum row:", " | ".join([f"{i}:{v}" for i,v in enumerate(next_row) if pd.notna(v)]))
                if col_ee_amount != -1 and col_ee_amount < len(next_row):
                    ee_taxes = parse_amt(next_row[col_ee_amount])
                if col_er_amount != -1 and col_er_amount < len(next_row):
                    er_taxes = parse_amt(next_row[col_er_amount])
            break

print(f"EE Taxes: {ee_taxes}, ER Taxes: {er_taxes}")
