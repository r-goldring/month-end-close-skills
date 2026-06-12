import pandas as pd
import sys

raw_xls = r"c:\Users\Accountant\Documents\Finance's Requests\Antigravity\Canada Payroll Automation\G9D2KT000060 (raw file).xls"
dept_map_csv = r"c:\Users\Accountant\Documents\Finance's Requests\Antigravity\Canada Payroll Automation\Canada Payroll Department Mapping File.csv"

dept_df = pd.read_csv(dept_map_csv)
dept_map = {}
for _, row in dept_df.iterrows():
    cc_full = str(row['Payroll Report Cost Center']).strip()
    cc = cc_full.split("Cost Center: ")[1].strip() if "Cost Center: " in cc_full else cc_full
    dept_map[cc] = 1

df1 = pd.read_excel(raw_xls, sheet_name='Sheet1', header=None)

missing = []
for row_idx in range(len(df1)):
    row_str_joined = " ".join([str(v) for v in df1.iloc[row_idx].values if pd.notna(v)])
    if "Group Summary for:" in row_str_joined and "Cost Center:" in row_str_joined:
        cc = row_str_joined.split("Cost Center:")[1].strip()
        found = False
        for k in dept_map:
            if k in cc: found = True
        if not found and cc != "Revenue":
            missing.append(cc)

print("Missing Cost Centers in Map:", missing)
