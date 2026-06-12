import pandas as pd

raw_xls = r"c:\Users\Accountant\Documents\Finance's Requests\Antigravity\Canada Payroll Automation\G9D2KT000060 (raw file).xls"
df2 = pd.read_excel(raw_xls, sheet_name='Sheet2', header=None)

count_rt = 0
for row_idx in range(len(df2)):
    row_str_joined = " ".join([str(v) for v in df2.iloc[row_idx].values if pd.notna(v)])
    if "Report Totals:" in row_str_joined:
        count_rt += 1
        if count_rt == 2:
            print(f"Found second Report Totals at row_idx {row_idx}")
            for offset in range(-1, 5):
                if row_idx + offset < len(df2):
                    row_vals = df2.iloc[row_idx + offset].values
                    joined = " ".join([str(v) for v in row_vals if pd.notna(v)])
                    print(f"row {row_idx+offset}: {joined}")
            break
