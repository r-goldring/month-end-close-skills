import pandas as pd
import sys

xls_path = r"c:\Users\Accountant\Documents\Finance's Requests\Antigravity\Canada Payroll Automation\G9D2KT000060 (raw file).xls"

df = pd.read_excel(xls_path, sheet_name='Sheet2', header=None)
print("Sheet2 rows 5 to 13:")
print(df.iloc[5:13].to_string())
