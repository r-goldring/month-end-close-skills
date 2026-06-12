import pandas as pd

xls_path = r"c:\Users\Accountant\Documents\Finance's Requests\Antigravity\Canada Payroll Automation\G9D2KT000060 (raw file).xls"
df1 = pd.read_excel(xls_path, sheet_name='Sheet1', header=None)

row_8_idx = -1
for r in range(5, 12):
    row_str = " ".join([str(x) for x in df1.iloc[r].values if pd.notna(x)])
    if "Earnings" in row_str and "Employee Deds" in row_str:
        row_8_idx = r
        break

if row_8_idx != -1:
    row8 = df1.iloc[row_8_idx].values
    row9 = df1.iloc[row_8_idx + 1].values
    print("Row 8 Non-Nulls:")
    for i, v in enumerate(row8):
        if pd.notna(v): print(f"  {i}: {v}")
    
    print("Row 9 Non-Nulls:")
    for i, v in enumerate(row9):
        if pd.notna(v): print(f"  {i}: {v}")

    categories = ["Earnings", "Employee Deds", "Employee Taxes", "Employer Ded Exp", "Employer Tax Exp"]
    col_pairs = []
    for c in range(len(row8)):
        val = str(row8[c]).strip()
        for cat in categories:
            if cat in val:
                code_col = -1
                amt_col = -1
                print(f"--> Found category '{cat}' at col {c}")
                for c2 in range(c, min(c+10, len(row9))):
                    cat2 = str(row9[c2]).strip()
                    if cat2 == "Code" and code_col == -1:
                        code_col = c2
                        print(f"    Code at {c2}")
                    elif cat2 in ["Amount", "Current Amt"] and amt_col == -1:
                        amt_col = c2
                        print(f"    Amount at {c2} (Value: {cat2})")
                if code_col != -1 and amt_col != -1:
                    col_pairs.append({'code_col': code_col, 'amt_col': amt_col, 'cat': cat})
    
    print("\nFinal col_pairs:", col_pairs)
