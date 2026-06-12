import pandas as pd

xls_path = r"c:\Users\Accountant\Documents\Finance's Requests\Antigravity\Canada Payroll Automation\G9D2KT000060 (raw file).xls"

df1 = pd.read_excel(xls_path, sheet_name='Sheet1', header=None)

# Let's find Business Operations
start_idx = df1[df1[0].astype(str).str.contains("Cost Center: BUSOPS - Business Operations", na=False)].index.to_list()
idx = start_idx[0]
print("--- Sheet 1 Headers (Rows 7-9) ---")
print(df1.iloc[7:10].to_string())

print("\n--- Sheet 1 BUSOPS Block ---")
block = df1.iloc[idx:idx+20]
for r in range(len(block)):
    row = block.iloc[r]
    # find all non-null values and their original column indices
    cells = []
    for c in range(len(row)):
        if pd.notna(row[c]) and str(row[c]).strip() != '':
            cells.append(f"col{c}: {row[c]}")
    if cells:
        print(f"Row {r}: " + " | ".join(cells))

