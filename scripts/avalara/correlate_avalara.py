import pandas as pd

path = r"c:\Users\Accountant\Documents\Finance's Requests\Antigravity\Avalara JE\Historical Examples\January 2026\TaxLiabilityWorksheetSummaryReturnDetail (10).xlsx"
df = pd.read_excel(path, header=None)

header_idx = -1
for i, row in df.iterrows():
    if 'State' in [str(x) for x in row.values]:
        header_idx = i
        break

df.columns = df.iloc[header_idx].values
df = df.iloc[header_idx+1:].reset_index(drop=True)
df = df.dropna(subset=['State'])

for state_code in ['OH', 'PA', 'SC']:
    print(f"\n--- {state_code} ---")
    state_df = df[df['State'].astype(str).str.contains(state_code)]
    # keep only relevant numeric columns
    cols = ['Return Month Sales Tax Liability', 'Current Period Vendor Discount', 'Prior Period Vendor Discount', 'Other Adjustments', 'Amount Due To Avalara']
    for _, row in state_df.iterrows():
        for c in cols:
            if c in row:
                print(f"{c}: {row[c]}")
