import pandas as pd
import os

paths = [
    r"c:\Users\Accountant\Documents\Finance's Requests\Antigravity\Avalara JE\Historical Examples\January 2026\TaxLiabilityWorksheetSummaryReturnDetail (10).xlsx",
    r"c:\Users\Accountant\Documents\Finance's Requests\Antigravity\Avalara JE\Historical Examples\January 2026\TransactionImpactResults355.xls",
    r"c:\Users\Accountant\Documents\Finance's Requests\Antigravity\Avalara JE\Historical Examples\December 2025\TaxLiabilityWorksheetSummaryReturnDetail (11).xlsx",
    r"c:\Users\Accountant\Documents\Finance's Requests\Antigravity\Avalara JE\Historical Examples\December 2025\TransactionImpactResults382.xls"
]

for p in paths:
    print(f"\n\n--- {os.path.basename(p)} ---")
    df = None
    try:
        # Some .xls files are actually HTML or CSV, pandas read_html or read_csv might be needed
        # Let's try read_excel first, then read_csv
        try:
            df = pd.read_excel(p)
        except Exception as e1:
            try:
                df = pd.read_csv(p, sep='\t') # sometimes .xls are tab separated
            except Exception as e2:
                try:
                    df = pd.read_csv(p)
                except Exception as e3:
                    try:
                        df = pd.read_html(p)[0]
                    except Exception as e4:
                        print(f"Failed to read. e1:{e1}")
                        continue
        if df is not None:
            print("Columns:", df.columns.tolist())
            print(df.head(10).to_string())
    except Exception as e:
        print("Error:", e)
