import re
import os

files = [
    r"c:\Users\Accountant\Documents\Finance's Requests\Antigravity\Avalara JE\Historical Examples\January 2026\TransactionImpactResults355.xls",
    r"c:\Users\Accountant\Documents\Finance's Requests\Antigravity\Avalara JE\Historical Examples\December 2025\TransactionImpactResults382.xls",
    r"c:\Users\Accountant\Documents\Finance's Requests\Antigravity\Avalara JE\Historical Examples\February 2026\TransactionImpactResults371.xls"
]

for path in files:
    print(f"\n--- {os.path.basename(path)} ---")
    with open(path, 'r', encoding='utf-8', errors='replace') as f:
        text = f.read()

    rows = re.split(r'<Row[^>]*>', text)
    for i, row in enumerate(rows[1:]): # skip first split
        cells = re.split(r'<Cell[^>]*>', row)
        row_data = []
        for cell in cells[1:]:
            match = re.search(r'<Data[^>]*>(.*?)</Data>', cell, re.DOTALL)
            if match:
                val = match.group(1).replace('&#10;', ' ').strip()
                val = re.sub(r'<[^>]+>', '', val)
                row_data.append(val)
            else:
                row_data.append('')
        
        if len(row_data) > 0 and '651210' in row_data[0]:
            print(row_data)
