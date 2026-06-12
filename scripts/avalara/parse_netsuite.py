import re

path = r"c:\Users\Accountant\Documents\Finance's Requests\Antigravity\Avalara JE\Historical Examples\January 2026\TransactionImpactResults355.xls"

with open(path, 'r', encoding='utf-8', errors='replace') as f:
    text = f.read()

# Split by <Row> or <Row ...>
rows = re.split(r'<Row[^>]*>', text)
print(f"Found {len(rows)-1} rows.")
for i, row in enumerate(rows[1:20]): # skip first split
    cells = re.split(r'<Cell[^>]*>', row)
    row_data = []
    for cell in cells[1:]:
        # extract content between <Data> and </Data>
        match = re.search(r'<Data[^>]*>(.*?)</Data>', cell, re.DOTALL)
        if match:
            # clean up any xml entities
            val = match.group(1).replace('&#10;', ' ').strip()
            val = re.sub(r'<[^>]+>', '', val) # remove inner tags just in case
            row_data.append(val)
        else:
            row_data.append('')
    if any(row_data):
        print(f"Row {i}:", row_data)
