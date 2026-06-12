# Check-and-Balance Pattern

Every skill that writes to NetSuite (posts JEs, applies payments, creates records) MUST follow this exact 4-step pattern.

## The pattern

### Step 1 — Build in memory
Construct the full operation (all JE lines, all field values) before touching NetSuite. Do not make any write calls until after the user confirms.

### Step 2 — Display preview
Show a formatted table of everything that will be posted. Minimum columns:
- For JEs: Line | Account | Department | Debit | Credit | Memo
- For payments: Customer | Invoice | Amount | Date | GL Account

Format numbers with commas and 2 decimal places. Show totals (debits must equal credits).

Example:
```
PREVIEW — April 2026 Accruals JE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 # | Account              | Department        | Debit      | Credit
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 1 | 671100 Software Sub  | Engineering       |  12,500.00 |
 2 | 231100 Accrued Liab  |                   |            |  12,500.00
 3 | 651100 Prof Fees     | Legal             |   8,000.00 |
 4 | 231100 Accrued Liab  |                   |            |   8,000.00
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   TOTALS                                     |  20,500.00 |  20,500.00
```

### Step 3 — Ask for confirmation
After the preview, always print this exact line:

```
Review the above carefully. Post to NetSuite? Type 'yes' to confirm or anything else to cancel.
```

Wait for the user's response. Only proceed if they type exactly `yes` (case-insensitive). Any other response cancels the operation.

### Step 4 — Post and log
After confirmed `yes`:
1. Execute the NetSuite MCP call
2. Capture the returned transaction ID
3. Append to `audit_log.json`:

```python
import json, datetime, os

entry = {
    "timestamp": datetime.datetime.now().isoformat(timespec='seconds'),
    "skill": "skill-name-here",
    "action": "POST_JE",  # or APPLY_PAYMENT, CREATE_RECORD, etc.
    "description": "One-line description",
    "netsuite_id": "JE#####"  # from MCP response
}

log_path = os.path.join(os.path.dirname(__file__), "../../../../audit_log.json")
with open(log_path, "r") as f:
    log = json.load(f)
log.append(entry)
with open(log_path, "w") as f:
    json.dump(log, f, indent=2)
```

4. Confirm to the user: "Posted successfully. NetSuite ID: JE#####"

## What NOT to do
- Never post silently without previewing first
- Never skip the confirmation prompt even if you're "confident" the data is correct
- Never proceed if the user types anything other than `yes` (including "yes please", "y", "go ahead")
