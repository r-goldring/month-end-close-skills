---
name: je-error-fix
description: >
  Fix journal entry errors in NetSuite. Handles the three most common error types:
  (1) wrong department — builds a correcting JE that moves the amount to the right dept,
  (2) period error — builds a reversal in the wrong period + re-entry in the correct period,
  (3) GL reclass — builds a correcting JE to move an amount between two GL accounts.
  Use this skill when the user mentions: fix JE, wrong department, period error, reclass
  JE, correcting entry, JE mistake, coded wrong, wrong account, needs to be reversed, or
  asks to move an expense from one account/department/period to another.
---

# JE Error Fix Skill

## Overview

Builds correcting journal entries for the most common NetSuite posting errors. Always
queries the original JE first, shows both the original and the correction, then posts
after explicit confirmation.

**References:**
- SuiteQL patterns: `../_shared/netsuite-queries.md`
- Check-and-balance: `../_shared/check-and-balance.md`
- Subsidiary constants: `../_shared/subsidiary-constants.md`
- ID lookup: `../_shared/id-lookup-guide.md`
- **Approval workflow: `../_shared/approval-required.md`** — **MCP-based JE
  posting is suspended (effective 2026-05-01)**. Correcting JEs from this
  skill MUST be generated as a CSV for upload via the NetSuite UI. Do NOT
  call `ns_createRecord` for journal entries. The CSV upload routes the JE
  through the controller's Pending Approval queue.

---

## Step 1 — Gather error details

Ask the user for:
1. The JE or transaction number (e.g., `JE#####`, `A-163913`)
2. What type of error: wrong department / period error / GL reclass
3. What the correction should be (e.g., "should be Legal not Marketing", "should be March not April")

If the user describes the error without a JE number, ask for it.

---

## Step 2 — Query the original transaction

Use the transaction lookup SuiteQL from `../_shared/netsuite-queries.md`:

```sql
SELECT
    t.tranid, t.trandate, t.recordtype, t.memo AS header_memo,
    a.acctnumber, a.fullname AS account_name,
    tl.memo AS line_memo,
    tl.debitforeignamount AS debit,
    tl.creditforeignamount AS credit,
    sub.fullname AS subsidiary,
    d.fullname AS department,
    tl.id AS line_id, tl.linesequencenumber
FROM transaction t
JOIN transactionline tl  ON t.id = tl.transaction AND tl.mainline = 'F'
JOIN account a           ON tl.account = a.id
LEFT JOIN subsidiary sub ON tl.subsidiary = sub.id
LEFT JOIN department d   ON tl.department = d.id
WHERE t.tranid = '{JE_NUMBER}' AND t.voided = 'F'
ORDER BY tl.linesequencenumber
```

Display the original JE:
```
ORIGINAL — JE##### (March 31, 2026)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 # | Account             | Department       | Debit    | Credit   | Memo
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 1 | 651150 Legal Fees   | Sales & Mktg     | 8,000.00 |          | External Auditor - March legal fees
 2 | 211100 AP           |                  |          | 8,000.00 | External Auditor - March legal fees
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

Ask user to confirm this is the correct transaction and confirm the correction details.

---

## Error Type A — Wrong Department

### When to use
User says: "the department on [JE] should be X, not Y" or "it was coded to the wrong department."

### Correction approach
Build a correcting JE that:
- Credits the wrong line (reverses it out of the wrong department)
- Debits the same account + amount to the correct department

No period change, no account change.

### JE structure

```
CORRECTING JE for JE##### — Wrong Department
Header Memo: Correct Dept - JE##### - {Account}
Date: today (or last day of current open period)
Subsidiary: same as original

 1 | 651150 Legal Fees | Sales & Mktg  |          | 8,000.00 | Correct dept - JE#####
 2 | 651150 Legal Fees | Legal         | 8,000.00 |          | Correct dept - JE#####
```

### Check-and-balance
Show the correction → ask for `yes` → post → log.

---

## Error Type B — Period Error

### When to use
User says: "JE was posted in [wrong month], should be in [correct month]" or "needs to be reversed out of April and re-entered in March."

### Correction approach
Two JEs:
1. **Reversal JE**: All lines flipped (debits → credits, credits → debits), dated **last day of the original wrong period**. Header Memo: `Reverse JE##### - wrong period`
2. **Re-entry JE**: Identical to the original, dated **last day of the correct period**. Header Memo: `Re-enter JE##### - correct period: {Mon-YY}`

### JE structure

```
JE 1 — REVERSAL (original period: March 2026)
Date: 3/31/2026
 1 | 651150 Legal Fees | Legal |          | 8,000.00 | Reverse JE##### - wrong period
 2 | 211100 AP         |       | 8,000.00 |          | Reverse JE##### - wrong period

JE 2 — RE-ENTRY (correct period: February 2026)
Date: 2/28/2026
 1 | 651150 Legal Fees | Legal | 8,000.00 |          | Re-enter JE##### - correct period: Feb-26
 2 | 211100 AP         |       |          | 8,000.00 | Re-enter JE##### - correct period: Feb-26
```

### Check-and-balance
Show both JEs together → ask for `yes` → post both → log two entries.

> **Warning**: If the original period is already closed, NetSuite may not allow posting the reversal there. Warn the user and ask if they want to reverse in the current open period instead.

---

## Error Type C — GL Reclass (Wrong Account)

### When to use
User says: "it was coded to 651100 but should be 671100" or "needs to move from Professional Fees to Software."

### Correction approach
Build a correcting JE in the **same period as the original**:
- Reverse the original account line (debit → credit or credit → debit)
- Post to the correct account in the same direction as the original

### JE structure

```
CORRECTING JE for JE##### — GL Reclass
Header Memo: Reclass JE##### - {wrong account} to {correct account}
Date: same period end as original

 1 | 651100 Prof Fees  | Legal |          | 8,000.00 | Reclass JE##### - 651100 to 671100
 2 | 671100 Software   | Legal | 8,000.00 |          | Reclass JE##### - 651100 to 671100
```

### Check-and-balance
Show correction → ask for `yes` → post → log.

---

## Combination errors

If a JE has both a wrong department AND a wrong account, handle both in a single correcting JE with appropriate lines for each change. Ask the user to confirm all corrections before building.

---

## Multiple affected lines

If the original JE has multiple lines that need correction (e.g., wrong department on several lines), build one correcting JE that addresses all of them. Show the full before/after picture.

---

## Audit log entry format

```json
{
  "timestamp": "...",
  "skill": "je-error-fix",
  "action": "POST_CORRECTING_JE",
  "description": "Corrected JE##### — wrong department: Sales to Legal",
  "netsuite_id": "JE#####"
}
```

For period errors, log both JEs separately.
