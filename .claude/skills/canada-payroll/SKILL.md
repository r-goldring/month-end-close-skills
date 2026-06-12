---
name: canada-payroll
description: >
  Process Canada bi-weekly ADP payroll into a balanced NetSuite journal entry import CSV.
  Reads the raw .xls export from ADP, validates mappings via preflight, generates a
  Backup.xlsx + JE Import CSV (NetSuite Import Assistant format).
  Use this skill when the user mentions: Canada payroll, Canadian payroll, Canada ADP,
  Canada payroll JE, RRSP, CRRSP, CRRVF, CRSPV, or drops a G*000060.xls file into
  Monthly Payroll/Pay Runs/Canada/{MM.DD.YYYY}/.
---

# Canada Payroll JE Skill

## Overview

Processes the bi-weekly ADP Canada payroll export into a balanced NetSuite journal
entry. Mirrors the US/foreign payroll skill pattern: preflight check for new cost
centers and codes, then the main script that aggregates per-employee Earnings into
the COGS/OpEx GL split (with 70/30 Infrastructure split) and reads liability totals
from Sheet2.

**Cadence:** Canada is **bi-weekly** — two pay runs per month (15th + last day).
Each pay run gets its own folder and its own JE.

**Subsidiary:** `Acme Holdings : Acme, Inc. : Acme Canada`.

**Currency:** CAD throughout — no USD conversion in the JE.

**Scripts:**
- [`scripts/canada-payroll/check_mappings.py`](../../../scripts/canada-payroll/check_mappings.py) — preflight
- [`scripts/canada-payroll/process_canada_payroll.py`](../../../scripts/canada-payroll/process_canada_payroll.py) — JE generator

**Mapping files** (authoritative knowledge base — update here when new codes appear):
- `scripts/canada-payroll/Canada Payroll Department Mapping File.csv` — Cost Center to Department
- `scripts/canada-payroll/Payroll_Mapping_with_GL_Accounts (final).csv` — code to GL account (shared file with US payroll)

## Step 1 — the accountant drops the raw file

the accountant drops the ADP export into:
```
Monthly Payroll/Pay Runs/Canada/{MM.DD.YYYY}/
```
(e.g., `04.30.2026`). Filename pattern: `G*000060.xls` (ADP job number).

Sheets:
- `Sheet1` — per-employee Earnings, Employee Deds, Employee Taxes, Employer Ded Exp, Employer Tax Exp (grouped by Cost Center)
- `Sheet2` — Report Totals: aggregate code totals + Total Net pay + RRSP totals (CRRSP/CRRVF/CRSPV) + EE/ER tax totals

## Step 2 — REQUIRED: Run the preflight

Always run the mapping preflight before the main script. It scans the raw file for
Cost Centers and payroll Codes and flags any that aren't mapped, so they can be
added to the knowledge base instead of being silently dropped from the JE.

```bash
cd scripts/canada-payroll
python check_mappings.py "../../Monthly Payroll/Pay Runs/Canada/MM.DD.YYYY/G*.xls"
```

Categories:
- **OK**: mapped — no action needed
- **NOTE**: informational (`Revenue` cost center handled by script fallback; new codes in skipped categories like Employee Taxes don't affect the JE)
- **MISS**: BLOCKING — new codes in `Earnings` or `Employer Ded Exp` that would be silently dropped

Exit 0 = safe to proceed. Exit 1 = new items need review.

### If the preflight surfaces new items

| Finding | Fix |
|---------|-----|
| New Cost Center | Add row to `Canada Payroll Department Mapping File.csv` mapping the cost center to a NetSuite department. Re-run preflight. |
| New Earnings code | Ask the accountant how to route the code (Salaries / Bonus / Commission / Severance / etc.). Add row to `Payroll_Mapping_with_GL_Accounts (final).csv` with both OpEx (611xxx) and COGS (511xxx) GL accounts. Re-run preflight. |
| New Employer Ded Exp code | Ask the accountant if it's a 401K-style match expense. Map to OpEx (611350) and COGS (511250). Add row to GL CSV. |
| New Earnings code that should be ignored (like TBRSP — voluntary RRSP that nets out via Employee Deds) | Add to `EARNINGS_INTENTIONAL_SKIPS` in `check_mappings.py`. |

## Step 3 — Run the main script

Only when preflight is clean:

```bash
cd scripts/canada-payroll
python process_canada_payroll.py "../../Monthly Payroll/Pay Runs/Canada/MM.DD.YYYY/G*.xls"
```

Script behavior:
- Auto-detects the raw file passed; reads pay date from folder name (`MM.DD.YYYY`)
- Iterates Sheet1 by Cost Center block; routes Earnings + ER Match + ER Tax to OpEx/COGS expense by department (with 70/30 split for `Engineering : Infrastructure`)
- Negative Earnings (retro pay reversals) become CR to Salaries (preserved sign)
- Reads Sheet2 Report Totals for Total Net (231200), RRSP (231250), and consolidated EE+ER taxes (231350)
- CLTD (Canadian Long-Term Disability) deduction: booked as CR to Other Benefits (recovers EE-paid premium)
- Aggregates expense lines by (account, department, memo)
- Writes:
  - `{MM.DD.YYYY} Canada Payroll Backup.xlsx` — `raw_Sheet1` + `raw_Sheet2` + `JE` tabs
  - `{MM.DD.YYYY} Canada Payroll JE Import.csv` — NetSuite Import Assistant format

## Step 4 — Handle imbalances

If the script reports `Balanced: False`, the CSV is **not written**. Common causes:

| Symptom | Likely cause |
|--------|--------------|
| Imbalance equals 2x a negative Earnings amount | Older script used `abs()` on Earnings amounts; fixed in the current version. If you re-run an old version, this is the regression. |
| Imbalance equals an Earnings code total | A code was added to Sheet2 totals but no expense row was generated (preflight should catch this). |
| Imbalance is small (< $50) | Likely rounding or a code in `EARNINGS_INTENTIONAL_SKIPS` that shouldn't be skipped. |

## Step 5 — Review the JE

Open `{MM.DD.YYYY} Canada Payroll Backup.xlsx` and inspect the `JE` tab. Standard layout:

DR side (per Cost Center, COGS/OpEx split):
| GL | Memo | Source |
|----|------|--------|
| 511100 / 611100 Salaries and Wages | `CAN PAYROLL - Salary` | CREG, CFLEX, CSICK, CCELB, CMPS, CRETR (retro) |
| 511150 / 611150 Bonus | `CAN PAYROLL - Bonus` | CBONU, CINCE |
| 611200 Commissions (OpEx only) | `CAN PAYROLL - Commission` | CCOMM |
| 511175 / 611250 Severance | `CAN PAYROLL - Severance` | CSEVR |
| 511250 / 611350 401k Match | `CAN PAYROLL - 401K` | CRRSP (Employer Ded Exp) |
| 511350 / 611450 Payroll Taxes | `CAN PAYROLL - Payroll Taxes` | Consolidated ER Tax total per dept |

DR/CR Other Benefits:
| GL | Side | Source |
|----|------|--------|
| 511300 / 611400 Other Benefits | CR (recovers EE premium) | CLTD |

CR side (single line each, no department):
| GL | Memo | Source |
|----|------|--------|
| 231200 Payroll Liability | `Canada Payroll Liability` | Sheet2 Total Net |
| 231250 401K payable | `Canada 401K Liability` | Sheet2 RRSP totals (CRRSP + CRRVF + CRSPV, both EE and ER) |
| 231350 Payroll Tax Liability | `Canada Payroll Tax Liability` | Sheet2 EE Tax + ER Tax totals |

Display a preview of the JE in chat (account, debit, credit, department, memo) with totals, then ask the accountant:

> **"Review the above carefully. Upload to NetSuite via Import Assistant? Type 'yes' to confirm or anything else to cancel."**

## Step 6 — Pre-upload gut-check review (auto)

Run `/gut-check` against the Backup.xlsx before telling the accountant to upload. It compares
the new JE against the prior two Canada-payroll JEs in NetSuite, flagging sign
flips, dept misroutes, and abnormal variances. Gate the upload instruction on the
gut-check result:
- PASS/WARN: tell the accountant it's safe to upload.
- FAIL: tell the accountant "DO NOT UPLOAD. Fix the CSV first; re-run /gut-check after edits."

## Step 7 — Hand off to the accountant + log

Tell the accountant:
> "CSV ready at `{MM.DD.YYYY} Canada Payroll JE Import.csv`. Upload via NetSuite UI:
> **Lists → Import Assistant → Import Type: Transactions → Record Type: Journal Entry**
> → upload the CSV → confirm field mapping → run import. The JE will land in the
> controller's Pending Approval queue."

Append to `audit_log.json` with `action: "GENERATE_CSV"`. After the accountant reports the JE
number post-import, append a follow-up entry with both `je_number` and
`netsuite_internal_id`.

**DO NOT** call `ns_createRecord` for journal entries — JE-style transactions go
through CSV upload only (per `_shared/approval-required.md`, effective 2026-05-01).

## Key differences from US payroll

- **Cadence:** identical (bi-weekly, 15th + last day of month)
- **Input format:** `.xls` (legacy Excel) with two sheets; US is `.xlsx` with one sheet
- **Currency:** CAD (US is USD)
- **RRSP vs 401k:** Canada has CRRSP (EE + ER match), CRRVF (vacation pay RRSP), CRSPV (spousal RRSP) — all CR to 231250. Multiple codes total a single liability.
- **Employer Taxes:** Consolidated single total per department (231350); not split by code (CPP/EI/etc.)
- **Subsidiary:** `Acme Holdings : Acme, Inc. : Acme Canada` (US is `Acme, Inc.`)
- **Revenue cost center fallback:** `Revenue` → `Sales & Marketing : Revenue`
- **Infrastructure split:** 70% OpEx / 30% COGS for `Engineering : Infrastructure` department
