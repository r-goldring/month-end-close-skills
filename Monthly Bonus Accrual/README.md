# Monthly Bonus Accrual

Source files from FP&A (the FP&A Lead) for the monthly bonus accrual JE.

## How to use

1. Each month, FP&A sends a workbook named like
   `2026-05_Bonus Accrual_vAccounting_vF.xlsx`.
2. Drop it into a new folder here:
   `Monthly Bonus Accrual/{YYYY-MM} {Month Name} {Year}/`
   (e.g. `Monthly Bonus Accrual/2026-05 May 2026/`).
3. Tell Claude to run the bonus accrual skill. The skill auto-detects the
   workbook and writes the upload-ready CSV alongside it as
   `{YYYY-MM} Bonus Accrual JE Import.csv`.
4. Upload the CSV via NetSuite UI: Setup -> Import/Export -> Import CSV
   Records -> Journal Entries. Each of the 5 External IDs
   (`USB{YYMM}`, `CANB{YYMM}`, `NLB{YYMM}`, `UKB{YYMM}`, `URYB{YYMM}`)
   becomes a separate JE in the controller's Pending Approval queue.

## Skill

- Skill instructions: `.claude/skills/bonus-accrual/SKILL.md`
- Build script: `scripts/bonus-accrual/build_bonus_je.py`

## Git

The month subfolders here are gitignored (raw FP&A data + generated CSVs).
This README is the only tracked file in this folder.
