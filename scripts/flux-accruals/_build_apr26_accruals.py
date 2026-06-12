"""
April 2026 accrual JE Import CSV — mirrors the Mar-26 accrual (JE#####)
formatting exactly:
  - Line memo style mirrors Mar lines (e.g., "April estimate based on Mar Activity")
  - Name column populated on BOTH debit and credit lines
  - Department populated on BOTH lines (full NS path)
  - Reversal Date 5/1/2026 on every line
  - Journal Entry Memo: "Apr-26 Accruals"

Output: Monthly Flux Analysis/2026/2026-04/2026-04 Accruals JE Import.csv
"""

import datetime as dt
import os
import sys
from pathlib import Path

SCRIPT_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, str(SCRIPT_DIR.parent / "_shared"))

from je_csv_writer import write_je_csv  # noqa: E402

PERIOD = "2026-04"
PERIOD_DATE = dt.date(2026, 4, 30)
REVERSAL_DATE = dt.date(2026, 5, 1)
SUBSIDIARY = "Acme Holdings : Acme, Inc."
ACCRUED_LIAB = "231100 Accrued Liabilities"
JE_HEADER_MEMO = "Apr-26 Accruals"

# Full NS department paths (verified via SuiteQL on 2026-05-08)
DEPT = {
    "Infrastructure": "Engineering : Infrastructure",
    "Engineering": "Research & Development : Technology : Engineering",
    "Data Science Analytics": "Research & Development : Technology : Data Science Analytics",
    "Info Security Privacy": "Research & Development : Technology : Info Security Privacy",
    "Customer Success": "Sales & Marketing : Customer Success",
    "Revenue": "Sales & Marketing : Revenue",
    "Business Operations": "Sales & Marketing : Revenue : Business Operations",
    "Marketing": "Sales & Marketing : Marketing",
    "Sales Development": "Sales & Marketing : Marketing : Sales Development",
    "Information Technology": "General & Administrative : Information Technology",
    "GA": "General & Administrative : GA",
    "People": "General & Administrative : People",
    "Professional Services": "COGS : Professional Services",
}

# Each accrual: (name, expense_account, dept_short, amount, line_memo)
# Line memos mirror Mar-26 (JE#####) phrasing style:
#   - "April estimate based on Mar Activity"  (IT Hardware Partner-like)
#   - "Sept-Apr accrual (Renewal started Sept)"  (DNS Provider A cumulative)
#   - "Aug-Apr accrual of healthcare benchmarking billed in arrears ($110k annually)"
#   - "April estimate based on recent activity"  (DPX contractor-like)
#   - "April accrual estimate."  (External Auditor-like)
#   - "Invoice {N} not yet received."  (401k Advisory Firm-like)
ACCRUALS = [
    # COGS
    ("IT Hardware Partner Networked Solutions Group, LLC", "511400 COGS - Hosting", "Infrastructure", 377090.88,
     "April estimate based on Mar Activity"),
    ("DNS Provider A", "511400 COGS - Hosting", "Infrastructure", 17332.00,
     "Sept-Apr accrual (Renewal started Sept)"),
    ("Epiphany Group LLC", "511600 COGS - Other", "Professional Services", 82500.00,
     "Aug-Apr accrual of healthcare benchmarking billed in arrears ($110k annually)"),

    # Contractors
    ("DataPlatformExperts LLC", "511370 COGS - Contractor Payroll", "Infrastructure", 2064.60,
     "April estimate based on recent activity"),
    ("DataPlatformExperts LLC", "611700 Contractor Payroll - OpEx", "Infrastructure", 20327.40,
     "April estimate based on recent activity"),
    ("Szymon Nieznanski", "611700 Contractor Payroll - OpEx", "Infrastructure", 8472.00,
     "April contractor accrual (AP confirmed)"),
    ("Paul Stiniguta", "511370 COGS - Contractor Payroll", "Infrastructure", 390.00,
     "April contractor accrual (50% COGS split per historical pattern)"),
    ("Paul Stiniguta", "611700 Contractor Payroll - OpEx", "Infrastructure", 390.00,
     "April contractor accrual (50% OpEx split per historical pattern)"),
    ("Denis Shaposhnikov", "611700 Contractor Payroll - OpEx", "Engineering", 7500.00,
     "April contractor accrual (AP confirmed)"),

    # Professional Fees
    ("401k Advisory Firm, LLC", "651100 Professional Fees", "GA", 4273.00,
     "April estimate based on Mar Activity; bill not yet posted to 651100"),
    ("External Auditor U.K. LLP", "651100 Professional Fees", "GA", 5000.00,
     "April accrual estimate."),

    # Software (recurring SaaS where Apr bill not yet hit)
    ("WalkMe", "671100 Software Subscriptions", "Customer Success", 2098.38,
     "April accrual for FY26 renewal"),
    ("Twilio Inc.", "671100 Software Subscriptions", "Engineering", 5903.91,
     "April estimate based on Mar Activity; HWNADG monthly invoice not yet posted"),
    ("Gong.io Inc.", "671100 Software Subscriptions", "Revenue", 5401.02,
     "April estimate based on Mar Activity"),

    # Investigation finding: Leniolabs Team 2 invoice missing in April
    ("Leniolabs LLC", "611700 Contractor Payroll - OpEx", "Data Science Analytics", 4800.00,
     "April estimate; Team 2 invoice not yet received"),
]

EXTRA_COLS = [
    ("Name", "Name"),
    ("Reversal Date", "Reversal Date"),
]


def main():
    repo = SCRIPT_DIR.parent.parent
    out_path = repo / "Monthly Flux Analysis" / "2026" / PERIOD / f"{PERIOD} Accruals JE Import.csv"

    rows = []
    grand = 0.0
    for name, account, dept_short, amount, line_memo in ACCRUALS:
        line_memo_a = line_memo.encode("ascii", errors="replace").decode("ascii").replace("?", "")
        name_a = name.encode("ascii", errors="replace").decode("ascii").replace("?", "")
        dept_full = DEPT[dept_short]

        # Both lines get Name + full Department (matches JE##### / Mar-26 pattern).
        common = {
            "Date": PERIOD_DATE,
            "Journal Entry Memo": JE_HEADER_MEMO,
            "Subsidiary": SUBSIDIARY,
            "Line Memo": line_memo_a,
            "Name": name_a,
            "Department": dept_full,
            "Reversal Date": REVERSAL_DATE,
        }
        # Debit expense
        rows.append({**common, "Account": account, "Debit": amount, "Credit": ""})
        # Credit Accrued Liabilities (same Name + Dept on this line)
        rows.append({**common, "Account": ACCRUED_LIAB, "Debit": "", "Credit": amount})
        grand += amount

    write_je_csv(rows, out_path, currency="USD",
                 default_external_id=f"{PERIOD}-ACCRUALS",
                 extra_columns=EXTRA_COLS)

    print(f"Wrote: {out_path.relative_to(repo)}")
    print(f"  Lines: {len(ACCRUALS)} accruals (= {len(rows)} CSV rows incl. offsets)")
    print(f"  Total debits: ${grand:,.2f}  (= total credits to 231100)")
    print(f"  Reversal Date: {REVERSAL_DATE.month}/{REVERSAL_DATE.day}/{REVERSAL_DATE.year} on every line")
    print(f"  Name + Department populated on BOTH lines per JE##### pattern")


if __name__ == "__main__":
    main()
