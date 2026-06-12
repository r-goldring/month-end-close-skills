"""
April 2026 Software Reclass JE Import CSV.

Per the accountant's spec:
  - Name column (vendor in NS format) on every line
  - Same line memo on both lines: "Reclass {Vendor} from {From} to {To}"
  - No "Apr-26" prefix in memo
  - No the accountant name references
  - Full NS department paths in Department column
  - No Reversal Date (reclasses are permanent)

14 reclasses across Apr 2026 software activity:
  Account reclasses (671100 OpEx Software -> 511425 COGS Software):
    - SSO Provider, iPaaS Vendor A
  Pure dept reclasses (671100 stays):
    - Lucid, Pagerduty, GitLab Federal, Docusign, CDW, Sales Engagement SaaS, Adobe,
      ChatGPT, Freshworks, MONDAY.COM, Webflow, Google
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
SUBSIDIARY = "Acme Holdings : Acme, Inc."
JE_HEADER_MEMO = "Apr-26 Software Reclass"

# NS department full paths
DEPT = {
    "Infrastructure": "Engineering : Infrastructure",
    "Engineering": "Research & Development : Technology : Engineering",
    "Info Security Privacy": "Research & Development : Technology : Info Security Privacy",
    "Technology": "Research & Development : Technology",
    "Customer Success": "Sales & Marketing : Customer Success",
    "Revenue": "Sales & Marketing : Revenue",
    "Business Operations": "Sales & Marketing : Revenue : Business Operations",
    "Marketing": "Sales & Marketing : Marketing",
    "Sales Development": "Sales & Marketing : Marketing : Sales Development",
    "Information Technology": "General & Administrative : Information Technology",
    "People": "General & Administrative : People",
    "Professional Services": "COGS : Professional Services",
}
SW_OPEX = "671100 Software Subscriptions"
SW_COGS = "511425 COGS - Software Subscriptions"

# Each reclass entry is a list of (vendor_name, account, dept_short, debit, credit, line_memo)
# Same line memo applied to all lines of a reclass block.
RECLASSES = [
    # 1. SSO Provider - Apr OpEx Software (Infra $X,XXX.XX + Eng $X,XXX.XX) -> COGS Software / Engineering
    {
        "vendor": "SSO Provider",
        "memo": "Reclass SSO Provider from OpEx Software to COGS Software",
        "lines": [
            (SW_COGS, "Engineering", 17010.36, ""),
            (SW_OPEX, "Infrastructure", "", 11340.37),
            (SW_OPEX, "Engineering", "", 5669.99),
        ],
    },
    # 2. iPaaS Vendor A - Apr OpEx Software / Engineering -> COGS Software / Engineering
    {
        "vendor": "iPaaS Vendor A",
        "memo": "Reclass iPaaS Vendor A from OpEx Software to COGS Software",
        "lines": [
            (SW_COGS, "Engineering", 6366.35, ""),
            (SW_OPEX, "Engineering", "", 6366.35),
        ],
    },
    # 3. Lucid Software - 671100 / Info Sec Privacy -> 671100 / Engineering
    {
        "vendor": "Lucid Software Inc.",
        "memo": "Reclass Lucid Software from Info Security Privacy to Engineering",
        "lines": [
            (SW_OPEX, "Engineering", 1664.57, ""),
            (SW_OPEX, "Info Security Privacy", "", 1664.57),
        ],
    },
    # 4. Pagerduty - 671100 / Infrastructure -> 671100 / Engineering
    {
        "vendor": "Pagerduty, Inc.",
        "memo": "Reclass Pagerduty from Infrastructure to Engineering",
        "lines": [
            (SW_OPEX, "Engineering", 5040.00, ""),
            (SW_OPEX, "Infrastructure", "", 5040.00),
        ],
    },
    # 5. GitLab Federal - 671100 / Infrastructure -> 671100 / Engineering
    {
        "vendor": "GitLab Federal, LLC",
        "memo": "Reclass GitLab Federal from Infrastructure to Engineering",
        "lines": [
            (SW_OPEX, "Engineering", 4154.82, ""),
            (SW_OPEX, "Infrastructure", "", 4154.82),
        ],
    },
    # 6. Docusign - 671100 / Business Operations -> 671100 / Revenue
    {
        "vendor": "Docusign",
        "memo": "Reclass Docusign from Business Operations to Revenue",
        "lines": [
            (SW_OPEX, "Revenue", 3977.91, ""),
            (SW_OPEX, "Business Operations", "", 3977.91),
        ],
    },
    # 7. CDW - 671100 / Infrastructure -> 671100 / Information Technology
    {
        "vendor": "CDW",
        "memo": "Reclass CDW from Infrastructure to Information Technology",
        "lines": [
            (SW_OPEX, "Information Technology", 3608.20, ""),
            (SW_OPEX, "Infrastructure", "", 3608.20),
        ],
    },
    # 8. Sales Engagement SaaS - 671100 / Revenue -> 671100 / Sales Development
    {
        "vendor": "Sales Engagement SaaS",
        "memo": "Reclass Sales Engagement SaaS from Revenue to Sales Development",
        "lines": [
            (SW_OPEX, "Sales Development", 3600.00, ""),
            (SW_OPEX, "Revenue", "", 3600.00),
        ],
    },
    # 9. Adobe - 671100 / Professional Services -> 671100 / Marketing
    {
        "vendor": "ADOBE",
        "memo": "Reclass Adobe from Professional Services to Marketing",
        "lines": [
            (SW_OPEX, "Marketing", 3422.20, ""),
            (SW_OPEX, "Professional Services", "", 3422.20),
        ],
    },
    # 10. ChatGPT - 671100 / Information Technology -> 671100 / Infrastructure
    {
        "vendor": "Chatgpt",
        "memo": "Reclass ChatGPT from Information Technology to Infrastructure",
        "lines": [
            (SW_OPEX, "Infrastructure", 1336.55, ""),
            (SW_OPEX, "Information Technology", "", 1336.55),
        ],
    },
    # 11. Freshworks - 671100 / Professional Services -> 671100 / Customer Success
    {
        "vendor": "Freshworks Inc",
        "memo": "Reclass Freshworks from Professional Services to Customer Success",
        "lines": [
            (SW_OPEX, "Customer Success", 464.60, ""),
            (SW_OPEX, "Professional Services", "", 464.60),
        ],
    },
    # 12. MONDAY.COM - 671100 / People -> 671100 / Marketing
    {
        "vendor": "MONDAY.COM",
        "memo": "Reclass Monday.com from People to Marketing",
        "lines": [
            (SW_OPEX, "Marketing", 139.61, ""),
            (SW_OPEX, "People", "", 139.61),
        ],
    },
    # 13. Webflow - 671100 / Infrastructure -> 671100 / Technology (R&D : Technology parent)
    # Mar reclass JE##### confirmed destination is the Technology parent dept (not a leaf).
    {
        "vendor": "Webflow",
        "memo": "Reclass Webflow from Infrastructure to Technology",
        "lines": [
            (SW_OPEX, "Technology", 102.00, ""),
            (SW_OPEX, "Infrastructure", "", 102.00),
        ],
    },
    # 14. Google - small drift from two depts back to IT
    {
        "vendor": "Google",
        "memo": "Reclass Google from Engineering and Business Operations to Information Technology",
        "lines": [
            (SW_OPEX, "Information Technology", 109.40, ""),
            (SW_OPEX, "Engineering", "", 9.40),
            (SW_OPEX, "Business Operations", "", 100.00),
        ],
    },
]

EXTRA_COLS = [("Name", "Name")]


def main():
    repo = SCRIPT_DIR.parent.parent
    out_path = repo / "Monthly Flux Analysis" / "2026" / PERIOD / f"{PERIOD} Software Reclass JE Import.csv"

    rows = []
    total_debits = 0.0
    total_credits = 0.0
    for block in RECLASSES:
        vendor = block["vendor"].encode("ascii", errors="replace").decode("ascii").replace("?", "")
        memo = block["memo"].encode("ascii", errors="replace").decode("ascii").replace("?", "")
        for account, dept_short, debit, credit in block["lines"]:
            rows.append({
                "Date": PERIOD_DATE,
                "Journal Entry Memo": JE_HEADER_MEMO,
                "Account": account,
                "Debit": debit,
                "Credit": credit,
                "Line Memo": memo,
                "Subsidiary": SUBSIDIARY,
                "Department": DEPT[dept_short],
                "Name": vendor,
            })
            if isinstance(debit, (int, float)):
                total_debits += debit
            if isinstance(credit, (int, float)):
                total_credits += credit

    write_je_csv(rows, out_path, currency="USD",
                 default_external_id=f"{PERIOD}-SW-RECLASS",
                 extra_columns=EXTRA_COLS)

    print(f"Wrote: {out_path.relative_to(repo)}")
    print(f"  Reclasses: {len(RECLASSES)}")
    print(f"  Total CSV rows: {len(rows)}")
    print(f"  Debits:  ${total_debits:,.2f}")
    print(f"  Credits: ${total_credits:,.2f}")
    if abs(total_debits - total_credits) > 0.01:
        print(f"  WARNING: JE not balanced! diff = ${total_debits - total_credits:,.2f}")
    else:
        print(f"  Balanced.")


if __name__ == "__main__":
    main()
