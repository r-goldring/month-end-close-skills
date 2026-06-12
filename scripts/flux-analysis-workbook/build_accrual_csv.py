"""
build_accrual_csv.py

Generate the final `{YYYY-MM} Accruals JE Import.csv` from the accountant-approved
candidates in `accrual_candidates.json`.

Input: `Monthly Flux Analysis/{YYYY}/{YYYY-MM}/State/accrual_candidates.json`
       — only candidates with `approved: true` are included.

Output: `Monthly Flux Analysis/{YYYY}/{YYYY-MM}/JE Imports/{YYYY-MM} Accruals JE Import.csv`
       — NetSuite Import Assistant format with reversal date = 1st of next month.

CSV columns (matching production accrual CSV format from Mar/Apr 2026):
    External ID, Date, Journal Entry Memo, Currency, Account,
    Debit, Credit, Line Memo, Subsidiary, Department, Name, Reversal Date

For each approved vendor: emits one DR line (to the expense account) and one
CR line (to 231100 Accrued Liabilities).

Usage:
    python build_accrual_csv.py 2026-05
"""
from __future__ import annotations

import argparse
import calendar
import csv
import json
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "Monthly Flux Analysis"

DEFAULT_SUBSIDIARY = "Acme Holdings : Acme, Inc."
ACCRUED_LIABILITIES_ACCT = "231100"

CSV_HEADERS = [
    "External ID",
    "Date",
    "Journal Entry Memo",
    "Currency",
    "Account",
    "Debit",
    "Credit",
    "Line Memo",
    "Subsidiary",
    "Department",
    "Name",
    "Reversal Date",
]


def month_endpoints(year_month: str) -> tuple[str, str, str]:
    """
    Given 'YYYY-MM' return (date_for_je_str, reversal_date_str, external_id_prefix).
    Date is end-of-month; reversal is first-of-next-month.
    """
    year, month = (int(p) for p in year_month.split("-"))
    last_day = calendar.monthrange(year, month)[1]
    je_date = date(year, month, last_day)
    if month == 12:
        reversal = date(year + 1, 1, 1)
    else:
        reversal = date(year, month + 1, 1)

    return (
        je_date.strftime("%-m/%-d/%Y") if sys.platform != "win32" else je_date.strftime("%#m/%#d/%Y"),
        reversal.strftime("%-m/%-d/%Y") if sys.platform != "win32" else reversal.strftime("%#m/%#d/%Y"),
        f"{year_month}-ACCRUALS",
    )


def find_month_dir(year_month: str) -> Path:
    year = year_month[:4]
    return ROOT / year / year_month


def build_csv(year_month: str, dry_run: bool = False) -> dict:
    month_dir = find_month_dir(year_month)
    candidates_path = month_dir / "State" / "accrual_candidates.json"
    if not candidates_path.exists():
        raise FileNotFoundError(f"No accrual_candidates.json at {candidates_path}. "
                                f"Run the accrual builder first.")

    with open(candidates_path, "r", encoding="utf-8") as f:
        payload = json.load(f)

    approved = [c for c in payload.get("candidates", []) if c.get("approved")]
    if not approved:
        print(f"No approved candidates in {candidates_path}. Nothing to write.")
        return {"rows_written": 0, "approved_count": 0, "csv_path": None}

    je_date, reversal_date, external_id = month_endpoints(year_month)
    je_memo = f"{date.fromisoformat(year_month + '-01').strftime('%b-%y')} Accruals"

    rows: list[dict] = []
    for c in approved:
        amount = round(float(c["amount"]), 2)
        if amount <= 0:
            continue
        common = {
            "External ID": external_id,
            "Date": je_date,
            "Journal Entry Memo": je_memo,
            "Currency": "USD",
            "Line Memo": c.get("notes") or c.get("reason", "")[:200],
            "Subsidiary": c.get("subsidiary", DEFAULT_SUBSIDIARY),
            "Department": c.get("department", ""),
            "Name": c["vendor"],
            "Reversal Date": reversal_date,
        }
        # DR expense account
        rows.append({
            **common,
            "Account": c["account"],
            "Debit": f"{amount:.2f}",
            "Credit": "",
        })
        # CR accrued liabilities
        rows.append({
            **common,
            "Account": ACCRUED_LIABILITIES_ACCT,
            "Debit": "",
            "Credit": f"{amount:.2f}",
        })

    if not rows:
        return {"rows_written": 0, "approved_count": len(approved), "csv_path": None}

    out_dir = month_dir / "JE Imports"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{year_month} Accruals JE Import.csv"

    if dry_run:
        print(f"[DRY RUN] would write {len(rows)} rows to {out_path}")
        for r in rows[:6]:
            print(f"  {r}")
        return {"rows_written": len(rows), "approved_count": len(approved), "csv_path": str(out_path)}

    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_HEADERS)
        writer.writeheader()
        writer.writerows(rows)

    print(f"Wrote {len(rows)} rows ({len(approved)} approved vendors) -> {out_path}")
    return {"rows_written": len(rows), "approved_count": len(approved), "csv_path": str(out_path)}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("year_month", help="Month, e.g. 2026-05")
    parser.add_argument("--dry-run", action="store_true", help="Print without writing")
    args = parser.parse_args(argv)
    try:
        build_csv(args.year_month, dry_run=args.dry_run)
        return 0
    except FileNotFoundError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
