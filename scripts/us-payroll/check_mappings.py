"""
US Payroll - Pre-flight mapping check.

Scans the raw ADP XLSX/XLS file inside a pay-run folder for all cost centers
and payroll codes and compares against the US department and GL mapping CSVs.
Reports any unmapped items BEFORE the main script would tag them as
UNMAPPED_DEPT / UNMAPPED_ACCOUNT.

Usage:
  python check_mappings.py <pay-run folder | raw .xlsx/.xls>

Exit codes:
  0 = all required items mapped (safe to proceed with payroll_mapper.py)
  1 = unmapped items found in mapping-required categories (review required)

Note for US specifically: exit 1 is advisory, not a hard halt. the accountant sometimes
runs the mapper anyway and uses the unbalance to triangulate where a new
code belongs. The skill prompts for a decision on non-zero exit.
"""

import glob
import os
import re
import sys
import pandas as pd

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(SCRIPT_DIR, "..", "_shared"))
from preflight_report import print_report  # noqa: E402


# Hardcoded in payroll_mapper.py - these codes are known but intentionally ignored
CODES_TO_IGNORE = {"DPIID", "DPIIM", "DPIIV", "GTL", "GIFT"}


# How the US payroll_mapper.py treats each category
CATEGORY_INFO = {
    "Earnings": {"note": "every earnings code is mapped to a GL account",
                 "required": True},
    "Employee Deds": {"note": "every deduction code is mapped",
                      "required": True},
    "Employee Taxes": {"note": "not mapped per-code; 231350 is built from the ADP footer total (with NYPFLEE/NYSDIEE backout)",
                       "required": False},
    "Employer Ded Exp": {"note": "only 401KM is extracted; others silently dropped",
                         "required": True},
    "Employer Tax Exp": {"note": "not mapped per-code; all codes route uniformly to 611450/511350 by department",
                         "required": False},
}


# US payroll codes - same pattern as Canada
CODE_PATTERN = re.compile(r"^[A-Z][A-Z0-9]{2,9}$")


def looks_like_code(cell_val: str) -> bool:
    if not cell_val or cell_val.lower() == "nan":
        return False
    if len(cell_val) < 3 or len(cell_val) > 10:
        return False
    if " " in cell_val:
        return False
    return bool(CODE_PATTERN.match(cell_val))


def scan_input(raw_xlsx: str):
    """Read the US payroll XLSX and return (cost_centers, codes_by_category)."""
    xl = pd.ExcelFile(raw_xlsx)
    sheets = [s for s in xl.sheet_names if s.startswith("Sheet")]

    cost_centers = set()
    codes_by_category = {
        "Earnings": set(),
        "Employee Deds": set(),
        "Employee Taxes": set(),
        "Employer Ded Exp": set(),
        "Employer Tax Exp": set(),
    }

    for sheet in sheets:
        df = pd.read_excel(xl, sheet_name=sheet, header=None)

        # Find the header row (contains "Employee Deds")
        header_row_idx = -1
        for idx, row in df.head(50).iterrows():
            row_str = " ".join(str(x) for x in row.values if pd.notna(x))
            if "Employee Deds" in row_str and "Code" not in row_str:
                header_row_idx = idx
                break
        if header_row_idx == -1:
            continue

        categories = df.iloc[header_row_idx].ffill().fillna("").astype(str).str.strip().tolist()
        col_names = df.iloc[header_row_idx + 1].fillna("").astype(str).str.strip().tolist()

        # Find the "Code" column index for each category
        code_col_by_category = {}
        for i, cat in enumerate(categories):
            if col_names[i] == "Code" and cat in codes_by_category:
                code_col_by_category[cat] = i

        # Walk rows
        for idx in range(header_row_idx + 2, len(df)):
            row = df.iloc[idx]

            val0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            if val0.startswith("Cost Center:"):
                cc = val0.replace("Cost Center: ", "").strip()
                if cc:
                    cost_centers.add(cc)
                continue

            for cat, code_idx in code_col_by_category.items():
                if code_idx >= len(row):
                    continue
                cell_val = str(row.iloc[code_idx]).strip() if pd.notna(row.iloc[code_idx]) else ""
                if not looks_like_code(cell_val):
                    continue
                if cell_val in CODES_TO_IGNORE:
                    continue
                codes_by_category[cat].add(cell_val)

    return cost_centers, codes_by_category


def find_raw_file(target: str) -> str:
    """Resolve `target` to an input ADP payroll file path.

    Accepts either a pay-run folder (auto-discovers the file inside) or
    the file path directly. Discovery priority matches payroll_mapper.find_raw_file:
    Backup.xls (controller drop, includes Reclasses) > Backup.xlsx (re-run) >
    other ADP-named file (legacy G* fallback).
    """
    if os.path.isfile(target):
        return target
    if not os.path.isdir(target):
        raise FileNotFoundError(f"Not a file or folder: {target!r}")

    backup_candidates = []
    other_candidates = []
    for ext in ("*.xlsx", "*.xls"):
        for f in glob.glob(os.path.join(target, ext)):
            base = os.path.basename(f)
            if base.startswith("~"):
                continue
            if "Review_Report" in base:
                continue
            if "Desired Output" in base:
                continue
            if "(verify)" in base:
                continue
            if "Backup" in base:
                backup_candidates.append(f)
            else:
                other_candidates.append(f)

    if backup_candidates:
        xls_first = sorted(backup_candidates, key=lambda p: 0 if p.lower().endswith(".xls") else 1)
        if len(xls_first) > 1 and not xls_first[0].lower().endswith(".xls"):
            raise RuntimeError(
                f"Multiple Backup files in {target}: {xls_first}. "
                "Leave only one Backup file in the folder."
            )
        return xls_first[0]

    if not other_candidates:
        raise FileNotFoundError(f"No input file found in {target}")
    if len(other_candidates) > 1:
        raise RuntimeError(
            f"Multiple candidate input files in {target}: {other_candidates}. "
            "Leave only one ADP file in the folder."
        )
    return other_candidates[0]


def main():
    if len(sys.argv) < 2:
        print("Usage: python check_mappings.py <pay-run folder | raw .xlsx/.xls>")
        sys.exit(2)

    raw_xlsx = find_raw_file(sys.argv[1])
    dept_map_path = os.path.join(SCRIPT_DIR, "US Payroll Department Mapping File.csv")
    gl_map_path = os.path.join(SCRIPT_DIR, "Payroll_Mapping_with_GL_Accounts (final).csv")

    pay_date_folder = os.path.basename(os.path.dirname(os.path.abspath(raw_xlsx)))

    dept_df = pd.read_csv(dept_map_path)
    gl_df = pd.read_csv(gl_map_path)

    mapped_cost_centers = set()
    for val in dept_df["Payroll Report Cost Center"].astype(str):
        cc = val.replace("Cost Center: ", "").strip()
        mapped_cost_centers.add(cc)

    mapped_codes = set(gl_df["Payroll Code"].astype(str).str.strip().unique())

    cost_centers_found, codes_by_category = scan_input(raw_xlsx)

    cost_centers_mapped = cost_centers_found & mapped_cost_centers
    cost_centers_unmapped = cost_centers_found - mapped_cost_centers

    codes_unmapped_by_category = {
        cat: codes - mapped_codes for cat, codes in codes_by_category.items()
    }

    exit_code = print_report(
        country="US",
        pay_date_folder=pay_date_folder,
        input_file=os.path.basename(raw_xlsx),
        dept_map_file=os.path.basename(dept_map_path),
        gl_map_file=os.path.basename(gl_map_path),
        dept_map_entries=len(dept_df),
        gl_map_entries=len(gl_df),
        cost_centers_found=sorted(cost_centers_found),
        cost_centers_mapped=sorted(cost_centers_mapped),
        cost_centers_unmapped=sorted(cost_centers_unmapped),
        cost_centers_special_case=[],
        codes_by_category=codes_by_category,
        codes_mapped=mapped_codes,
        codes_unmapped_by_category=codes_unmapped_by_category,
        category_info=CATEGORY_INFO,
    )

    sys.exit(exit_code)


if __name__ == "__main__":
    main()
