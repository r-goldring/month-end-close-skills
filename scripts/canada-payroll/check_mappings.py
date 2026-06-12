"""
Canada Payroll - Pre-flight mapping check.

Scans the input ADP XLS file for all cost centers and payroll codes and compares
against the Canada department mapping and GL mapping CSVs. Reports any unmapped
items BEFORE the main script would silently skip them.

Usage:
  python check_mappings.py <input.xls> [dept_map.csv] [gl_map.csv]

Exit codes:
  0 = all required items mapped (safe to proceed with process_canada_payroll.py)
  1 = unmapped items found in mapping-required categories (review required)
"""

import os
import re
import sys
import pandas as pd

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(SCRIPT_DIR, "..", "_shared"))
from preflight_report import print_report  # noqa: E402


# Describes how process_canada_payroll.py actually uses each category.
# required=True  -> unmapped codes in this category get silently dropped
#                   (will cause JE to be under-stated). MUST be fixed.
# required=False -> category is intentionally skipped/consolidated;
#                   new codes here are informational only.
CATEGORY_INFO = {
    "Earnings": {"note": "every earnings code is mapped to a GL account",
                 "required": True},
    "Employee Deds": {"note": "only CLTD is extracted; other employee deduction codes are intentionally skipped",
                      "required": False},
    "Employee Taxes": {"note": "whole category is skipped by the main script",
                       "required": False},
    "Employer Ded Exp": {"note": "only 401KM-style codes are extracted",
                         "required": True},
    "Employer Tax Exp": {"note": "all codes summed into a single consolidated total per department",
                         "required": False},
}

# Cost centers that appear in the XLS but are handled by the script's
# fallback logic and don't need a CSV mapping entry.
COST_CENTER_SPECIAL_CASES = {"Revenue"}

# Earnings codes that appear in the raw file but are intentionally not posted.
# process_canada_payroll.py silently drops any unmapped Earnings code; this set
# makes that intent explicit in the preflight so they don't show up as MISS.
# If a code starts mattering later, remove it here and add a row to the GL
# mapping CSV.
EARNINGS_INTENTIONAL_SKIPS = {"TBRSP"}


# A valid ADP payroll code is a short all-caps alphanumeric token,
# typically 4-8 characters. This regex weeds out single-char artifacts
# like "Z" that appear in the raw cells.
CODE_PATTERN = re.compile(r"^[A-Z][A-Z0-9]{2,9}$")


def looks_like_code(cell_val: str) -> bool:
    """Heuristic: does this cell look like an ADP payroll code?"""
    if not cell_val or cell_val.lower() == "nan":
        return False
    if len(cell_val) < 3 or len(cell_val) > 10:
        return False
    if " " in cell_val:
        return False
    return bool(CODE_PATTERN.match(cell_val))


def scan_input(raw_xls: str):
    """Read the Canada XLS and return (cost_centers, codes_by_category)."""
    df1 = pd.read_excel(raw_xls, sheet_name="Sheet1", header=None)

    # Find the category header row
    row_8_idx = -1
    for r in range(5, 12):
        row_str = " ".join([str(x) for x in df1.iloc[r].values if pd.notna(x)])
        if "Earnings" in row_str and "Employee Deds" in row_str:
            row_8_idx = r
            break
    if row_8_idx == -1:
        raise RuntimeError("Could not find the category header row in Sheet1")

    # Build category column ranges
    row8 = df1.iloc[row_8_idx].values
    categories = ["Earnings", "Employee Deds", "Employee Taxes", "Employer Ded Exp", "Employer Tax Exp"]
    found_cats = []
    for c in range(len(row8)):
        val = str(row8[c]).strip()
        for cat in categories:
            if cat in val:
                found_cats.append((c, cat))
    found_cats.sort(key=lambda x: x[0])

    cat_ranges = {}
    for i in range(len(found_cats)):
        start_col = found_cats[i][0]
        end_col = found_cats[i + 1][0] - 1 if i < len(found_cats) - 1 else len(row8) - 1
        cat_ranges[found_cats[i][1]] = (start_col, end_col)

    cost_centers = set()
    codes_by_category = {cat: set() for cat in cat_ranges}

    current_cc = None
    in_block = False

    for row_idx in range(row_8_idx + 2, len(df1)):
        row_vals = df1.iloc[row_idx].values
        row_str = " ".join([str(v) for v in row_vals if pd.notna(v)])

        if "Group Summary for:" in row_str and "Cost Center:" in row_str:
            cc = row_str.split("Cost Center:")[1].strip() if len(row_str.split("Cost Center:")) > 1 else None
            if cc:
                cost_centers.add(cc)
                current_cc = cc
                in_block = True
            continue
        if "Group Totals:" in row_str:
            in_block = False
            current_cc = None
            continue

        if not in_block or not current_cc:
            continue

        for cat, (start_col, end_col) in cat_ranges.items():
            for c in range(start_col, min(end_col + 1, len(row_vals))):
                cell_val = str(row_vals[c]).strip() if pd.notna(row_vals[c]) else ""
                if looks_like_code(cell_val):
                    codes_by_category[cat].add(cell_val)

    return cost_centers, codes_by_category


def main():
    if len(sys.argv) < 2:
        print("Usage: python check_mappings.py <input.xls> [dept_map.csv] [gl_map.csv]")
        sys.exit(2)

    raw_xls = sys.argv[1]
    dept_map_path = sys.argv[2] if len(sys.argv) > 2 else os.path.join(SCRIPT_DIR, "Canada Payroll Department Mapping File.csv")
    gl_map_path = sys.argv[3] if len(sys.argv) > 3 else os.path.join(SCRIPT_DIR, "Payroll_Mapping_with_GL_Accounts (final).csv")

    pay_date_folder = os.path.basename(os.path.dirname(os.path.abspath(raw_xls)))

    dept_df = pd.read_csv(dept_map_path)
    gl_df = pd.read_csv(gl_map_path)

    mapped_cost_centers = set()
    for val in dept_df["Payroll Report Cost Center"].astype(str):
        cc = val.split("Cost Center: ")[1].strip() if "Cost Center: " in val else val.strip()
        mapped_cost_centers.add(cc)

    mapped_codes = set(gl_df["Payroll Code"].astype(str).str.strip().unique())

    cost_centers_found, codes_by_category = scan_input(raw_xls)

    cost_centers_mapped = cost_centers_found & mapped_cost_centers
    cost_centers_special = (cost_centers_found - mapped_cost_centers) & COST_CENTER_SPECIAL_CASES
    cost_centers_unmapped = cost_centers_found - mapped_cost_centers - COST_CENTER_SPECIAL_CASES

    codes_unmapped_by_category = {
        cat: codes - mapped_codes for cat, codes in codes_by_category.items()
    }
    # Strip intentional skips so they don't show as MISS in the report.
    if "Earnings" in codes_unmapped_by_category:
        codes_unmapped_by_category["Earnings"] -= EARNINGS_INTENTIONAL_SKIPS

    exit_code = print_report(
        country="Canada",
        pay_date_folder=pay_date_folder,
        input_file=os.path.basename(raw_xls),
        dept_map_file=os.path.basename(dept_map_path),
        gl_map_file=os.path.basename(gl_map_path),
        dept_map_entries=len(dept_df),
        gl_map_entries=len(gl_df),
        cost_centers_found=sorted(cost_centers_found),
        cost_centers_mapped=sorted(cost_centers_mapped),
        cost_centers_unmapped=sorted(cost_centers_unmapped),
        cost_centers_special_case=sorted(cost_centers_special),
        codes_by_category=codes_by_category,
        codes_mapped=mapped_codes,
        codes_unmapped_by_category=codes_unmapped_by_category,
        category_info=CATEGORY_INFO,
    )

    sys.exit(exit_code)


if __name__ == "__main__":
    main()
