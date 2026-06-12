"""
Poland Payroll -> NetSuite JE generator.

Reads the monthly raw spreadsheet from the Polish payroll provider (single sheet
'Acme Corp {MM} {YYYY}', English headers on row 1, Polish headers on row 2, one
row per employee). Maps each earning/benefit column to a GL account, routes
COGS vs OpEx based on each employee's department (from Poland Payroll Employee
Mapping.csv), extracts liability totals from the raw file's summary table, and
writes an Excel workbook with a 'raw' + 'JE' tab matching the accountant's historical
backup format.

Severance (Odprawa bez ZUS, column Z) is auto-detected and routes to 511175 +
EBITDA Adjustments department, regardless of the employee's normal department.

Unknown employees or unknown columns halt the script with a clear flag.

Usage:
    python process_poland_payroll.py <input.xlsx> [output.xlsx] [emp_map.csv] [gl_map.csv]
"""

import os
import sys
import calendar
import datetime as dt
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(SCRIPT_DIR, "..", "_shared"))
from je_csv_writer import write_je_csv, make_external_id  # noqa: E402

SUBSIDIARY = "Acme Holdings : Acme, Inc. : Acme Poland"
LIAB_NET_ACCT = "231200 Payroll Liability"
LIAB_TAX_ACCT = "231350 Payroll Tax Liability"
LIAB_PPK_ACCT = "231250 401K payable"
SEVERANCE_DEPT = "EBITDA Adjustments"

SUMMARY_LABEL_NET = "NET SALARY for employees"
SUMMARY_LABEL_SOCIAL = "EE & ER Social Insurance & Health Insurance"
SUMMARY_LABEL_TAX = "TAX"
SUMMARY_LABEL_PPK = "PPK"


def _to_float(v):
    if pd.isna(v):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("$", "").replace(" ", "")
    if not s:
        return 0.0
    if s.startswith("(") and s.endswith(")"):
        return -float(s[1:-1])
    try:
        return float(s)
    except ValueError:
        return 0.0


def _last_day_of_month(d: dt.date) -> dt.date:
    return d.replace(day=calendar.monthrange(d.year, d.month)[1])


def _find_header_row(df):
    for r in range(min(10, len(df))):
        vals = [str(x).strip() for x in df.iloc[r].values if pd.notna(x)]
        if "Name & Surname" in vals and "Code EE" in vals:
            return r
    raise ValueError("Could not locate header row (needs 'Name & Surname' + 'Code EE').")


def _load_employee_map(csv_path):
    df = pd.read_csv(csv_path)
    return {str(r["Employee Name"]).strip(): str(r["Department"]).strip() for _, r in df.iterrows()}


def _load_gl_map(csv_path):
    df = pd.read_csv(csv_path).fillna("")
    m = {}
    for _, r in df.iterrows():
        col = str(r["Raw Column"]).strip()
        m[col] = {
            "role": str(r["Role"]).strip(),
            "cogs": str(r["GL Account - COGS"]).strip(),
            "opex": str(r["GL Account - OpEx"]).strip(),
            "memo": str(r["Line Memo"]).strip(),
        }
    return m


def _is_cogs(department: str) -> bool:
    return department.strip().startswith("COGS :") or department.strip().startswith("COGS:")


def _find_summary_value(raw, header_row, label):
    """Scan rows below employee data for a summary row whose column F (index 5) matches the label.
    Returns the numeric value in column J (index 9), or None if not found."""
    for r in range(header_row + 1, len(raw)):
        row = raw.iloc[r]
        if len(row) <= 9:
            continue
        cell = row[5] if len(row) > 5 else None
        if pd.notna(cell) and str(cell).strip() == label:
            return _to_float(row[9])
    return None


def generate_je(raw_xlsx, emp_map_csv, gl_map_csv, output_xlsx):
    emp_map = _load_employee_map(emp_map_csv)
    gl_map = _load_gl_map(gl_map_csv)

    raw = pd.read_excel(raw_xlsx, sheet_name=0, header=None)

    header_row = _find_header_row(raw)
    headers = [str(x).strip() if pd.notna(x) else "" for x in raw.iloc[header_row].values]
    col_idx = {h: i for i, h in enumerate(headers) if h}

    unknown_cols = [h for h in headers if h and h not in gl_map]
    if unknown_cols:
        raise RuntimeError(
            "Unknown columns in raw file (not in Poland Payroll GL Mapping.csv): "
            + ", ".join(unknown_cols)
            + "\nAdd rows for these columns to the GL mapping and re-run."
        )

    employee_rows = []
    for r in range(header_row + 2, len(raw)):
        row = raw.iloc[r]
        code = row[col_idx["Code EE"]] if "Code EE" in col_idx else None
        name = row[col_idx["Name & Surname"]] if "Name & Surname" in col_idx else None
        if pd.isna(code) or pd.isna(name) or not str(name).strip():
            break
        if str(code).strip().upper() == "TOTAL":
            break
        employee_rows.append(r)

    if not employee_rows:
        raise RuntimeError("No employee rows found below header row.")

    payroll_date_val = raw.iloc[employee_rows[0]][col_idx["Payroll Date"]]
    if isinstance(payroll_date_val, dt.datetime):
        payroll_date = payroll_date_val.date()
    else:
        payroll_date = pd.to_datetime(str(payroll_date_val)).date()
    je_date = _last_day_of_month(payroll_date)
    ym_prefix = je_date.strftime("%Y-%m")
    je_memo = f"{ym_prefix} Poland Payroll"

    unknown_emps = []
    for r in employee_rows:
        name = str(raw.iloc[r][col_idx["Name & Surname"]]).strip()
        if name not in emp_map:
            unknown_emps.append(name)
    if unknown_emps:
        raise RuntimeError(
            "Unknown employees (not in Poland Payroll Employee Mapping.csv): "
            + ", ".join(unknown_emps)
            + "\nAdd rows for each employee with their Department and re-run."
        )

    lines = []

    def add(acct, dept, memo, debit=0.0, credit=0.0):
        if debit == 0.0 and credit == 0.0:
            return
        lines.append(
            {
                "Date": je_date.strftime("%Y-%m-%d"),
                "Journal Entry Memo": je_memo,
                "Account": acct,
                "Debit": round(debit, 2) if debit else "",
                "Credit": round(credit, 2) if credit else "",
                "Line Memo": memo,
                "Subsidiary": SUBSIDIARY,
                "Department": dept,
            }
        )

    # First pass: emit all fixed-expense rows. Track per-employee Variable totals
    # so we can decide after summary parsing whether Y / Z contributed to this
    # month's cash flow (they're sometimes informational, sometimes cash).
    variable_expense_lines = []     # (emp_name, col_name, mapping, amt)
    variable_severance_lines = []   # (emp_name, col_name, mapping, amt)

    for r in employee_rows:
        row = raw.iloc[r]
        name = str(row[col_idx["Name & Surname"]]).strip()
        dept = emp_map[name]
        is_cogs = _is_cogs(dept)

        for col_name, mapping in gl_map.items():
            role = mapping["role"]
            if col_name not in col_idx:
                continue
            amt = _to_float(row[col_idx[col_name]])
            if amt == 0.0:
                continue

            if role == "Expense":
                acct = mapping["cogs"] if is_cogs else mapping["opex"]
                if not acct:
                    acct = mapping["opex"] or mapping["cogs"]
                if not acct:
                    raise RuntimeError(
                        f"No GL account configured for column '{col_name}' on {'COGS' if is_cogs else 'OpEx'} side "
                        f"(employee {name}, dept {dept})."
                    )
                add(acct, dept, f"{ym_prefix} {mapping['memo']}", debit=amt)
            elif role == "VariableExpense":
                variable_expense_lines.append((name, dept, col_name, mapping, amt))
            elif role == "VariableSeverance":
                variable_severance_lines.append((name, dept, col_name, mapping, amt))
            elif role == "IgnoreIfZero":
                print(f"  [note] Ignoring '{col_name}' amount {amt:.2f} for {name} (per current policy).")

    net_total = _find_summary_value(raw, header_row, SUMMARY_LABEL_NET)
    social_total = _find_summary_value(raw, header_row, SUMMARY_LABEL_SOCIAL)
    tax_total = _find_summary_value(raw, header_row, SUMMARY_LABEL_TAX)
    ppk_total = _find_summary_value(raw, header_row, SUMMARY_LABEL_PPK)

    missing = []
    if net_total is None: missing.append(SUMMARY_LABEL_NET)
    if social_total is None: missing.append(SUMMARY_LABEL_SOCIAL)
    if tax_total is None: missing.append(SUMMARY_LABEL_TAX)
    if ppk_total is None: missing.append(SUMMARY_LABEL_PPK)
    if missing:
        raise RuntimeError(
            "Could not locate summary rows in raw file: "
            + ", ".join(missing)
            + "\nExpected each label in column F with the amount in column J."
        )

    tax_liability_total = (social_total or 0.0) + (tax_total or 0.0)
    cr_total = net_total + tax_liability_total + ppk_total

    fixed_dr_total = sum(float(l["Debit"]) for l in lines if l["Debit"] != "")
    y_total = sum(amt for (_, _, _, _, amt) in variable_expense_lines)
    z_total = sum(amt for (_, _, _, _, amt) in variable_severance_lines)
    residual = round(cr_total - fixed_dr_total, 2)

    include_y = False
    include_z = False
    if abs(residual) < 0.01:
        pass
    elif abs(residual - y_total) < 0.01:
        include_y = True
    elif abs(residual - z_total) < 0.01:
        include_z = True
    elif abs(residual - y_total - z_total) < 0.01:
        include_y = True
        include_z = True
    else:
        raise RuntimeError(
            f"Cannot reconcile DR to CR. Fixed DR={fixed_dr_total:.2f}, CR={cr_total:.2f}, "
            f"residual={residual:.2f}, Y (vacation payout)={y_total:.2f}, Z (severance)={z_total:.2f}. "
            "Residual did not match Y, Z, or Y+Z - a non-standard column may be included in the provider summary."
        )

    if include_y:
        print(f"  [info] Vacation payout (Y) is cash this month: {y_total:.2f} PLN -> Bonus bucket.")
        for name, dept, col_name, mapping, amt in variable_expense_lines:
            is_cogs = _is_cogs(dept)
            acct = (mapping["cogs"] if is_cogs else mapping["opex"]) or mapping["opex"] or mapping["cogs"]
            add(acct, dept, f"{ym_prefix} {mapping['memo']}", debit=amt)
    elif y_total > 0:
        print(f"  [info] Vacation payout (Y) is informational only this month: {y_total:.2f} PLN (not expensed).")

    if include_z:
        print(f"  [info] Severance (Z) is cash this month: {z_total:.2f} PLN -> EBITDA Adjustments.")
        for name, dept, col_name, mapping, amt in variable_severance_lines:
            acct = mapping["cogs"] or mapping["opex"]
            add(acct, SEVERANCE_DEPT, f"{ym_prefix} {mapping['memo']} ({name})", debit=amt)
    elif z_total > 0:
        print(f"  [info] Severance (Z) NOT in provider summary this month: {z_total:.2f} PLN (paid via separate mechanism - not in payroll JE).")

    add(LIAB_NET_ACCT, "", f"{ym_prefix} Poland Payroll Liability", credit=net_total)
    add(LIAB_TAX_ACCT, "", f"{ym_prefix} Poland Payroll Tax Liability", credit=tax_liability_total)
    add(LIAB_PPK_ACCT, "", f"{ym_prefix} Poland Payroll 401K Payable", credit=ppk_total)

    df = pd.DataFrame(lines)

    agg = (
        df.assign(Debit=pd.to_numeric(df["Debit"], errors="coerce").fillna(0.0),
                  Credit=pd.to_numeric(df["Credit"], errors="coerce").fillna(0.0))
          .groupby(["Date", "Journal Entry Memo", "Account", "Line Memo", "Subsidiary", "Department"],
                   dropna=False, as_index=False, sort=False)[["Debit", "Credit"]]
          .sum()
    )
    agg["Debit"] = agg["Debit"].apply(lambda x: round(x, 2) if x else "")
    agg["Credit"] = agg["Credit"].apply(lambda x: round(x, 2) if x else "")
    agg = agg[["Date", "Journal Entry Memo", "Account", "Debit", "Credit", "Line Memo", "Subsidiary", "Department"]]

    total_dr = sum(float(x) for x in agg["Debit"] if x != "")
    total_cr = sum(float(x) for x in agg["Credit"] if x != "")
    imbalance = round(total_dr - total_cr, 2)
    if abs(imbalance) > 0.01:
        raise RuntimeError(
            f"JE out of balance: Debits={total_dr:.2f}, Credits={total_cr:.2f}, Imbalance={imbalance:.2f}\n"
            f"  Net={net_total}, Social={social_total}, Tax={tax_total}, PPK={ppk_total}"
        )

    wb = Workbook()
    ws_raw = wb.active
    ws_raw.title = "raw"
    for row in dataframe_to_rows(raw, index=False, header=False):
        ws_raw.append(list(row))

    ws_je = wb.create_sheet("JE")
    ws_je.append(list(agg.columns))
    for _, row in agg.iterrows():
        ws_je.append([row[c] for c in agg.columns])

    wb.save(output_xlsx)

    csv_path = Path(output_xlsx).with_name(f"{ym_prefix} Poland Payroll JE Import.csv")
    write_je_csv(
        rows=agg.to_dict(orient="records"),
        output_path=csv_path,
        currency="PLN",
        default_external_id=make_external_id(ym_prefix, "PL"),
    )

    print(f"Wrote {output_xlsx}")
    print(f"Wrote {csv_path}")
    print(f"  JE date:     {je_date.isoformat()}")
    print(f"  JE memo:     {je_memo}")
    print(f"  Lines:       {len(agg)}")
    print(f"  Total Dr:    {total_dr:,.2f} PLN")
    print(f"  Total Cr:    {total_cr:,.2f} PLN")
    print(f"  Balanced:    {abs(imbalance) <= 0.01}")
    return agg


if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    emp_map_default = os.path.join(script_dir, "Poland Payroll Employee Mapping.csv")
    gl_map_default = os.path.join(script_dir, "Poland Payroll GL Mapping.csv")

    if len(sys.argv) < 2:
        print("Usage: python process_poland_payroll.py <input.xlsx> [output.xlsx] [emp_map.csv] [gl_map.csv]")
        sys.exit(1)

    raw_file = sys.argv[1]
    input_dir = os.path.dirname(os.path.abspath(raw_file))
    folder_name = os.path.basename(input_dir)
    try:
        mm, yyyy = folder_name.split("-")
        ym_tag = f"{yyyy}-{mm}"
    except ValueError:
        ym_tag = folder_name

    default_output = os.path.join(input_dir, f"{ym_tag} Poland Payroll Backup.xlsx")
    out_xlsx = sys.argv[2] if len(sys.argv) > 2 else default_output
    emp_map = sys.argv[3] if len(sys.argv) > 3 else emp_map_default
    gl_map = sys.argv[4] if len(sys.argv) > 4 else gl_map_default

    generate_je(raw_file, emp_map, gl_map, out_xlsx)
