"""
Germany Payroll -> NetSuite JE generator.

Reads the monthly raw spreadsheet from the provider (Tabelle1 sheet, header on row 3),
maps each earning/benefit column to a GL account, routes COGS vs OpEx based on
each employee's department (from Germany Payroll Employee Mapping.csv), and writes
an Excel workbook with a 'raw' tab + 'JE' tab matching the accountant's historical backup format.

Unknown employees or unknown non-zero columns halt the script with a clear flag.

Usage:
    python process_germany_payroll.py <input.xlsx> [output.xlsx] [emp_map.csv] [gl_map.csv]

Optional severance (the accountant specifies in-chat when applicable):
    generate_je(..., severance_spec={"employee": "Name, First", "amount": 5000.00, "taxes": 1200.00})
"""

import os
import sys
import calendar
import datetime as dt
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(SCRIPT_DIR, "..", "_shared"))
from je_csv_writer import write_je_csv, make_external_id  # noqa: E402

SUBSIDIARY = "Acme Holdings : Acme, Inc. : Acme Netherlands"
LIABILITY_ACCT = "231200 Payroll Liability"
SEVERANCE_ACCT = "511175 COGS - Salary and Compensation : COGS - Severance"
SEVERANCE_TAX_ACCT = "511350 COGS - Salary and Compensation : COGS - Payroll Taxes"
SEVERANCE_DEPT = "EBITDA Adjustments"
VBG_COGS_ACCT = "511350 COGS - Salary and Compensation : COGS - Payroll Taxes"
VBG_OPEX_ACCT = "611450 Salary and Compensation : Payroll Taxes"
VBG_MEMO = "Germany VBG (Workers Comp Insurance)"


def _to_float(v):
    if pd.isna(v):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("$", "")
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
        if "Employee Name" in vals and "Net Pay" in vals and "Cost Center" in vals:
            return r
    raise ValueError("Could not locate header row (needs 'Employee Name' + 'Net Pay' + 'Cost Center').")


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


def generate_je(raw_xlsx, emp_map_csv, gl_map_csv, output_xlsx, severance_spec=None):
    emp_map = _load_employee_map(emp_map_csv)
    gl_map = _load_gl_map(gl_map_csv)

    raw = pd.read_excel(raw_xlsx, sheet_name=0, header=None)

    header_row = _find_header_row(raw)
    headers = [str(x).strip() if pd.notna(x) else "" for x in raw.iloc[header_row].values]
    col_idx = {h: i for i, h in enumerate(headers) if h}

    unknown_cols = [h for h in headers if h and h not in gl_map]
    if unknown_cols:
        raise RuntimeError(
            "Unknown columns in raw file (not in Germany Payroll GL Mapping.csv): "
            + ", ".join(unknown_cols)
            + "\nAdd rows for these columns to the GL mapping and re-run."
        )

    employee_rows = []
    for r in range(header_row + 1, len(raw)):
        row = raw.iloc[r]
        name_cell = row[col_idx["Employee Name"]] if "Employee Name" in col_idx else None
        if pd.isna(name_cell) or not str(name_cell).strip():
            break
        employee_rows.append(r)

    if not employee_rows:
        raise RuntimeError("No employee rows found below header row.")

    check_date_val = raw.iloc[employee_rows[0]][col_idx["Check Date"]]
    if isinstance(check_date_val, dt.datetime):
        check_date = check_date_val.date()
    else:
        check_date = pd.to_datetime(str(check_date_val)).date()
    je_date = _last_day_of_month(check_date)
    ym_prefix = je_date.strftime("%Y-%m")
    je_memo = f"{ym_prefix} Germany Payroll"

    unknown_emps = []
    for r in employee_rows:
        name = str(raw.iloc[r][col_idx["Employee Name"]]).strip()
        if name not in emp_map:
            unknown_emps.append(name)
    if unknown_emps:
        raise RuntimeError(
            "Unknown employees (not in Germany Payroll Employee Mapping.csv): "
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

    for r in employee_rows:
        row = raw.iloc[r]
        name = str(row[col_idx["Employee Name"]]).strip()
        dept = emp_map[name]
        is_cogs = _is_cogs(dept)

        for col_name, mapping in gl_map.items():
            if mapping["role"] != "Expense":
                continue
            if col_name not in col_idx:
                continue
            amt = _to_float(row[col_idx[col_name]])
            if amt == 0.0:
                continue
            acct = mapping["cogs"] if is_cogs else mapping["opex"]
            if not acct:
                # Some columns may only have one side (e.g., expense reimbursement is OpEx-only)
                acct = mapping["opex"] or mapping["cogs"]
            if not acct:
                raise RuntimeError(
                    f"No GL account configured for column '{col_name}' on {'COGS' if is_cogs else 'OpEx'} side "
                    f"(employee {name}, dept {dept})."
                )
            line_memo = f"{ym_prefix} {mapping['memo']}"
            add(acct, dept, line_memo, debit=amt)

        gift_col = "Gift Card"
        if gift_col in col_idx:
            gift_amt = _to_float(row[col_idx[gift_col]])
            if gift_amt != 0.0:
                print(f"  [note] Ignoring Gift Card amount {gift_amt:.2f} for {name} (per current policy).")

    net_pay_total = 0.0
    ss_contrib_total = 0.0
    vbg_total = 0.0
    for r in range(header_row + 1, len(raw)):
        row = raw.iloc[r]
        label_col4 = row[4] if len(row) > 4 else None
        if pd.notna(label_col4):
            label = str(label_col4).strip()
            amt_cell = row[col_idx["Net Pay"]] if "Net Pay" in col_idx else None
            amt = _to_float(amt_cell)
            if label == "Netpay":
                net_pay_total = amt
            elif label == "Social Security contributions":
                ss_contrib_total = amt
            elif label in ("VBG", "VGB"):
                vbg_total = amt

    if vbg_total > 0:
        ss_col = "social security ER"
        ee_ss_weights = []
        for r in employee_rows:
            row = raw.iloc[r]
            name = str(row[col_idx["Employee Name"]]).strip()
            ss = _to_float(row[col_idx[ss_col]]) if ss_col in col_idx else 0.0
            ee_ss_weights.append((name, emp_map[name], ss))
        total_weight = sum(w for _, _, w in ee_ss_weights)
        if total_weight == 0.0:
            total_weight = float(len(ee_ss_weights))
            ee_ss_weights = [(n, d, 1.0) for n, d, _ in ee_ss_weights]
        allocated = 0.0
        for i, (name, dept, w) in enumerate(ee_ss_weights):
            if i == len(ee_ss_weights) - 1:
                share = round(vbg_total - allocated, 2)
            else:
                share = round(vbg_total * w / total_weight, 2)
                allocated += share
            if share == 0.0:
                continue
            is_cogs = _is_cogs(dept)
            acct = VBG_COGS_ACCT if is_cogs else VBG_OPEX_ACCT
            add(acct, dept, f"{ym_prefix} {VBG_MEMO}", debit=share)

    if net_pay_total == 0.0 or ss_contrib_total == 0.0:
        raise RuntimeError(
            f"Could not locate 'Netpay' / 'Social Security contributions' totals in raw file "
            f"(got net={net_pay_total}, ss={ss_contrib_total})."
        )

    severance_dr = 0.0
    severance_tax_dr = 0.0
    if severance_spec:
        sev_emp = severance_spec.get("employee", "unspecified")
        severance_dr = float(severance_spec.get("amount", 0.0))
        severance_tax_dr = float(severance_spec.get("taxes", 0.0))
        if severance_dr > 0:
            add(
                SEVERANCE_ACCT,
                SEVERANCE_DEPT,
                f"{ym_prefix} Germany Payroll - Severance ({sev_emp})",
                debit=severance_dr,
            )
        if severance_tax_dr > 0:
            add(
                SEVERANCE_TAX_ACCT,
                SEVERANCE_DEPT,
                f"{ym_prefix} Germany Payroll - Severance Taxes ({sev_emp})",
                debit=severance_tax_dr,
            )

    add(LIABILITY_ACCT, "", f"{ym_prefix} Germany Payroll Liability", credit=net_pay_total + severance_dr)
    add(LIABILITY_ACCT, "", f"{ym_prefix} Germany Payroll Taxes Liability", credit=ss_contrib_total + severance_tax_dr)
    if vbg_total > 0:
        add(LIABILITY_ACCT, "", f"{ym_prefix} Germany VBG Liability", credit=vbg_total)

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
            f"JE out of balance: Debits={total_dr:.2f}, Credits={total_cr:.2f}, Imbalance={imbalance:.2f}"
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

    csv_path = Path(output_xlsx).with_name(f"{ym_prefix} Germany Payroll JE Import.csv")
    write_je_csv(
        rows=agg.to_dict(orient="records"),
        output_path=csv_path,
        currency="EUR",
        default_external_id=make_external_id(ym_prefix, "DE"),
    )

    print(f"Wrote {output_xlsx}")
    print(f"Wrote {csv_path}")
    print(f"  JE date:     {je_date.isoformat()}")
    print(f"  JE memo:     {je_memo}")
    print(f"  Lines:       {len(agg)}")
    print(f"  Total Dr:    {total_dr:,.2f}")
    print(f"  Total Cr:    {total_cr:,.2f}")
    print(f"  Balanced:    {abs(imbalance) <= 0.01}")
    return agg


if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    emp_map_default = os.path.join(script_dir, "Germany Payroll Employee Mapping.csv")
    gl_map_default = os.path.join(script_dir, "Germany Payroll GL Mapping.csv")

    if len(sys.argv) < 2:
        print("Usage: python process_germany_payroll.py <input.xlsx> [output.xlsx] [emp_map.csv] [gl_map.csv]")
        sys.exit(1)

    raw_file = sys.argv[1]
    input_dir = os.path.dirname(os.path.abspath(raw_file))
    folder_name = os.path.basename(input_dir)
    try:
        mm, yyyy = folder_name.split("-")
        ym_tag = f"{yyyy}-{mm}"
    except ValueError:
        ym_tag = folder_name

    default_output = os.path.join(input_dir, f"{ym_tag} Germany Payroll Backup.xlsx")
    out_xlsx = sys.argv[2] if len(sys.argv) > 2 else default_output
    emp_map = sys.argv[3] if len(sys.argv) > 3 else emp_map_default
    gl_map = sys.argv[4] if len(sys.argv) > 4 else gl_map_default

    generate_je(raw_file, emp_map, gl_map, out_xlsx)
