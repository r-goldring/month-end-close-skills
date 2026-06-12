"""
UK Payroll -> NetSuite JE generator.

Reads the monthly raw provider PREReport (PreReport sheet, header row 1, one row
per employee) and a corresponding HMRC P32 Summary PDF for HMRC reconciliation.
Maps each earning/benefit column to a GL account, routes COGS vs OpEx based on
each employee's department, computes SMP Unrecovered + Apprenticeship Levy from
the P32 PDF, and writes an Excel workbook with a 'raw' + 'JE' tab matching
the accountant's historical backup format.

Severance routes to 511175 (COGS) / 611250 (OpEx) under EBITDA Adjustments.
SMP Unrecovered + App Levy post to the SMP recipient's department (defaulting
to Sales & Marketing : Revenue : New Business if no SMP this month).
Health Insurance is treated as an EE-paid premium recovery (CREDIT).

Usage:
    python process_uk_payroll.py <raw_PREReport.xlsx> [output.xlsx] [emp_map.csv] [gl_map.csv]

The P32 PDF is auto-discovered as `P32Summary_*.pdf` in the same folder as the raw.
"""

import os
import re
import sys
import glob
import calendar
import datetime as dt
from pathlib import Path
import pandas as pd
import pdfplumber
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(SCRIPT_DIR, "..", "_shared"))
from je_csv_writer import write_je_csv, make_external_id  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

SUBSIDIARY = "Acme Holdings : Acme, Inc. : Acme UK Ltd"
LIAB_ACCT = "231200 Payroll Liability"
SEVERANCE_DEPT = "EBITDA Adjustments"
DEFAULT_SMP_DEPT = "Sales & Marketing : Revenue : New Business"

# Health insurance (EE-paid recovery) credit account by side
HEALTH_COGS = "511200 COGS - Salary and Compensation : COGS - Health Benefits"
HEALTH_OPEX = "611300 Salary and Compensation : Health Benefits"

# App Levy account (employer payroll tax)
APP_LEVY_COGS = "511350 COGS - Salary and Compensation : COGS - Payroll Taxes"
APP_LEVY_OPEX = "611450 Salary and Compensation : Payroll Taxes"


def _to_float(v):
    if pd.isna(v):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("$", "").replace("£", "").replace(" ", "")
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
    for r in range(min(5, len(df))):
        vals = [str(x).strip() for x in df.iloc[r].values if pd.notna(x)]
        if "EeRef" in vals and "Name" in vals and "NetPay" in vals:
            return r
    raise ValueError("Could not locate header row (needs 'EeRef' + 'Name' + 'NetPay').")


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
    d = department.strip()
    return d.startswith("COGS :") or d.startswith("COGS:")


def parse_p32_pdf(pdf_path):
    """Parse a P32Summary PDF; return dict with values for the tax month
    indicated by the filename suffix (P32Summary_YYYY-MM.pdf -> month MM)."""
    fn = os.path.basename(pdf_path)
    m = re.search(r"P32Summary[_-](\d{4})[-_](\d{1,2})\.pdf$", fn, re.IGNORECASE)
    if not m:
        raise ValueError(f"Cannot derive tax month from PDF filename: {fn}")
    tax_year = int(m.group(1))
    tax_month = int(m.group(2))

    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join(pg.extract_text() or "" for pg in pdf.pages)

    pat = rf"TOTAL for Month {tax_month}\b([\d.\s]+)"
    match = re.search(pat, text)
    if not match:
        raise RuntimeError(f"Could not find 'TOTAL for Month {tax_month}' line in {fn}")

    nums = [float(x) for x in match.group(1).split() if re.fullmatch(r"-?\d+\.?\d*", x)]
    if len(nums) < 23:
        raise RuntimeError(
            f"Expected at least 23 numeric fields on TOTAL line, got {len(nums)}: {nums}"
        )

    return {
        "tax_year": tax_year,
        "tax_month": tax_month,
        "paye_tax": nums[0],
        "student_loans": nums[1],
        "pay_income_tax": nums[2],
        "gross_ee_ni": nums[3],
        "net_allow": nums[4],
        "smp_recovered": nums[5],
        "total_deductions": nums[-6],
        "app_levy": nums[-5],
        "net_nics": nums[-4],
        "funding_received": nums[-3],
        "net_credit": nums[-2],
        "total_due": nums[-1],
    }


def generate_je(raw_xlsx, p32_pdf, emp_map_csv, gl_map_csv, output_xlsx):
    emp_map = _load_employee_map(emp_map_csv)
    gl_map = _load_gl_map(gl_map_csv)

    raw = pd.read_excel(raw_xlsx, sheet_name="PreReport", header=None)

    header_row = _find_header_row(raw)
    headers = [str(x).strip() if pd.notna(x) else "" for x in raw.iloc[header_row].values]
    col_idx = {h: i for i, h in enumerate(headers) if h}

    unknown_cols = [h for h in headers if h and h not in gl_map]
    if unknown_cols:
        raise RuntimeError(
            "Unknown columns in raw file (not in UK Payroll GL Mapping.csv): "
            + ", ".join(unknown_cols)
            + "\nAdd rows for these columns to the GL mapping and re-run."
        )

    employee_rows = []
    for r in range(header_row + 1, len(raw)):
        row = raw.iloc[r]
        name = row[col_idx["Name"]] if "Name" in col_idx else None
        if pd.isna(name) or not str(name).strip() or str(name).strip().lower() == "totals":
            break
        employee_rows.append(r)

    if not employee_rows:
        raise RuntimeError("No employee rows found below header row.")

    # JE date from PayRunDate (last day of month)
    payrun_val = raw.iloc[employee_rows[0]][col_idx["PayRunDate"]]
    if isinstance(payrun_val, dt.datetime):
        payrun_date = payrun_val.date()
    else:
        try:
            payrun_date = pd.to_datetime(str(payrun_val), dayfirst=True).date()
        except Exception:
            payrun_date = pd.to_datetime(str(payrun_val)).date()
    je_date = _last_day_of_month(payrun_date)
    ym_prefix = je_date.strftime("%Y-%m")
    je_memo = f"{ym_prefix} UK Payroll"

    # Validate every employee is in the mapping
    unknown_emps = []
    for r in employee_rows:
        name = str(raw.iloc[r][col_idx["Name"]]).strip()
        if name not in emp_map:
            unknown_emps.append(name)
    if unknown_emps:
        # Dedupe while preserving order
        seen = set()
        unique = [n for n in unknown_emps if not (n in seen or seen.add(n))]
        raise RuntimeError(
            "Unknown employees (not in UK Payroll Employee Mapping.csv): "
            + ", ".join(unique)
            + "\nAdd rows for each employee with their Department and re-run."
        )

    # Parse the P32 PDF
    p32 = parse_p32_pdf(p32_pdf)
    print(f"P32 month {p32['tax_month']} (tax year {p32['tax_year']}/{p32['tax_year']+1}): "
          f"SMP Recovered={p32['smp_recovered']:,.2f}, App Levy={p32['app_levy']:,.2f}, "
          f"Total Due={p32['total_due']:,.2f}")

    lines = []

    def add(acct, dept, memo, debit=0.0, credit=0.0):
        if debit == 0.0 and credit == 0.0:
            return
        lines.append({
            "Date": je_date.strftime("%Y-%m-%d"),
            "Journal Entry Memo": je_memo,
            "Account": acct,
            "Debit": round(debit, 2) if debit else "",
            "Credit": round(credit, 2) if credit else "",
            "Line Memo": memo,
            "Subsidiary": SUBSIDIARY,
            "Department": dept,
        })

    # Pass 1: per-employee expense, severance, and health-recovery lines
    smp_per_emp = {}  # name -> SMP gross paid by provider
    smp_emp_dept = {}

    for r in employee_rows:
        row = raw.iloc[r]
        name = str(row[col_idx["Name"]]).strip()
        dept = emp_map[name]
        is_cogs = _is_cogs(dept)

        # Track SMP for later allocation
        if "SMP" in col_idx:
            smp_amt = _to_float(row[col_idx["SMP"]])
            if smp_amt != 0.0:
                smp_per_emp[name] = smp_per_emp.get(name, 0.0) + smp_amt
                smp_emp_dept[name] = dept

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
                        f"No GL account configured for column '{col_name}' on "
                        f"{'COGS' if is_cogs else 'OpEx'} side (employee {name})."
                    )
                add(acct, dept, f"{ym_prefix} {mapping['memo']}", debit=amt)

            elif role == "Severance":
                acct = mapping["cogs"] if is_cogs else mapping["opex"]
                if not acct:
                    acct = mapping["cogs"] or mapping["opex"]
                add(acct, SEVERANCE_DEPT, f"{ym_prefix} {mapping['memo']}", debit=amt)

            elif role == "HealthRecovery":
                # Provider stores deductions as negative; flip to positive recovery credit
                acct = mapping["cogs"] if is_cogs else mapping["opex"]
                add(acct, dept, f"{ym_prefix} {mapping['memo']}", credit=-amt)

    # SMP Unrecovered + App Levy from PDF
    provider_smp_total = sum(smp_per_emp.values())
    smp_unrecovered = round(provider_smp_total - p32["smp_recovered"], 2)
    if smp_unrecovered > 0.005:
        # Allocate to SMP recipient(s) by their share of SMP
        for name, smp in smp_per_emp.items():
            dept = smp_emp_dept[name]
            is_cogs = _is_cogs(dept)
            acct = "511100 COGS - Salary and Compensation : COGS - Salaries and Wages" if is_cogs \
                else "611100 Salary and Compensation : Salaries and Wages"
            share = round(smp_unrecovered * (smp / provider_smp_total), 2)
            add(acct, dept, f"{ym_prefix} UK Payroll - SMP Unrecovered", debit=share)
    elif smp_unrecovered < -0.005:
        print(f"  [info] SMP overcovered by {-smp_unrecovered:,.2f} (HMRC reclaim > provider gross). "
              "Not posting an offset; review manually if material.")

    if p32["app_levy"] != 0:
        # Pick department: SMP recipient(s) if any, else default
        if smp_per_emp:
            for name, smp in smp_per_emp.items():
                dept = smp_emp_dept[name]
                is_cogs = _is_cogs(dept)
                acct = APP_LEVY_COGS if is_cogs else APP_LEVY_OPEX
                share = round(p32["app_levy"] * (smp / provider_smp_total), 2)
                add(acct, dept, f"{ym_prefix} UK Payroll - Apprenticeship Levy", debit=share)
        else:
            # No SMP recipient this month; default to New Business per the accountant's convention
            dept = DEFAULT_SMP_DEPT
            is_cogs = _is_cogs(dept)
            acct = APP_LEVY_COGS if is_cogs else APP_LEVY_OPEX
            add(acct, dept, f"{ym_prefix} UK Payroll - Apprenticeship Levy", debit=p32["app_levy"])

    # Liability credits
    net_pay_total = sum(_to_float(raw.iloc[r][col_idx["NetPay"]]) for r in employee_rows)
    pension_total = sum(_to_float(raw.iloc[r][col_idx["TotalPens"]]) for r in employee_rows) \
        if "TotalPens" in col_idx else 0.0

    add(LIAB_ACCT, "", f"{ym_prefix} UK Payroll Liability", credit=net_pay_total)
    add(LIAB_ACCT, "", f"{ym_prefix} UK Payroll - Pension Liability (EE+ER)", credit=pension_total)
    add(LIAB_ACCT, "", f"{ym_prefix} UK Payroll - HMRC Liability", credit=p32["total_due"])

    # Aggregate by (account, dept, memo) to keep JE summarized
    df = pd.DataFrame(lines)
    agg = (
        df.assign(
            Debit=pd.to_numeric(df["Debit"], errors="coerce").fillna(0.0),
            Credit=pd.to_numeric(df["Credit"], errors="coerce").fillna(0.0),
        )
        .groupby(["Date", "Journal Entry Memo", "Account", "Line Memo", "Subsidiary", "Department"],
                 dropna=False, as_index=False, sort=False)[["Debit", "Credit"]]
        .sum()
    )
    agg["Debit"] = agg["Debit"].apply(lambda x: round(x, 2) if x else "")
    agg["Credit"] = agg["Credit"].apply(lambda x: round(x, 2) if x else "")
    agg = agg[["Date", "Journal Entry Memo", "Account", "Debit", "Credit",
               "Line Memo", "Subsidiary", "Department"]]

    total_dr = sum(float(x) for x in agg["Debit"] if x != "")
    total_cr = sum(float(x) for x in agg["Credit"] if x != "")
    imbalance = round(total_dr - total_cr, 2)
    if abs(imbalance) > 0.01:
        raise RuntimeError(
            f"JE out of balance: Debits={total_dr:,.2f}, Credits={total_cr:,.2f}, "
            f"Imbalance={imbalance:,.2f}\n"
            f"  Net Pay total: {net_pay_total:,.2f}\n"
            f"  Pension total: {pension_total:,.2f}\n"
            f"  HMRC Total Due: {p32['total_due']:,.2f}\n"
            f"  Provider SMP: {provider_smp_total:,.2f}, PDF SMP Recovered: {p32['smp_recovered']:,.2f}, "
            f"  SMP Unrecovered: {smp_unrecovered:,.2f}, App Levy: {p32['app_levy']:,.2f}"
        )

    # Write workbook
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

    csv_path = Path(output_xlsx).with_name(f"{ym_prefix} UK Payroll JE Import.csv")
    write_je_csv(
        rows=agg.to_dict(orient="records"),
        output_path=csv_path,
        currency="GBP",
        default_external_id=make_external_id(ym_prefix, "UK"),
    )

    print(f"Wrote {output_xlsx}")
    print(f"Wrote {csv_path}")
    print(f"  JE date:         {je_date.isoformat()}")
    print(f"  JE memo:         {je_memo}")
    print(f"  Lines:           {len(agg)}")
    print(f"  Total Dr:        {total_dr:>14,.2f} GBP")
    print(f"  Total Cr:        {total_cr:>14,.2f} GBP")
    print(f"  Balanced:        {abs(imbalance) <= 0.01}")
    if smp_unrecovered > 0.005:
        print(f"  SMP Unrecovered: {smp_unrecovered:>14,.2f} (provider {provider_smp_total:.2f} - PDF rec {p32['smp_recovered']:.2f})")
    if p32["app_levy"] != 0:
        print(f"  App Levy:        {p32['app_levy']:>14,.2f}")
    return agg


if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    emp_map_default = os.path.join(script_dir, "UK Payroll Employee Mapping.csv")
    gl_map_default = os.path.join(script_dir, "UK Payroll GL Mapping.csv")

    if len(sys.argv) < 2:
        print("Usage: python process_uk_payroll.py <raw_PREReport.xlsx> [output.xlsx] [emp_map.csv] [gl_map.csv]")
        sys.exit(1)

    raw_file = sys.argv[1]
    input_dir = os.path.dirname(os.path.abspath(raw_file))

    pdfs = sorted(glob.glob(os.path.join(input_dir, "P32Summary_*.pdf")))
    if not pdfs:
        print(f"ERROR: No P32Summary_*.pdf found in {input_dir}")
        print("Drop the HMRC P32 Summary PDF into the same folder and re-run.")
        sys.exit(1)
    p32_path = pdfs[0]

    folder_name = os.path.basename(input_dir)
    # Folder is YYYY-MM (UK convention) or MM-YYYY for some other countries.
    # Use the JE date computation from raw to derive output filename.
    if re.fullmatch(r"\d{4}-\d{2}", folder_name):
        ym_tag = folder_name
    elif re.fullmatch(r"\d{2}-\d{4}", folder_name):
        mm, yyyy = folder_name.split("-")
        ym_tag = f"{yyyy}-{mm}"
    else:
        ym_tag = folder_name

    default_output = os.path.join(input_dir, f"{ym_tag} UK Payroll Backup.xlsx")
    out_xlsx = sys.argv[2] if len(sys.argv) > 2 else default_output
    emp_map = sys.argv[3] if len(sys.argv) > 3 else emp_map_default
    gl_map = sys.argv[4] if len(sys.argv) > 4 else gl_map_default

    print(f"Raw:    {raw_file}")
    print(f"P32:    {p32_path}")
    print(f"Output: {out_xlsx}\n")
    generate_je(raw_file, p32_path, emp_map, gl_map, out_xlsx)
