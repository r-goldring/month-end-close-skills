"""
Shared NetSuite Journal Entry Import CSV writer.

Used by foreign payroll scripts (Germany, Netherlands, Poland, UK, Uruguay) to
emit a CSV alongside the Backup.xlsx workbook. The CSV is uploaded via NetSuite
Lists -> Import Assistant -> Transactions -> Journal Entry. JEs land in the
controller's Pending Approval queue, preserving separation of duties (the
auto-approver issue that prompted commit 14e19d7).

Column order:
  External ID, Date, Journal Entry Memo, Currency, Account, Debit, Credit,
  Line Memo, Subsidiary, Department

Format:
  - Date: M/D/YYYY (handles 'YYYY-MM-DD' string OR datetime/date object)
  - Account: number only (e.g., '611100'), extracted from the row's existing
    'Account' field by splitting on the first space
  - Debit/Credit: numeric or empty (no zeros, no quotes)
  - Subsidiary path quoted via QUOTE_MINIMAL (it contains commas)
  - ASCII only -- raises ValueError on any non-ASCII char so we never silently
    upload an em-dash or curly quote

Combined-CSV usage (Uruguay): pass an external_id_fn callback so each row's
External ID can vary by JE block. NetSuite Import Assistant bundles rows that
share an External ID into a single JE.
"""

import csv
import datetime as dt
import re
from pathlib import Path
from typing import Callable, Iterable, Optional


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
]


def _normalize_date(d) -> str:
    """Return M/D/YYYY for a datetime/date object or a 'YYYY-MM-DD' / 'YYYY-MM-DDTHH:MM:SS' string."""
    if isinstance(d, (dt.date, dt.datetime)):
        return f"{d.month}/{d.day}/{d.year}"
    if isinstance(d, str):
        s = d.strip()
        if not s:
            raise ValueError("Date field is empty")
        # Accept 'YYYY-MM-DD' or anything starting with that pattern (e.g. with time suffix)
        m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
        if m:
            yyyy, mm, dd = m.groups()
            return f"{int(mm)}/{int(dd)}/{int(yyyy)}"
        # Already M/D/YYYY?
        m2 = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
        if m2:
            return s
        raise ValueError(f"Unrecognized date format: {d!r}")
    raise ValueError(f"Unsupported date type: {type(d).__name__} ({d!r})")


def _account_number(account_str: str) -> str:
    """Extract leading account number from '611100 Salaries and Wages' -> '611100'."""
    if not account_str:
        return ""
    return str(account_str).split(" ", 1)[0]


def _coerce_amount(v) -> str:
    """Empty cell stays empty (no zeros). Numeric -> str with 2 decimals. Pre-formatted strings pass through."""
    if v is None:
        return ""
    if isinstance(v, str):
        s = v.strip()
        return s  # already formatted (e.g., "" or "1234.56")
    try:
        f = float(v)
    except (TypeError, ValueError):
        return ""
    if f == 0:
        return ""
    return f"{f:.2f}"


def _ascii_check(value: str, *, row_idx: int, column: str) -> str:
    try:
        value.encode("ascii")
    except UnicodeEncodeError as e:
        raise ValueError(
            f"Non-ASCII character in row {row_idx} column {column!r}: {value!r} ({e})"
        ) from None
    return value


def write_je_csv(
    rows: Iterable[dict],
    output_path,
    *,
    currency: str,
    external_id_fn: Optional[Callable[[dict], str]] = None,
    default_external_id: Optional[str] = None,
    extra_columns: Optional[list] = None,
) -> Path:
    """
    Write a NetSuite-Import-Assistant-compatible JE CSV.

    rows: each dict must have keys: Date, "Journal Entry Memo", Account, Debit,
          Credit, "Line Memo", Subsidiary, Department.
    output_path: str or Path.
    currency: ISO code (e.g. 'EUR', 'PLN', 'GBP', 'UYU'). Written on every row.
    external_id_fn: callback(row_dict) -> external_id. If None, default_external_id
                    is used for all rows.
    default_external_id: used when external_id_fn is None. Required if no fn given.
    extra_columns: optional list of (header_name, row_dict_key) tuples. Adds extra
                   columns to the CSV after Department. If a key is "Reversal Date",
                   the value is normalized via _normalize_date. Other extras are
                   ASCII-checked and written as-is.

    Returns: the Path written.
    Raises: ValueError on non-ASCII content or missing external ID.
    """
    rows = list(rows)
    if not rows:
        raise ValueError("write_je_csv called with no rows")

    if external_id_fn is None:
        if not default_external_id:
            raise ValueError("Must provide external_id_fn or default_external_id")
        external_id_fn = lambda _row: default_external_id  # noqa: E731

    extras = list(extra_columns or [])
    headers = list(CSV_HEADERS) + [h for (h, _k) in extras]

    output_path = Path(output_path)

    with open(output_path, "w", newline="", encoding="ascii") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
        writer.writerow(headers)
        line_count = 0
        for i, row in enumerate(rows, start=1):
            debit = _coerce_amount(row.get("Debit"))
            credit = _coerce_amount(row.get("Credit"))
            # Skip placeholder rows where both sides are empty (NL's fixed-template
            # padding for components that didn't apply this month). NetSuite Import
            # Assistant rejects rows without a debit or credit.
            if not debit and not credit:
                continue

            ext_id = external_id_fn(row)
            if not ext_id:
                raise ValueError(f"Row {i} has no External ID")

            date_str = _normalize_date(row["Date"])
            memo = str(row["Journal Entry Memo"])
            acct = _account_number(row["Account"])
            line_memo = str(row.get("Line Memo") or "")
            subsidiary = str(row.get("Subsidiary") or "")
            department = str(row.get("Department") or "")

            out = [
                _ascii_check(str(ext_id), row_idx=i, column="External ID"),
                _ascii_check(date_str, row_idx=i, column="Date"),
                _ascii_check(memo, row_idx=i, column="Journal Entry Memo"),
                _ascii_check(str(currency), row_idx=i, column="Currency"),
                _ascii_check(acct, row_idx=i, column="Account"),
                _ascii_check(debit, row_idx=i, column="Debit"),
                _ascii_check(credit, row_idx=i, column="Credit"),
                _ascii_check(line_memo, row_idx=i, column="Line Memo"),
                _ascii_check(subsidiary, row_idx=i, column="Subsidiary"),
                _ascii_check(department, row_idx=i, column="Department"),
            ]
            for header_name, key in extras:
                v = row.get(key) or row.get(header_name) or ""
                if v and key.lower().endswith("date"):
                    v = _normalize_date(v)
                out.append(_ascii_check(str(v), row_idx=i, column=header_name))
            writer.writerow(out)
            line_count += 1

    if line_count == 0:
        output_path.unlink(missing_ok=True)
        raise ValueError("write_je_csv produced no non-empty rows; nothing written")

    return output_path


def make_external_id(yyyy_mm: str, country_code: str, suffix: str = "") -> str:
    """
    Build a deterministic, ASCII External ID.
      yyyy_mm: '2026-04'
      country_code: 'DE' / 'NL' / 'PL' / 'UK' / 'UY'
      suffix: optional, e.g. 'PAYROLL' or 'PAYROLL-AGUINALDO-ACCRUAL'

    Examples:
      make_external_id('2026-04', 'DE')                          -> '2026-04-DE-PAYROLL'
      make_external_id('2026-04', 'UY', 'Aguinaldo Accrual')     -> '2026-04-UY-AGUINALDO-ACCRUAL'
    """
    suffix = (suffix or "PAYROLL").strip()
    suffix = re.sub(r"[^A-Za-z0-9]+", "-", suffix).strip("-").upper()
    if not suffix:
        suffix = "PAYROLL"
    return f"{yyyy_mm}-{country_code}-{suffix}"
