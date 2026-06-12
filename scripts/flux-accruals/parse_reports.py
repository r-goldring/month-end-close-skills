"""
Parse the 4 NetSuite reports (537/540/542/721) into a flat list of
TransactionRecord rows the accrual builder can join, group, and aggregate.

Claude pulls each report via ns_runReport and saves the JSON to
  Monthly Flux Analysis/{YYYY}/{YYYY-MM}/_cache/{report_id}.json

This module reads those JSONs and emits a uniform schema regardless of which
report produced them.

Schema is documented in flux-analysis-workbook/SKILL.md Step 3 and mirrored here.
"""

from __future__ import annotations

import datetime as dt
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple


CATEGORY_BY_REPORT_ID = {
    537: "COGS",
    540: "Contractors",
    542: "ProfFees",
    721: "Software",
}

# Account filters per category. Mirrors flux-analysis-workbook + CLAUDE.md.
ACCOUNT_FILTERS: Dict[str, set[str]] = {
    "COGS":        {"511400", "511425", "511450", "511510", "511520", "511550", "511600"},
    "Contractors": {"511370", "611700"},
    "ProfFees":    {"651100", "651101", "651150"},
    "Software":    {"671000", "671100", "511425"},
}


@dataclass
class TransactionRecord:
    category: str          # COGS / Contractors / ProfFees / Software
    period_yyyymm: str     # "2026-04"
    date: Optional[str]    # "2026-04-15" (ISO, may be None for some sub-records)
    vendor: str            # cleaned final name; "" if unknown
    account_number: str    # "671100" (parsed from Account or section header)
    account_name: str      # "Software Subscriptions"
    department: str        # cost center display name
    amount: float          # positive = expense; signed correctly per category rules


def parse_report_json(report_id: int, payload: dict) -> List[TransactionRecord]:
    """
    Parse one ns_runReport response into TransactionRecords.
    payload = the full JSON returned by ns_runReport (has reportData dict).
    """
    if report_id not in CATEGORY_BY_REPORT_ID:
        raise ValueError(f"Unknown report_id {report_id}; expected one of {list(CATEGORY_BY_REPORT_ID)}")

    category = CATEGORY_BY_REPORT_ID[report_id]
    report_data = payload.get("reportData", {})
    if not report_data:
        return []

    if category == "COGS":
        return _parse_cogs(report_data)
    return _parse_gl_detail(report_data, category)


def _parse_cogs(report_data: dict) -> List[TransactionRecord]:
    """
    COGS report (537) parsing:
      - Section header rows give us "Account (Line) context" via their `value`
      - Detail rows have detailLineValues with: Type, Date, Document Number, Name,
        Entity (Line), Clr, Split, Amount, Memo, Message, Account (Line): Name,
        Department: Name, Accounting Period: Name
      - Negate Amount (report shows expenses as negative)
      - Filter to ACCOUNT_FILTERS["COGS"]
    """
    out: List[TransactionRecord] = []
    current_account_section: Tuple[str, str] = ("", "")  # (number, name)

    rows_in_order = sorted(report_data.items(), key=lambda kv: int(kv[0]))
    for _, row in rows_in_order:
        is_detail = row.get("isDetailLine", False)
        if not is_detail:
            label = row.get("value") or row.get("label") or ""
            num, name = _parse_leaf_account(label)
            if num:
                current_account_section = (num, name)
            continue

        dlv = _values_to_dict(row.get("detailLineValues", []))
        rtype = dlv.get("Type")
        if not rtype or str(rtype).strip().lower() == "none":
            continue

        # Account: prefer the GL-style column, fall back to section header.
        # GL columns are parent\x01leaf - parse the leaf.
        acct_field = (dlv.get("Account (Line): Name (GL-style)")
                      or dlv.get("Account (Line): Name")
                      or dlv.get("Account (Line)") or "")
        acct_number, acct_name = _parse_leaf_account(acct_field)
        if not acct_number and current_account_section[0]:
            acct_number, acct_name = current_account_section

        if acct_number not in ACCOUNT_FILTERS["COGS"]:
            continue

        amount = _negate(_to_float(dlv.get("Amount")))
        if amount == 0:
            continue

        vendor = _coalesce(dlv.get("Entity (Line)"), dlv.get("Name"))
        date_iso = _parse_date(dlv.get("Date"))
        period = _period_yyyymm(dlv.get("Accounting Period: Name"), date_iso)
        dept = (dlv.get("Department: Name") or "").strip()

        out.append(TransactionRecord(
            category="COGS",
            period_yyyymm=period,
            date=date_iso,
            vendor=vendor,
            account_number=acct_number,
            account_name=acct_name,
            department=dept,
            amount=amount,
        ))
    return out


def _parse_gl_detail(report_data: dict, category: str) -> List[TransactionRecord]:
    """
    GL Detail reports (540/542/721) — share the same column layout.
      - detailLineValues columns: Account, Type, Date, Document Number, Name,
        Entity (Line): Name, "" (unnamed = net amount), Memo, Message,
        Department: Name, Subsidiary: Name, Accounting Period: Name, Balance
      - Net column "" is the signed amount (positive=Debit, negative=Credit)
      - For expenses, debits increase, credits decrease — so net amount IS the
        period expense impact (no negation needed).
    """
    out: List[TransactionRecord] = []
    allowed = ACCOUNT_FILTERS[category]
    rows_in_order = sorted(report_data.items(), key=lambda kv: int(kv[0]))

    for _, row in rows_in_order:
        if not row.get("isDetailLine", False):
            continue
        dlv = _values_to_dict(row.get("detailLineValues", []))
        rtype = dlv.get("Type")
        if not rtype or str(rtype).strip().lower() == "none":
            continue

        # Account is parent\x01leaf; leaf account is what we filter on.
        acct_field = dlv.get("Account") or ""
        acct_number, acct_name = _parse_leaf_account(acct_field)
        if acct_number not in allowed:
            continue

        # NetSuite GL Detail returns net under '_total'; older versions used '' or 'Amount'.
        amount = _to_float(dlv.get("_total") or dlv.get("") or dlv.get("Amount"))
        if amount == 0:
            continue

        vendor = _coalesce(dlv.get("Entity (Line): Name"), dlv.get("Name"))
        date_iso = _parse_date(dlv.get("Date"))
        period = _period_yyyymm(dlv.get("Accounting Period: Name"), date_iso)
        dept = (dlv.get("Department: Name") or "").strip()

        out.append(TransactionRecord(
            category=category,
            period_yyyymm=period,
            date=date_iso,
            vendor=vendor,
            account_number=acct_number,
            account_name=acct_name,
            department=dept,
            amount=amount,
        ))
    return out


def _values_to_dict(line_values) -> dict:
    """detailLineValues is a list of {columnName: value} singletons. Flatten."""
    out: dict = {}
    if not line_values:
        return out
    if isinstance(line_values, dict):
        return line_values
    for item in line_values:
        if isinstance(item, dict):
            out.update(item)
    return out


def _parse_account_label(label) -> Tuple[str, str]:
    """
    "671100 Software Subscriptions"          -> ("671100", "Software Subscriptions")
    "511400 - COGS - Hosting"                 -> ("511400", "COGS - Hosting")
    "Account 671100\x01Software Subscript..."-> ("671100", "Software Subscriptions")
    Returns ("", "") on no match.
    """
    if not label:
        return ("", "")
    s = str(label).replace("\x01", ":").strip()
    m = re.search(r"\b(\d{4,6})\b\s*[-:]?\s*(.*)$", s)
    if m:
        return (m.group(1), m.group(2).strip(" -:"))
    return ("", s)


def _parse_leaf_account(label) -> Tuple[str, str]:
    """
    NetSuite GL Detail report 'Account' is parent\x01leaf. Always parse the leaf:
      "511001 - COGS - Salary\x01511370 - COGS - Contractor Payroll"
        -> ("511370", "COGS - Contractor Payroll")
      "671100 Software Subscriptions"  (no parent)  -> ("671100", "Software Subscriptions")
    """
    if not label:
        return ("", "")
    s = str(label)
    parts = s.split("\x01")
    leaf = parts[-1].strip()
    return _parse_account_label(leaf)


def _coalesce(*args) -> str:
    for v in args:
        if v is None:
            continue
        s = str(v).strip()
        if s and s != "- No Entity -":
            return s
    return ""


def _to_float(v) -> float:
    if v is None or v == "":
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        # NetSuite sometimes gives strings like "$X,XXX.XX" or "($X,XXX.XX)"
        s = str(v).strip().replace(",", "").replace("$", "")
        if s.startswith("(") and s.endswith(")"):
            s = "-" + s[1:-1]
        try:
            return float(s)
        except (TypeError, ValueError):
            return 0.0


def _negate(v: float) -> float:
    return -v if v else 0.0


def _parse_date(v) -> Optional[str]:
    """Accept 'Mon, 01 Apr 2026 00:00:00 GMT' or ISO. Return YYYY-MM-DD or None."""
    if v is None or v == "":
        return None
    s = str(v).strip()
    # ISO already?
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    # RFC 2822-ish ("Mon, 01 Apr 2026 00:00:00 GMT")
    try:
        d = dt.datetime.strptime(s[:25], "%a, %d %b %Y %H:%M:%S")
        return d.strftime("%Y-%m-%d")
    except ValueError:
        pass
    # 'M/D/YYYY'
    m2 = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
    if m2:
        mm, dd, yyyy = m2.groups()
        return f"{int(yyyy):04d}-{int(mm):02d}-{int(dd):02d}"
    return None


_MONTH_TO_NUM = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


def _period_yyyymm(period_label, date_iso: Optional[str]) -> str:
    """
    Accounting Period: Name comes through as 'Apr 2026' or sometimes '2026-04'.
    Fall back to deriving from the row date.
    """
    if period_label:
        s = str(period_label).strip()
        m = re.match(r"^(\w{3,9})\s+(\d{4})$", s)
        if m:
            mname = m.group(1)[:3].lower()
            if mname in _MONTH_TO_NUM:
                return f"{int(m.group(2)):04d}-{_MONTH_TO_NUM[mname]:02d}"
        m2 = re.match(r"^(\d{4})-(\d{2})", s)
        if m2:
            return f"{m2.group(1)}-{m2.group(2)}"
    if date_iso:
        return date_iso[:7]
    return ""


def load_cache_dir(cache_dir: Path) -> List[TransactionRecord]:
    """
    Load all 4 reports from a month's _cache/ directory. Each file should be
    named '{report_id}.json' (e.g., '537.json'). Returns merged transaction list.
    """
    out: List[TransactionRecord] = []
    for report_id in CATEGORY_BY_REPORT_ID:
        path = cache_dir / f"{report_id}.json"
        if not path.exists():
            continue
        with open(path, "r", encoding="utf-8") as f:
            payload = json.load(f)
        out.extend(parse_report_json(report_id, payload))
    return out


def filter_by_period(records: Iterable[TransactionRecord], yyyymm: str) -> List[TransactionRecord]:
    return [r for r in records if r.period_yyyymm == yyyymm]


def aggregate_vendor_period(records: Iterable[TransactionRecord]) -> Dict[Tuple[str, str, str, str], float]:
    """
    Group by (period, category, normalized_vendor, account_number) -> summed amount.
    Used to compute monthly actuals per vendor for the budget-vs-actual diff.
    """
    from load_vendor_budget import _normalize_name  # local import to share normalizer
    bucket: Dict[Tuple[str, str, str, str], float] = {}
    for r in records:
        key = (r.period_yyyymm, r.category, _normalize_name(r.vendor), r.account_number)
        bucket[key] = bucket.get(key, 0.0) + r.amount
    return bucket
