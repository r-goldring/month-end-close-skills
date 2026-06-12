"""
Load and validate the BillFlow TransactionListByExportStatus CSV.

the accountant exports this monthly to:
  Monthly Flux Analysis/{YYYY}/{YYYY-MM}/BillFlow Export/TransactionListByExportStatus*.csv

The export is filtered to "not yet exported to NetSuite" but the accountant asked us to
double-check by querying NetSuite for already-posted vendor bills in the period.

CSV columns (confirmed from 2026-04 export):
  Approval Status, Type, Date, Due Date, Vendor Name, Invoice Number,
  Payment Reference, Amount, Allow Export

Filter rules on ingest:
  - Type == "Bill"
  - Amount > 0  (drop payments and credits)
"""

from __future__ import annotations

import csv
import datetime as dt
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple


@dataclass
class BillcomBill:
    approval_status: str
    date: str               # YYYY-MM-DD
    due_date: str
    vendor_name: str
    invoice_number: str
    amount: float
    already_in_netsuite: bool = False
    netsuite_tranid: str = ""    # filled if already_in_netsuite

    @property
    def vendor_normalized(self) -> str:
        from load_vendor_budget import _normalize_name
        return _normalize_name(self.vendor_name)


def find_billcom_csv(month_dir: Path) -> Optional[Path]:
    """
    Look for any TransactionListByExportStatus*.csv under {month_dir}/BillFlow Export/.
    Returns None if the folder or file is absent (skill skips BillFlow merge silently).
    """
    folder = month_dir / "BillFlow Export"
    if not folder.exists():
        return None
    candidates = sorted(folder.glob("TransactionListByExportStatus*.csv"),
                        key=lambda p: p.stat().st_mtime, reverse=True)
    if not candidates:
        candidates = sorted(folder.glob("*.csv"),
                            key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[0] if candidates else None


def load_billcom_csv(path: Path) -> List[BillcomBill]:
    """Load the CSV, filter to Type='Bill' AND Amount > 0."""
    bills: List[BillcomBill] = []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            bill_type = (row.get("Type") or "").strip()
            if bill_type != "Bill":
                continue
            amt = _to_float(row.get("Amount"))
            if amt <= 0:
                continue
            bills.append(BillcomBill(
                approval_status=(row.get("Approval Status") or "").strip(),
                date=_parse_date(row.get("Date")),
                due_date=_parse_date(row.get("Due Date")),
                vendor_name=(row.get("Vendor Name") or "").strip(),
                invoice_number=(row.get("Invoice Number") or "").strip(),
                amount=amt,
            ))
    return bills


def mark_already_in_netsuite(bills: List[BillcomBill],
                             ns_vendor_bills: Iterable[dict]) -> Tuple[List[BillcomBill], List[BillcomBill]]:
    """
    Cross-reference BillFlow bills against NS VendorBills posted in the period.
    ns_vendor_bills: each dict must have keys 'tranid' and 'companyname' (vendor display).

    Match rule: tranid matches BillFlow Invoice Number AND vendor names normalize-equal.
    Returns (open_bills, already_in_ns).
    """
    from load_vendor_budget import _normalize_name
    by_tranid: Dict[str, List[dict]] = {}
    for nb in ns_vendor_bills:
        tid = (nb.get("tranid") or "").strip()
        if not tid:
            continue
        by_tranid.setdefault(tid, []).append(nb)

    open_bills: List[BillcomBill] = []
    already: List[BillcomBill] = []

    for b in bills:
        match = None
        for cand in by_tranid.get(b.invoice_number, []):
            if _normalize_name(cand.get("companyname") or "") == b.vendor_normalized:
                match = cand
                break
        if match is not None:
            b.already_in_netsuite = True
            b.netsuite_tranid = match.get("tranid") or b.invoice_number
            already.append(b)
        else:
            open_bills.append(b)

    return open_bills, already


def aggregate_by_vendor(bills: Iterable[BillcomBill]) -> Dict[str, dict]:
    """
    Sum BillFlow amounts by normalized vendor name. Returns:
      {normalized_name: {"vendor": display_name,
                          "total": float,
                          "invoice_numbers": [str, ...]}}
    """
    out: Dict[str, dict] = {}
    for b in bills:
        key = b.vendor_normalized
        if not key:
            continue
        slot = out.setdefault(key, {
            "vendor": b.vendor_name,
            "total": 0.0,
            "invoice_numbers": [],
        })
        slot["total"] += b.amount
        slot["invoice_numbers"].append(b.invoice_number)
    return out


def _to_float(v) -> float:
    if v is None or v == "":
        return 0.0
    s = str(v).strip().replace(",", "").replace("$", "")
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return float(s)
    except (TypeError, ValueError):
        return 0.0


def _parse_date(v) -> str:
    """BillFlow gives M/D/YY or M/D/YYYY. Return YYYY-MM-DD."""
    if v is None or v == "":
        return ""
    s = str(v).strip()
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2,4})$", s)
    if m:
        mm, dd, yy = m.groups()
        year = int(yy)
        if year < 100:
            year += 2000
        return f"{year:04d}-{int(mm):02d}-{int(dd):02d}"
    # Already ISO?
    m2 = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m2:
        return s[:10]
    return s
