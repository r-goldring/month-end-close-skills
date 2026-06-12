"""
strip_pivot_cache.py

Force Excel to rebuild every pivot cache the next time the workbook opens.

openpyxl preserves PivotTable XML well but cannot regenerate
`pivotCacheRecords*.xml` to match new data sizes. The result is that Excel
either (a) crashes when copying the tab to another workbook because the cache
and data are out-of-sync, or (b) shows stale rows that no longer exist in the
data sheet.

Fix: replace every `pivotCacheRecords*.xml` with an empty stub and set
`refreshOnLoad="1"` + `recordCount="0"` on each `pivotCacheDefinition*.xml`.
Excel re-reads the data range from the worksheetSource on the pivot's next
refresh (which happens automatically on open when `refreshOnLoad="1"`) and
builds a fresh cache from scratch.

Idempotent. Safe to call on an xlsx with no pivots (it becomes a no-op).

Usage:
    from strip_pivot_cache import strip_pivot_cache
    wb.save(path)                # openpyxl save
    strip_pivot_cache(path)      # then strip caches in-place
"""
from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from pathlib import Path

_CACHE_RECORDS_RE = re.compile(r"^xl/pivotCache/pivotCacheRecords(\d+)\.xml$")
_CACHE_DEFINITION_RE = re.compile(r"^xl/pivotCache/pivotCacheDefinition\d+\.xml$")
_RECORD_COUNT_RE = re.compile(rb'recordCount="\d+"')
_REFRESH_ON_LOAD_RE = re.compile(rb'refreshOnLoad="[01]"')

_EMPTY_RECORDS_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<pivotCacheRecords '
    b'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    b'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    b'count="0"/>'
)


def _patch_definition_xml(content: bytes) -> bytes:
    """Force recordCount="0" and refreshOnLoad="1" on a pivotCacheDefinition."""
    # recordCount: set to 0 (cache is empty, will be rebuilt on next refresh)
    if _RECORD_COUNT_RE.search(content):
        content = _RECORD_COUNT_RE.sub(b'recordCount="0"', content)
    else:
        content = re.sub(
            rb"(<pivotCacheDefinition\b[^>]*?)(/?>)",
            rb'\1 recordCount="0"\2',
            content,
            count=1,
        )

    # refreshOnLoad: force to 1
    if _REFRESH_ON_LOAD_RE.search(content):
        content = _REFRESH_ON_LOAD_RE.sub(b'refreshOnLoad="1"', content)
    else:
        content = re.sub(
            rb"(<pivotCacheDefinition\b[^>]*?)(/?>)",
            rb'\1 refreshOnLoad="1"\2',
            content,
            count=1,
        )

    return content


def strip_pivot_cache(xlsx_path: str | Path) -> dict:
    """
    Empty every pivot cache so Excel rebuilds from the worksheetSource on open.
    In-place modification. Returns a summary dict.
    """
    src = Path(xlsx_path)
    if not src.exists():
        raise FileNotFoundError(src)

    records_emptied = 0
    definitions_patched = 0

    with tempfile.NamedTemporaryFile(
        suffix=".xlsx", dir=src.parent, delete=False
    ) as tmp:
        tmp_path = Path(tmp.name)

    try:
        with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(
            tmp_path, "w", zipfile.ZIP_DEFLATED
        ) as zout:
            for item in zin.infolist():
                name = item.filename

                if _CACHE_RECORDS_RE.match(name):
                    # Replace with empty stub; keep the file path so rels stay valid
                    zout.writestr(item, _EMPTY_RECORDS_XML)
                    records_emptied += 1
                    continue

                data = zin.read(name)

                if _CACHE_DEFINITION_RE.match(name):
                    data = _patch_definition_xml(data)
                    definitions_patched += 1

                zout.writestr(item, data)

        shutil.move(str(tmp_path), str(src))
    finally:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass

    return {
        "records_emptied": records_emptied,
        "definitions_patched": definitions_patched,
        "had_pivots": (records_emptied + definitions_patched) > 0,
    }


if __name__ == "__main__":
    import sys

    if len(sys.argv) != 2:
        print("Usage: python strip_pivot_cache.py <path/to/file.xlsx>")
        sys.exit(2)
    result = strip_pivot_cache(sys.argv[1])
    print(result)
