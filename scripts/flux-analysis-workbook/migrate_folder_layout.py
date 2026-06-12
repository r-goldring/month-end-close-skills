"""
migrate_folder_layout.py

Reorganize a Monthly Flux Analysis/{YYYY}/{YYYY-MM}/ folder into the v2 layout:

    {YYYY-MM}/
    ├── Flux Workbook/         (IS/BS drops, flux analysis xlsx versions, transcripts)
    ├── JE Imports/            (NetSuite CSV uploads)
    ├── Schedules/             (prepaid items, accrual candidates, support xlsx)
    ├── State/                 (auto-managed JSON: flux_notes, accrual_candidates, pending_jes)
    ├── _cache/                (NetSuite report JSON)
    └── Supporting/            (BillFlow export, other team inputs)

Classification rules (first match wins):

| Pattern                                                | Destination       |
|--------------------------------------------------------|-------------------|
| name endswith ".csv" AND contains "JE Import"          | JE Imports/       |
| name == "_cache" (dir)                                 | _cache/           |
| name endswith "_notes.json" / "accrual_candidates.json"| State/            |
| name endswith "pending_jes.json"                       | State/            |
| contains "Flux Analysis" (xlsx)                        | Flux Workbook/    |
| contains "Flux Pivot Template" (xlsx)                  | Flux Workbook/    |
| contains "Income Statement" OR "Balance Sheet" (xlsx)  | Flux Workbook/    |
| contains "Notes by Gemini" OR "Transcript" (md/txt)    | Flux Workbook/    |
| contains "Prepaid Schedule"                            | Schedules/        |
| contains "Accrual_Reclass_Candidates"                  | Schedules/        |
| starts with "Accrual Support Workbook"                 | Schedules/        |
| name in {BillFlow Export, Mar-26 Software Reclasses}   | Supporting/       |
| anything else                                          | Supporting/       |

Dry-run by default. Use --apply to actually move files.

Idempotent: re-running on an already-migrated folder is a no-op.

Usage:
    python migrate_folder_layout.py 2026-04                # dry run
    python migrate_folder_layout.py 2026-04 --apply        # do it
    python migrate_folder_layout.py 2026-03 2026-04 --apply
"""
from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / "Monthly Flux Analysis"

NEW_DIRS = ["Flux Workbook", "JE Imports", "Schedules", "State", "_cache", "Supporting"]


def classify(item: Path) -> str:
    """
    Return the destination subfolder name for an item (file or dir) sitting at
    the month-root level. Returns "" to mean "leave it alone" (already
    classified, or it IS a destination dir).
    """
    name = item.name
    lower = name.lower()

    # Skip Excel temp lock files
    if name.startswith("~$"):
        return ""

    # Already a destination dir
    if item.is_dir() and name in NEW_DIRS:
        return ""

    # _cache stays where it is (it's already its own dir)
    if name == "_cache":
        return ""

    # Existing supporting subfolders -> Supporting/
    if item.is_dir() and name in ("BillFlow Export", "Updated FS", "Updated Support",
                                  "Accruals Support", "Mar-26 Software Reclasses"):
        return "Supporting"

    # JSON state files
    if name.endswith(".json") and any(
        s in lower for s in ("flux_notes", "accrual_candidates", "pending_jes")
    ):
        return "State"

    # JE Imports
    if name.endswith(".csv") and "je import" in lower:
        return "JE Imports"

    # Flux Workbook contents
    if any(s in name for s in ("Flux Analysis", "Flux Pivot Template",
                                "Income Statement", "Balance Sheet")):
        return "Flux Workbook"
    if name.endswith((".md", ".txt")) and ("notes by gemini" in lower
                                             or "transcript" in lower):
        return "Flux Workbook"

    # Schedules
    if "prepaid schedule" in lower:
        return "Schedules"
    if "accrual_reclass_candidates" in lower or "accrual support" in lower:
        return "Schedules"
    if "p&l" in lower or "p&amp;l" in lower:
        return "Schedules"

    # Default: leave non-obvious files in Supporting
    return "Supporting"


def plan_moves(month_dir: Path) -> list[tuple[Path, Path]]:
    """Return a list of (src, dst) moves needed."""
    moves = []
    for item in sorted(month_dir.iterdir()):
        target_subdir = classify(item)
        if not target_subdir:
            continue
        dst = month_dir / target_subdir / item.name
        if item.resolve() == dst.resolve():
            continue
        moves.append((item, dst))
    return moves


def _files_identical(a: Path, b: Path) -> bool:
    """Cheap byte-identical check for two file paths (size first, then content)."""
    if not a.is_file() or not b.is_file():
        return False
    if a.stat().st_size != b.stat().st_size:
        return False
    with open(a, "rb") as fa, open(b, "rb") as fb:
        while True:
            ca = fa.read(8192)
            cb = fb.read(8192)
            if ca != cb:
                return False
            if not ca:
                return True


def execute_moves(moves: list[tuple[Path, Path]], month_dir: Path) -> dict:
    """
    Execute moves, tolerating locked-file failures.
    Returns dict with 'moved', 'skipped_duplicate', 'failed_locked' counts and
    a list of failed source paths.
    """
    # Pre-create all destination dirs (only for moves we'll actually do)
    needed_dirs = {dst.parent for _, dst in moves}
    for d in needed_dirs:
        d.mkdir(parents=True, exist_ok=True)

    moved = 0
    skipped_dup = 0
    failed: list[Path] = []

    for src, dst in moves:
        try:
            # If dst already exists, check if src is identical (left over from
            # a prior aborted run). If yes, just delete src; if no, append (1).
            if dst.exists():
                if src.is_file() and dst.is_file() and _files_identical(src, dst):
                    try:
                        src.unlink()
                        skipped_dup += 1
                        continue
                    except PermissionError:
                        failed.append(src)
                        continue
                # Different content: append (1), (2), ...
                stem = dst.stem
                suffix = dst.suffix
                i = 1
                while True:
                    candidate = dst.with_name(f"{stem} ({i}){suffix}")
                    if not candidate.exists():
                        dst = candidate
                        break
                    i += 1
            shutil.move(str(src), str(dst))
            moved += 1
        except PermissionError:
            failed.append(src)
            continue

    return {
        "moved": moved,
        "skipped_duplicate": skipped_dup,
        "failed_locked": len(failed),
        "failed_paths": [str(p) for p in failed],
    }


def migrate(month_label: str, apply: bool) -> int:
    month_dir = ROOT / month_label[:4] / month_label
    if not month_dir.exists():
        # Try without year prefix
        candidates = list(ROOT.glob(f"*/{month_label}"))
        if not candidates:
            print(f"NOT FOUND: {month_label} (looked in {ROOT})")
            return 1
        month_dir = candidates[0]

    moves = plan_moves(month_dir)
    if not moves:
        print(f"[{month_label}] Already migrated. Nothing to do.")
        return 0

    print(f"[{month_label}] Planned moves ({len(moves)}):")
    for src, dst in moves:
        rel_src = src.relative_to(month_dir)
        rel_dst = dst.relative_to(month_dir)
        print(f"  {rel_src}  ->  {rel_dst}")

    if not apply:
        print(f"\n[{month_label}] Dry run. Re-run with --apply to execute.")
        return 0

    result = execute_moves(moves, month_dir)
    print(
        f"\n[{month_label}] moved={result['moved']}, "
        f"skipped_duplicate={result['skipped_duplicate']}, "
        f"failed_locked={result['failed_locked']}"
    )
    if result["failed_locked"]:
        print("  Locked files (likely open in Excel / another process):")
        for p in result["failed_paths"]:
            print(f"    {p}")
        print("  Close those files and re-run to finish.")
        return 1
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "months", nargs="+", help="Month folders to migrate (e.g. 2026-03 2026-04)"
    )
    parser.add_argument(
        "--apply", action="store_true", help="Actually move files (default: dry run)"
    )
    args = parser.parse_args(argv)

    rc = 0
    for month in args.months:
        rc |= migrate(month, args.apply)
    return rc


if __name__ == "__main__":
    sys.exit(main())
