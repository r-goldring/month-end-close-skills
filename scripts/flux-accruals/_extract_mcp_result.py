"""
Helper: extract the embedded NetSuite ns_runReport JSON payload from a saved
MCP tool-results file (when the MCP wrapper truncates due to size) and write
it to the flux-accruals cache directory.

Usage:
    python _extract_mcp_result.py <mcp_result_file.txt> <out_path.json>
"""
import json
import sys
from pathlib import Path


def main():
    if len(sys.argv) != 3:
        print("Usage: _extract_mcp_result.py <src.txt> <dst.json>")
        sys.exit(2)
    src = Path(sys.argv[1])
    dst = Path(sys.argv[2])
    with open(src, "r", encoding="utf-8") as f:
        wrapper = json.load(f)
    if not isinstance(wrapper, list) or not wrapper:
        raise ValueError(f"Unexpected MCP wrapper structure: {type(wrapper)}")
    inner_text = wrapper[0].get("text") or wrapper[0].get("value")
    if not inner_text:
        raise ValueError(f"No 'text' field in MCP wrapper element: {list(wrapper[0].keys())}")
    payload = json.loads(inner_text)
    dst.parent.mkdir(parents=True, exist_ok=True)
    with open(dst, "w", encoding="utf-8") as f:
        json.dump(payload, f)
    rows = len(payload.get("reportData", payload.get("rows", [])))
    print(f"OK -> {dst}")
    print(f"  reportData rows: {rows}")


if __name__ == "__main__":
    main()
