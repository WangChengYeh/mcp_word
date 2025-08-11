#!/usr/bin/env python3
"""
Parse debug.log lines of the form:
    [in/out time] <json>

Outputs human-readable JSON containing one object per line with:
    {
      "direction": "in" | "out",
      "time": "...",
      "header": "[in/out time]",
      "data": <parsed JSON or raw string if parse fails>
    }

Usage:
  python debug.py                 # reads ./debug.log, prints pretty JSON array
  python debug.py -o output.json  # writes pretty JSON array to file
  python debug.py --ndjson        # prints newline-delimited JSON entries
  python debug.py -i path/to.log  # specify a different input file
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

LINE_RE = re.compile(r"^\[(in|out)\s+([^\]]+)\]\s*(.*)$", re.IGNORECASE)


def parse_line(line: str):
    line = line.rstrip("\n")
    if not line.strip():
        return None
    m = LINE_RE.match(line)
    if not m:
        # Not matching the expected prefix; treat entire line as raw data
        return {
            "direction": None,
            "time": None,
            "header": None,
            "data": line.strip(),
        }
    direction, when, payload = m.groups()
    data = _parse_json(payload)
    return {
        "direction": direction.lower(),
        "time": when.strip(),
        "header": f"[{direction} {when}]",
        "data": data,
    }


def _parse_json(s: str):
    s = s.strip()
    if not s:
        return None
    try:
        return json.loads(s)
    except Exception:
        # If it isn't valid JSON, return the raw string for visibility
        return s


def main(argv=None):
    p = argparse.ArgumentParser(description="Pretty-print entries from debug.log")
    p.add_argument("-i", "--input", default="debug.log", help="log file path (default: debug.log)")
    p.add_argument("-o", "--output", default=None, help="output file path (default: stdout)")
    p.add_argument("--ndjson", action="store_true", help="emit newline-delimited JSON entries instead of an array")
    args = p.parse_args(argv)

    in_path = Path(args.input)
    try:
        with in_path.open("r", encoding="utf-8") as f:
            entries = []
            for idx, line in enumerate(f, 1):
                parsed = parse_line(line)
                if parsed is None:
                    continue
                entries.append(parsed)
    except FileNotFoundError:
        print(f"error: input file not found: {in_path}", file=sys.stderr)
        return 1

    if args.output:
        out_fh = Path(args.output).open("w", encoding="utf-8")
        should_close = True
    else:
        out_fh = sys.stdout
        should_close = False

    try:
        if args.ndjson:
            for entry in entries:
                json.dump(entry, out_fh, ensure_ascii=False)
                out_fh.write("\n")
        else:
            json.dump(entries, out_fh, ensure_ascii=False, indent=2)
            if out_fh is sys.stdout:
                out_fh.write("\n")
    finally:
        if should_close:
            out_fh.close()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
