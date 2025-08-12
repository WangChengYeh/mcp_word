#!/usr/bin/env python3
"""
Beautify debug.log into structured JSON.

Under the refined logging, debug lines look like:
  [2025-08-12T23:31:22.036Z] [DEBUG stdin] { ... }
  [2025-08-12T23:31:22.036Z] [DEBUG stdout] { ... }
  [2025-08-12T23:31:22.036Z] [DEBUG socket:send] {"event":"...","payload":{...}}
  [2025-08-12T23:31:22.036Z] [DEBUG socket.recv] {"event":"...","payload":{...}}

Older logs may contain legacy frames:
  [in 2025-08-12T23:00:53.414Z] { ... }
  [out 2025-08-12T23:00:53.414Z] { ... }

This script normalizes all of the above into entries like:
  {
    "type": "stdin" | "stdout" | "socket_send" | "socket_recv" | "debug" | "log" | "raw",
    "time": "ISO-8601 timestamp when available",
    "header": "[DEBUG socket.recv]" (when present),
    "event": "event name for socket logs (if present)",
    "payload": { ... } (for socket logs),
    "data": <parsed JSON or raw string>,
    "raw": "original line"
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

LEGACY_INOUT_RE = re.compile(r"^\[(in|out)\s+([^\]]+)\]\s*(.*)$", re.IGNORECASE)
TS_PREFIX_RE = re.compile(r"^\[(\d{4}-\d{2}-\d{2}T[^\]]+)\]\s*(.*)$")
DEBUG_TAG_RE = re.compile(r"^\[(DEBUG [^\]]+)\]\s*(.*)$")


def parse_line(line: str):
    line = line.rstrip("\n")
    if not line.strip():
        return None
    raw = line

    # 1) Timestamped lines: [ISO] <rest>
    m_ts = TS_PREFIX_RE.match(line)
    if m_ts:
        ts, rest = m_ts.groups()
        # 1a) Timestamped DEBUG-tagged lines: [DEBUG ...] payload
        m_dbg = DEBUG_TAG_RE.match(rest)
        if m_dbg:
            tag, payload = m_dbg.groups()
            tag_lower = tag.lower()
            entry = {
                "type": "debug",
                "time": ts,
                "header": f"[{tag}]",
                "data": None,
                "raw": raw,
            }
            parsed = _parse_json(payload)
            entry["data"] = parsed

            # Specialize known tags
            if tag_lower == "debug stdin":
                entry["type"] = "stdin"
            elif tag_lower == "debug stdout":
                entry["type"] = "stdout"
            elif tag_lower == "debug socket:send":
                entry["type"] = "socket_send"
                if isinstance(parsed, dict):
                    entry["event"] = parsed.get("event")
                    entry["payload"] = parsed.get("payload")
            elif tag_lower == "debug socket.recv":
                entry["type"] = "socket_recv"
                if isinstance(parsed, dict):
                    entry["event"] = parsed.get("event")
                    entry["payload"] = parsed.get("payload")
            return entry

        # 1b) Other timestamped lines: treat as generic log
        return {
            "type": "log",
            "time": ts,
            "header": None,
            "message": rest.strip(),
            "raw": raw,
        }

    # 2) Non-timestamped DEBUG-tagged lines (older runs)
    m_dbg = DEBUG_TAG_RE.match(line)
    if m_dbg:
        tag, payload = m_dbg.groups()
        tag_lower = tag.lower()
        entry = {
            "type": "debug",
            "time": None,
            "header": f"[{tag}]",
            "data": _parse_json(payload),
            "raw": raw,
        }
        if tag_lower == "debug stdin":
            entry["type"] = "stdin"
        elif tag_lower == "debug stdout":
            entry["type"] = "stdout"
        elif tag_lower == "debug socket:send":
            entry["type"] = "socket_send"
            if isinstance(entry["data"], dict):
                entry["event"] = entry["data"].get("event")
                entry["payload"] = entry["data"].get("payload")
        elif tag_lower == "debug socket.recv":
            entry["type"] = "socket_recv"
            if isinstance(entry["data"], dict):
                entry["event"] = entry["data"].get("event")
                entry["payload"] = entry["data"].get("payload")
        return entry

    # 3) Legacy [in/out time] payload lines
    m_legacy = LEGACY_INOUT_RE.match(line)
    if m_legacy:
        direction, when, payload = m_legacy.groups()
        data = _parse_json(payload)
        typ = "stdin" if direction.lower() == "in" else "stdout"
        return {
            "type": typ,
            "time": when.strip(),
            "header": f"[{direction} {when}]",
            "data": data,
            "raw": raw,
        }

    # 4) Fallback: raw message
    return {
        "type": "raw",
        "time": None,
        "header": None,
        "data": line.strip(),
        "raw": raw,
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
