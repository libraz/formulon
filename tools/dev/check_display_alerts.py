#!/usr/bin/env python3
"""Reject turning Excel `display_alerts` off outside the shared helper.

The property is application-wide and survives the driver's exit, so turning it
off without restoring it silences the user's own Excel dialogs. Drivers use
`alerts_off` from `tools/oracle/drivers/excel_alerts.py`, which restores it.

Stdlib only; run from the repository root.
"""

from __future__ import annotations

import re
import subprocess
import sys

HELPER = "tools/oracle/drivers/excel_alerts.py"
ASSIGNMENT = re.compile(r"^\s*[\w.\[\]]*display_alerts\s*=\s*(?:False|0)\b", re.M)


def tracked_python() -> list[str]:
    out = subprocess.run(["git", "ls-files", "-z", "*.py"], check=True, capture_output=True).stdout
    return [p for p in out.decode("utf-8").split("\0") if p]


def main() -> int:
    hits: list[str] = []
    for path in tracked_python():
        if path == HELPER:
            continue
        try:
            with open(path, encoding="utf-8") as fh:
                text = fh.read()
        except (UnicodeDecodeError, FileNotFoundError):
            continue
        for match in ASSIGNMENT.finditer(text):
            hits.append(f"{path}:{text.count(chr(10), 0, match.start()) + 1}")
    if hits:
        print("check_display_alerts: Excel display_alerts turned off outside the shared helper:")
        for hit in hits:
            print(f"  {hit}")
        print(f"Use `alerts_off(app)` from {HELPER}; it turns alerts back on when the block exits.")
        return 1
    print("check_display_alerts: ok")
    return 0


if __name__ == "__main__":
    sys.exit(main())
