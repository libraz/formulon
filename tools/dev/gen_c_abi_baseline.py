#!/usr/bin/env python3
"""Regenerates the released C ABI baseline from a git tag.

The baseline (`c_abi_baseline.txt`) is the surface a third-party consumer
compiled against, which is what makes a deletion or rename of a base entry
point a real break rather than a refactor. It is generated rather than
hand-written so it cannot drift into "whatever the header says today", which
would make the guard vacuous.

It holds one generation: the most recently released tag. The file name is
fixed so that cutting a release does not require editing the checker, and
the tag it was taken from is named in its first line. Rebase it *after* the
tag exists, as the last step of a release -- at which point every entry in
`c_abi_breaks.txt` that the new baseline absorbs can be dropped, since the
ledger only has to explain breaks against the current baseline.

Run this only to rebase onto a newly released tag, never to silence a
failing check -- a deliberate break belongs in `c_abi_breaks.txt`.

Usage:
  python3 tools/dev/gen_c_abi_baseline.py [tag]      # default: v0.10.0
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from check_binding_drift import c_abi_declarations  # noqa: E402

HERE = Path(__file__).resolve().parent
DEFAULT_TAG = "v0.10.0"

PREAMBLE = """# Released C ABI surface, as shipped in {tag}.
#
# This is the compatibility commitment a third-party C/C++ consumer built
# against: every entry point below existed in the released header and in the
# shared library's export list. Nothing in the repo consumes several of them
# -- five families are reached only through their `_ex` variant -- so deleting
# or renaming one leaves all four surfaces green and the break ships silently.
# That is the gap this file closes.
#
# Regenerate: tools/dev/gen_c_abi_baseline.py <tag> (reads the tag, writes this
# file). One generation only -- rebasing onto a newer tag replaces it.
# Do NOT hand-edit to make a check pass. A deliberate break is recorded in
# tools/dev/c_abi_breaks.txt instead, so it stays visible in review and can be
# carried into the CHANGELOG and the migration notes.
#
# Format: one declaration per line, `<return-type> <name>(<param-types>)`,
# parameter names stripped, sorted by name. Struct *layout* is not covered
# here -- a by-value struct that keeps its name but changes size is a calling
# convention change this file cannot see. Those are pinned separately by the
# `static_assert(sizeof(...))` tripwires in the C ABI tests.
"""


def main(argv: list[str]) -> int:
    tag = argv[0] if argv else DEFAULT_TAG
    source = subprocess.run(
        ["git", "show", f"{tag}:src/c_api/formulon_c.h"],
        capture_output=True,
        text=True,
        check=True,
    ).stdout
    declarations = c_abi_declarations(source)
    lines = [
        f"{ret} {name}({', '.join(params) if params else 'void'})"
        for name, (ret, params) in sorted(declarations.items())
    ]
    out = HERE / "c_abi_baseline.txt"
    out.write_text(PREAMBLE.format(tag=tag) + "\n" + "\n".join(lines) + "\n", encoding="utf-8")
    print(f"gen_c_abi_baseline: wrote {len(lines)} declarations to {out.relative_to(HERE.parents[1])}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
