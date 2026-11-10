#!/usr/bin/env python3
"""Fixture tests for the tristate-reads check.

Run from the repo root: `python3 -m tools.dev.binding_drift.tristate_reads_test`.
"""

from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from tools.dev.binding_drift.tristate_reads import check_tristate_reads, tristate_fields

HEADER = """
typedef struct fm_page_break_t { uint32_t row; int32_t manual; /* fm_tristate_t */ } fm_page_break_t;
typedef struct fm_cfvo_t { int32_t gte; /* fm_tristate_t */ int32_t hidden; } fm_cfvo_t;
typedef struct fm_other_t { int32_t hidden; /* fm_tristate_t */ } fm_other_t;
"""

PYTHON = """
ok = S.tristate_flag(brk["manual"])
ok2 = d["manual"] == S.TRISTATE_TRUE
bad = bool(brk["manual"])
bad2 = d["gte"] != 0
"""

GLUE = """
o.set("manual", brk.manual == FM_TRISTATE_TRUE);
out.gte = reader.boolean(v, "gte", true, "cfvo.gte") ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
flag("gte", &p.gte_op, &p.gte);
brk.manual = span.manual;
item.Set("gte", JsTristate(env, v.gte));
o.set("manual", val(brk.manual != 0));
"""


class TristateReadsTest(unittest.TestCase):
    def test_fields_exclude_names_that_are_plain_elsewhere(self) -> None:
        self.assertEqual(tristate_fields(HEADER), {"manual", "gte"})

    def test_truthiness_reads_are_flagged(self) -> None:
        with TemporaryDirectory() as tmp:
            py, glue = Path(tmp) / "py", Path(tmp) / "glue"
            py.mkdir()
            glue.mkdir()
            (py / "m.py").write_text(PYTHON, encoding="utf-8")
            (glue / "g.cpp").write_text(GLUE, encoding="utf-8")
            problems = check_tristate_reads(HEADER, [(glue, False), (py, True)])
        lines = sorted((p.split(":")[1].rsplit("/", 1)[-1], int(p.split(":")[2])) for p in problems)
        self.assertEqual(lines, [("g.cpp", 7), ("m.py", 4), ("m.py", 5)])


if __name__ == "__main__":
    unittest.main()
