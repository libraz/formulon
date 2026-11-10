#!/usr/bin/env python3
"""Fixture tests for the tristate-reads check.

Run from the repo root: `python3 -m tools.dev.binding_drift.tristate_reads_test`.
"""

from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from tools.dev.binding_drift.tristate_reads import check_tristate_reads, tristate_fields, unmarked_tristate_fields

HEADER = """
typedef struct fm_page_break_t { uint32_t row; int32_t manual; /* fm_tristate_t */ } fm_page_break_t;
typedef struct fm_cfvo_t { int32_t gte; /* fm_tristate_t */ int32_t hidden; } fm_cfvo_t;
typedef struct fm_other_t { int32_t hidden; /* fm_tristate_t */ } fm_other_t;
"""

STRUCTS = """
CFVO = Struct("fm_cfvo_t", [("gte", I32), ("hidden", I32)])
OTHER = Struct(
    "fm_other_t",
    [("hidden", I32)],
)
"""

AMBIGUOUS_PY = """
def plain(lib, ptr):
    d = S.CFVO.unpack(lib, ptr)
    return bool(d["hidden"])


def tristate(lib, ptr):
    d = S.OTHER.unpack(lib, ptr)
    return bool(d["hidden"])


def documented(d):
    \"\"\"Decode an ``fm_other_t``.\"\"\"
    return bool(d["hidden"])


def unresolved(attrs):
    return bool(attrs["hidden"])
"""

AMBIGUOUS_GLUE = """
void Plain(const fm_cfvo_t& c) {
  o.set("hidden", c.hidden != 0);
}
void Tristate(const fm_other_t& c) {
  o.set("hidden", c.hidden != 0);
}
void Both(const fm_cfvo_t& a, const fm_other_t& b) {
  o.set("hidden", a.hidden != 0);
}
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

    def test_a_name_plain_elsewhere_is_checked_where_only_its_tristate_struct_is_in_scope(self) -> None:
        with TemporaryDirectory() as tmp:
            py, glue = Path(tmp) / "py", Path(tmp) / "glue"
            py.mkdir()
            glue.mkdir()
            (py / "_structs.py").write_text(STRUCTS, encoding="utf-8")
            (py / "m.py").write_text(AMBIGUOUS_PY, encoding="utf-8")
            (glue / "g.cpp").write_text(AMBIGUOUS_GLUE, encoding="utf-8")
            problems = check_tristate_reads(HEADER, [(glue, False), (py, True)])
        lines = sorted((p.split(":")[1].rsplit("/", 1)[-1], int(p.split(":")[2])) for p in problems)
        self.assertEqual(lines, [("g.cpp", 6), ("m.py", 9), ("m.py", 14)])

    def test_a_field_the_c_implementation_treats_as_tristate_must_be_marked(self) -> None:
        impl = """
out->manual = brk.manual ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
out->subtotal_top = tristate_of(field.subtotal_top);
if (spec->visible == FM_TRISTATE_FALSE) {}
bool on = tristate_to_bool(spec->styled);
out->count = 3;
"""
        self.assertEqual(
            unmarked_tristate_fields(HEADER, [("x.cpp", impl)]),
            ["x.cpp:3: subtotal_top", "x.cpp:4: visible", "x.cpp:5: styled"],
        )


if __name__ == "__main__":
    unittest.main()
