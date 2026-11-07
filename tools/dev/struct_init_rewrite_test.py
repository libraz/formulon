#!/usr/bin/env python3
"""Fixture tests for struct_init_rewrite."""

from __future__ import annotations

import shutil
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from tools.dev import struct_init_rewrite as sir

FIX = Path(__file__).parent / "struct_init_rewrite_fixtures"


class StructInitRewriteTest(unittest.TestCase):
    def setUp(self) -> None:
        self.types = sir.record_types((FIX / "header.h").read_text())

    def test_record_set_excludes_value_types(self) -> None:
        self.assertEqual(self.types, {"fm_rec_t", "fm_other_t"})

    def test_empty_inits_rewritten_value_types_skipped(self) -> None:
        new, _ = sir.scan((FIX / "input.cpp").read_text(), self.types)
        self.assertEqual(new, (FIX / "expected.cpp").read_text())

    def test_non_empty_listed(self) -> None:
        _, manual = sir.scan((FIX / "input.cpp").read_text(), self.types)
        self.assertEqual(
            [(m.line, m.kind) for m in manual],
            [(7, "positional"), (8, "designated"), (9, "array"), (10, "array"), (12, "temporary")],
        )

    def test_idempotent(self) -> None:
        once, _ = sir.scan((FIX / "input.cpp").read_text(), self.types)
        twice, _ = sir.scan(once, self.types)
        self.assertEqual(once, twice)

    def test_cli_check_and_write(self) -> None:
        with TemporaryDirectory() as d:
            f = Path(d) / "input.cpp"
            shutil.copy(FIX / "input.cpp", f)
            hdr = str(FIX / "header.h")
            self.assertEqual(sir.main(["--check", "--header", hdr, str(f)]), 1)
            self.assertEqual(sir.main(["--write", "--header", hdr, str(f)]), 0)
            self.assertEqual(f.read_text(), (FIX / "expected.cpp").read_text())
            self.assertEqual(sir.main(["--check", "--header", hdr, str(f)]), 0)


if __name__ == "__main__":
    unittest.main()
