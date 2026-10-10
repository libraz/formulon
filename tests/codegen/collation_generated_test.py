#!/usr/bin/env python3
"""Tests for the text collation table generator."""

from __future__ import annotations

import contextlib
import copy
import importlib.util
import io
import sys
import tempfile
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
GENERATOR_PATH = REPO_ROOT / "tools" / "collation" / "generate_tables.py"

# Code points each locale's tables approximate instead of reproducing (the BMP outside
# the exact ranges; on zh_CN also the ideographs outside GB 2312 level 1).
EXPECTED_OUT_OF_RANGE = {
    "ja_JP": 26986,
    "en_US": 26986,
    "de_DE": 26986,
    "fr_FR": 26986,
    "zh_CN": 44301,
    "ko_KR": 26986,
    "th_TH": 26986,
    "es_ES": 26986,
    "es_MX": 26986,
    "pt_BR": 26986,
    "ru_RU": 26986,
    "zh_TW": 26986,
    "it_IT": 26986,
    "nl_NL": 26986,
}


def _load_generator():
    spec = importlib.util.spec_from_file_location("formulon_collation_generator", GENERATOR_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {GENERATOR_PATH}")
    module = importlib.util.module_from_spec(spec)
    # The generator's dataclasses resolve their module through sys.modules.
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


class CollationGeneratedTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.generator = _load_generator()
        cls.generated = cls.generator.generate()

    def test_outputs_match_tracked_files(self) -> None:
        with tempfile.TemporaryDirectory(prefix="formulon-collation-") as directory:
            for path, text in self.generated.files.items():
                regenerated = Path(directory) / path.name
                regenerated.write_text(text, encoding="utf-8")
                with self.subTest(path=path.name):
                    self.assertEqual(regenerated.read_bytes(), path.read_bytes())

    def test_out_of_range_counts(self) -> None:
        self.assertEqual(self.generated.out_of_range, EXPECTED_OUT_OF_RANGE)

    def test_overflowing_weight_fails(self) -> None:
        gen = self.generator
        tables = copy.deepcopy(self.generated.tables)
        loc = tables.locales["en_US"]
        loc.overrides[0x41] = (gen.BASE, gen.MAX_WEIGHT + 1, 0, 0, 0)
        with self.assertRaises(SystemExit):
            gen.check_ranges(tables)
        with self.assertRaises(SystemExit):
            gen.weight_bytes(gen.MAX_WEIGHT + 1)
        with self.assertRaises(SystemExit):
            gen.leb128([-1])
        self.assertEqual(gen.weight_bytes(gen.MAX_WEIGHT), b"\xff\xff")

    def test_check_rejects_drifted_temporary_fixture(self) -> None:
        gen = self.generator
        with tempfile.TemporaryDirectory(prefix="formulon-collation-") as directory:
            fixture = Path(directory) / gen.SOURCE_PATH.name
            fixture.write_bytes(gen.SOURCE_PATH.read_bytes() + b"// drift\n")
            with contextlib.redirect_stderr(io.StringIO()):
                self.assertEqual(gen.check({gen.SOURCE_PATH: fixture}, self.generated), 1)
                missing = fixture.with_name("missing.cpp")
                self.assertEqual(gen.check({gen.SOURCE_PATH: missing}, self.generated), 2)
                self.assertEqual(gen.check(None, self.generated), 0)


if __name__ == "__main__":
    unittest.main()
