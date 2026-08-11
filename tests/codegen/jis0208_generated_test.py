#!/usr/bin/env python3
"""Tests for the JIS X 0208 data-table generator."""

from __future__ import annotations

import contextlib
import importlib.util
import io
import tempfile
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
GENERATOR_PATH = REPO_ROOT / "tools" / "jis0208" / "generate_table.py"


def _load_generator():
    spec = importlib.util.spec_from_file_location("formulon_jis0208_generator", GENERATOR_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {GENERATOR_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class Jis0208GeneratedTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.generator = _load_generator()

    def test_render_matches_tracked_snapshot(self) -> None:
        expected = self.generator.TABLE_PATH.read_bytes()
        self.assertEqual(self.generator.render().encode("utf-8"), expected)

    def test_check_rejects_drifted_temporary_fixture(self) -> None:
        with tempfile.TemporaryDirectory(prefix="formulon-jis0208-") as directory:
            fixture = Path(directory) / "jis0208_table.cpp"
            fixture.write_bytes(self.generator.TABLE_PATH.read_bytes() + b"// drift\n")
            with contextlib.redirect_stderr(io.StringIO()):
                self.assertEqual(self.generator.check(fixture), 1)
                self.assertEqual(self.generator.check(fixture.with_name("missing.cpp")), 2)


if __name__ == "__main__":
    unittest.main()
