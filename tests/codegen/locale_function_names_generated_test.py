#!/usr/bin/env python3
"""Tests for the localized function-name table generator."""

from __future__ import annotations

import contextlib
import importlib.util
import io
import sys
import tempfile
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
GENERATOR_PATH = REPO_ROOT / "tools" / "locale" / "generate_function_names.py"


def _load_generator():
    spec = importlib.util.spec_from_file_location("formulon_locale_function_names", GENERATOR_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {GENERATOR_PATH}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


class LocaleFunctionNamesGeneratedTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.generator = _load_generator()
        cls.rendered, cls.gaps = cls.generator.outputs()

    def test_outputs_match_tracked_files(self) -> None:
        for path, text in self.rendered.items():
            with self.subTest(path=path.name):
                self.assertEqual(text, path.read_text(encoding="utf-8"))

    def test_every_catalog_function_has_a_captured_name(self) -> None:
        self.assertEqual(self.gaps, {}, "capture the missing names into tools/locale/data")

    def test_check_rejects_drifted_temporary_fixture(self) -> None:
        source = self.generator.SOURCE_PATH
        with tempfile.TemporaryDirectory(prefix="formulon-fn-names-") as directory:
            fixture = Path(directory) / source.name
            fixture.write_bytes(source.read_bytes() + b"// drift\n")
            with contextlib.redirect_stderr(io.StringIO()):
                self.assertEqual(self.generator.check({source: fixture}), 1)


if __name__ == "__main__":
    unittest.main()
