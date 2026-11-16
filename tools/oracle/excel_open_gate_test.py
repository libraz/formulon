#!/usr/bin/env python3
"""Tests for the open gate's check that formulon_cli is not older than the sources it was built from."""

from __future__ import annotations

import os
import sys
import tempfile
import types
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    import xlwings  # noqa: F401
except Exception:  # noqa: BLE001 - the check under test never touches Excel
    sys.modules["xlwings"] = types.ModuleType("xlwings")
import excel_open_gate as gate  # noqa: E402


def _touch(path: Path, mtime: float) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("x")
    os.utime(path, (mtime, mtime))
    return path


class StaleCliTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.base = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_export_root_passes_while_the_repo_has_a_newer_edit(self) -> None:
        export = self.base / "export"
        repo = self.base / "repo"
        _touch(export / "src" / "a.cpp", 1000)
        cli = _touch(self.base / "formulon_cli", 2000)
        _touch(repo / "src" / "a.cpp", 3000)
        self.assertIsNone(gate.stale_cli(cli, export))
        self.assertIn("older than", gate.stale_cli(cli, repo))

    def test_root_with_a_file_newer_than_the_cli_fails(self) -> None:
        root = self.base / "root"
        _touch(root / "src" / "a.cpp", 1000)
        newer = _touch(root / "src" / "sub" / "b.h", 3000)
        cli = _touch(self.base / "formulon_cli", 2000)
        reason = gate.stale_cli(cli, root)
        self.assertIsNotNone(reason)
        self.assertIn(str(newer), reason)

    def test_missing_cli_or_sources_fail(self) -> None:
        root = self.base / "root"
        _touch(root / "src" / "a.cpp", 1000)
        self.assertIn("not found", gate.stale_cli(self.base / "absent", root))
        cli = _touch(self.base / "formulon_cli", 2000)
        self.assertIn("no sources", gate.stale_cli(cli, self.base / "empty"))


if __name__ == "__main__":
    unittest.main()
