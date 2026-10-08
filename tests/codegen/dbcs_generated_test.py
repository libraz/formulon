#!/usr/bin/env python3
"""Tests for the DBCS and Mac Thai code-page table generator."""

from __future__ import annotations

import contextlib
import importlib.util
import io
import sys
import tempfile
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
GENERATOR_PATH = REPO_ROOT / "tools" / "dbcs" / "generate_tables.py"


def _load_generator():
    spec = importlib.util.spec_from_file_location("formulon_dbcs_generator", GENERATOR_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {GENERATOR_PATH}")
    module = importlib.util.module_from_spec(spec)
    # The generator's dataclass resolves its module through sys.modules.
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


class DbcsGeneratedTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.generator = _load_generator()

    def test_outputs_match_tracked_files(self) -> None:
        sbcs = self.generator.SBCS_PATH.read_text(encoding="utf-8")
        for path, text in self.generator.outputs(sbcs).items():
            with self.subTest(path=path.name):
                self.assertEqual(text, path.read_text(encoding="utf-8"))

    def test_encoding_round_trips(self) -> None:
        # Mirrors the C++ decoder in src/eval/dbcs_table.cpp.
        gen = self.generator
        for table in gen.TABLES:
            cells = gen.cells_for(table.codec)
            data = gen.encode_cells(cells)
            bits = "".join(f"{b:08b}" for b in data)
            pos = 0

            def read(width: int) -> int:
                nonlocal pos
                value = int(bits[pos : pos + width] or "0", 2) if width else 0
                pos += width
                return value

            def exp_golomb(k: int) -> int:
                nonlocal pos
                length = 0
                while bits[pos] == "0":
                    length += 1
                    pos += 1
                w = read(length + 1)
                return ((w - 1) << k) | read(k)

            decoded = [0] * len(cells)
            slot, prev, acc = 0, 0, gen.ACC_INIT
            while slot < len(cells):
                k = max(0, (acc >> gen.ACC_SHIFT).bit_length() - gen.K_BIAS)
                v = exp_golomb(k)
                if v == 0:
                    slot += exp_golomb(0) + 1
                    continue
                prev += -((v + 1) >> 1) if v & 1 else v >> 1
                decoded[slot] = prev
                slot += 1
                acc += v - (acc >> gen.ACC_SHIFT)
            with self.subTest(table=table.name):
                self.assertEqual(decoded, cells)

    def test_check_rejects_drifted_temporary_fixture(self) -> None:
        table = self.generator.TABLES[0]
        with tempfile.TemporaryDirectory(prefix="formulon-dbcs-") as directory:
            fixture = Path(directory) / table.path.name
            fixture.write_bytes(table.path.read_bytes() + b"// drift\n")
            with contextlib.redirect_stderr(io.StringIO()):
                self.assertEqual(self.generator.check({table.path: fixture}), 1)
                self.assertEqual(self.generator.check({table.path: fixture.with_name("missing.cpp")}), 2)
                self.assertEqual(self.generator.check(), 0)


if __name__ == "__main__":
    unittest.main()
