"""Integer arguments take integers only: an absent value is ``None``, never any falsy value."""

from __future__ import annotations

import pathlib
import re
import unittest

from formulon import (
    Workbook,
)


class PageArgumentTests(unittest.TestCase):
    def test_cursor_and_limit_refuse_a_bool_and_a_float(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for kwargs in ({"limit": False}, {"limit": 0.0}, {"cursor": False}, {"cursor": 0.0}):
                with self.subTest(kwargs=kwargs):
                    with self.assertRaises(TypeError):
                        ws.get_invalid_cells(**kwargs)
                    with self.assertRaises(TypeError):
                        ws.get_cells_in_range("A1:B2", **kwargs)


class FalsyFoldTests(unittest.TestCase):
    def test_no_integer_argument_folds_a_falsy_value_to_zero(self) -> None:
        # `x or 0` would pass False and 0.0 as 0 instead of refusing them; an absent value is tested with `is None`.
        package = pathlib.Path(__file__).resolve().parent.parent / "formulon"
        folded = re.compile(r"_[us]int\([^,()]+ or 0,")
        hits = [
            f"{path.name}:{n}"
            for path in sorted(package.glob("*.py"))
            for n, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1)
            if folded.search(line)
        ]
        self.assertEqual(hits, [])


if __name__ == "__main__":
    unittest.main()
