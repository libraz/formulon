"""A table range naming a sheet the workbook does not have is NotFound."""

from __future__ import annotations

import unittest

from formulon import (
    NotFoundError,
    TableInput,
    Workbook,
)


class TableRangeTests(unittest.TestCase):
    def test_table_range_on_an_unknown_sheet_is_not_found(self) -> None:
        with Workbook.create_default() as wb:
            table = wb.sheet(0).add_table(TableInput(range="A1:B3", name="Sales"))
            with self.assertRaises(NotFoundError):
                table.update(range="Nope!A1:B4")


if __name__ == "__main__":
    unittest.main()
