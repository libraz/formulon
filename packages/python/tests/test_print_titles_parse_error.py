"""A print-titles span string that does not parse fails as every A1 argument does."""

from __future__ import annotations

import unittest

from formulon import (
    FormulonError,
    PrintTitlesInput,
    StatusCode,
    Workbook,
)


class SpanTextTests(unittest.TestCase):
    def test_print_titles_span_that_does_not_parse_is_a_parser_error(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError) as cm:
                wb.sheet(0).set_print_titles(PrintTitlesInput(rows="1:!", columns=None))
            self.assertEqual(cm.exception.code, StatusCode.PARSER_INVALID_REFERENCE)


if __name__ == "__main__":
    unittest.main()
