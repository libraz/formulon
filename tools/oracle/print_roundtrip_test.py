#!/usr/bin/env python3
"""Tests for the print round-trip fixture author.

Split in two. The address and spec handling is pure Python and always
runs. The authoring tests need the Formulon binding and skip without it,
because this file also runs on the oracle venv, which deliberately does
not install the wheel.
"""

from __future__ import annotations

import sys
import unittest
import zipfile
from io import BytesIO
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import print_roundtrip as rt  # noqa: E402

try:
    sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "packages" / "python"))
    import formulon as _formulon
except Exception:  # noqa: BLE001 - any import failure means "not available here"
    _formulon = None


def _case(roundtrip, sheets=None):
    return {
        "id": "case",
        "sheets": sheets if sheets is not None else {"Sheet1": {"A1": {"kind": "text", "value": "x"}}},
        "roundtrip": roundtrip,
    }


class AddressTest(unittest.TestCase):
    def test_column_letters_convert_to_zero_based(self):
        self.assertEqual(rt._column_index("A"), 0)
        self.assertEqual(rt._column_index("Z"), 25)
        self.assertEqual(rt._column_index("AA"), 26)
        self.assertEqual(rt._column_index("XFD"), 16383)

    def test_a1_splits_to_zero_based_row_and_column(self):
        self.assertEqual(rt._split_a1("A1"), (0, 0))
        self.assertEqual(rt._split_a1("D20"), (19, 3))

    def test_malformed_address_is_rejected(self):
        with self.assertRaises(rt.RoundtripSpecError):
            rt._split_a1("1A")


class SpecTest(unittest.TestCase):
    def test_case_without_a_roundtrip_block_is_rejected(self):
        with self.assertRaisesRegex(rt.RoundtripSpecError, "no roundtrip block"):
            rt.author_case({"id": "case", "sheets": {}}, module=object())

    def test_roundtrip_ids_are_listed(self):
        doc = {
            "cases": [
                {"id": "plain", "print": {}},
                {"id": "rt", "roundtrip": {"sheet": "Sheet1"}},
            ]
        }
        self.assertEqual(rt.roundtrip_case_ids(doc), ["rt"])


class _RecordingWorkbook:
    """Records the sheet calls `_apply_sheets` makes, in order."""

    def __init__(self, default_name="Sheet1"):
        self.default_name = default_name
        self.order = []

    def rename_sheet(self, index, name):
        self.order.append(name)

    def add_sheet(self, name):
        self.order.append(name)

    def sheet_name(self, index):
        return self.default_name

    def set_blank(self, *args):
        pass

    def set_number(self, *args):
        pass

    def set_bool(self, *args):
        pass

    def set_text(self, *args):
        pass

    def set_formula(self, *args):
        pass


class SheetIndexTest(unittest.TestCase):
    """Sheet indices follow the spec's declaration order, not its names.

    The capture half and the C++ authoring half have to land on the same
    name -> index map or they author and inspect different sheets, and a
    comparison against a field that happens to match the printer default
    passes anyway. Pinning the order here is what gives the other half a
    fixed contract to match.
    """

    def test_indices_follow_declaration_order(self):
        wb = _RecordingWorkbook()
        indices = rt._apply_sheets(wb, {"sheets": {"Report": {}, "Data": {}}})
        self.assertEqual(indices, {"Report": 0, "Data": 1})
        self.assertEqual(wb.order, ["Report", "Data"])

    def test_reverse_declaration_order_is_honoured_too(self):
        # The same two names the other way round: an implementation that
        # sorts by name returns the same map for both, and only this pair
        # tells the two apart.
        indices = rt._apply_sheets(_RecordingWorkbook(), {"sheets": {"Data": {}, "Report": {}}})
        self.assertEqual(indices, {"Data": 0, "Report": 1})

    def test_empty_sheets_block_falls_back_to_the_default_sheet(self):
        indices = rt._apply_sheets(_RecordingWorkbook("Sheet1"), {"sheets": {}})
        self.assertEqual(indices, {"Sheet1": 0})

    def test_absent_sheets_block_falls_back_to_the_default_sheet(self):
        indices = rt._apply_sheets(_RecordingWorkbook("Sheet1"), {})
        self.assertEqual(indices, {"Sheet1": 0})


@unittest.skipIf(_formulon is None, "the Formulon Python binding is not importable here")
class AuthoringTest(unittest.TestCase):
    """Drives the real binding, so these assert on the bytes we would ship."""

    def _sheet_xml(self, case):
        data = rt.author_case(case, module=_formulon)
        with zipfile.ZipFile(BytesIO(data)) as zf:
            return zf.read("xl/worksheets/sheet1.xml").decode("utf-8")

    def _workbook_xml(self, case):
        data = rt.author_case(case, module=_formulon)
        with zipfile.ZipFile(BytesIO(data)) as zf:
            return zf.read("xl/workbook.xml").decode("utf-8")

    def test_fit_to_page_writes_both_halves(self):
        # The sheetPr flag and the pageSetup attributes are separate
        # elements; writing one without the other is the mistake this
        # whole capture mode exists to catch, so pin it at authoring time
        # too rather than waiting for Excel to disagree.
        xml = self._sheet_xml(
            _case(
                {
                    "sheet": "Sheet1",
                    "page_setup": {"fit_to_page": True, "fit_to_width": 1, "fit_to_height": 2},
                }
            )
        )
        self.assertIn('fitToPage="true"', xml)
        self.assertIn('fitToWidth="1"', xml)
        self.assertIn('fitToHeight="2"', xml)

    def test_literal_ampersand_is_xml_escaped(self):
        # Two escaping layers stack: "&&" is Excel's spelling of a literal
        # ampersand in a header, and every ampersand is then XML-escaped.
        # A raw "&" on disk would make the part ill-formed.
        xml = self._sheet_xml(_case({"sheet": "Sheet1", "header_footer": {"odd_header": "&Ca && b"}}))
        self.assertIn("<oddHeader>&amp;Ca &amp;&amp; b</oddHeader>", xml)
        self.assertNotIn("<oddHeader>&C", xml)

    def test_print_area_and_titles_become_defined_names(self):
        xml = self._workbook_xml(
            _case(
                {
                    "sheet": "Sheet1",
                    "print_area": "A1:F40",
                    "print_titles": {"repeat_rows": "1:2", "repeat_cols": "A:A"},
                }
            )
        )
        self.assertIn("_xlnm.Print_Area", xml)
        self.assertIn("$A$1:$F$40", xml)
        self.assertIn("_xlnm.Print_Titles", xml)
        # A whole-axis token keeps its shape rather than being expanded to
        # the full grid, which is how Excel writes it.
        self.assertIn("$A:$A", xml)

    def test_breaks_are_authored_as_manual(self):
        data = rt.author_case(_case({"sheet": "Sheet1", "row_breaks": [21], "col_breaks": ["D"]}), module=_formulon)
        reloaded = _formulon.Workbook.load(data)
        try:
            rows = reloaded.get_row_breaks(0)
            cols = reloaded.get_col_breaks(0)
        finally:
            reloaded.close()
        # A case states 1-based rows, matching the rest of the schema; the
        # break sits before that row, so it reads back at row - 1.
        self.assertEqual([(b.id, b.manual) for b in rows], [(20, True)])
        self.assertEqual([(b.id, b.manual) for b in cols], [(3, True)])

    def test_named_sheet_selects_the_right_index(self):
        case = _case(
            {"sheet": "Report", "page_setup": {"orientation": 2}},
            sheets={
                "Data": {"A1": {"kind": "number", "value": 1.0}},
                "Report": {"A1": {"kind": "text", "value": "r"}},
            },
        )
        data = rt.author_case(case, module=_formulon)
        with zipfile.ZipFile(BytesIO(data)) as zf:
            first = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
            second = zf.read("xl/worksheets/sheet2.xml").decode("utf-8")
        self.assertNotIn("landscape", first)
        self.assertIn("landscape", second)

    def test_unknown_sheet_name_is_rejected(self):
        with self.assertRaisesRegex(rt.RoundtripSpecError, "is not one of"):
            rt.author_case(_case({"sheet": "Nope"}), module=_formulon)

    def test_payload_carries_the_digest_of_the_bytes(self):
        import base64
        import hashlib

        case = _case({"sheet": "Sheet1", "page_setup": {"scale": 75}})
        payload = rt.fixture_payload(case, module=_formulon)
        raw = base64.b64decode(payload["xlsx_base64"])
        self.assertEqual(payload["xlsx_bytes"], len(raw))
        self.assertEqual(payload["xlsx_sha256"], hashlib.sha256(raw).hexdigest())


def _target_sheet_index(case):
    """Resolves a case's `roundtrip.sheet` to the 0-based index it authors.

    Mirrors `_apply_sheets`'s declaration-order contract (see
    `SheetIndexTest` above) without needing a live workbook: a string name
    is looked up in the declared `sheets` order, an int is used as-is, and
    an absent `sheets` block falls back to sheet 0.
    """
    names = list((case.get("sheets") or {}).keys())
    sheet_key = case["roundtrip"].get("sheet", 0)
    if isinstance(sheet_key, str):
        return names.index(sheet_key) if names else 0
    return int(sheet_key)


class CommittedSuiteTest(unittest.TestCase):
    """The committed round-trip suite has to stay authorable."""

    CASE_FILE = Path(__file__).resolve().parents[2] / "tests" / "oracle" / "cases_wb" / "print_roundtrip.case.json"

    def test_every_case_declares_a_roundtrip_block(self):
        import json

        doc = json.loads(self.CASE_FILE.read_text(encoding="utf-8"))
        ids = [c["id"] for c in doc["cases"]]
        self.assertEqual(rt.roundtrip_case_ids(doc), ids)

    @unittest.skipIf(_formulon is None, "the Formulon Python binding is not importable here")
    def test_every_case_authors_a_readable_workbook(self):
        import base64
        import json

        doc = json.loads(self.CASE_FILE.read_text(encoding="utf-8"))
        for case in doc["cases"]:
            with self.subTest(case=case["id"]):
                payload = rt.fixture_payload(case, module=_formulon)
                self.assertGreater(payload["xlsx_bytes"], 0)
                raw = base64.b64decode(payload["xlsx_base64"])
                # `xlsx_bytes > 0` alone would still pass for a truncated or
                # non-zip blob; open it as a zip and require the parts a
                # reader actually needs -- the workbook manifest and the
                # worksheet part the case's `roundtrip.sheet` targets.
                with zipfile.ZipFile(BytesIO(raw)) as zf:
                    names = zf.namelist()
                    self.assertIn("xl/workbook.xml", names)
                    sheet_part = f"xl/worksheets/sheet{_target_sheet_index(case) + 1}.xml"
                    self.assertIn(
                        sheet_part,
                        names,
                        f"case {case['id']!r} targets sheet index {_target_sheet_index(case)} "
                        f"but the archive has no {sheet_part}",
                    )
                    self.assertGreater(len(zf.read(sheet_part)), 0)


if __name__ == "__main__":
    unittest.main()
