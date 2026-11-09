"""Workbook and sheet core: sheet positions, names, visibility, spans, required arguments and id lists."""

from __future__ import annotations

import datetime
import io
import unittest
import zipfile

import formulon
from formulon import (
    AlreadyExistsError,
    Color,
    ColorContext,
    ColorKind,
    FormulonError,
    IndexOutOfRangeError,
    InvalidArgumentError,
    NotFoundError,
    Person,
    ProtectedRangeInput,
    SheetVisibility,
    StatusCode,
    TableInput,
    ThreadedCommentInput,
    Workbook,
)

_PERSON = "{0A1B2C3D-0000-4000-8000-000000000001}"
_GUID_SHAPE = "{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}"


def _with_external_link() -> bytes:
    """A saved default workbook plus one external-link part that points at ``other.xlsx``."""
    with Workbook.create_default() as wb:
        saved = wb.save()
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(saved)) as zi, zipfile.ZipFile(out, "w") as zo:
        for name in zi.namelist():
            text = zi.read(name).decode()
            if name == "[Content_Types].xml":
                text = text.replace(
                    "</Types>",
                    '<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType='
                    '"application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>',
                )
            elif name == "xl/workbook.xml":
                text = text.replace(
                    "</sheets>", '</sheets><externalReferences><externalReference r:id="rId9"/></externalReferences>'
                )
            elif name == "xl/_rels/workbook.xml.rels":
                text = text.replace(
                    "</Relationships>",
                    '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                    'relationships/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>',
                )
            zo.writestr(name, text)
        zo.writestr(
            "xl/externalLinks/externalLink1.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<externalBook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1">'
            '<sheetNames><sheetName val="Sheet1"/></sheetNames></externalBook></externalLink>',
        )
        zo.writestr(
            "xl/externalLinks/_rels/externalLink1.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
            'externalLinkPath" Target="other.xlsx" TargetMode="External"/></Relationships>',
        )
    return out.getvalue()


def _names(wb: Workbook) -> list:
    return [s.name for s in wb.list_sheets()]


class SheetPositionTests(unittest.TestCase):
    def test_add_sheet_insert_before_accepts_a_name_a_handle_and_a_position(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.add_sheet("ByName", insert_before="Sheet1").index, 0)
            first = wb.sheet(0)
            self.assertEqual(wb.add_sheet("ByHandle", insert_before=first).index, 0)
            self.assertEqual(wb.add_sheet("ByPosition", insert_before=1).index, 1)
            self.assertEqual(wb.add_sheet("Tail", insert_before=None).index, 4)
            self.assertEqual(wb.add_sheet("Appended", insert_before=len(wb)).index, 5)
            self.assertEqual(_names(wb), ["ByHandle", "ByPosition", "ByName", "Sheet1", "Tail", "Appended"])

    def test_add_sheet_insert_before_unknown_sheet_is_not_found_or_out_of_range(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(NotFoundError):
                wb.add_sheet("X", insert_before="Missing")
            with self.assertRaises(IndexOutOfRangeError) as cm:
                wb.add_sheet("X", insert_before=len(wb) + 1)
            self.assertEqual(cm.exception.code, StatusCode.SHEET_INDEX_OUT_OF_RANGE)
            self.assertEqual(_names(wb), ["Sheet1"])

    def test_add_and_copy_sheet_options_are_keyword_only(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(TypeError):
                wb.add_sheet("X", 0)  # type: ignore[misc]
            with self.assertRaises(TypeError):
                wb.copy_sheet(0, 0)  # type: ignore[misc]
            self.assertEqual(len(wb), 1)

    def test_copy_sheet_insert_before_and_new_name(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_sheet("Other")
            self.assertEqual(wb.copy_sheet("Other", insert_before="Sheet1").index, 0)
            self.assertEqual(wb.sheet(0).name, "Other (2)")
            self.assertEqual(wb.copy_sheet(0, insert_before=None, new_name="Tail").index, 3)
            self.assertEqual(wb.copy_sheet(0, insert_before=len(wb)).index, 4)
            self.assertEqual(_names(wb), ["Other (2)", "Sheet1", "Other", "Tail", "Other (3)"])
            with self.assertRaises(NotFoundError):
                wb.copy_sheet(0, insert_before="Missing")
            with self.assertRaises(AlreadyExistsError):
                wb.copy_sheet(0, new_name="TAIL")

    def test_move_uses_the_list_before_the_removal(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_sheet("B")
            wb.add_sheet("C")
            wb.sheet(0).move(2)
            self.assertEqual(_names(wb), ["B", "Sheet1", "C"])
            wb.sheet("B").move(None)
            self.assertEqual(_names(wb), ["Sheet1", "C", "B"])
            wb.sheet("Sheet1").move(3)
            self.assertEqual(_names(wb), ["C", "B", "Sheet1"])

    def test_move_before_itself_or_before_the_next_sheet_changes_nothing(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_sheet("B")
            wb.add_sheet("C")
            first = wb.sheet(0)
            first.move(first)
            first.move(1)
            self.assertEqual(_names(wb), ["Sheet1", "B", "C"])
            self.assertEqual(first.index, 0)
            with self.assertRaises(IndexOutOfRangeError):
                first.move(4)

    def test_negative_subscript_counts_from_the_end(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_sheet("B")
            wb.add_sheet("C")
            self.assertIs(wb[-1], wb.sheet(2))
            self.assertIs(wb[-len(wb)], wb.sheet(0))
            with self.assertRaises(IndexOutOfRangeError):
                wb[-len(wb) - 1]
            with self.assertRaises(IndexOutOfRangeError):
                wb.sheet(-1)
            with self.assertRaises(TypeError):
                wb[0:1]  # type: ignore[index]

    def test_workbook_membership_and_length_stand_in_for_has_sheet_and_sheet_count(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_sheet("Data")
            self.assertEqual(len(wb), len(wb.list_sheets()))
            self.assertIn("Data", wb)
            self.assertIn("data", wb)
            self.assertIn(1, wb)
            self.assertNotIn("Missing", wb)
            self.assertNotIn(99, wb)


class SheetNameTests(unittest.TestCase):
    def test_add_sheet_without_a_name_takes_the_smallest_free_number(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.add_sheet().name, "Sheet2")
            wb.add_sheet("sheet4")
            self.assertEqual(wb.add_sheet().name, "Sheet3")
            self.assertEqual(wb.add_sheet(None).name, "Sheet5")

    def test_add_sheet_without_a_name_follows_the_profile_base_name(self) -> None:
        with Workbook.create_for_profile("win-365-de_DE") as wb:
            self.assertEqual(wb.sheet(0).name, "Tabelle1")
            self.assertEqual(wb.add_sheet().name, "Tabelle2")
            wb.sheet("Tabelle2").rename("keep")
            wb.add_sheet("tabelle3")
            self.assertEqual(wb.add_sheet().name, "Tabelle2")
            self.assertEqual(wb.add_sheet().name, "Tabelle4")

    def test_an_empty_sheet_name_is_still_refused(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError) as cm:
                wb.add_sheet("")
            self.assertEqual(cm.exception.code, StatusCode.INVALID_SHEET_NAME)
            self.assertEqual(len(wb), 1)

    def test_protected_range_names_compare_by_unicode_case_fold(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.add_protected_range(ProtectedRangeInput(name="Budget", ranges="A1:B2"))
            with self.assertRaises(AlreadyExistsError):
                ws.add_protected_range(ProtectedRangeInput(name="BUDGET", ranges="C1:C2"))
            ws.remove_protected_range("budget")
            self.assertEqual(ws.list_protected_ranges(), [])
            ws.add_protected_range(ProtectedRangeInput(name="Übersicht", ranges="A1:B2"))
            with self.assertRaises(AlreadyExistsError):
                ws.add_protected_range(ProtectedRangeInput(name="übersicht", ranges="C1:C2"))

    def test_table_names_compare_by_unicode_case_fold(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for r, header in enumerate(("a", "b")):
                ws.set_text(0, r, header)
                ws.set_number(1, r, 1.0)
            ws.set_text(4, 0, "a")
            ws.set_text(4, 1, "b")
            ws.add_table(TableInput(range="A1:B2", name="Übersicht"))
            with self.assertRaises(AlreadyExistsError):
                ws.add_table(TableInput(range="A5:B5", name="übersicht"))

    def test_cell_style_lookup_ignores_case(self) -> None:
        with Workbook.create_default() as wb:
            style = wb.get_cell_style("Normal")
            self.assertIsNotNone(style)
            assert style is not None
            lowered = wb.get_cell_style("normal")
            self.assertIsNotNone(lowered)
            assert lowered is not None
            self.assertEqual(lowered.name, style.name)


class SheetVisibilityTests(unittest.TestCase):
    def test_set_visibility_reads_back(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_sheet("Shown")
            ws = wb.sheet(0)
            self.assertIs(ws.get_visibility(), SheetVisibility.VISIBLE)
            for state in (SheetVisibility.HIDDEN, SheetVisibility.VERY_HIDDEN, SheetVisibility.VISIBLE):
                ws.set_visibility(state)
                self.assertIs(ws.get_visibility(), state)

    def test_hiding_the_last_visible_sheet_is_precondition_failed(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for state in (SheetVisibility.HIDDEN, SheetVisibility.VERY_HIDDEN):
                with self.assertRaises(FormulonError) as cm:
                    ws.set_visibility(state)
                self.assertEqual(cm.exception.code, StatusCode.PRECONDITION_FAILED)
            self.assertIs(ws.get_visibility(), SheetVisibility.VISIBLE)

    def test_a_visibility_that_is_not_a_member_is_invalid(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_sheet("Shown")
            with self.assertRaises(InvalidArgumentError):
                wb.sheet(0).set_visibility(3)  # type: ignore[arg-type]
            self.assertIs(wb.sheet(0).get_visibility(), SheetVisibility.VISIBLE)

    def test_the_sheet_view_no_longer_carries_visibility(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(TypeError):
                wb.sheet(0).update_view(visibility=SheetVisibility.HIDDEN)  # type: ignore[call-arg]
            self.assertFalse(hasattr(wb.sheet(0).get_view(), "visibility"))


class SheetSpanTests(unittest.TestCase):
    def test_insert_rows_first_last_is_inclusive(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_number(3, 0, 9.0)  # A4
            ws.insert_rows(3, 5)
            self.assertEqual(ws.get_typed_value(6, 0).number, 9.0)  # three rows inserted
            self.assertEqual(ws.get_typed_value(3, 0).kind, formulon.ValueKind.BLANK)

    def test_delete_columns_first_last_and_the_span_string_agree(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for c in range(6):
                ws.set_number(0, c, float(c))
            ws.delete_columns(2, 3)
            self.assertEqual([ws.get_typed_value(0, c).number for c in range(4)], [0.0, 1.0, 4.0, 5.0])
            ws.insert_columns("A:B")
            self.assertEqual(ws.get_typed_value(0, 2).number, 0.0)
            ws.delete_columns("A:B")
            self.assertEqual(ws.get_typed_value(0, 0).number, 0.0)

    def test_first_past_last_is_invalid(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for call in (ws.insert_rows, ws.delete_rows, ws.insert_columns, ws.delete_columns):
                with self.subTest(call=call.__name__), self.assertRaises(InvalidArgumentError):
                    call(5, 3)


class RequiredArgumentTests(unittest.TestCase):
    def test_resolve_color_requires_its_context(self) -> None:
        with Workbook.create_default() as wb:
            color = Color(kind=ColorKind.RGB, argb=0xFF123456)
            with self.assertRaises(TypeError):
                wb.resolve_color(color)  # type: ignore[call-arg]
            self.assertEqual(wb.resolve_color(color, ColorContext.FONT).argb, 0xFF123456)

    def test_format_value_requires_its_number_format(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(TypeError):
                wb.format_value(1)  # type: ignore[call-arg]
            self.assertEqual(
                wb.format_value(formulon.Value(kind=formulon.ValueKind.NUMBER, number=1.5), "0.00").text, "1.50"
            )

    def test_the_second_argument_is_named_number_format(self) -> None:
        with Workbook.create_default() as wb:
            value = formulon.Value(kind=formulon.ValueKind.NUMBER, number=1.5)
            self.assertEqual(wb.format_value(value, number_format="0.0").text, "1.5")
            with self.assertRaises(TypeError):
                wb.format_value(value, format_code="0.0")  # type: ignore[call-arg]


class ProfileIdListTests(unittest.TestCase):
    def test_every_listed_profile_id_creates_a_workbook(self) -> None:
        ids = formulon.list_profile_ids()
        self.assertEqual(len(ids), 28)
        self.assertEqual(len(set(ids)), 28)
        self.assertIn("win-365-en_US", ids)
        self.assertIn("mac-365-ja_JP", ids)
        for profile_id in ids:
            with self.subTest(profile_id=profile_id), Workbook.create_for_profile(profile_id) as wb:
                self.assertEqual(wb.get_profile_id(), profile_id)

    def test_the_list_is_a_fresh_copy(self) -> None:
        ids = formulon.list_profile_ids()
        ids.clear()
        self.assertEqual(len(formulon.list_profile_ids()), 28)


class SheetIterationTests(unittest.TestCase):
    def test_a_sheet_refuses_iteration_and_membership(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            with self.assertRaises(TypeError):
                iter(ws)
            with self.assertRaises(TypeError):
                _ = "A1" in ws
            with self.assertRaises(TypeError):
                [row for row in ws]


class ExternalLinkBookNumberTests(unittest.TestCase):
    def test_book_number_is_the_one_based_bracket_index_of_formula_text(self) -> None:
        with Workbook.load(_with_external_link()) as wb:
            links = wb.list_external_links()
            self.assertEqual([link.book_number for link in links], [1])
            self.assertFalse(hasattr(links[0], "index"))
            wb.sheet(0).set_formula("A1", "=[1]Sheet1!A1")
            self.assertEqual(wb.sheet(0).get_formula("A1"), "=[1]Sheet1!A1")


class ThreadedCommentIdTests(unittest.TestCase):
    def _comment(self, **overrides: object) -> ThreadedCommentInput:
        return ThreadedCommentInput(person_id=_PERSON, text="x", **overrides)  # type: ignore[arg-type]

    def test_an_omitted_id_is_generated_and_returned(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_person(Person(id=_PERSON, display_name="Alice"))
            ws = wb.sheet(0)
            new_id = ws.add_threaded_comment(self._comment())
            self.assertEqual(len(new_id), len(_GUID_SHAPE))
            self.assertTrue(new_id.startswith("{") and new_id.endswith("}"))
            self.assertEqual(ws.get_threaded_comment(new_id).text, "x")
            other = ws.add_threaded_comment(self._comment(cell="B2"))
            self.assertNotEqual(other, new_id)

    def test_an_omitted_created_takes_the_pinned_clock(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_person(Person(id=_PERSON, display_name="Alice"))
            wb.set_pinned_now(datetime.datetime(2026, 5, 6, 7, 8, 9))
            ws = wb.sheet(0)
            new_id = ws.add_threaded_comment(self._comment())
            self.assertEqual(ws.get_threaded_comment(new_id).created, datetime.datetime(2026, 5, 6, 7, 8, 9))

    def test_a_pinned_clock_makes_generated_ids_repeatable(self) -> None:
        ids = []
        for _ in range(2):
            with Workbook.create_default() as wb:
                wb.add_person(Person(id=_PERSON, display_name="Alice"))
                wb.set_pinned_now(datetime.datetime(2026, 5, 6, 7, 8, 9))
                ids.append(wb.sheet(0).add_threaded_comment(self._comment()))
        self.assertEqual(ids[0], ids[1])

    def test_an_explicit_id_is_kept_and_a_duplicate_already_exists(self) -> None:
        explicit = "{6F1E0D2C-1111-4111-8111-000000000001}"
        with Workbook.create_default() as wb:
            wb.add_person(Person(id=_PERSON, display_name="Alice"))
            ws = wb.sheet(0)
            self.assertEqual(ws.add_threaded_comment(self._comment(id=explicit)), explicit)
            with self.assertRaises(AlreadyExistsError):
                ws.add_threaded_comment(self._comment(id=explicit, cell="B2"))

    def test_an_unknown_parent_is_not_found(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_person(Person(id=_PERSON, display_name="Alice"))
            with self.assertRaises(NotFoundError):
                wb.sheet(0).add_threaded_comment(self._comment(parent_id="{6F1E0D2C-1111-4111-8111-0000000000FF}"))


class TypedOoxmlAccessorTests(unittest.TestCase):
    def test_ooxml_is_a_property_declared_in_the_class_body(self) -> None:
        self.assertIsInstance(formulon.Workbook.__dict__["ooxml"], property)
        hints = formulon.Workbook.ooxml.fget.__annotations__  # type: ignore[union-attr]
        self.assertIn("return", hints)
        with Workbook.create_default() as wb:
            self.assertIsInstance(wb.ooxml, formulon.WorkbookOoxml)


class SatisfiedSurfacePins(unittest.TestCase):
    def test_handle_classes_refuse_direct_construction(self) -> None:
        for cls in (formulon.Sheet, formulon.Table, formulon.Pivot, formulon.PivotCache):
            with self.subTest(cls=cls.__name__), self.assertRaises(TypeError):
                cls()  # type: ignore[call-arg]

    def test_a_table_names_its_sheet(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_text(0, 0, "a")
            ws.set_number(1, 0, 1.0)
            table = ws.add_table(TableInput(range="A1:A2", name="Orders"))
            self.assertIs(table.sheet, ws)

    def test_totals_row_function_is_exported_and_typed_on_columns(self) -> None:
        self.assertTrue(hasattr(formulon, "TableTotalsRowFunction"))
        self.assertIn("TableTotalsRowFunction", formulon.__all__)
        hint = str(formulon.TableColumn.__dataclass_fields__["totals_row_function"].type)
        self.assertIn("TableTotalsRowFunction", hint)


if __name__ == "__main__":
    unittest.main()
