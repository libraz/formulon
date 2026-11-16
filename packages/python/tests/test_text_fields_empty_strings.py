"""Empty-value rules of the comment, person, hyperlink, validation, defined-name, protection, table and format_reference
text fields (api-conventions §5 rule 30).

``""`` is refused where it is not a legal value, means the same as the absent value where the file holds one state, and
is a value where the file keeps it.
"""

from __future__ import annotations

import re
import unittest

from formulon import (
    CellRect,
    CommentInput,
    DataValidationInput,
    DefinedNameInput,
    DxfInput,
    HyperlinkInput,
    InvalidArgumentError,
    MentionInput,
    PersonInput,
    ProtectedRangeInput,
    SheetProtectionInput,
    TableInput,
    ThreadedCommentInput,
    ValidationType,
    Workbook,
    WorkbookProtectionInput,
)

GUID = re.compile(r"^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$")


class _Base(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)


class FormatReferenceTests(_Base):
    def test_empty_sheet_name_is_refused_and_none_writes_no_qualifier(self) -> None:
        with self.assertRaises(InvalidArgumentError):
            self.wb.format_reference(CellRect(0, 0, 0, 0), sheet_name="")
        self.assertEqual(self.wb.format_reference(CellRect(0, 0, 0, 0), sheet_name=None), "A1")
        self.assertEqual(self.wb.format_reference(CellRect(0, 0, 0, 0)), "A1")
        self.assertEqual(self.wb.format_reference(self.wb.parse_reference("B2:C3")), "B2:C3")


class HyperlinkTests(_Base):
    def test_empty_target_and_location_are_refused_and_add_nothing(self) -> None:
        with self.assertRaises(InvalidArgumentError):
            self.ws.set_hyperlink(HyperlinkInput(range="A1", target=""))
        with self.assertRaises(InvalidArgumentError):
            self.ws.set_hyperlink(HyperlinkInput(range="A1", target="https://example.com/", location=""))
        with self.assertRaises(InvalidArgumentError):
            self.ws.set_hyperlink(HyperlinkInput(range="A1", target=None, location=""))
        with self.assertRaises(InvalidArgumentError):
            self.ws.insert_hyperlink_at(0, HyperlinkInput(range="A1", target=""))
        self.assertEqual(self.ws.list_hyperlinks(), [])

    def test_none_target_or_location_is_absent(self) -> None:
        self.ws.set_hyperlink(HyperlinkInput(range="A1", target="https://example.com/", location=None))
        self.ws.set_hyperlink(HyperlinkInput(range="B1", target=None, location="Sheet1!C3"))
        self.assertIsNone(self.ws.get_hyperlink("A1").location)
        self.assertIsNone(self.ws.get_hyperlink("B1").target)
        self.assertEqual(self.ws.get_hyperlink("B1").location, "Sheet1!C3")

    def test_empty_display_and_tooltip_are_the_absent_value(self) -> None:
        self.ws.set_hyperlink(HyperlinkInput(range="A1", target="https://example.com/", display="", tooltip=""))
        self.ws.set_hyperlink(HyperlinkInput(range="B1", target="https://example.com/"))
        empty = self.ws.get_hyperlink("A1")
        omitted = self.ws.get_hyperlink("B1")
        self.assertEqual(empty.display, omitted.display)
        self.assertIsNone(empty.tooltip)
        self.assertIsNone(omitted.tooltip)


class ValidationTests(_Base):
    @staticmethod
    def _rule(**extra: object) -> DataValidationInput:
        fields: dict = {"type": ValidationType.WHOLE, "ranges": "D1:D3", "formula1": "1", "formula2": "5"}
        fields.update(extra)
        return DataValidationInput(**fields)

    def test_empty_formulas_are_refused_and_add_nothing(self) -> None:
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_validation(self._rule(formula1=""))
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_validation(self._rule(formula2=""))
        self.assertEqual(self.ws.list_validations(), [])

    def test_none_formula_is_absent(self) -> None:
        rule_id = self.ws.add_validation(self._rule(formula2=None))
        self.assertIsNone(self.ws.get_validation(rule_id).formula2)
        self.assertIsNotNone(self.ws.get_validation(rule_id).formula1)

    def test_update_refuses_an_empty_formula_and_none_clears_it(self) -> None:
        rule_id = self.ws.add_validation(self._rule())
        with self.assertRaises(InvalidArgumentError):
            self.ws.update_validation(rule_id, formula1="")
        with self.assertRaises(InvalidArgumentError):
            self.ws.update_validation(rule_id, formula2="")
        self.assertIsNotNone(self.ws.get_validation(rule_id).formula1)
        self.assertIsNotNone(self.ws.get_validation(rule_id).formula2)
        self.ws.update_validation(rule_id, formula2=None)
        self.assertIsNone(self.ws.get_validation(rule_id).formula2)

    def test_empty_titles_and_messages_are_the_absent_value(self) -> None:
        rule_id = self.ws.add_validation(
            self._rule(error_title="", error_message="", prompt_title="", prompt_message="")
        )
        got = self.ws.get_validation(rule_id)
        self.assertIsNone(got.error_title)
        self.assertIsNone(got.error_message)
        self.assertIsNone(got.prompt_title)
        self.assertIsNone(got.prompt_message)


class DefinedNameTests(_Base):
    def test_empty_comment_and_omitted_are_the_same_state(self) -> None:
        self.wb.set_defined_name(DefinedNameInput(name="WithEmpty", formula="=Sheet1!$A$1", comment=""))
        self.wb.set_defined_name(DefinedNameInput(name="WithNone", formula="=Sheet1!$A$1"))
        self.assertIsNone(self.wb.get_defined_name("WithEmpty").comment)
        self.assertIsNone(self.wb.get_defined_name("WithNone").comment)


class NoteTests(_Base):
    def test_empty_text_writes_a_note_with_no_text_and_none_removes_it(self) -> None:
        self.ws.set_comment(1, 1, CommentInput(""))
        note = self.ws.get_comment(1, 1)
        self.assertIsNotNone(note)
        self.assertEqual(note.text, "")
        self.assertEqual(len(self.ws.list_comments()), 1)
        self.ws.set_comment(1, 1, None)  # type: ignore[arg-type]
        self.assertIsNone(self.ws.get_comment(1, 1))
        self.assertEqual(self.ws.list_comments(), [])

    def test_empty_author_and_omitted_are_the_same_state(self) -> None:
        self.ws.set_comment(1, 1, CommentInput("a", author=""))
        self.ws.set_comment(2, 2, CommentInput("a"))
        self.assertIsNone(self.ws.get_comment(1, 1).author)
        self.assertIsNone(self.ws.get_comment(2, 2).author)


class PeopleAndThreadsTests(_Base):
    def test_person_empty_id_and_provider_id_are_refused_and_user_id_is_absent(self) -> None:
        with self.assertRaises(InvalidArgumentError):
            self.wb.add_person(PersonInput(display_name="Ann", id=""))
        with self.assertRaises(InvalidArgumentError):
            self.wb.add_person(PersonInput(display_name="Ann", provider_id=""))
        self.assertEqual(self.wb.list_people(), [])
        person_id = self.wb.add_person(PersonInput(display_name="Ann", user_id=""))
        self.assertIsNone(self.wb.get_person(person_id).user_id)
        self.assertIsNone(self.wb.get_person(person_id).provider_id)

    def test_threaded_comment_empty_ids_are_refused_and_omitted_generates(self) -> None:
        person = self.wb.add_person(PersonInput(display_name="Ann"))

        def make(**extra: object) -> ThreadedCommentInput:
            return ThreadedCommentInput(person_id=person, text="@Ann hi", cell="B2", **extra)  # type: ignore[arg-type]

        with self.assertRaises(InvalidArgumentError):
            self.ws.add_threaded_comment(make(id=""))
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_threaded_comment(make(mentions=[MentionInput(person_id="", start=0, end=4)]))
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_threaded_comment(make(mentions=[MentionInput(person_id=person, start=0, end=4, mention_id="")]))
        self.assertEqual(self.ws.list_threaded_comments(), [])

        comment_id = self.ws.add_threaded_comment(make(mentions=[MentionInput(person_id=person, start=0, end=4)]))
        (comment,) = self.ws.list_threaded_comments()
        self.assertEqual(comment.id, comment_id)
        self.assertRegex(comment.mentions[0].mention_id, GUID)


class TableTests(_Base):
    def test_empty_name_and_empty_column_name_are_refused(self) -> None:
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_table(TableInput(range="A1:B3", name=""))
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_table(TableInput(range="A1:B3", name="Sales", columns=["", ""]))
        with self.assertRaises(InvalidArgumentError):
            self.ws.add_table(TableInput(range="A1:B3", name="Sales", columns=["A", ""]))
        # Nothing was created, so the name is still free; omitted columns derive their names.
        table = self.ws.add_table(TableInput(range="A1:B3", name="Sales"))
        self.assertEqual(table.name, "Sales")


PROTECTED_RANGE_TEXT = ("legacy_password_hash", "algorithm_name", "hash_value", "salt_value", "security_descriptor")


class ProtectionTests(_Base):
    def test_protected_range_empty_strings_are_refused_and_none_clears(self) -> None:
        for field in PROTECTED_RANGE_TEXT:
            with self.subTest(field=field):
                with self.assertRaises(InvalidArgumentError):
                    self.ws.add_protected_range(ProtectedRangeInput(name="R", ranges="A1:B2", **{field: ""}))
                self.assertEqual(self.ws.list_protected_ranges(), [])
                self.ws.add_protected_range(ProtectedRangeInput(name="R", ranges="A1:B2", **{field: "x"}))
                self.assertEqual(getattr(self.ws.list_protected_ranges()[0], field), "x")
                with self.assertRaises(InvalidArgumentError):
                    self.ws.update_protected_range("R", **{field: ""})
                self.assertEqual(getattr(self.ws.list_protected_ranges()[0], field), "x")
                self.ws.update_protected_range("R", **{field: None})
                self.assertIsNone(getattr(self.ws.list_protected_ranges()[0], field))
                self.ws.remove_protected_range("R")

    def test_protected_range_omitted_and_none_strings_are_absent(self) -> None:
        self.ws.add_protected_range(ProtectedRangeInput(name="R", ranges="A1:B2"))
        self.ws.add_protected_range(ProtectedRangeInput(name="S", ranges="C1:D2", hash_value=None))
        for protected in self.ws.list_protected_ranges():
            for field in PROTECTED_RANGE_TEXT:
                self.assertIsNone(getattr(protected, field))

    def test_sheet_protection_empty_strings_are_refused_and_none_clears(self) -> None:
        for field in ("algorithm_name", "hash_value", "salt_value", "legacy_password_hash"):
            with self.subTest(field=field):
                with self.assertRaises(InvalidArgumentError):
                    self.ws.set_sheet_protection(SheetProtectionInput(**{field: ""}))
                self.assertIsNone(self.ws.get_sheet_protection())
                self.ws.set_sheet_protection(SheetProtectionInput(**{field: "x"}))
                with self.assertRaises(InvalidArgumentError):
                    self.ws.update_sheet_protection(**{field: ""})
                self.assertEqual(getattr(self.ws.get_sheet_protection(), field), "x")
                self.ws.update_sheet_protection(**{field: None})
                self.assertIsNone(getattr(self.ws.get_sheet_protection(), field))
                self.ws.set_sheet_protection(None)  # type: ignore[arg-type]

    def test_workbook_protection_empty_strings_are_refused_and_none_is_absent(self) -> None:
        fields = (
            "workbook_legacy_password_hash",
            "workbook_algorithm_name",
            "workbook_hash_value",
            "workbook_salt_value",
            "revisions_legacy_password_hash",
            "revisions_algorithm_name",
            "revisions_hash_value",
            "revisions_salt_value",
        )
        for field in fields:
            with self.subTest(field=field):
                with self.assertRaises(InvalidArgumentError):
                    self.wb.set_workbook_protection(WorkbookProtectionInput(lock_structure=True, **{field: ""}))
                self.assertIsNone(self.wb.get_workbook_protection())
                self.wb.set_workbook_protection(WorkbookProtectionInput(lock_structure=True, **{field: None}))
                self.assertIsNone(getattr(self.wb.get_workbook_protection(), field))
                self.wb.set_workbook_protection(None)  # type: ignore[arg-type]


class DxfTests(_Base):
    def test_empty_number_format_is_the_absent_value(self) -> None:
        empty = self.wb.ooxml.add_dxf(DxfInput(number_format=""))
        omitted = self.wb.ooxml.add_dxf(DxfInput())
        self.assertEqual(self.wb.ooxml.get_dxf(empty).number_format, self.wb.ooxml.get_dxf(omitted).number_format)
        self.assertIsNone(self.wb.ooxml.get_dxf(empty).number_format)


if __name__ == "__main__":
    unittest.main()
