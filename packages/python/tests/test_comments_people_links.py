"""Notes, hyperlinks, people and threaded-comment mentions through the public API."""

from __future__ import annotations

import re
import unittest

import formulon
from formulon import (
    CellAddress,
    CellRect,
    Comment,
    CommentInput,
    FormulonError,
    HyperlinkInput,
    StatusCode,
    ThreadedCommentInput,
    Workbook,
)

ALICE = "{11111111-1111-4111-8111-111111111111}"
THREAD = "{33333333-3333-4333-8333-333333333333}"
GUID = re.compile(r"^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$")


class _Base(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)


class NoteTests(_Base):
    def test_set_comment_none_removes(self) -> None:
        self.ws.set_comment("B2", Comment("hello", author="ann"))
        self.assertEqual(self.ws.get_comment("B2"), Comment("hello", author="ann"))
        self.ws.set_comment("B2", None)  # type: ignore[arg-type]
        self.assertIsNone(self.ws.get_comment("B2"))
        self.ws.set_comment(1, 1, None)  # type: ignore[arg-type]
        self.assertIsNone(self.ws.get_comment(1, 1))

    def test_removed_members_are_gone(self) -> None:
        self.assertFalse(hasattr(self.ws, "remove_comment"))
        self.assertFalse(hasattr(self.ws, "remove_hyperlink"))
        self.assertFalse(hasattr(self.ws, "set_hyperlink_range"))

    def test_set_cell_phonetic_properties_none_resets(self) -> None:
        self.ws.set_value("A1", "text")
        self.ws.set_cell_phonetic_properties("A1", formulon.PhoneticProperties(font_index=0, stated=True))
        self.assertTrue(self.ws.get_cell_phonetic_properties("A1").stated)
        self.ws.set_cell_phonetic_properties("A1", None)  # type: ignore[arg-type]
        self.assertFalse(self.ws.get_cell_phonetic_properties("A1").stated)

    def test_set_comment_takes_comment_input_and_a_read_comment(self) -> None:
        self.ws.set_comment("A1", CommentInput("plain"))
        self.assertEqual(self.ws.get_comment("A1"), Comment("plain", author=""))
        self.ws.set_comment(0, 1, CommentInput("signed", "ann"))
        self.assertEqual(self.ws.get_comment(0, 1), Comment("signed", author="ann"))
        read = self.ws.get_comment(0, 1)
        assert read is not None
        self.ws.set_comment("C1", read)
        self.assertEqual(self.ws.get_comment("C1"), read)

    def test_set_comment_rejects_another_record(self) -> None:
        with self.assertRaisesRegex(TypeError, r"must be a CommentInput, a Comment or null"):
            self.ws.set_comment("A1", HyperlinkInput(target="https://example.test/"))  # type: ignore[arg-type]

    def test_list_comments_cell(self) -> None:
        self.ws.set_comment("B2", Comment("hi", author="ann"))
        (entry,) = self.ws.list_comments()
        self.assertEqual(entry.cell, CellAddress(1, 1))
        self.assertEqual((entry.author, entry.text), ("ann", "hi"))
        self.assertFalse(hasattr(entry, "row"))


class HyperlinkTests(_Base):
    def test_range_kept(self) -> None:
        self.ws.set_hyperlink("A1", HyperlinkInput(target="https://example.com/", range="A1:B2"))
        link = self.ws.get_hyperlink("A1")
        self.assertEqual(link.range, CellRect(0, 0, 1, 1))
        self.ws.set_hyperlink("A1", link)
        self.assertEqual(self.ws.get_hyperlink("A1").range, CellRect(0, 0, 1, 1))
        self.ws.set_hyperlink("D4", HyperlinkInput(target="https://example.com/"))
        self.assertEqual(self.ws.get_hyperlink("D4").range, CellRect(3, 3, 3, 3))

    def test_range_not_at_key_invalid(self) -> None:
        with self.assertRaises(FormulonError) as caught:
            self.ws.set_hyperlink("A1", HyperlinkInput(target="https://example.com/", range="B2:C3"))
        self.assertEqual(caught.exception.code, StatusCode.INVALID_ARGUMENT)
        self.assertIsNone(self.ws.get_hyperlink("A1"))

    def test_none_removes(self) -> None:
        self.ws.set_hyperlink("A1", HyperlinkInput(target="https://example.com/"))
        self.ws.set_hyperlink("A1", None)  # type: ignore[arg-type]
        self.assertIsNone(self.ws.get_hyperlink("A1"))


class PeopleTests(_Base):
    def test_add_person_without_id_generates_guid(self) -> None:
        stored = self.wb.add_person(formulon.PersonInput(display_name="Zed"))
        self.assertRegex(stored, GUID)
        self.assertEqual(self.wb.get_person(stored).display_name, "Zed")
        self.assertEqual(self.wb.add_person(formulon.PersonInput(id=ALICE, display_name="Alice")), ALICE)
        self.assertNotEqual(stored, self.wb.add_person(formulon.PersonInput(display_name="Zed")))

    def test_add_person_bad_id_invalid_argument(self) -> None:
        with self.assertRaises(FormulonError) as caught:
            self.wb.add_person(formulon.PersonInput(id="not-a-guid", display_name="Zed"))
        self.assertEqual(caught.exception.code, StatusCode.INVALID_ARGUMENT)

    def test_absent_ids_none(self) -> None:
        stored = self.wb.add_person(formulon.PersonInput(display_name="Zed"))
        person = self.wb.get_person(stored)
        self.assertIsNone(person.user_id)
        self.assertIsNone(person.provider_id)
        self.assertEqual([p.user_id for p in self.wb.list_people()], [None])


class ThreadedCommentTests(_Base):
    def setUp(self) -> None:
        super().setUp()
        self.wb.add_person(formulon.PersonInput(id=ALICE, display_name="Alice"))

    def test_mention_end(self) -> None:
        text = "@Alice please look"
        self.ws.add_threaded_comment(
            ThreadedCommentInput(
                id=THREAD,
                cell="B2",
                person_id=ALICE,
                text=text,
                mentions=[formulon.MentionInput(person_id=ALICE, start=0, end=6)],
            )
        )
        (mention,) = self.ws.get_threaded_comment(THREAD).mentions
        self.assertEqual((mention.start, mention.end), (0, 6))
        self.assertEqual(text[mention.start : mention.end], "@Alice")
        self.assertFalse(hasattr(mention, "length"))

    def test_mention_end_before_start_invalid(self) -> None:
        with self.assertRaises(FormulonError) as caught:
            self.ws.add_threaded_comment(
                ThreadedCommentInput(
                    id=THREAD,
                    cell="B2",
                    person_id=ALICE,
                    text="@Alice",
                    mentions=[formulon.MentionInput(person_id=ALICE, start=4, end=2)],
                )
            )
        self.assertEqual(caught.exception.code, StatusCode.INVALID_ARGUMENT)

    def test_mention_input_generates_id(self) -> None:
        self.ws.add_threaded_comment(
            ThreadedCommentInput(
                id=THREAD,
                cell="B2",
                person_id=ALICE,
                text="@Alice hi",
                mentions=[formulon.MentionInput(person_id=ALICE, start=0, end=6)],
            )
        )
        (mention,) = self.ws.get_threaded_comment(THREAD).mentions
        self.assertRegex(mention.mention_id, GUID)
        # A read Mention feeds back into the update.
        self.ws.update_threaded_comment(THREAD, mentions=[mention])
        self.assertEqual(self.ws.get_threaded_comment(THREAD).mentions, [mention])


if __name__ == "__main__":
    unittest.main()
