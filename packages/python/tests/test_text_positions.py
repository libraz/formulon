"""Text positions are code points in Python; the file and the JavaScript binding keep UTF-16 code units."""

from __future__ import annotations

import io
import re
import unittest
import zipfile

import formulon
from formulon import FormulaSyntaxError, MentionInput, PhoneticRun, ThreadedCommentInput, Workbook

ALICE = "{11111111-1111-4111-8111-111111111111}"
THREAD = "{33333333-3333-4333-8333-333333333333}"
# One astral character (two UTF-16 units) ahead of every span.
EMOJI = "\U0001f600"


class ParseErrorPositionTests(unittest.TestCase):
    def test_start_end_and_column_are_code_points(self) -> None:
        for formula, column in ((f'="{EMOJI}{EMOJI}"+)', 7), (f'="{EMOJI}"&\n1+)', 3)):
            with self.subTest(formula=formula), Workbook.create_default() as wb:
                with self.assertRaises(FormulaSyntaxError) as caught:
                    wb.sheet(0).set_formula(0, 0, formula)
                error = caught.exception.parse_error
                self.assertEqual(formula[error.start : error.end], ")")
                self.assertEqual(error.column, column)


class PhoneticRunPositionTests(unittest.TestCase):
    def test_runs_are_code_points_on_both_sides(self) -> None:
        runs = [PhoneticRun(1, 3, "かんじ")]
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_text(0, 0, f"{EMOJI}漢字")
            ws.set_cell_phonetic_runs(0, 0, runs)
            self.assertEqual(ws.get_cell_phonetic_runs(0, 0), runs)
            # PHONETIC (a Japanese-locale function) substitutes exactly the span the run covers.
            wb.set_profile_id("mac-365-ja_JP")
            ws.set_formula(0, 1, "=PHONETIC(A1)")
            self.assertEqual(ws.get_value(0, 1), f"{EMOJI}かんじ")
            data = wb.save()
        with Workbook.load(data) as reloaded:
            self.assertEqual(reloaded.sheet(0).get_cell_phonetic_runs(0, 0), runs)


class MentionPositionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)
        self.wb.add_person(formulon.PersonInput(id=ALICE, display_name="Alice"))

    def _saved_mention_spans(self) -> list:
        with zipfile.ZipFile(io.BytesIO(self.wb.save())) as package:
            (part,) = [n for n in package.namelist() if n.startswith("xl/threadedComments/")]
            xml = package.read(part).decode("utf-8")
        return [
            (int(start), int(length))
            for start, length in re.findall(r'<mention [^>]*startIndex="(\d+)" length="(\d+)"', xml)
        ]

    def test_mentions_are_code_points_and_the_file_keeps_utf16(self) -> None:
        text = f"{EMOJI} @Alice hi"
        self.ws.add_threaded_comment(
            ThreadedCommentInput(
                id=THREAD,
                cell="A1",
                person_id=ALICE,
                text=text,
                mentions=[MentionInput(person_id=ALICE, start=2, end=8)],
            )
        )
        (mention,) = self.ws.get_threaded_comment(THREAD).mentions
        self.assertEqual((mention.start, mention.end), (2, 8))
        self.assertEqual(text[mention.start : mention.end], "@Alice")
        self.assertEqual(self._saved_mention_spans(), [(3, 6)])

    def test_update_converts_against_the_text_the_comment_ends_with(self) -> None:
        self.ws.add_threaded_comment(ThreadedCommentInput(id=THREAD, cell="A1", person_id=ALICE, text="@Alice"))
        # Mentions alone: the stored text is the one they index.
        self.ws.update_threaded_comment(THREAD, mentions=[MentionInput(person_id=ALICE, start=0, end=6)])
        self.assertEqual(self._saved_mention_spans(), [(0, 6)])
        # New text with mentions: the new text is the one they index.
        text = f"{EMOJI}{EMOJI} @Alice"
        self.ws.update_threaded_comment(THREAD, text=text, mentions=[MentionInput(person_id=ALICE, start=3, end=9)])
        (mention,) = self.ws.get_threaded_comment(THREAD).mentions
        self.assertEqual(text[mention.start : mention.end], "@Alice")
        self.assertEqual(self._saved_mention_spans(), [(5, 6)])
        # A read mention passed back unchanged keeps its place.
        self.ws.update_threaded_comment(THREAD, mentions=[mention])
        self.assertEqual(self.ws.get_threaded_comment(THREAD).mentions, [mention])
        self.assertEqual(self._saved_mention_spans(), [(5, 6)])


if __name__ == "__main__":
    unittest.main()
