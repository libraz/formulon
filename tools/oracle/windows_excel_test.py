#!/usr/bin/env python3
"""Unit tests for the Windows Excel driver's COM-side helpers."""

from __future__ import annotations

import base64
import io
import os
import unittest
from unittest.mock import patch

from tools.oracle import divergence_check
from tools.oracle.drivers import windows_excel


class _FakeVersionApi:
    def __init__(self, version, build):
        self.Version = version
        self.Build = build


class _FakeVersionApp:
    def __init__(self, version, build):
        self.version = version
        self.api = _FakeVersionApi(version, build)


class WindowsVersionStampTest(unittest.TestCase):
    """The product stamp a capture is promoted (or refused) on.

    `Application.Version` is the bare Office major on every SKU, so the
    build has to reach the stamp -- but `promote_capture` refuses anything
    `divergence_check.is_pending_stamp` rejects, and that allows at most
    `16.<n>.<n>`. Every shape the channels report has to land inside it.
    """

    def _stamp(self, version, build) -> str:
        oracle = object.__new__(windows_excel.WindowsExcelOracle)
        oracle._app = _FakeVersionApp(version, build)
        with patch.object(windows_excel, "detect_locale_from_app", return_value="ja-JP"):
            return oracle.probe_environment().excel_version

    def test_bare_build_is_appended(self) -> None:
        self.assertEqual(self._stamp("16.0", "18025"), "16.0.18025")

    def test_build_carrying_a_sub_build_is_cut_to_three_components(self) -> None:
        # This ja-JP Microsoft 365 install reports Build "20228.0"; the
        # naive join produced "16.0.20228.0", which promotion refused.
        self.assertEqual(self._stamp("16.0", "20228.0"), "16.0.20228")

    def test_whole_dotted_build_replaces_the_version_without_doubling(self) -> None:
        self.assertEqual(self._stamp("16.0", "16.0.20228.20190"), "16.0.20228")

    def test_build_already_present_is_not_appended_twice(self) -> None:
        self.assertEqual(self._stamp("16.0.20228", "20228"), "16.0.20228")

    def test_every_stamp_survives_the_promotion_check(self) -> None:
        for version, build in (("16.0", "18025"), ("16.0", "20228.0"), ("16.0", "16.0.20228.20190")):
            with self.subTest(build=build):
                self.assertFalse(divergence_check.is_pending_stamp(self._stamp(version, build)))


class _FakeAxisTarget:
    """A column / row handle whose size property Excel may or may not take."""

    def __init__(self, default: float, *, stubborn: bool = False, snap: float = 0.0) -> None:
        self._value = default
        self._stubborn = stubborn
        self._snap = snap

    @property
    def ColumnWidth(self) -> float:  # noqa: N802 - COM property name
        return self._value

    @ColumnWidth.setter
    def ColumnWidth(self, value: float) -> None:  # noqa: N802 - COM property name
        if not self._stubborn:
            self._value = value - self._snap

    @property
    def RowHeight(self) -> float:  # noqa: N802 - COM property name
        return self._value

    @RowHeight.setter
    def RowHeight(self, value: float) -> None:  # noqa: N802 - COM property name
        if not self._stubborn:
            self._value = value


class _FakeSizedApi:
    def __init__(self, target: _FakeAxisTarget) -> None:
        self._target = target
        self.column_keys: list = []
        self.row_keys: list = []

    def Columns(self, key):  # noqa: N802 - COM method name
        self.column_keys.append(key)
        return self._target

    def Rows(self, key):  # noqa: N802 - COM method name
        self.row_keys.append(key)
        return self._target


class _FakeSizedSheet:
    def __init__(self, target: _FakeAxisTarget) -> None:
        self.api = _FakeSizedApi(target)


class WindowsAxisSizeTest(unittest.TestCase):
    """Column-width / row-height application for print cases.

    A print case's break counts are a function of these sizes, so a size
    that never reached Excel produces a golden that looks ordinary while
    describing a default-width sheet the case never asked for.
    """

    def test_span_key_is_passed_to_columns_verbatim(self) -> None:
        sheet = _FakeSizedSheet(_FakeAxisTarget(8.43))

        windows_excel._apply_axis_size(sheet, "column", "A:H", 28.0)

        self.assertEqual(sheet.api.column_keys, ["A:H"])
        self.assertEqual(sheet.api.Columns("A:H").ColumnWidth, 28.0)

    def test_row_span_key_is_passed_to_rows_verbatim(self) -> None:
        sheet = _FakeSizedSheet(_FakeAxisTarget(18.75))

        windows_excel._apply_axis_size(sheet, "row", "3:5", 60.0)

        self.assertEqual(sheet.api.row_keys, ["3:5"])
        self.assertEqual(sheet.api.Rows("3:5").RowHeight, 60.0)

    def test_size_excel_did_not_take_is_refused_not_swallowed(self) -> None:
        sheet = _FakeSizedSheet(_FakeAxisTarget(8.43, stubborn=True))

        with self.assertRaisesRegex(RuntimeError, "does not match"):
            windows_excel._apply_axis_size(sheet, "column", "A:H", 28.0)

    def test_snapping_within_tolerance_is_accepted(self) -> None:
        # Excel rounds a size to its own grid; that is not a failed apply.
        sheet = _FakeSizedSheet(_FakeAxisTarget(8.43, snap=0.14))

        windows_excel._apply_axis_size(sheet, "column", "A", 28.0)

        self.assertAlmostEqual(sheet.api.Columns("A").ColumnWidth, 27.86)


class _FakePrinterApi:
    """`Application` as far as the printer pin is concerned.

    Models the workbook gate Excel enforces: assigning `ActivePrinter`
    while no book is open raises "cannot set the ActivePrinter property
    of the Application class" regardless of how the device is spelled.
    Without that gate a fake accepts a pin the real Application refuses,
    which is exactly how the print track came to capture every suite
    against the host's default device while its tests stayed green.
    """

    def __init__(self, accepted, current: str = "FUJIFILM Apeos C2360 on Ne01:", *, books_open: int = 1) -> None:
        self._accepted = set(accepted)
        self._current = current
        self.books_open = books_open
        self.attempts: list = []

    @property
    def ActivePrinter(self) -> str:  # noqa: N802 - COM property name
        return self._current

    @ActivePrinter.setter
    def ActivePrinter(self, value: str) -> None:  # noqa: N802 - COM property name
        self.attempts.append(value)
        if not self.books_open:
            raise RuntimeError("Cannot set the ActivePrinter property of the Application class")
        if value not in self._accepted:
            raise RuntimeError(f"Excel rejects {value}")
        self._current = value


class WindowsPrinterPinTest(unittest.TestCase):
    """Which printer's metrics a print capture paginates against.

    Excel derives automatic breaks from the active printer driver, so an
    unpinned capture records whichever device the host happened to
    default to -- and a network device that renegotiates its capabilities
    moves the breaks between two otherwise identical runs.
    """

    def _oracle(self, accepted, *, books_open: int = 1, **env):
        oracle = object.__new__(windows_excel.WindowsExcelOracle)
        oracle._printer_pinned = False
        oracle._printer_warned = False
        oracle._app = type("_App", (), {"api": _FakePrinterApi(accepted, books_open=books_open)})()
        self._env = patch.dict(os.environ, env, clear=False)
        self._env.start()
        self.addCleanup(self._env.stop)
        if not env:
            os.environ.pop(windows_excel._PRINTER_ENV_VAR, None)
        return oracle

    def test_default_device_is_pinned_when_nothing_is_requested(self) -> None:
        oracle = self._oracle({"Microsoft Print to PDF on Ne00:"})

        oracle._ensure_printer_pinned()

        self.assertEqual(oracle._app.api.ActivePrinter, "Microsoft Print to PDF on Ne00:")

    def test_port_suffix_is_searched_because_it_differs_per_machine(self) -> None:
        oracle = self._oracle({"Microsoft Print to PDF on Ne03:"})

        oracle._ensure_printer_pinned()

        self.assertEqual(oracle._app.api.ActivePrinter, "Microsoft Print to PDF on Ne03:")
        self.assertGreater(len(oracle._app.api.attempts), 1)

    def test_requested_device_excel_rejects_is_an_error(self) -> None:
        oracle = self._oracle(set(), **{windows_excel._PRINTER_ENV_VAR: "No Such Printer"})

        with self.assertRaisesRegex(RuntimeError, "names no printer"):
            oracle._ensure_printer_pinned()

    def test_absent_default_falls_back_to_the_host_device(self) -> None:
        oracle = self._oracle(set())

        oracle._ensure_printer_pinned()

        # No exception: an unset variable is a preference, not a demand.
        self.assertEqual(oracle._app.api.ActivePrinter, "FUJIFILM Apeos C2360 on Ne01:")

    def test_pin_is_attempted_once_per_session(self) -> None:
        oracle = self._oracle({"Microsoft Print to PDF on Ne00:"})

        oracle._ensure_printer_pinned()
        before = len(oracle._app.api.attempts)
        oracle._ensure_printer_pinned()

        self.assertEqual(len(oracle._app.api.attempts), before)

    def test_a_pin_no_workbook_gated_away_is_retried(self) -> None:
        """A failed attempt must not count as a pinned session.

        `ActivePrinter` is workbook-gated, so a pin attempted before
        `books.add()` fails for every candidate. Latching the "pinned"
        flag on the attempt rather than on a device that took is what let
        that failure stand for a whole capture: the fallback warning
        printed once and every subsequent print case paginated against
        the host default without retrying.
        """

        oracle = self._oracle({"Microsoft Print to PDF on Ne00:"}, books_open=0)

        oracle._ensure_printer_pinned()
        self.assertFalse(oracle._printer_pinned)
        self.assertEqual(oracle._app.api.ActivePrinter, "FUJIFILM Apeos C2360 on Ne01:")

        oracle._app.api.books_open = 1
        oracle._ensure_printer_pinned()

        self.assertTrue(oracle._printer_pinned)
        self.assertEqual(oracle._app.api.ActivePrinter, "Microsoft Print to PDF on Ne00:")

    def test_a_print_case_pins_after_it_has_opened_its_workbook(self) -> None:
        """Ordering, not just retry: the gate makes `books.add()` a prerequisite.

        `_run_print_case` used to pin before opening its workbook, so the
        pin could never take on any host -- the capture ran against the
        default device and said so only on a stderr channel the WSL
        bridge swallowed.
        """

        oracle = self._oracle({"Microsoft Print to PDF on Ne00:"}, books_open=0)
        api = oracle._app.api

        class _FakeBooks:
            @staticmethod
            def add():
                api.books_open += 1
                return type("_Wb", (), {"close": staticmethod(lambda: None)})()

        oracle._app.books = _FakeBooks()

        with (
            patch.object(windows_excel.WindowsExcelOracle, "_build_workbook_sheets", lambda self, wb, case: None),
            patch.object(windows_excel, "_apply_and_read_print", lambda wb, spec: {"pages": 1}),
        ):
            oracle._run_print_case({"id": "c"}, {"sheet": "Sheet1"})

        self.assertEqual(api.ActivePrinter, "Microsoft Print to PDF on Ne00:")

    def test_a_roundtrip_case_pins_after_it_has_opened_its_workbook(self) -> None:
        """The same ordering, on the capture path that opens a file.

        A round-trip case reaches its workbook through `books.open`
        rather than `books.add()`, so it needs its own guard: the gate is
        about whether a book is open, not about how it got there.
        """

        oracle = self._oracle({"Microsoft Print to PDF on Ne00:"}, books_open=0)
        api = oracle._app.api

        class _FakeBooks:
            @staticmethod
            def open(path):  # noqa: A003 - mirrors the xlwings method name
                api.books_open += 1
                return type(
                    "_Wb",
                    (),
                    {"sheets": {"Sheet1": object()}, "close": staticmethod(lambda: None)},
                )()

        oracle._app.books = _FakeBooks()

        with patch.object(windows_excel, "_read_roundtrip", lambda sht: {}):
            oracle._run_roundtrip_case(
                {"id": "c"},
                {"sheet": "Sheet1", "xlsx_base64": base64.b64encode(b"not-a-real-xlsx").decode("ascii")},
            )

        self.assertEqual(api.ActivePrinter, "Microsoft Print to PDF on Ne00:")

    def test_the_fallback_warning_is_not_repeated_per_case(self) -> None:
        oracle = self._oracle(set())

        with patch("sys.stderr", new_callable=io.StringIO) as err:
            oracle._ensure_printer_pinned()
            oracle._ensure_printer_pinned()

        self.assertEqual(err.getvalue().count("is unavailable"), 1)


class _FakeBreakCollection:
    """`HPageBreaks` / `VPageBreaks`: a callable with a `.Count`."""

    def __init__(self, locations, axis: str) -> None:
        self._locations = list(locations)
        self._axis = axis

    @property
    def Count(self) -> int:  # noqa: N802 - COM property name
        return len(self._locations)

    def __call__(self, index: int):
        one_based = self._locations[index - 1] + 1
        return type("_Loc", (), {"Location": type("_Cell", (), {self._axis: one_based})()})()


class _FakePaginationSheet:
    """A worksheet whose successive break reads follow a script."""

    def __init__(self, readings) -> None:
        self._readings = list(readings)
        self._index = 0
        self.settles = 0

    def _current(self):
        return self._readings[min(self._index, len(self._readings) - 1)]

    @property
    def HPageBreaks(self):  # noqa: N802 - COM property name
        return _FakeBreakCollection(self._current()[0], "Row")

    @property
    def VPageBreaks(self):  # noqa: N802 - COM property name
        return _FakeBreakCollection(self._current()[1], "Column")

    def Calculate(self) -> None:  # noqa: N802 - COM method name
        # The settle between reads is what advances the script.
        self.settles += 1
        self._index += 1


class _FakePagesPageSetup:
    def __init__(self, sheet: _FakePaginationSheet) -> None:
        self._sheet = sheet

    @property
    def Pages(self):  # noqa: N802 - COM property name
        return type("_Pages", (), {"Count": self._sheet._current()[2]})()


class _FakePaginationApp:
    class _Api:
        ScreenUpdating = True

        @staticmethod
        def ExecuteExcel4Macro(_expression):  # noqa: N802 - COM method name
            raise RuntimeError("unused")

    api = _Api()


class _FakePaginationWorkbook:
    app = type("_App", (), {"api": _FakePaginationApp._Api()})()


class WindowsPaginationSettleTest(unittest.TestCase):
    """The guard against a stale page-break read.

    Excel answers the break collections from its last layout pass, which
    is not always the current case's: a reading that does not reproduce
    would otherwise be written into a golden that looks ordinary.
    """

    def setUp(self) -> None:
        # The real settle waits for Excel's layout thread; the fakes here
        # answer instantly, so the wait is only dead time.
        patcher = patch.object(windows_excel, "_LAYOUT_SETTLE_SECONDS", 0.0)
        patcher.start()
        self.addCleanup(patcher.stop)

    def _read(self, readings):
        sheet = _FakePaginationSheet(readings)
        value = windows_excel._read_pagination_settled(sheet, _FakePagesPageSetup(sheet), _FakePaginationWorkbook())
        return value, sheet

    def test_stable_case_is_read_twice_and_returned(self) -> None:
        value, sheet = self._read([([39], [], 2)])

        self.assertEqual(value, ([39], [], 2))
        self.assertEqual(sheet.settles, 1)

    def test_stale_first_reading_is_discarded_for_the_one_that_repeats(self) -> None:
        # The observed shape: the preceding case's break, then this case's.
        value, _sheet = self._read([([26], [], 2), ([39], [], 2), ([39], [], 2)])

        self.assertEqual(value, ([39], [], 2))

    def test_reading_that_never_repeats_is_refused(self) -> None:
        never_settles = [([i], [], 2) for i in range(windows_excel._BREAK_READ_ATTEMPTS + 2)]

        with self.assertRaisesRegex(RuntimeError, "did not settle"):
            self._read(never_settles)

    def test_break_locations_are_zero_based_and_sorted(self) -> None:
        value, _sheet = self._read([([39, 12], [7, 2], 6)])

        self.assertEqual(value, ([12, 39], [2, 7], 6))


class _FakeResultCell:
    """An xlwings-like result cell: a value read plus an external address."""

    def __init__(self, value, address="'[Book1]Sheet1'!$A$1", offsets=None) -> None:
        self.value = value
        self.api = None
        self._address = address
        self._offsets = offsets or {}

    def get_address(self, *, external=False):
        if not external:
            raise AssertionError("the oracle must request an external address")
        return self._address

    def offset(self, row, column):
        return self._offsets[(row, column)]


class _FakeEvaluateApi:
    def __init__(self, answers) -> None:
        self._answers = answers
        self.calls = []

    def Evaluate(self, expression):  # noqa: N802 - COM method name
        self.calls.append(expression)
        return self._answers(expression)


class _FakeEvaluateApp:
    def __init__(self, answers) -> None:
        self.api = _FakeEvaluateApi(answers)


class WindowsEmptyTextReadbackTest(unittest.TestCase):
    """COM reads "" and a blank cell alike; TYPE() tells them apart, as on Mac."""

    def test_empty_read_back_is_split_into_empty_text_and_blank_by_type(self) -> None:
        for value in ("", None):
            for type_code, expected in ((2.0, ("text", "")), (1.0, ("blank", None))):
                with self.subTest(value=value, type_code=type_code):
                    app = _FakeEvaluateApp(lambda _expression, code=type_code: code)
                    evaluate = windows_excel._app_evaluate(app)
                    result = windows_excel._classify_value(_FakeResultCell(value), evaluate)
                    self.assertEqual((result.kind, result.value), expected)
                    self.assertEqual(app.api.calls, ["TYPE('[Book1]Sheet1'!$A$1)"])

    def test_non_empty_value_does_not_probe_type(self) -> None:
        def evaluate(expression):
            raise AssertionError(f"unexpected probe: {expression}")

        for value, kind in ((0, "number"), ("x", "text"), (False, "bool")):
            with self.subTest(value=value):
                self.assertEqual(windows_excel._classify_value(_FakeResultCell(value), evaluate).kind, kind)

    def test_without_an_evaluator_an_empty_read_back_stays_blank(self) -> None:
        self.assertEqual(windows_excel._classify_value(_FakeResultCell(None)).kind, "blank")

    def test_spilled_empty_text_cells_are_probed_per_cell(self) -> None:
        blank = _FakeResultCell(None, address="'[Book1]Sheet1'!$A$2")
        text = _FakeResultCell("", address="'[Book1]Sheet1'!$A$1")
        anchor = _FakeResultCell("", offsets={(0, 0): text, (1, 0): blank})

        def answers(expression):
            if expression.startswith("TYPE("):
                return 2.0 if "$A$1" in expression else 1.0
            raise AssertionError(expression)

        app = _FakeEvaluateApp(answers)

        class _Sheet:
            def range(self, _addr):
                return anchor

        with patch.object(windows_excel, "_evaluate_spill_shape", return_value=(2, 1)):
            result = windows_excel._classify_result_cell(app, _Sheet(), "A1")
        self.assertEqual((result.kind, result.value, result.array_shape), ("array", ["", None], [2, 1]))


if __name__ == "__main__":
    unittest.main()
