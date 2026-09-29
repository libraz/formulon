"""Deterministic call-count tests for the Python diagnostic bridge."""

from __future__ import annotations

import threading
import unittest
from typing import Optional

import wasmtime

from formulon import _c


class _Export:
    def __init__(self, result: int, calls: dict[str, int], name: Optional[str] = None) -> None:
        self.result = result
        self.calls = calls
        self.name = name

    def __call__(self, _store, *_args: int) -> int:
        if self.name is not None:
            self.calls[self.name] += 1
        return self.result


class _TrappingExport:
    """A raw export stand-in that raises like an aborted WASM call."""

    def __init__(self, message: str) -> None:
        self.message = message
        self.calls = 0

    def __call__(self, _store, *_args: int) -> int:
        self.calls += 1
        raise wasmtime.Trap(self.message)


def _fake_instance(function_name: str, function_result: int) -> tuple[_c._WasmInstance, dict[str, int]]:
    calls = {"message": 0, "context": 0}
    exports = {
        function_name: _Export(function_result, calls),
        "fm_last_error_message": _Export(0x100, calls, "message"),
        "fm_last_error_context": _Export(0x200, calls, "context"),
    }
    instance = object.__new__(_c._WasmInstance)
    instance._engine = None
    instance._store = object()
    instance._instance = object()
    instance._memory = object()
    instance._exports = exports
    instance._init_lock = threading.Lock()
    instance._call_lock = threading.RLock()
    instance._last_diagnostic = threading.local()
    instance._poison_message = None
    instance._read_cstr_unlocked = lambda ptr: {0x100: "captured message", 0x200: "captured context"}[ptr]
    return instance, calls


class DiagnosticCaptureTests(unittest.TestCase):
    def test_non_status_nonzero_result_does_not_read_or_clear_pending(self) -> None:
        instance, calls = _fake_instance("fm_workbook_sheet_count", 17)
        instance._last_diagnostic.value = (99, "pending message", "pending context")

        self.assertEqual(instance.fm_workbook_sheet_count(0), 17)
        self.assertEqual(calls, {"message": 0, "context": 0})
        self.assertEqual(instance.last_diagnostic(99), ("pending message", "pending context"))

    def test_pointer_status_string_does_not_read_or_clear_pending(self) -> None:
        instance, calls = _fake_instance("fm_status_string", 0x1234)
        instance._last_diagnostic.value = (7, "pending message", "pending context")

        self.assertEqual(instance.fm_status_string(7), 0x1234)
        self.assertEqual(calls, {"message": 0, "context": 0})
        self.assertEqual(instance.last_diagnostic(7), ("pending message", "pending context"))

    def test_unknown_export_fails_closed(self) -> None:
        instance, calls = _fake_instance("fm_unknown_export", 1)
        instance._last_diagnostic.value = (1, "pending message", "pending context")

        self.assertEqual(instance.fm_unknown_export(), 1)
        self.assertEqual(calls, {"message": 0, "context": 0})
        self.assertEqual(instance.last_diagnostic(1), ("pending message", "pending context"))

    def test_status_failure_captures_both_diagnostics_once_and_take_is_one_shot(self) -> None:
        instance, calls = _fake_instance("fm_workbook_remove_sheet", 5050)

        self.assertEqual(instance.fm_workbook_remove_sheet(0, 99), 5050)
        self.assertEqual(calls, {"message": 1, "context": 1})
        self.assertEqual(instance.last_diagnostic(5050), ("captured message", "captured context"))
        self.assertEqual(instance.last_diagnostic(5050), ("", ""))
        self.assertEqual(calls, {"message": 1, "context": 1})

    def test_status_success_clears_prior_pending_without_reads(self) -> None:
        instance, calls = _fake_instance("fm_workbook_remove_sheet", 0)
        instance._last_diagnostic.value = (5050, "stale message", "stale context")

        self.assertEqual(instance.fm_workbook_remove_sheet(0, 99), 0)
        self.assertEqual(calls, {"message": 0, "context": 0})
        self.assertEqual(instance.last_diagnostic(5050), ("", ""))

    def test_absent_or_mismatched_take_does_not_raw_read(self) -> None:
        instance, calls = _fake_instance("fm_workbook_remove_sheet", 0)

        self.assertEqual(instance.last_diagnostic(5050), ("", ""))
        instance._last_diagnostic.value = (5050, "message", "context")
        self.assertEqual(instance.last_diagnostic(0), ("", ""))
        self.assertEqual(calls, {"message": 0, "context": 0})


class WasmTrapPoisoningTests(unittest.TestCase):
    """A `wasmtime.Trap` (e.g. an OOM abort) must surface as
    `FormulonError`, not the raw trap type, and must poison the shared
    instance so later calls fail the same way instead of touching the
    module again."""

    def test_trap_during_export_call_raises_formulon_error(self) -> None:
        instance, _calls = _fake_instance("fm_workbook_recalc", 0)
        instance._exports["fm_workbook_recalc"] = _TrappingExport("out of memory")

        with self.assertRaises(_c.FormulonError) as caught:
            instance.fm_workbook_recalc(0)
        self.assertEqual(caught.exception.status_name, "kInternalError")
        self.assertIn("out of memory", caught.exception.message)
        self.assertIn("fm_workbook_recalc", str(caught.exception))

    def test_trap_poisons_instance_for_every_later_call(self) -> None:
        instance, _calls = _fake_instance("fm_workbook_recalc", 0)
        trapping = _TrappingExport("out of memory")
        instance._exports["fm_workbook_recalc"] = trapping
        other_calls: dict[str, int] = {"other": 0}
        instance._exports["fm_workbook_sheet_count"] = _Export(0, other_calls, "other")

        with self.assertRaises(_c.FormulonError):
            instance.fm_workbook_recalc(0)
        self.assertEqual(trapping.calls, 1)

        # A second call, even to an export that never itself traps, is
        # rejected without touching the (possibly corrupted) module.
        with self.assertRaises(_c.FormulonError) as caught:
            instance.fm_workbook_sheet_count(0)
        self.assertEqual(other_calls, {"other": 0})
        self.assertEqual(caught.exception.status_name, "kInternalError")

    def test_trap_from_malloc_raises_formulon_error_not_memory_error(self) -> None:
        instance, _calls = _fake_instance("fm_workbook_recalc", 0)
        instance._exports["malloc"] = _TrappingExport("out of memory")

        with self.assertRaises(_c.FormulonError) as caught:
            instance.alloc(64)
        self.assertEqual(caught.exception.status_name, "kInternalError")

        # The instance is poisoned: `free` also fails closed rather than
        # calling into the module again.
        instance._exports["free"] = _Export(0, {"free": 0}, "free")
        with self.assertRaises(_c.FormulonError):
            instance.free(0x1000)


if __name__ == "__main__":
    unittest.main()
