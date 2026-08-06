#!/usr/bin/env python3
"""Unit tests for the cross-language parity gate's outcome semantics."""

from __future__ import annotations

import contextlib
import io
import sys
import unittest
from pathlib import Path
from typing import Dict

sys.path.insert(0, str(Path(__file__).resolve().parent))

import run_parity


class StubChannel(run_parity.Channel):
    def __init__(self, name: str, result: run_parity.ChannelResult, active: bool = True) -> None:
        self.name = name
        self._result = result
        self._active = active

    def available(self) -> bool:
        return self._active

    def availability_reason(self) -> str:
        return "test channel"

    def evaluate(self, formula: str) -> run_parity.ChannelResult:
        del formula
        return self._result


def number_fixture(value: float = 6) -> Dict[str, object]:
    return {"id": "sum", "formula": "=SUM(1,2,3)", "expect": {"kind": "number", "value": value}}


def number_result(value: float) -> run_parity.ChannelResult:
    return run_parity.ChannelResult(run_parity.make_number_record(value), value, None)


class ParityRunnerTest(unittest.TestCase):
    def run_gate(self, fixtures: list[Dict[str, object]], channels: list[StubChannel]) -> tuple[int, str]:
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            status = run_parity.run(fixtures, channels, verbose=False)
        return status, output.getvalue()

    def test_single_channel_is_skipped(self) -> None:
        status, output = self.run_gate([number_fixture()], [StubChannel("cli", number_result(6))])
        self.assertEqual(status, run_parity.SKIP_RETURN_CODE)
        self.assertIn("SKIPPED", output)

    def test_channel_failure_fails_gate(self) -> None:
        failure = run_parity.ChannelResult(None, None, "transport failed")
        status, output = self.run_gate(
            [number_fixture()], [StubChannel("cli", number_result(6)), StubChannel("npm", failure)]
        )
        self.assertEqual(status, 1)
        self.assertIn("channel-failure", output)

    def test_fixture_expectation_is_checked_for_each_channel(self) -> None:
        status, output = self.run_gate(
            [number_fixture(6)], [StubChannel("cli", number_result(7)), StubChannel("npm", number_result(7))]
        )
        self.assertEqual(status, 1)
        self.assertIn("expectation-failure", output)

    def test_divergent_channel_fails_gate(self) -> None:
        status, output = self.run_gate(
            [number_fixture(6)], [StubChannel("cli", number_result(6)), StubChannel("npm", number_result(7))]
        )
        self.assertEqual(status, 1)
        self.assertIn("DIVERGENCE", output)

    def test_matching_channels_and_expectation_pass(self) -> None:
        status, output = self.run_gate(
            [number_fixture()], [StubChannel("cli", number_result(6)), StubChannel("npm", number_result(6))]
        )
        self.assertEqual(status, 0)
        self.assertIn("expectation_failures=0", output)


if __name__ == "__main__":
    unittest.main()
