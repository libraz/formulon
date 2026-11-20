"""Conditional-format rules addressed by position, and priority placement, on a file whose two rules share one id."""

from __future__ import annotations

import dataclasses
import unittest
from pathlib import Path
from typing import List, Tuple

from formulon import (
    CfExpressionRuleInput,
    IndexOutOfRangeError,
    InvalidArgumentError,
    Sheet,
    Workbook,
)

_FIXTURE = Path(__file__).resolve().parents[3] / "tests/fixtures/excel/cf_x14_shared_id.xlsx"
SHARED_ID = "{00000000-000E-0000-0000-000001000000}"
INT32_MAX = 2147483647


def _view(ws: Sheet) -> List[Tuple[str, int, int, bool]]:
    return [(r.id, r.priority, r.ranges[0].first_col, r.stop_if_true) for r in ws.list_cf_rules()]


def _priorities(ws: Sheet) -> List[int]:
    return [r.priority for r in ws.list_cf_rules()]


def _by_col(ws: Sheet, col: int) -> Tuple[str, int, int, bool]:
    return next(v for v in _view(ws) if v[2] == col)


def _expression(**extra: object) -> CfExpressionRuleInput:
    return CfExpressionRuleInput(ranges="D1:D10", formula1="TRUE", **extra)  # type: ignore[arg-type]


class _Base(unittest.TestCase):
    """Two data-bar rules sharing one id: A1:A10 at priority 1 and B1:B10 at priority 2."""

    def setUp(self) -> None:
        self.wb = Workbook.load(_FIXTURE)
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)


class PositionAddressingTests(_Base):
    def test_fixture_lists_two_rules_sharing_an_id(self) -> None:
        self.assertEqual([(i, p, c) for i, p, c, _ in _view(self.ws)], [(SHARED_ID, 1, 0), (SHARED_ID, 2, 1)])

    def test_update_with_position_changes_that_rule_not_the_first_sharing_its_id(self) -> None:
        before = _view(self.ws)
        self.ws.update_cf_rule(1, ranges="F1:F10")
        after = _view(self.ws)
        self.assertEqual(after[0], before[0])
        self.assertEqual(after[1][:3], (SHARED_ID, 2, 5))

    def test_update_with_position_zero_changes_the_first_rule_only(self) -> None:
        before = _view(self.ws)
        self.ws.update_cf_rule(0, ranges="G1:G10")
        after = _view(self.ws)
        self.assertEqual(after[0][2], 6)
        self.assertEqual(after[1], before[1])

    def test_remove_position_zero_leaves_the_second_with_its_priority_and_shared_id(self) -> None:
        self.ws.remove_cf_rule(0)
        self.assertEqual([v[:3] for v in _view(self.ws)], [(SHARED_ID, 2, 1)])

    def test_remove_position_one_leaves_the_first_rule(self) -> None:
        self.ws.remove_cf_rule(1)
        self.assertEqual([v[:3] for v in _view(self.ws)], [(SHARED_ID, 1, 0)])

    def test_an_id_names_the_first_of_the_rules_sharing_it(self) -> None:
        got = self.ws.get_cf_rule(SHARED_ID)
        assert got is not None
        self.assertEqual(got.ranges[0].first_col, 0)
        self.ws.update_cf_rule(SHARED_ID, stop_if_true=True)
        self.assertTrue(_by_col(self.ws, 0)[3])
        self.assertFalse(_by_col(self.ws, 1)[3])
        self.ws.remove_cf_rule(SHARED_ID)
        self.assertEqual([v[2] for v in _view(self.ws)], [1])
        # Once the first is gone the id reaches the remaining one.
        self.ws.remove_cf_rule(SHARED_ID)
        self.assertEqual(self.ws.list_cf_rules(), [])


class PositionErrorTests(_Base):
    def test_a_position_outside_the_list_is_index_out_of_range_and_changes_nothing(self) -> None:
        before = _view(self.ws)
        calls = {
            "update": lambda p: self.ws.update_cf_rule(p, stop_if_true=True),
            "remove": lambda p: self.ws.remove_cf_rule(p),
            "priority": lambda p: self.ws.set_cf_rule_priority(p, 1),
        }
        for position in (2, 5, -1):
            for name, call in calls.items():
                with self.subTest(call=name, position=position):
                    with self.assertRaises(IndexOutOfRangeError):
                        call(position)
        self.assertEqual(_view(self.ws), before)

    def test_a_bool_is_not_a_position(self) -> None:
        before = _view(self.ws)
        with self.assertRaises(TypeError):
            self.ws.update_cf_rule(True, stop_if_true=True)  # type: ignore[arg-type]
        with self.assertRaises(TypeError):
            self.ws.remove_cf_rule(False)  # type: ignore[arg-type]
        with self.assertRaises(TypeError):
            self.ws.set_cf_rule_priority(True, 1)  # type: ignore[arg-type]
        self.assertEqual(_view(self.ws), before)


class DuplicatePriorityTests(_Base):
    def test_explicit_duplicate_priority_is_accepted_on_add_and_update(self) -> None:
        rule_id = self.ws.add_cf_rule(_expression(priority=2))
        got = self.ws.get_cf_rule(rule_id)
        assert got is not None
        self.assertEqual(got.priority, 2)
        self.assertEqual(_priorities(self.ws), [1, 2, 2])
        self.ws.update_cf_rule(1, priority=1)
        self.assertEqual(_priorities(self.ws), [1, 1, 2])

    def test_a_listed_rule_added_back_with_its_id_cleared_keeps_its_fields(self) -> None:
        copy = dataclasses.replace(self.ws.list_cf_rules()[0], id=None)  # type: ignore[arg-type]
        rule_id = self.ws.add_cf_rule(copy)
        self.assertNotEqual(rule_id, SHARED_ID)
        added = self.ws.get_cf_rule(rule_id)
        assert added is not None
        self.assertEqual(added.priority, 1)
        self.assertEqual(added.ranges[0].first_col, 0)
        self.assertEqual(len(self.ws.list_cf_rules()), 3)

    def test_an_automatic_priority_is_one_past_the_highest(self) -> None:
        rule_id = self.ws.add_cf_rule(_expression())
        got = self.ws.get_cf_rule(rule_id)
        assert got is not None
        self.assertEqual(got.priority, 3)


class SetPriorityTests(_Base):
    def test_priority_one_shifts_every_other_rule_up(self) -> None:
        rule_id = self.ws.add_cf_rule(_expression())
        self.ws.set_cf_rule_priority(rule_id, 1)
        got = self.ws.get_cf_rule(rule_id)
        assert got is not None
        self.assertEqual(got.priority, 1)
        self.assertEqual(_by_col(self.ws, 0)[1], 2)
        self.assertEqual(_by_col(self.ws, 1)[1], 3)

    def test_tied_rules_stay_together_and_in_order(self) -> None:
        self.ws.update_cf_rule(1, priority=1)
        rule_id = self.ws.add_cf_rule(_expression())
        self.ws.set_cf_rule_priority(len(self.ws.list_cf_rules()) - 1, 1)
        got = self.ws.get_cf_rule(rule_id)
        assert got is not None
        self.assertEqual(got.priority, 1)
        self.assertEqual([(v[2], v[1]) for v in _view(self.ws) if v[2] < 2], [(0, 2), (1, 2)])

    def test_lower_numbers_and_the_old_slot_are_left_alone(self) -> None:
        rule_id = self.ws.add_cf_rule(_expression())
        self.ws.set_cf_rule_priority(rule_id, 2)
        self.assertEqual(_by_col(self.ws, 0)[1], 1)
        self.assertEqual(_by_col(self.ws, 1)[1], 3)
        self.ws.set_cf_rule_priority(rule_id, 9)
        got = self.ws.get_cf_rule(rule_id)
        assert got is not None
        self.assertEqual(got.priority, 9)
        self.assertEqual(_by_col(self.ws, 1)[1], 3)

    def test_above_the_maximum_shifts_nothing(self) -> None:
        self.ws.set_cf_rule_priority(0, 10)
        self.assertEqual(_by_col(self.ws, 0)[1], 10)
        self.assertEqual(_by_col(self.ws, 1)[1], 2)

    def test_the_current_priority_is_a_no_op(self) -> None:
        before = _view(self.ws)
        for position, view in enumerate(before):
            self.ws.set_cf_rule_priority(position, view[1])
            self.assertEqual(_view(self.ws), before)

    def test_an_id_acts_on_the_first_rule_sharing_it(self) -> None:
        self.ws.set_cf_rule_priority(SHARED_ID, 5)
        self.assertEqual(_by_col(self.ws, 0)[1], 5)
        self.assertEqual(_by_col(self.ws, 1)[1], 2)

    def test_below_one_is_invalid_and_changes_nothing(self) -> None:
        before = _view(self.ws)
        for bad in (0, -1):
            with self.subTest(priority=bad):
                with self.assertRaises(InvalidArgumentError):
                    self.ws.set_cf_rule_priority(1, bad)
        self.assertEqual(_view(self.ws), before)

    def test_a_shift_that_would_overflow_is_invalid_and_changes_nothing(self) -> None:
        self.ws.update_cf_rule(1, priority=INT32_MAX)
        before = _view(self.ws)
        for target in (INT32_MAX, 2):
            with self.subTest(priority=target):
                with self.assertRaises(InvalidArgumentError):
                    self.ws.set_cf_rule_priority(0, target)
        self.assertEqual(_view(self.ws), before)

    def test_priorities_survive_save_and_reload(self) -> None:
        self.ws.set_cf_rule_priority(1, 1)
        with Workbook.load(self.wb.save()) as reloaded:
            rules = reloaded.sheet(0).list_cf_rules()
            self.assertEqual(sorted((r.ranges[0].first_col, r.priority) for r in rules), [(0, 2), (1, 1)])


if __name__ == "__main__":
    unittest.main()
