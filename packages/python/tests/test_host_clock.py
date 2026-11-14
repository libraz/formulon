"""NOW() reads the host's local time, under whatever TZ the process has."""

from __future__ import annotations

import datetime
import os
import time
import unittest

from formulon import Workbook

EPOCH = datetime.datetime(1899, 12, 30)
TOLERANCE_DAYS = 120.0 / 86400.0


def _serial(moment: datetime.datetime) -> float:
    return (moment - EPOCH).total_seconds() / 86400.0


def _now_serial() -> float:
    with Workbook.create_default() as wb:
        wb.sheet(0).set_formula(0, 0, "=NOW()")
        wb.recalc()
        value = wb.sheet(0).get_value(0, 0)
    assert isinstance(value, float)
    return value


@unittest.skipUnless(hasattr(time, "tzset"), "time.tzset is POSIX only")
class HostClockTests(unittest.TestCase):
    def setUp(self) -> None:
        self._saved_tz = os.environ.get("TZ")

    def tearDown(self) -> None:
        if self._saved_tz is None:
            os.environ.pop("TZ", None)
        else:
            os.environ["TZ"] = self._saved_tz
        time.tzset()

    def _set_tz(self, zone: str) -> None:
        os.environ["TZ"] = zone
        time.tzset()

    def test_now_is_local_time_in_each_zone(self) -> None:
        for zone in ("Asia/Tokyo", "America/New_York", "UTC"):
            with self.subTest(zone=zone):
                self._set_tz(zone)
                self.assertAlmostEqual(_now_serial(), _serial(datetime.datetime.now()), delta=TOLERANCE_DAYS)

    def test_zones_differ_by_their_offset(self) -> None:
        self._set_tz("Asia/Tokyo")
        tokyo = _now_serial()
        self._set_tz("UTC")
        utc = _now_serial()
        self.assertAlmostEqual(tokyo - utc, 9.0 / 24.0, delta=TOLERANCE_DAYS)

    def test_pinned_clock_ignores_the_zone(self) -> None:
        pinned = datetime.datetime(2026, 4, 23, 15, 30, 45)
        for zone in ("Asia/Tokyo", "America/New_York"):
            with self.subTest(zone=zone):
                self._set_tz(zone)
                with Workbook.create_default() as wb:
                    wb.set_pinned_now(pinned)
                    wb.sheet(0).set_formula(0, 0, "=NOW()")
                    wb.recalc()
                    self.assertEqual(wb.sheet(0).get_value(0, 0), _serial(pinned))


if __name__ == "__main__":
    unittest.main()
