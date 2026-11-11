// Sheet.getDate and the pivot date-group bounds keep a stored wall-clock time in a zone with a daylight-saving gap.

import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const ZONE = 'America/New_York';
const CHILD = 'FORMULON_DST_GAP_CHILD';

// The staged entry named by the package's own "main" field.
async function loadPackage() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  if (!pkg.main) throw new Error('package.json is missing "main"');
  return import(pathToFileURL(path.resolve(pkgRoot, pkg.main)).href);
}

if (process.env[CHILD] !== '1') {
  // The zone is fixed at process start, so the checks run in a child process pinned to a zone with a gap.
  describe('daylight-saving gap', () => {
    it('keeps a stored wall-clock time unchanged in a zone with a gap', () => {
      const { NODE_TEST_CONTEXT: _context, ...inherited } = process.env;
      const run = spawnSync(process.execPath, [fileURLToPath(import.meta.url)], {
        env: { ...inherited, TZ: ZONE, [CHILD]: '1' },
        encoding: 'utf8',
      });
      assert.equal(run.status, 0, `${run.stdout}\n${run.stderr}`);
    });
  });
} else {
  const f = await (await loadPackage()).createFormulon();

  /** Runs `body(ws, wb)` on a default workbook, then disposes it. */
  const withSheet = (body) => {
    const wb = f.Workbook.createDefault();
    try {
      return body(wb.sheet(0), wb);
    } finally {
      if (!wb.disposed) wb.dispose();
    }
  };

  const serialOf = (year, month, day, hour = 0, minute = 0, second = 0) =>
    f.convertCivilToSerial({ year, month, day, hour, minute, second });

  /** A pivot over a date field, grouped by months; returns the pivot and the date field index. */
  const datePivot = (wb) => {
    const ws = wb.sheet(0);
    ws.setText('A1', 'Region');
    ws.setText('B1', 'Date');
    const rows = [
      ['East', 45292],
      ['West', 45323],
    ];
    rows.forEach(([region, date], i) => {
      ws.setText(i + 1, 0, region);
      ws.setNumber(i + 1, 1, date);
    });
    const cache = wb.addPivotCache();
    cache.setWorksheetSource({ range: 'A1:B3', sheet: 'Sheet1' });
    for (const name of ['Region', 'Date']) cache.addField(name);
    for (const item of ['East', 'West']) cache.addFieldSharedItem(0, item);
    rows.forEach(([region, date]) => {
      const record = cache.addRecord();
      cache.setRecordValue(record, 0, region === 'East' ? 0 : 1);
      cache.setRecordValue(record, 1, new Date(1899, 11, 30 + date));
    });
    const pivot = ws.addPivot({ name: 'Pivot1', cacheId: cache.cacheId, anchor: { row: 0, col: 4 } });
    pivot.setFieldAxis('Region', f.PivotAxis.Row);
    return { pivot, date: 1 };
  };

  describe('Sheet.getDate in a zone with a daylight-saving gap', () => {
    it('runs in the pinned zone', () => {
      assert.equal(new Date(2024, 2, 10, 2, 30).getHours(), 3);
    });

    it('returns the Date of the wall-clock time wherever that local time exists', () => {
      withSheet((ws) => {
        for (const [month, day] of [
          [3, 9],
          [3, 10],
          [3, 11],
          [11, 2],
          [11, 3],
          [11, 4],
        ]) {
          for (let hour = 0; hour < 24; hour++) {
            for (const minute of [0, 30, 59]) {
              ws.setNumber(0, 0, serialOf(2024, month, day, hour, minute, 15));
              const read = ws.getDate(0, 0);
              const expected = new Date(2024, month - 1, day, hour, minute, 15);
              assert.equal(read.getTime(), expected.getTime(), `${month}/${day} ${hour}:${minute}`);
              if (expected.getHours() === hour) {
                assert.deepEqual(
                  [read.getFullYear(), read.getMonth() + 1, read.getDate(), read.getHours(), read.getMinutes()],
                  [2024, month, day, hour, minute],
                );
              }
            }
          }
        }
      });
    });

    it('moves a time that does not exist forward by the length of the gap, as the Date constructor does', () => {
      withSheet((ws) => {
        ws.setNumber(0, 0, serialOf(2024, 3, 10, 2, 30, 0));
        const read = ws.getDate(0, 0);
        assert.equal(read.getTime(), new Date(2024, 2, 10, 2, 30, 0).getTime());
        assert.deepEqual([read.getHours(), read.getMinutes()], [3, 30]);
        ws.setNumber(1, 0, serialOf(2024, 3, 10, 1, 59, 59));
        assert.deepEqual([ws.getDate(1, 0).getHours(), ws.getDate(1, 0).getMinutes()], [1, 59]);
        ws.setNumber(2, 0, serialOf(2024, 3, 10, 3, 0, 0));
        assert.deepEqual([ws.getDate(2, 0).getHours(), ws.getDate(2, 0).getMinutes()], [3, 0]);
      });
    });
  });

  describe('a pivot date group in a zone with a daylight-saving gap', () => {
    const bounds = {
      start: { year: 2024, month: 3, day: 10, hour: 2, minute: 30, second: 0 },
      end: { year: 2025, month: 3, day: 9, hour: 2, minute: 59, second: 59 },
    };

    it('reads a bound inside the gap back as the stored civil fields', () => {
      withSheet((_ws, wb) => {
        const { pivot, date } = datePivot(wb);
        pivot.setFieldDateGroup(date, { periods: f.PivotDatePeriod.Months, ...bounds });
        const read = pivot.getFieldDateGroup(date);
        assert.deepEqual({ ...read.start }, bounds.start);
        assert.deepEqual({ ...read.end }, bounds.end);
        assert.equal(read.autoStart, false);
        assert.equal(read.autoEnd, false);
      });
    });

    it('writes the read record back unchanged', () => {
      withSheet((_ws, wb) => {
        const { pivot, date } = datePivot(wb);
        pivot.setFieldDateGroup(date, { periods: f.PivotDatePeriod.Months, ...bounds });
        const read = pivot.getFieldDateGroup(date);
        pivot.setFieldDateGroup(date, read);
        assert.deepEqual(pivot.getFieldDateGroup(date), read);
      });
    });
  });
}
