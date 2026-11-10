// Spill.cause: what blocks a formula's spill or a PivotTable report, on Sheet.getSpill and Pivot.getSpill.

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

// The staged entry named by the package's own "main" field.
async function loadPackage() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  if (!pkg.main) throw new Error('package.json is missing "main"');
  return import(pathToFileURL(path.resolve(pkgRoot, pkg.main)).href);
}

const mod = await loadPackage();
const f = await mod.createFormulon();
const { PivotAxis, PivotAggregation, SpillCause } = f;

const LAST_ROW = 1048575;
const LAST_COL = 16383;

/** Runs `body(ws, wb)` on a default workbook, then disposes it. */
function withSheet(body) {
  const wb = f.Workbook.createDefault();
  try {
    return body(wb.sheet(0), wb);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

const rect = (firstRow, firstCol, lastRow, lastCol) => ({ firstRow, firstCol, lastRow, lastCol });
const plain = (spill) => (spill === null ? null : { ...spill, anchor: { ...spill.anchor }, range: { ...spill.range } });

/** A pivot summing Amount by Region at E1, drawn over E1:F4; `grow()` makes it one row longer, into E5. */
function buildPivot(ws, wb) {
  ws.setText('A1', 'Region');
  ws.setText('B1', 'Amount');
  const rows = [
    ['East', 10],
    ['West', 30],
    ['East', 5],
  ];
  rows.forEach(([region, amount], i) => {
    ws.setText(i + 1, 0, region);
    ws.setNumber(i + 1, 1, amount);
  });
  const cache = wb.addPivotCache();
  cache.setWorksheetSource({ range: 'A1:B4', sheet: 'Sheet1' });
  for (const name of ['Region', 'Amount']) cache.addField(name);
  for (const item of ['East', 'West']) cache.addFieldSharedItem(0, item);
  rows.forEach(([region, amount]) => {
    const record = cache.addRecord();
    cache.setRecordValue(record, 0, region === 'East' ? 0 : 1);
    cache.setRecordValue(record, 1, amount);
  });
  const pivot = ws.addPivot({ name: 'Pivot1', cacheId: cache.cacheId, anchor: 'E1' });
  pivot.setFieldAxis('Region', PivotAxis.Row);
  pivot.addDataField({ fieldIndex: 'Amount', aggregation: PivotAggregation.Sum });
  pivot.project();
  const grow = () => {
    ws.setText('A4', 'North');
    cache.refresh();
  };
  return { pivot, grow };
}

describe('Spill.cause on a formula', () => {
  it('a drawn spill reads blocked false and no cause, at its anchor and its phantoms', () => {
    withSheet((ws) => {
      ws.setFormula('A1', '=SEQUENCE(3)');
      const drawn = { anchor: { row: 0, col: 0 }, blocked: false, cause: null, range: rect(0, 0, 2, 0) };
      assert.deepEqual(plain(ws.getSpill('A1')), drawn);
      assert.deepEqual(plain(ws.getSpill('A3')), drawn);
    });
  });

  it('a value in the way is Occupied and a merge is Merged', () => {
    withSheet((ws) => {
      ws.setFormula('C1', '=SEQUENCE(3)');
      ws.setText('C3', 'x');
      ws.addMerge('D2:E2');
      ws.setFormula('D1', '=SEQUENCE(3)');
      assert.deepEqual(plain(ws.getSpill('C1')), {
        anchor: { row: 0, col: 2 },
        blocked: true,
        cause: SpillCause.Occupied,
        range: rect(0, 2, 2, 2),
      });
      assert.deepEqual(plain(ws.getSpill('D1')), {
        anchor: { row: 0, col: 3 },
        blocked: true,
        cause: SpillCause.Merged,
        range: rect(0, 3, 2, 3),
      });
      assert.equal(ws.getSpill('C2'), null);
    });
  });

  it('a spill past the sheet edge is BeyondEdge with its range clipped to the grid', () => {
    withSheet((ws) => {
      ws.setFormula(LAST_ROW - 1, 0, '=SEQUENCE(5)');
      ws.setFormula(0, LAST_COL - 1, '=SEQUENCE(1,5)');
      assert.deepEqual(plain(ws.getSpill(LAST_ROW - 1, 0)), {
        anchor: { row: LAST_ROW - 1, col: 0 },
        blocked: true,
        cause: SpillCause.BeyondEdge,
        range: rect(LAST_ROW - 1, 0, LAST_ROW, 0),
      });
      assert.deepEqual(plain(ws.getSpill(0, LAST_COL - 1)), {
        anchor: { row: 0, col: LAST_COL - 1 },
        blocked: true,
        cause: SpillCause.BeyondEdge,
        range: rect(0, LAST_COL - 1, 0, LAST_COL),
      });
    });
  });

  it('a phantom cell reports the region anchor', () => {
    withSheet((ws) => {
      ws.setFormula('B2', '=SEQUENCE(3)');
      assert.deepEqual({ ...ws.getSpill('B4').anchor }, { row: 1, col: 1 });
    });
  });

  it('a cell in no spill reads null', () => {
    withSheet((ws) => {
      ws.setNumber('Z1', 1);
      assert.equal(ws.getSpill('Z1'), null);
    });
  });
});

describe('Spill.cause on a PivotTable report', () => {
  it('a drawn report reads its extent at its anchor, and its body cells read null', () => {
    withSheet((ws, wb) => {
      const { pivot } = buildPivot(ws, wb);
      const drawn = { anchor: { row: 0, col: 4 }, blocked: false, cause: null, range: rect(0, 4, 3, 5) };
      assert.deepEqual(plain(pivot.getSpill()), drawn);
      assert.deepEqual(plain(ws.getSpill('E1')), drawn);
      assert.equal(ws.getSpill('E2'), null);
      assert.equal(ws.getSpill('F4'), null);
    });
  });

  it('a value in the way is Occupied', () => {
    withSheet((ws, wb) => {
      const { pivot, grow } = buildPivot(ws, wb);
      ws.setText('E5', 'x');
      grow();
      const blocked = {
        anchor: { row: 0, col: 4 },
        blocked: true,
        cause: SpillCause.Occupied,
        range: rect(0, 4, 4, 5),
      };
      assert.deepEqual(plain(pivot.getSpill()), blocked);
      assert.deepEqual(plain(ws.getSpill('E1')), blocked);
    });
  });

  it('a merge in the way is Merged', () => {
    withSheet((ws, wb) => {
      const { pivot, grow } = buildPivot(ws, wb);
      ws.addMerge('E5:F5');
      grow();
      assert.deepEqual(plain(pivot.getSpill()), {
        anchor: { row: 0, col: 4 },
        blocked: true,
        cause: SpillCause.Merged,
        range: rect(0, 4, 4, 5),
      });
    });
  });

  it('a report growing past the sheet edge is BeyondEdge with its range clipped to the grid', () => {
    withSheet((ws, wb) => {
      const { pivot, grow } = buildPivot(ws, wb);
      pivot.setAnchor({ row: LAST_ROW - 3, col: 4 });
      grow();
      const blocked = {
        anchor: { row: LAST_ROW - 3, col: 4 },
        blocked: true,
        cause: SpillCause.BeyondEdge,
        range: rect(LAST_ROW - 3, 4, LAST_ROW, 5),
      };
      assert.deepEqual(plain(pivot.getSpill()), blocked);
      assert.deepEqual(plain(ws.getSpill(LAST_ROW - 3, 4)), blocked);
    });
  });

  it('a drawn report read from a file reads the location it was read with', () => {
    withSheet((ws, wb) => {
      buildPivot(ws, wb);
      const loaded = f.Workbook.load(wb.save());
      try {
        const drawn = { anchor: { row: 0, col: 4 }, blocked: false, cause: null, range: rect(0, 4, 3, 5) };
        assert.deepEqual(plain(loaded.sheet(0).pivot(0).getSpill()), drawn);
        assert.deepEqual(plain(loaded.sheet(0).getSpill('E1')), drawn);
        assert.equal(loaded.sheet(0).getSpill('E2'), null);
      } finally {
        loaded.dispose();
      }
    });
  });

  it('SpillCause has no member for absence', () => {
    assert.deepEqual({ ...mod.SpillCause }, { Occupied: 0, Merged: 1, BeyondEdge: 2 });
    assert.equal(mod.SpillCause, SpillCause);
  });
});
