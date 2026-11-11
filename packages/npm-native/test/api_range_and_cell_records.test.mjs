// Range defaults and cell-valued fields: the block reads default to the used range, a view, a threaded comment, a
// sparkline and a pivot worksheet source carry cells and ranges as records, and the XF remap pair takes cells and a
// Map.

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

const f = await (await loadPackage()).createFormulon();

/** Runs `body(ws, wb)` on a fresh default workbook and disposes it. */
function withSheet(body) {
  const wb = f.Workbook.createDefault();
  try {
    return body(wb.sheet(0), wb);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

function expectCode(fn, code) {
  assert.throws(fn, (e) => {
    assert.equal(e.name, 'FormulonError', `expected a FormulonError, got ${e}`);
    assert.equal(e.code, code, `code ${e.code}`);
    return true;
  });
}

const expectInvalid = (fn) => expectCode(fn, f.StatusCode.InvalidArgument);
const expectTypeError = (fn) => assert.throws(fn, (e) => e instanceof TypeError, 'expected a TypeError');

const rect = (firstRow, firstCol, lastRow, lastCol) => ({ firstRow, firstCol, lastRow, lastCol });
const plain = (value) => JSON.parse(JSON.stringify(value) ?? 'null');

const PERSON = '{00000000-0000-4000-8000-0000000000A1}';
const guid = (n) => `{00000000-0000-4000-8000-${String(n).padStart(12, '0')}}`;

function fillBlock(ws) {
  ws.setText('A1', 'Region');
  ws.setText('B1', 'Amount');
  ws.setText('A2', 'East');
  ws.setNumber('B2', 10);
  ws.setText('A3', 'West');
  ws.setNumber('B3', 30);
}

describe('the block reads default to the used range', () => {
  it('every block read of an empty sheet is empty', () => {
    withSheet((ws) => {
      assert.deepEqual(plain(ws.getValues()), []);
      assert.deepEqual(plain(ws.getValues(null)), []);
      assert.deepEqual(plain(ws.getTypedValues()), []);
      assert.deepEqual(plain(ws.getTypedValues(null)), []);
      assert.deepEqual(plain(ws.getCellsInRange()), { cells: [], nextCursor: null });
      assert.deepEqual(plain(ws.getCellsInRange(null)), { cells: [], nextCursor: null });
      assert.deepEqual([...ws.iterRows()], []);
      assert.deepEqual(plain(ws.listCells()), []);
    });
  });

  it('an omitted or null range equals the used range on a filled sheet', () => {
    withSheet((ws) => {
      fillBlock(ws);
      const used = ws.getUsedRange();
      assert.deepEqual(plain(ws.getValues()), plain(ws.getValues(used)));
      assert.deepEqual(plain(ws.getValues(null)), plain(ws.getValues(used)));
      assert.deepEqual(plain(ws.getTypedValues()), plain(ws.getTypedValues(used)));
      assert.deepEqual(plain(ws.getCellsInRange()), plain(ws.getCellsInRange(used)));
      assert.deepEqual(plain(ws.listCells()), plain(ws.listCells(used)));
      assert.deepEqual(plain([...ws.iterRows()]), plain(ws.getValues(used)));
    });
  });
});

describe('a sheet view carries its scroll cells as addresses', () => {
  it('reads topLeft and paneTopLeft as cell addresses without the row / col pairs', () => {
    withSheet((ws) => {
      const view = ws.getView();
      assert.deepEqual(plain(view.topLeft), { row: 0, col: 0 });
      assert.deepEqual(plain(view.paneTopLeft), { row: 0, col: 0 });
      for (const key of ['topLeftRow', 'topLeftCol', 'paneTopLeftRow', 'paneTopLeftCol']) {
        assert.equal(view[key], undefined, key);
      }
    });
  });

  it('updateView takes an A1 string or an address and reads it back', () => {
    withSheet((ws) => {
      ws.updateView({ topLeft: 'C5' });
      assert.deepEqual(plain(ws.getView().topLeft), { row: 4, col: 2 });
      ws.updateView({ topLeft: { row: 1, col: 1 } });
      assert.deepEqual(plain(ws.getView().topLeft), { row: 1, col: 1 });
    });
  });

  it('the view a read returned passes back and no key of it is clearable', () => {
    withSheet((ws) => {
      ws.updateView({ topLeft: 'B2' });
      const read = ws.getView();
      ws.updateView({ topLeft: 'A1' });
      ws.updateView(read);
      assert.deepEqual(plain(ws.getView().topLeft), { row: 1, col: 1 });
      expectTypeError(() => ws.updateView({ zoomScale: null }));
      expectTypeError(() => ws.updateView({ topLeft: null }));
    });
  });
});

describe('threaded comments carry their anchor as a cell', () => {
  function addThreads(ws, wb) {
    wb.addPerson({ id: PERSON, displayName: 'Alice', userId: 'alice', providerId: 'None' });
    const created = { year: 2026, month: 1, day: 2, hour: 3, minute: 4, second: 5 };
    ws.addThreadedComment({ id: guid(1), text: 'first', cell: 'B2', personId: PERSON, created });
    ws.addThreadedComment({ id: guid(2), text: 'reply', parentId: guid(1), personId: PERSON, created });
    ws.addThreadedComment({ id: guid(4), text: 'other', cell: { row: 2, col: 2 }, personId: PERSON, created });
  }

  it('a comment reads cell as an address and no row / col', () => {
    withSheet((ws, wb) => {
      addThreads(ws, wb);
      const read = ws.getThreadedComment(guid(1));
      assert.deepEqual(plain(read.cell), { row: 1, col: 1 });
      assert.equal(read.row, undefined);
      assert.equal(read.col, undefined);
    });
  });

  it('the cell list returns the openers and replies of that cell only, in list order', () => {
    withSheet((ws, wb) => {
      addThreads(ws, wb);
      assert.deepEqual(
        ws.listThreadedComments('B2').map((c) => c.id),
        [guid(1), guid(2)],
      );
      assert.deepEqual(
        ws.listThreadedComments('C3').map((c) => c.id),
        [guid(4)],
      );
      assert.deepEqual(plain(ws.listThreadedComments('D4')), []);
    });
  });

  it('the row / col and A1 forms agree and the unfiltered list holds every comment', () => {
    withSheet((ws, wb) => {
      addThreads(ws, wb);
      assert.deepEqual(plain(ws.listThreadedComments(1, 1)), plain(ws.listThreadedComments('B2')));
      assert.equal(ws.listThreadedComments().length, 3);
    });
  });

  it('the old row / col input keys are unknown keys', () => {
    withSheet((ws, wb) => {
      wb.addPerson({ id: PERSON, displayName: 'Alice', userId: 'alice', providerId: 'None' });
      expectTypeError(() => ws.addThreadedComment({ id: guid(1), text: 'x', row: 1, col: 1, personId: PERSON }));
    });
  });
});

describe('a pivot worksheet source is a range and a name', () => {
  function cacheWithData() {
    const wb = f.Workbook.createDefault();
    wb.addSheet('Data');
    return { wb, cache: wb.addPivotCache() };
  }

  it('an unqualified range takes its sheet from the sheet key', () => {
    const { wb, cache } = cacheWithData();
    try {
      cache.setWorksheetSource({ range: 'A1:C9', sheet: 'Data' });
      const read = cache.getWorksheetSource();
      assert.deepEqual(plain(read.range), { sheet: 1, rect: rect(0, 0, 8, 2) });
      assert.equal(read.name, null);
    } finally {
      wb.dispose();
    }
  });

  it('a qualified string and a SheetRect read back the same range', () => {
    const { wb, cache } = cacheWithData();
    try {
      cache.setWorksheetSource({ range: 'Data!A1:C9' });
      const qualified = plain(cache.getWorksheetSource().range);
      cache.setWorksheetSource({ range: { sheet: 1, rect: rect(0, 0, 8, 2) } });
      assert.deepEqual(plain(cache.getWorksheetSource().range), qualified);
      assert.deepEqual(qualified, { sheet: 1, rect: rect(0, 0, 8, 2) });
    } finally {
      wb.dispose();
    }
  });

  it('a sheet that disagrees with the range is InvalidArgument', () => {
    const { wb, cache } = cacheWithData();
    try {
      expectInvalid(() => cache.setWorksheetSource({ range: 'Data!A1:C9', sheet: 'Sheet1' }));
      expectInvalid(() => cache.setWorksheetSource({ range: { sheet: 1, rect: rect(0, 0, 8, 2) }, sheet: 0 }));
    } finally {
      wb.dispose();
    }
  });

  it('the read has no present, sheet or ref key, passes back, and null clears', () => {
    const { wb, cache } = cacheWithData();
    try {
      cache.setWorksheetSource({ range: 'A1:C9', sheet: 'Data' });
      const read = cache.getWorksheetSource();
      for (const key of ['present', 'sheet', 'ref']) assert.equal(key in read, false, key);
      cache.setWorksheetSource(null);
      assert.equal(cache.getWorksheetSource(), null);
      cache.setWorksheetSource(read);
      assert.deepEqual(plain(cache.getWorksheetSource().range), plain(read.range));
    } finally {
      wb.dispose();
    }
  });

  it('the old ref key is an unknown key', () => {
    const { wb, cache } = cacheWithData();
    try {
      expectTypeError(() => cache.setWorksheetSource({ ref: 'A1:C9', sheet: 'Data' }));
    } finally {
      wb.dispose();
    }
  });
});

describe('the XF remap pair takes cells and a Map', () => {
  /** A1 bold, B2 italic, the rest default. */
  function styled(ws) {
    ws.updateCellFormat('A1', { font: { bold: true } });
    ws.updateCellFormat('B2', { font: { italic: true } });
  }

  it('getRangeXfBasesExcluding takes addresses and A1 strings alike', () => {
    withSheet((ws) => {
      styled(ws);
      const all = ws.ooxml.getRangeXfBases('A1:C3');
      const mixed = ws.ooxml.getRangeXfBasesExcluding('A1:C3', [{ row: 0, col: 0 }, 'B2']);
      const addresses = ws.ooxml.getRangeXfBasesExcluding('A1:C3', [
        { row: 0, col: 0 },
        { row: 1, col: 1 },
      ]);
      assert.deepEqual(plain(mixed), plain(addresses));
      assert.ok(mixed.length < all.length);
    });
  });

  it('excluding every cell yields an empty list and an empty exclusion keeps every base', () => {
    withSheet((ws) => {
      styled(ws);
      const all = ws.ooxml.getRangeXfBases('A1:B2');
      assert.deepEqual(plain(ws.ooxml.getRangeXfBasesExcluding('A1:B2', ['A1', 'B1', 'A2', 'B2'])), []);
      assert.deepEqual(plain(ws.ooxml.getRangeXfBasesExcluding('A1:B2', [])), plain(all));
    });
  });

  it('a malformed excluded entry is refused', () => {
    withSheet((ws) => {
      expectTypeError(() => ws.ooxml.getRangeXfBasesExcluding('A1:B2', [3]));
      expectTypeError(() => ws.ooxml.getRangeXfBasesExcluding('A1:B2', 'A1'));
    });
  });

  it('remapRangeXfs applies a Map of base to patched XF', () => {
    withSheet((ws, wb) => {
      styled(ws);
      const base = ws.ooxml.getCellXfIndex('A1');
      const patched = wb.ooxml.addCellXf({ ...wb.ooxml.getCellXf(base), numFmtId: 2 });
      assert.notEqual(patched, base);
      ws.ooxml.remapRangeXfs('A1:B2', new Map([[base, patched]]));
      assert.equal(ws.ooxml.getCellXfIndex('A1'), patched);
    });
  });

  it('a plain object, a non-index key and a non-index value are TypeErrors', () => {
    withSheet((ws) => {
      expectTypeError(() => ws.ooxml.remapRangeXfs('A1:B2', { 0: 1 }));
      expectTypeError(() => ws.ooxml.remapRangeXfs('A1:B2', new Map([['0', 1]])));
      expectTypeError(() => ws.ooxml.remapRangeXfs('A1:B2', new Map([[0, 1.5]])));
    });
  });

  it('the parallel-array form is no longer the public signature', () => {
    withSheet((ws) => {
      expectTypeError(() => ws.ooxml.remapRangeXfs('A1:B2', [0], [1]));
    });
  });
});

describe('getCellXfIndices', () => {
  /** Two XFs beyond the default, and a mix of styled blanks, styled content, unstyled content and an explicit 0. */
  function styledCells(ws, wb) {
    const a = wb.ooxml.addCellXf({ ...wb.ooxml.getCellXf(0), numFmtId: 2 });
    const b = wb.ooxml.addCellXf({ ...wb.ooxml.getCellXf(0), numFmtId: 3 });
    ws.ooxml.setCellXfIndex('U1001', a); // styled blank, far from the content
    ws.ooxml.setCellXfIndex('D6', b); // styled blank
    ws.setNumber('H3', 4);
    ws.ooxml.setCellXfIndex('H3', a); // content with a style
    ws.setNumber('B5', 9); // content, unstyled
    ws.ooxml.setCellXfIndex('G7', 0); // explicit default
    return { a, b };
  }

  it('lists styled blank and styled content cells row-major, and leaves out xf 0 and unstyled content', () => {
    withSheet((ws, wb) => {
      const { a, b } = styledCells(ws, wb);
      assert.deepEqual(plain(ws.ooxml.getCellXfIndices()), [
        { row: 2, col: 7, xfIndex: a },
        { row: 5, col: 3, xfIndex: b },
        { row: 1000, col: 20, xfIndex: a },
      ]);
    });
  });

  it('the default range covers a format-only cell outside the used range', () => {
    withSheet((ws, wb) => {
      styledCells(ws, wb);
      const used = ws.getUsedRange();
      assert.ok(used.lastRow < 1000);
      assert.equal(ws.ooxml.getCellXfIndices(null).length, 3);
      assert.equal(ws.ooxml.getCellXfIndices(undefined).length, 3);
      const whole = plain(ws.ooxml.getCellXfIndices());
      assert.deepEqual(plain(ws.ooxml.getCellXfIndices(rect(0, 0, 1048575, 16383))), whole);
      assert.deepEqual(plain(ws.ooxml.getCellXfIndices('A1:XFD1048576')), whole);
    });
  });

  it('clips to a CellRect or an A1 range', () => {
    withSheet((ws, wb) => {
      const { a, b } = styledCells(ws, wb);
      const clipped = [
        { row: 2, col: 7, xfIndex: a },
        { row: 5, col: 3, xfIndex: b },
      ];
      assert.deepEqual(plain(ws.ooxml.getCellXfIndices(rect(0, 0, 5, 1000))), clipped);
      assert.deepEqual(plain(ws.ooxml.getCellXfIndices('A1:Z6')), clipped);
      assert.deepEqual(plain(ws.ooxml.getCellXfIndices('D6')), [{ row: 5, col: 3, xfIndex: b }]);
      assert.deepEqual(plain(ws.ooxml.getCellXfIndices('A8:C10')), []);
    });
  });

  it('survives a save and a load', () => {
    withSheet((ws, wb) => {
      const before = plain(ws.ooxml.getCellXfIndices());
      assert.equal(before.length, 0);
      styledCells(ws, wb);
      const expected = plain(ws.ooxml.getCellXfIndices());
      const reloaded = f.Workbook.load(wb.save());
      try {
        assert.deepEqual(plain(reloaded.sheet(0).ooxml.getCellXfIndices()), expected);
      } finally {
        reloaded.dispose();
      }
    });
  });

  it('refuses a rectangle outside the grid', () => {
    withSheet((ws) => {
      assert.throws(() => ws.ooxml.getCellXfIndices(rect(0, 0, 1048576, 0)));
    });
  });
});

describe('getShowPhoneticCells', () => {
  /** A flagged constant, a flag-only blank, a flagged text, a far flagged blank, and unflagged content and blank. */
  function flaggedCells(ws) {
    ws.setBlank('U1001');
    ws.setCellShowPhonetic('U1001', true); // flag-only blank, far from the content
    ws.setNumber('B2', 5);
    ws.setCellShowPhonetic('B2', true); // flagged constant
    ws.setCellShowPhonetic('A3', true); // flag-only blank
    ws.setText('H5', 'x');
    ws.setCellShowPhonetic('H5', true); // flagged text
    ws.setNumber('C2', 1); // content, no flag
    ws.setBlank('D6'); // stored blank, no flag
  }
  const all = [
    { row: 1, col: 1 },
    { row: 2, col: 0 },
    { row: 4, col: 7 },
    { row: 1000, col: 20 },
  ];

  it('lists flagged constants and flag-only blanks row-major and leaves out unflagged cells', () => {
    withSheet((ws) => {
      flaggedCells(ws);
      assert.deepEqual(plain(ws.getShowPhoneticCells()), all);
      for (const { row, col } of all) assert.equal(ws.getCellShowPhonetic(row, col), true);
    });
  });

  it('the default range covers a flagged cell outside the used range', () => {
    withSheet((ws) => {
      flaggedCells(ws);
      assert.ok(ws.getUsedRange().lastRow < 1000);
      assert.equal(ws.getShowPhoneticCells(null).length, 4);
      assert.equal(ws.getShowPhoneticCells(undefined).length, 4);
      const whole = plain(ws.getShowPhoneticCells());
      assert.deepEqual(plain(ws.getShowPhoneticCells(rect(0, 0, 1048575, 16383))), whole);
      assert.deepEqual(plain(ws.getShowPhoneticCells('A1:XFD1048576')), whole);
    });
  });

  it('clips to a CellRect or an A1 range', () => {
    withSheet((ws) => {
      flaggedCells(ws);
      const clipped = [
        { row: 1, col: 1 },
        { row: 2, col: 0 },
      ];
      assert.deepEqual(plain(ws.getShowPhoneticCells(rect(0, 0, 2, 1))), clipped);
      assert.deepEqual(plain(ws.getShowPhoneticCells('A1:B3')), clipped);
      assert.deepEqual(plain(ws.getShowPhoneticCells('H5')), [{ row: 4, col: 7 }]);
      assert.deepEqual(plain(ws.getShowPhoneticCells('A8:C10')), []);
    });
  });

  it('survives a save and a load, flag-only A3 included', () => {
    withSheet((ws, wb) => {
      assert.equal(ws.getShowPhoneticCells().length, 0);
      flaggedCells(ws);
      const reloaded = f.Workbook.load(wb.save());
      try {
        assert.deepEqual(plain(reloaded.sheet(0).getShowPhoneticCells()), all);
      } finally {
        reloaded.dispose();
      }
    });
  });

  it('refuses a rectangle outside the grid', () => {
    withSheet((ws) => {
      assert.throws(() => ws.getShowPhoneticCells(rect(0, 0, 1048576, 0)));
    });
  });
});
