// A name given to a facade method resolves as the engine resolves sheet names: exact, then case-only, then a unique
// width fold; two entries that collide only under the width fold resolve nothing.
// FORMULON_NPM_ENTRY selects the staged subpath: "." (default) or "./threads".

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const entry = process.env.FORMULON_NPM_ENTRY ?? '.';

async function loadPackage() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  const target = pkg.exports?.[entry]?.import;
  if (!target) throw new Error(`package.json "exports" has no import target for ${entry}`);
  return import(pathToFileURL(path.resolve(pkgRoot, target)).href);
}

const f = await (await loadPackage()).createFormulon();

const { StatusCode } = f;

// Full-width spellings: 'Sales' and its lower-case form fold to these under the width fold only.
const WIDE_LOWER = 'ｓａｌｅｓ';

function withWorkbook(body) {
  const wb = f.Workbook.createDefault();
  try {
    return body(wb, wb.sheet(0));
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

function expectNotFound(fn) {
  assert.throws(fn, (e) => {
    assert.equal(e.name, 'FormulonError', `expected a FormulonError, got ${e}`);
    assert.equal(e.code, StatusCode.NotFound, `code ${e.code}`);
    return true;
  });
}

/** Header cells `names` on row 1 of `ws`, one data row beneath. */
function fillHeaders(ws, names) {
  names.forEach((name, col) => {
    ws.setText(0, col, name);
    ws.setNumber(1, col, col + 1);
  });
}

/** A pivot cache with a field per name and no worksheet source, so the names are taken as given. */
function addBareCache(wb, names) {
  const cache = wb.addPivotCache();
  for (const name of names) cache.addField(name);
  return cache;
}

/** A pivot cache over the headers of `ws` with a field per name. */
function addCache(wb, ws, names) {
  fillHeaders(ws, names);
  const cache = wb.addPivotCache();
  cache.setWorksheetSource({
    range: { sheet: 0, rect: { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: names.length - 1 } },
  });
  for (const name of names) cache.addField(name);
  return cache;
}

describe('a pivot name resolves as sheet names do', () => {
  function pivots(wb, ws, names) {
    const cache = addCache(wb, ws, ['Region']);
    return names.map((name, i) => ws.addPivot({ name, cacheId: cache.cacheId, anchor: { row: 5, col: i * 4 } }));
  }

  it('exact, case-only and a unique width fold find the pivot', () => {
    withWorkbook((wb, ws) => {
      const [pivot] = pivots(wb, ws, ['Sales']);
      for (const name of ['Sales', 'sALES', WIDE_LOWER]) assert.equal(ws.pivot(name), pivot, name);
    });
  });

  it('a miss is NotFound', () => {
    withWorkbook((wb, ws) => {
      pivots(wb, ws, ['Sales']);
      expectNotFound(() => ws.pivot('Nope'));
    });
  });

  it('removePivot takes the same names', () => {
    withWorkbook((wb, ws) => {
      pivots(wb, ws, ['Sales']);
      ws.removePivot(WIDE_LOWER);
      assert.equal(ws.listPivots().length, 0);
    });
  });
});

describe('a table name resolves as sheet names do', () => {
  function tables(ws, names) {
    return names.map((name, i) => {
      ws.setText(0, i * 3, 'Key');
      ws.setNumber(1, i * 3, 1);
      return ws.addTable({ name, range: { firstRow: 0, firstCol: i * 3, lastRow: 1, lastCol: i * 3 } });
    });
  }

  it('exact, case-only and a unique width fold find the table', () => {
    withWorkbook((_wb, ws) => {
      const [table] = tables(ws, ['Sales']);
      for (const name of ['Sales', 'sALES', WIDE_LOWER]) assert.equal(ws.table(name), table, name);
    });
  });

  it('a miss is NotFound', () => {
    withWorkbook((_wb, ws) => {
      tables(ws, ['Sales']);
      expectNotFound(() => ws.table('Nope'));
    });
  });

  it('a table on another sheet is not found by this sheet', () => {
    withWorkbook((wb, ws) => {
      const other = wb.addSheet('Other');
      tables(other, ['Sales']);
      expectNotFound(() => ws.table('Sales'));
      assert.equal(wb.table('Sales').name, 'Sales');
    });
  });

  it('removeTable takes the same names', () => {
    withWorkbook((_wb, ws) => {
      tables(ws, ['Sales']);
      ws.removeTable(WIDE_LOWER);
      assert.equal(ws.listTables().length, 0);
    });
  });
});

describe('a pivot field name resolves as sheet names do', () => {
  const sharedCount = (cache, field) => cache.listFieldSharedItems(field).length;

  it('exact, case-only and a unique width fold find the field', () => {
    withWorkbook((wb) => {
      const cache = addBareCache(wb, ['Sales', 'Other']);
      cache.addFieldSharedItem(0, 'x');
      for (const name of ['Sales', 'sALES', WIDE_LOWER]) assert.equal(sharedCount(cache, name), 1, name);
      assert.equal(sharedCount(cache, 'Other'), 0);
    });
  });

  it('an exact name wins over a case-only one', () => {
    withWorkbook((wb) => {
      const cache = addBareCache(wb, ['a', 'A']);
      cache.addFieldSharedItem(1, 'x');
      assert.equal(sharedCount(cache, 'A'), 1);
      assert.equal(sharedCount(cache, 'a'), 0);
    });
  });

  it('two names that differ only by width resolve nothing, and so does a miss', () => {
    withWorkbook((wb) => {
      // Each twin is 'sales' with one letter fullwidth, so 'sales' reaches them only through the width fold.
      const cache = addBareCache(wb, ['\uFF53ales', 's\uFF41les']);
      expectNotFound(() => cache.listFieldSharedItems('sales'));
      expectNotFound(() => cache.listFieldSharedItems('Nope'));
    });
  });

  it('a pivot method takes the same names', () => {
    withWorkbook((wb, ws) => {
      const cache = addCache(wb, ws, ['Sales']);
      const pivot = ws.addPivot({ name: 'P', cacheId: cache.cacheId, anchor: { row: 5, col: 0 } });
      assert.equal(pivot.getField(WIDE_LOWER).cacheFieldIndex, 0);
      expectNotFound(() => pivot.getField('Nope'));
    });
  });
});

describe('a slicer column name resolves as sheet names do', () => {
  const ANCHOR = { fromRow: 20, fromCol: 8, toRow: 30, toCol: 10 };

  function slicerColumn(headers, column) {
    return withWorkbook((_wb, ws) => {
      headers.forEach((name, col) => {
        ws.setText(0, col, name);
        ws.setNumber(1, col, col);
      });
      const table = ws.addTable({
        name: 'T',
        range: { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: headers.length - 1 },
      });
      return ws.addSlicer(table, column, { anchor: ANCHOR }).getProperties().columnName;
    });
  }

  it('exact, case-only and a unique width fold find the column', () => {
    for (const name of ['Sales', 'sALES', WIDE_LOWER])
      assert.equal(slicerColumn(['Key', 'Sales'], name), 'Sales', name);
  });

  it('a miss is NotFound', () => {
    expectNotFound(() => slicerColumn(['Sales'], 'Nope'));
  });
});

describe('a conditional-format rule id resolves as ids do', () => {
  function withRule(body) {
    withWorkbook((_wb, ws) => {
      const id = ws.addCfRule({ type: f.CfRuleType.Expression, ranges: 'A1', formula1: 'A1>1' });
      body(ws, id);
    });
  }

  it('an exact id and its other-case spelling find the rule', () => {
    withRule((ws, id) => {
      for (const spelling of [id, id.toLowerCase(), id.toUpperCase()]) {
        ws.updateCfRule(spelling, { formula1: 'A1>2' });
        assert.equal(ws.getCfRule(id).formula1, '=A1>2', spelling);
        ws.updateCfRule(id, { formula1: 'A1>1' });
      }
    });
  });

  it('a width-folded spelling and an unknown id are NotFound', () => {
    withRule((ws, id) => {
      const wide = id.replace(/[0-9A-Za-z]/g, (c) => String.fromCharCode(c.charCodeAt(0) + 0xfee0));
      expectNotFound(() => ws.updateCfRule(wide, { formula1: 'A1>2' }));
      expectNotFound(() => ws.updateCfRule('{00000000-0000-4000-8000-000000000000}', { formula1: 'A1>2' }));
    });
  });
});

describe('a pivot worksheet source writes a resolved sheet in its stored spelling and an unresolved one verbatim', () => {
  it('a name that resolves finds the sheet', () => {
    withWorkbook((wb) => {
      wb.addSheet('Sales');
      const cache = wb.addPivotCache();
      for (const name of ['Sales', 'sALES', WIDE_LOWER]) {
        cache.setWorksheetSource({ range: 'A1:B3', sheet: name });
        assert.equal(cache.getWorksheetSource().range.sheet, 1, name);
      }
    });
  });

  it('an unresolved name is accepted and reads back without a range', () => {
    withWorkbook((wb) => {
      wb.addSheet('Sales');
      const cache = wb.addPivotCache();
      cache.setWorksheetSource({ range: 'A1:B3', sheet: 'Gone' });
      assert.equal(cache.getWorksheetSource().range, null);
    });
  });

  it('a qualified range and a sheet that resolve to one sheet agree, and two sheets disagree', () => {
    withWorkbook((wb) => {
      wb.addSheet('Sales');
      wb.addSheet('Other');
      const cache = wb.addPivotCache();
      cache.setWorksheetSource({ range: 'Sales!A1:B3', sheet: WIDE_LOWER });
      assert.equal(cache.getWorksheetSource().range.sheet, 1);
      assert.throws(
        () => cache.setWorksheetSource({ range: 'Sales!A1:B3', sheet: 'Other' }),
        (e) => e.code === StatusCode.InvalidArgument,
      );
    });
  });
});
