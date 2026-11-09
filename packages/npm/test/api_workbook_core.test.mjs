// Workbook and sheet core on the unified surface: sheet insert positions, default sheet names,
// sheet lookup, row / column spans, the profile list, the named exports, generated comment ids,
// external-link book numbers, one case-folding rule for named collections and byte inputs.
// FORMULON_NPM_ENTRY selects the staged subpath: "." (default) or "./threads".

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const pkgRoot = path.resolve(here, '..');
const entry = process.env.FORMULON_NPM_ENTRY ?? '.';

let packagePromise;
function loadPackage() {
  packagePromise ??= readFile(path.join(pkgRoot, 'package.json'), 'utf8').then((text) => {
    const pkg = JSON.parse(text);
    const target = pkg.exports?.[entry]?.import;
    if (!target) throw new Error(`package.json "exports" has no import target for ${entry}`);
    return import(pathToFileURL(path.resolve(pkgRoot, target)).href);
  });
  return packagePromise;
}

let formulonPromise;
function getFormulon() {
  formulonPromise ??= loadPackage().then((mod) => mod.createFormulon());
  return formulonPromise;
}

function expectCode(fn, code) {
  assert.throws(fn, (e) => {
    assert.equal(e.name, 'FormulonError', `expected a FormulonError, got ${e}`);
    assert.equal(e.code, code, `code ${e.code} (${e.message})`);
    return true;
  });
}

async function withWorkbook(fn, profile) {
  const formulon = await getFormulon();
  const wb = profile ? formulon.Workbook.createForProfile(profile) : formulon.Workbook.createDefault();
  try {
    return await fn(wb, formulon);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

const names = (wb) => wb.listSheets().map((sheet) => sheet.name);

// ---- Insert position on add, copy and move ---------------------------------------------------

test('addSheet appends by default and lands before the sheet insertBefore names', async () => {
  await withWorkbook((wb) => {
    wb.addSheet('Tail');
    wb.addSheet('ByName', { insertBefore: 'Sheet1' });
    assert.deepEqual(names(wb), ['ByName', 'Sheet1', 'Tail']);
    const tail = wb.sheet('Tail');
    wb.addSheet('ByHandle', { insertBefore: tail });
    wb.addSheet('ByIndex', { insertBefore: 0 });
    assert.deepEqual(names(wb), ['ByIndex', 'ByName', 'Sheet1', 'ByHandle', 'Tail']);
    assert.equal(tail.index, 4);
  });
});

test('addSheet appends when insertBefore is omitted, null or the sheet count', async () => {
  await withWorkbook((wb) => {
    wb.addSheet('A', {});
    wb.addSheet('B', { insertBefore: null });
    wb.addSheet('C', { insertBefore: wb.sheetCount });
    assert.deepEqual(names(wb), ['Sheet1', 'A', 'B', 'C']);
  });
});

test('an insertBefore that names no sheet is a status error and an unknown option key is a TypeError', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    expectCode(() => wb.addSheet('X', { insertBefore: 'Nope' }), StatusCode.NotFound);
    expectCode(() => wb.addSheet('X', { insertBefore: 9 }), StatusCode.SheetIndexOutOfRange);
    assert.throws(() => wb.addSheet('X', { insertAfter: 0 }), TypeError);
    assert.throws(() => wb.addSheet('X', 0), TypeError);
    assert.deepEqual(names(wb), ['Sheet1']);
  });
});

test('copySheet takes insertBefore and newName in one options object and returns the copy', async () => {
  await withWorkbook((wb) => {
    wb.addSheet('Second');
    const source = wb.sheet('Sheet1');
    source.setNumber('A1', 7);
    const front = wb.copySheet(source, { insertBefore: 0, newName: 'Front' });
    assert.equal(front.index, 0);
    assert.equal(source.index, 1);
    assert.equal(front.getTypedValue('A1').number, 7);
    const end = wb.copySheet('Second', { insertBefore: 'Front' });
    assert.equal(end.index, 0);
    assert.equal(wb.copySheet(source).index, wb.sheetCount - 1);
    assert.equal(wb.copySheet(source, { insertBefore: null }).index, wb.sheetCount - 1);
  });
});

test('Sheet.move takes the sheet it lands before, or null for the end', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    for (const name of ['B', 'C', 'D']) wb.addSheet(name);
    const first = wb.sheet(0);
    first.move(wb.sheet('D'));
    assert.deepEqual(names(wb), ['B', 'C', 'Sheet1', 'D']);
    assert.equal(first.index, 2);
    first.move(null);
    assert.deepEqual(names(wb), ['B', 'C', 'D', 'Sheet1']);
    first.move('B');
    assert.deepEqual(names(wb), ['Sheet1', 'B', 'C', 'D']);
    first.move(first);
    first.move(1);
    assert.deepEqual(names(wb), ['Sheet1', 'B', 'C', 'D']);
    first.move(wb.sheetCount);
    assert.equal(first.index, 3);
    expectCode(() => first.move(9), StatusCode.SheetIndexOutOfRange);
    expectCode(() => first.move('Nope'), StatusCode.NotFound);
  });
});

// ---- Default sheet names -----------------------------------------------------------------------

test('addSheet without a name takes the next free default name of the profile', async () => {
  await withWorkbook((wb) => {
    assert.equal(wb.addSheet().name, 'Sheet2');
    assert.equal(wb.addSheet(null).name, 'Sheet3');
    assert.equal(wb.addSheet('Data').name, 'Data');
  });
  await withWorkbook((wb) => {
    assert.equal(wb.addSheet().name, 'Tabelle2');
    wb.addSheet('tabelle3');
    assert.equal(wb.addSheet().name, 'Tabelle4');
  }, 'win-365-de_DE');
});

test('an empty sheet name is refused where a missing one takes the default', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    expectCode(() => wb.addSheet(''), StatusCode.InvalidSheetName);
    assert.throws(() => wb.addSheet(5), TypeError);
    assert.equal(wb.sheetCount, 1);
  });
});

// ---- sheetCount and hasSheet -------------------------------------------------------------------

test('sheetCount follows the sheet list and hasSheet answers a SheetRef without throwing', async () => {
  await withWorkbook((wb) => {
    const other = wb.addSheet('Other');
    assert.equal(wb.sheetCount, 2);
    assert.equal(wb.hasSheet('other'), true);
    assert.equal(wb.hasSheet('Missing'), false);
    assert.equal(wb.hasSheet(1), true);
    assert.equal(wb.hasSheet(2), false);
    assert.equal(wb.hasSheet(other), true);
    wb.removeSheet(other);
    assert.equal(wb.sheetCount, 1);
    assert.equal(wb.hasSheet(other), false);
    assert.equal(wb.hasSheet('Other'), false);
  });
});

test('hasSheet is false for a Sheet of another workbook and refuses a value that is no SheetRef', async () => {
  await withWorkbook(async (wb, formulon) => {
    const foreign = formulon.Workbook.createDefault();
    try {
      assert.equal(wb.hasSheet(foreign.sheet(0)), false);
    } finally {
      foreign.dispose();
    }
    assert.throws(() => wb.hasSheet(null), TypeError);
    assert.throws(() => wb.hasSheet({}), TypeError);
    assert.throws(() => wb.hasSheet(1.5), RangeError);
  });
});

// ---- Row and column spans ----------------------------------------------------------------------

test('insertRows and deleteRows take the closed row span first..last or an A1 span', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    for (let r = 0; r < 6; r += 1) ws.setNumber(r, 0, r);
    ws.insertRows(1, 2);
    assert.deepEqual(
      [0, 1, 2, 3, 4].map((r) => ws.getTypedValue(r, 0).number ?? null),
      [0, null, null, 1, 2],
    );
    ws.deleteRows(1, 2);
    assert.equal(ws.getTypedValue(1, 0).number, 1);
    ws.insertRows('2:4');
    assert.equal(ws.getTypedValue(4, 0).number, 1);
    ws.deleteRows('2:4');
    assert.equal(ws.getTypedValue(1, 0).number, 1);
    ws.insertRows(0, 0);
    assert.equal(ws.getTypedValue(1, 0).number, 0);
  });
});

test('insertColumns and deleteColumns take the closed column span or an A1 span', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    for (let c = 0; c < 5; c += 1) ws.setNumber(0, c, c);
    ws.insertColumns(1, 2);
    assert.equal(ws.getTypedValue(0, 3).number, 1);
    ws.deleteColumns('B:C');
    assert.equal(ws.getTypedValue(0, 1).number, 1);
    ws.deleteColumns(1, 1);
    assert.equal(ws.getTypedValue(0, 1).number, 2);
    ws.insertColumns('C');
    assert.equal(ws.getTypedValue(0, 3).number, 3);
  });
});

test('a span whose first is past its last is InvalidArgument and changes nothing', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    ws.setNumber(0, 0, 1);
    expectCode(() => ws.insertRows(3, 1), StatusCode.InvalidArgument);
    expectCode(() => ws.deleteRows(3, 1), StatusCode.InvalidArgument);
    expectCode(() => ws.insertColumns(3, 1), StatusCode.InvalidArgument);
    expectCode(() => ws.deleteColumns(3, 1), StatusCode.InvalidArgument);
    assert.equal(ws.getTypedValue(0, 0).number, 1);
    assert.throws(() => ws.insertRows(1.5, 2), RangeError);
    assert.throws(() => ws.deleteRows(), TypeError);
  });
});

// ---- Profile list ------------------------------------------------------------------------------

test('listProfileIds lists the 28 profile ids and each one creates a workbook', async () => {
  const formulon = await getFormulon();
  const ids = formulon.listProfileIds();
  assert.equal(ids.length, 28);
  assert.equal(new Set(ids).size, 28);
  assert.ok(ids.includes('win-365-en_US'));
  assert.ok(ids.includes('mac-365-ja_JP'));
  for (const id of ids) {
    const wb = formulon.Workbook.createForProfile(id);
    try {
      assert.equal(wb.getProfileId(), id);
    } finally {
      wb.dispose();
    }
  }
});

// ---- Named exports -----------------------------------------------------------------------------

test('the package exports toNative, mergeFunctionMetadata and enumName as the Formulon members do', async () => {
  const mod = await loadPackage();
  const formulon = await getFormulon();
  for (const name of ['toNative', 'mergeFunctionMetadata', 'enumName']) {
    assert.equal(typeof mod[name], 'function', name);
    assert.equal(mod[name], formulon[name], name);
  }
  assert.equal(mod.enumName(formulon.SheetVisibility, formulon.SheetVisibility.Hidden), 'Hidden');
  await withWorkbook((wb) => {
    wb.sheet(0).setNumber('A1', 5);
    assert.equal(mod.toNative(wb.sheet(0).getTypedValue('A1')), 5);
  });
});

// ---- Threaded comment ids ----------------------------------------------------------------------

const PERSON = '{00000000-0000-4000-8000-0000000000A1}';
const GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/;

test('addThreadedComment generates an id and a created time from the pinned clock and returns the id', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    wb.addPerson({ id: PERSON, displayName: 'A', userId: 'a', providerId: 'None' });
    const clock = { year: 2024, month: 1, day: 2, hour: 3, minute: 4, second: 5 };
    wb.setPinnedNow(clock);
    const ws = wb.sheet(0);
    const id = ws.addThreadedComment({ text: 'x', cell: 'A1', personId: PERSON });
    assert.match(id, GUID);
    const read = ws.getThreadedComment(id);
    assert.equal(read.id, id);
    assert.deepEqual(read.created, clock);
    const second = ws.addThreadedComment({ text: 'y', cell: 'A2', personId: PERSON });
    assert.notEqual(second, id);
    const given = '{00000000-0000-4000-8000-000000000001}';
    assert.equal(ws.addThreadedComment({ id: given, text: 'z', cell: 'A3', personId: PERSON }), given);
    expectCode(
      () => ws.addThreadedComment({ id: given, text: 'w', cell: 'A4', personId: PERSON }),
      StatusCode.AlreadyExists,
    );
    assert.equal(ws.listThreadedComments().length, 3);
  });
});

// ---- External-link book number -----------------------------------------------------------------

test('an external link reads its bookNumber, the [n] of formula text', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setFormula('A1', "='[Book2.xlsx]Sheet1'!A1");
    const [link] = wb.listExternalLinks();
    assert.equal(link.bookNumber, 1);
    assert.equal('index' in link, false);
    assert.match(ws.getFormula('A1'), /\[Book2\.xlsx\]/);
  });
});

test('book numbers follow the document order of the external links', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setFormula('A1', "='[First.xlsx]Sheet1'!A1");
    ws.setFormula('A2', "='[Second.xlsx]Sheet1'!A1");
    assert.deepEqual(
      wb.listExternalLinks().map((link) => [link.bookNumber, link.target]),
      [
        [1, 'First.xlsx'],
        [2, 'Second.xlsx'],
      ],
    );
  });
});

// ---- One case-folding rule for named collections -----------------------------------------------

test('protected range names compare case-insensitively for uniqueness and lookup', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    ws.addProtectedRange({ name: 'Budget', ranges: 'A1' });
    expectCode(() => ws.addProtectedRange({ name: 'BUDGET', ranges: 'B1' }), StatusCode.AlreadyExists);
    ws.removeProtectedRange('budget');
    assert.equal(ws.listProtectedRanges().length, 0);
  });
});

test('cell style names compare case-insensitively', async () => {
  await withWorkbook((wb) => {
    assert.equal(wb.getCellStyle('NORMAL').name, 'Normal');
    assert.equal(wb.getCellStyle('normal').name, 'Normal');
    assert.equal(wb.getCellStyle('Missing'), null);
  });
});

test('a non-ASCII table name collides with itself in another case', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    for (const cell of ['A1', 'A2', 'C1', 'C2']) ws.setText(cell, 'x');
    ws.addTable({ range: 'A1:A2', name: 'Ärger' });
    expectCode(() => ws.addTable({ range: 'C1:C2', name: 'ärger' }), StatusCode.AlreadyExists);
    assert.equal(ws.listTables().length, 1);
  });
});

test('a sheet name compares by Unicode simple case folding', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    wb.addSheet('Ä');
    expectCode(() => wb.addSheet('ä'), StatusCode.AlreadyExists);
    assert.equal(wb.sheet('ä').name, 'Ä');
  });
});

// ---- Byte inputs -------------------------------------------------------------------------------

// A 2 x 3 RGB PNG header: enough for the image probe, which reads the size only.
const PNG = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 2, 0, 0, 0, 3, 8, 2, 0,
  0, 0, 0, 0, 0, 0,
]);

test('getImageInfo takes every byte input load takes and refuses anything else', async () => {
  const formulon = await getFormulon();
  const expected = formulon.getImageInfo(PNG);
  assert.equal(expected.widthPx, 2);
  assert.equal(expected.heightPx, 3);
  assert.deepEqual(formulon.getImageInfo(PNG.buffer.slice(0)), expected);
  assert.deepEqual(formulon.getImageInfo(new DataView(PNG.buffer.slice(0))), expected);
  assert.deepEqual(formulon.getImageInfo(Buffer.from(PNG)), expected);
  assert.throws(() => formulon.getImageInfo('not bytes'), TypeError);
  assert.throws(() => formulon.getImageInfo(null), TypeError);
});

test('addPicture takes an ArrayBuffer and a DataView as it takes a Uint8Array', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.addPicture(PNG);
    ws.addPicture(PNG.buffer.slice(0));
    ws.addPicture(new DataView(PNG.buffer.slice(0)));
    assert.equal(ws.listDrawings().length, 3);
    assert.throws(() => ws.addPicture('not bytes'), TypeError);
  });
});

// ---- Handles are never constructed -------------------------------------------------------------

test('handle classes have no public constructor and a Table reads its Sheet', async () => {
  await withWorkbook((wb, formulon) => {
    const ws = wb.sheet(0);
    ws.setText('A1', 'h');
    ws.setText('A2', '1');
    const table = ws.addTable({ range: 'A1:A2', name: 'T' });
    assert.equal(table.sheet, ws);
    assert.equal(wb.table('T').sheet, ws);
    for (const handle of [table, ws]) {
      assert.throws(() => new handle.constructor(wb, 3), TypeError);
    }
    for (const name of ['Table', 'Pivot', 'PivotCache', 'Drawing', 'Picture', 'Chart', 'Shape', 'Slicer']) {
      assert.equal(name in formulon, false, name);
    }
  });
});
