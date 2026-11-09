import test from 'node:test';
import { assert, expectFailure, getModule } from './smoke_support.mjs';

// ---- Lookups by natural key ------------------------------------------------

const release = (wb) => wb.dispose();

const ALICE = '{0A1B2C3D-0000-4000-8000-000000000001}';
const THREAD = '{6F1E0D2C-1111-4111-8111-000000000001}';

// A found key returns its record; a missing one returns null.
const found = (r) => {
  assert.notEqual(r, null);
  return r;
};
const missing = (r) => {
  assert.equal(r, null);
};

test('defined names are read and removed by name and scope', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setDefinedName({ name: 'Rate', formula: 'Sheet1!$A$1' });
  wb.setDefinedName({ name: 'Rate', formula: 'Sheet1!$B$1', scope: 0 });
  const name = found(wb.getDefinedName('rate', null));
  assert.deepEqual(
    [name.name, name.formula, name.scope, name.comment, name.hidden],
    ['Rate', '=Sheet1!$A$1', null, null, false],
  );
  assert.equal(found(wb.getDefinedName('Rate', 0)).formula, '=Sheet1!$B$1');
  missing(wb.getDefinedName('Missing', null));
  wb.removeDefinedName('Rate', null);
  missing(wb.getDefinedName('Rate', null));
  assert.equal(found(wb.getDefinedName('Rate', 0)).formula, '=Sheet1!$B$1');
  expectFailure(() => wb.removeDefinedName('Rate', null), mod.StatusCode.NotFound);
  release(wb);
});

test('hyperlinks, people and threaded comments are read by key', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setHyperlink(2, 1, { target: 'https://example.com/' });
  assert.equal(found(wb.sheet(0).getHyperlink(2, 1)).target, 'https://example.com/');
  missing(wb.sheet(0).getHyperlink(2, 2));
  expectFailure(() => wb.sheet(9).getHyperlink(2, 1), mod.StatusCode.SheetIndexOutOfRange);

  wb.addPerson({ id: ALICE, displayName: 'Alice', userId: 'alice', providerId: 'None' });
  assert.equal(found(wb.getPerson(ALICE)).displayName, 'Alice');
  missing(wb.getPerson(THREAD));
  wb.sheet(0).addThreadedComment({
    id: THREAD,
    cell: { row: 4, col: 2 },
    personId: ALICE,
    created: { year: 2026, month: 1, day: 2, hour: 3, minute: 4, second: 5 },
    text: 'Check this',
    parentId: '',
  });
  const thread = found(wb.sheet(0).getThreadedComment(THREAD));
  assert.deepEqual([thread.cell.row, thread.cell.col, thread.text], [4, 2, 'Check this']);
  missing(wb.sheet(0).getThreadedComment(ALICE));
  release(wb);
});

test('shapes and drawing objects are read by object id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const shape = wb.sheet(0).addShape({
    presetGeometry: 'rect',
    anchor: { kind: mod.AnchorKind.OneCell, fromRow: 2, widthEmu: 914400, heightEmu: 457200 },
  });
  const added = shape.objectId;
  assert.equal(shape.getProperties().presetGeometry, 'rect');
  assert.equal(shape.getDrawingProperties().anchor.fromRow, 2);
  assert.equal(wb.sheet(0).drawing(added), shape);
  // An id the sheet does not hold is NotFound.
  expectFailure(() => wb.sheet(0).shape(added + 1), mod.StatusCode.NotFound);
  expectFailure(() => wb.sheet(0).drawing(added + 1), mod.StatusCode.NotFound);
  release(wb);
});

test('theme script fonts and passthrough parts are queried by path', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(found(wb.getThemeScriptFont({ scheme: mod.FontScheme.Major, script: 'Jpan' })).typeface.length > 0);
  wb.setThemeScriptFont({ scheme: mod.FontScheme.Minor, script: 'Zzzz', typeface: 'Test Face' });
  assert.deepEqual(found(wb.getThemeScriptFont({ scheme: mod.FontScheme.Minor, script: 'Zzzz' })), {
    scheme: mod.FontScheme.Minor,
    script: 'Zzzz',
    typeface: 'Test Face',
  });
  missing(wb.getThemeScriptFont({ scheme: mod.FontScheme.Major, script: 'Zzzz' }));
  const path = wb.ooxml.listPassthroughs()[0];
  assert.equal(wb.ooxml.hasPassthrough(path), true);
  assert.equal(wb.ooxml.hasPassthrough(`/${path}`), false);
  release(wb);
});

test('adds return the index or the id they allocate', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const sheet = wb.addSheet('Second');
  assert.equal(sheet.index, 1);
  assert.equal(sheet, wb.sheet(1));
  const rule = {
    ranges: [{ firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 0 }],
    type: 1,
    formula1: '1',
    formula2: '9',
  };
  const first = wb.sheet(1).addValidation(rule);
  const second = wb.sheet(1).addValidation(rule);
  assert.equal(typeof first, 'string');
  assert.notEqual(first, second);
  assert.deepEqual(
    wb
      .sheet(1)
      .listValidations()
      .map((v) => v.id),
    [first, second],
  );
  release(wb);
});
