import test from 'node:test';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';

// ---- Lookups by natural key ------------------------------------------------

const release = (wb) => wb.dispose();

const ALICE = '{0A1B2C3D-0000-4000-8000-000000000001}';
const THREAD = '{6F1E0D2C-1111-4111-8111-000000000001}';

// A found key carries its record; a missing one is ok with a null value.
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
  wb.setDefinedName({ name: 'Rate', formula: 'Sheet1!$B$1', localSheetId: 0 });
  const name = found(wb.getDefinedName('rate', null));
  assert.deepEqual(
    [name.name, name.formula, name.localSheetId, name.comment, name.hidden],
    ['Rate', '=Sheet1!$A$1', null, null, false],
  );
  assert.equal(found(wb.getDefinedName('Rate', 0)).formula, '=Sheet1!$B$1');
  assert.equal(found(wb.getDefinedName('Rate', 'Sheet1')).formula, '=Sheet1!$B$1');
  assert.equal(found(wb.getDefinedName('Rate')).formula, '=Sheet1!$A$1');
  missing(wb.getDefinedName('Missing', null));
  wb.removeDefinedName('Rate');
  missing(wb.getDefinedName('Rate'));
  assert.equal(found(wb.getDefinedName('Rate', 0)).formula, '=Sheet1!$B$1');
  assert.throws(
    () => wb.removeDefinedName('Rate', null),
    (e) => isFormulonError(e, mod.StatusCode.NotFound),
  );
  release(wb);
});

test('hyperlinks, people and threaded comments are read by key', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).addHyperlink(2, 1, { target: 'https://example.com/' });
  assert.equal(found(wb.sheet(0).getHyperlink(2, 1)).target, 'https://example.com/');
  missing(wb.sheet(0).getHyperlink(2, 2));
  assert.throws(() => wb.sheet(9).getHyperlink(2, 1), isFormulonError);

  wb.addPerson({ id: ALICE, displayName: 'Alice', userId: 'alice', providerId: 'None' });
  assert.equal(found(wb.getPerson(ALICE)).displayName, 'Alice');
  missing(wb.getPerson(THREAD));
  wb.sheet(0).addThreadedComment({
    id: THREAD,
    row: 4,
    col: 2,
    personId: ALICE,
    created: '2026-01-02T03:04:05.67',
    text: 'Check this',
    parentId: '',
  });
  const thread = found(wb.sheet(0).getThreadedComment(THREAD));
  assert.deepEqual([thread.row, thread.col, thread.text], [4, 2, 'Check this']);
  missing(wb.sheet(0).getThreadedComment(ALICE));
  release(wb);
});

test('shapes and drawing objects are read by object id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // anchorKind 0 is OneCell.
  const added = wb.sheet(0).addShape({
    presetGeometry: 'rect',
    anchorKind: 0,
    fromRow: 2,
    widthEmu: 914400,
    heightEmu: 457200,
  });
  assert.equal(wb.sheet(0).shape(added).getProperties().presetGeometry, 'rect');
  assert.equal(wb.sheet(0).drawing(added).getDrawingProperties().fromRow, 2);
  // An id the sheet does not hold is NotFound, and so is a removed object.
  const absent = (read) => assert.throws(read, (e) => isFormulonError(e, mod.StatusCode.NotFound));
  absent(() => wb.sheet(0).shape(added.objectId + 1));
  absent(() => wb.sheet(0).drawing(added.objectId + 1));
  wb.sheet(0).removeDrawing(added);
  absent(() => added.getProperties());
  absent(() => added.getDrawingProperties());
  release(wb);
});

test('theme script fonts and passthrough parts are read by key', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(found(wb.getThemeScriptFont({ major: true, script: 'Jpan' })).typeface.length > 0);
  wb.setThemeScriptFont({ major: false, script: 'Zzzz', typeface: 'Test Face' });
  assert.deepEqual(found(wb.getThemeScriptFont({ major: false, script: 'Zzzz' })), {
    major: false,
    script: 'Zzzz',
    typeface: 'Test Face',
  });
  missing(wb.getThemeScriptFont({ major: true, script: 'Zzzz' }));
  const path = wb.listPassthroughs()[0];
  assert.equal(found(wb.getPassthrough(path)), path);
  missing(wb.getPassthrough(`/${path}`));
  release(wb);
});

test('index-keyed adds return the index they allocate', async () => {
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
  assert.equal(wb.sheet(1).addValidation(rule), 0);
  assert.equal(wb.sheet(1).addValidation(rule), 1);
  release(wb);
});
