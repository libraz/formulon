import test from 'node:test';
import { assert, getModule } from './smoke_support.mjs';

// ---- Lookups by natural key ------------------------------------------------

const release = (wb) => wb.dispose();

const ALICE = '{0A1B2C3D-0000-4000-8000-000000000001}';
const THREAD = '{6F1E0D2C-1111-4111-8111-000000000001}';

// A found key carries its record; a missing one is ok with a null value.
const found = (r) => {
  assert.ok(r.status.ok, JSON.stringify(r.status));
  assert.notEqual(r.value, null);
  return r.value;
};
const missing = (r) => {
  assert.ok(r.status.ok, JSON.stringify(r.status));
  assert.equal(r.value, null);
};

test('defined names are read and removed by name and scope', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault().value;
  assert.ok(wb.setDefinedName('Rate', 'Sheet1!$A$1').ok);
  assert.ok(wb.setDefinedNameScoped('Rate', 'Sheet1!$B$1', 0).ok);
  const name = found(wb.getDefinedName('rate', -1));
  assert.deepEqual(
    [name.name, name.formula, name.localSheetId, name.comment, name.hidden],
    ['Rate', 'Sheet1!$A$1', -1, '', false],
  );
  assert.equal(found(wb.getDefinedName('Rate', 0)).formula, 'Sheet1!$B$1');
  missing(wb.getDefinedName('Missing', -1));
  assert.ok(wb.removeDefinedName('Rate', -1).ok);
  missing(wb.getDefinedName('Rate', -1));
  assert.equal(found(wb.getDefinedName('Rate', 0)).formula, 'Sheet1!$B$1');
  assert.equal(wb.removeDefinedName('Rate', -1).ok, false);
  release(wb);
});

test('hyperlinks, people and threaded comments are read by key', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault().value;
  assert.ok(wb.addHyperlink(0, 2, 1, 'https://example.com/', '', '', '').ok);
  assert.equal(found(wb.getHyperlink(0, 2, 1)).target, 'https://example.com/');
  missing(wb.getHyperlink(0, 2, 2));
  assert.equal(wb.getHyperlink(9, 2, 1).status.ok, false);

  assert.ok(wb.addPerson({ id: ALICE, displayName: 'Alice', userId: 'alice', providerId: 'None' }).ok);
  assert.equal(found(wb.getPerson(ALICE)).displayName, 'Alice');
  missing(wb.getPerson(THREAD));
  assert.ok(
    wb.addThreadedComment(0, {
      id: THREAD,
      row: 4,
      col: 2,
      personId: ALICE,
      created: '2026-01-02T03:04:05.67',
      text: 'Check this',
      parentId: '',
    }).ok,
  );
  const thread = found(wb.getThreadedComment(0, THREAD));
  assert.deepEqual([thread.row, thread.col, thread.text], [4, 2, 'Check this']);
  missing(wb.getThreadedComment(0, ALICE));
  release(wb);
});

test('shapes and drawing objects are read by object id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault().value;
  // anchorKind 0 is OneCell.
  const added = wb.addShape(0, { prst: 'rect', anchorKind: 0, fromRow: 2, cx: 914400, cy: 457200 });
  assert.ok(added.status.ok, JSON.stringify(added.status));
  assert.equal(found(wb.getShape(0, added.objectId)).prst, 'rect');
  assert.equal(found(wb.getDrawingObject(0, added.objectId)).fromRow, 2);
  missing(wb.getShape(0, added.objectId + 1));
  missing(wb.getDrawingObject(0, added.objectId + 1));
  release(wb);
});

test('theme script fonts and passthrough parts are read by key', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault().value;
  assert.ok(found(wb.getThemeScriptFont(true, 'Jpan')).typeface.length > 0);
  assert.ok(wb.setThemeScriptFont(false, 'Zzzz', 'Test Face').ok);
  assert.deepEqual(found(wb.getThemeScriptFont(false, 'Zzzz')), { script: 'Zzzz', typeface: 'Test Face' });
  missing(wb.getThemeScriptFont(true, 'Zzzz'));
  const path = wb.listPassthroughs()[0];
  assert.equal(found(wb.getPassthrough(path)), path);
  missing(wb.getPassthrough(`/${path}`));
  release(wb);
});

test('index-keyed adds return the index they allocate', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault().value;
  const sheet = wb.addSheet('Second');
  assert.ok(sheet.status.ok);
  assert.equal(sheet.index, 1);
  const rule = {
    ranges: [{ firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 0 }],
    type: 1,
    formula1: '1',
    formula2: '9',
  };
  assert.equal(wb.addValidation(1, rule).index, 0);
  assert.equal(wb.addValidation(1, rule).index, 1);
  release(wb);
});
