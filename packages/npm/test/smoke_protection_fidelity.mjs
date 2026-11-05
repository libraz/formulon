import test from 'node:test';
import { assert, getModule, readAll } from './smoke_support.mjs';

// ---- Workbook protection, protected ranges and note visibility ---------------

const release = (wb) => wb.delete();

const roundTrip = (mod, wb) => {
  const saved = wb.save();
  assert.ok(saved.status.ok, JSON.stringify(saved.status));
  return mod.Workbook.load(saved.bytes).value;
};

const fullProtection = {
  lockStructure: 1,
  lockWindows: 1,
  lockRevision: 0,
  workbookPassword: 'CF1A',
  workbookAlgorithmName: 'SHA-512',
  workbookHashValue: 'aGFzaA==',
  workbookSaltValue: 'c2FsdA==',
  workbookSpinCount: 100000,
  revisionsPassword: '8D2B',
  revisionsAlgorithmName: 'SHA-256',
  revisionsHashValue: 'cmV2',
  revisionsSaltValue: 'c2Fs',
  revisionsSpinCount: 50000,
};

test('workbook protection is absent by default and every field round-trips', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault().value;
  const empty = wb.getWorkbookProtection();
  assert.ok(empty.status.ok);
  assert.equal(empty.protection.lockStructure, 0);
  assert.equal(empty.protection.workbookHashValue, '');
  assert.ok(wb.setWorkbookProtection(fullProtection).ok);
  assert.deepEqual(wb.getWorkbookProtection().protection, fullProtection);
  const loaded = roundTrip(mod, wb);
  assert.deepEqual(loaded.getWorkbookProtection().protection, fullProtection);
  // An all-empty record removes the element.
  assert.ok(loaded.setWorkbookProtection({}).ok);
  assert.equal(roundTrip(mod, loaded).getWorkbookProtection().protection.lockStructure, 0);
  release(loaded);
  release(wb);
});

test('protected ranges add, list, update, remove and follow row edits', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault().value;
  assert.ok(
    wb.addProtectedRange(0, {
      name: 'Inputs',
      ranges: [
        { firstRow: 1, firstCol: 1, lastRow: 2, lastCol: 2 },
        { firstRow: 4, firstCol: 4, lastRow: 4, lastCol: 4 },
      ],
      password: 'CC1A',
    }).status.ok,
  );
  assert.ok(
    wb.addProtectedRange(0, {
      name: 'Totals',
      ranges: [{ firstRow: 9, firstCol: 0, lastRow: 11, lastCol: 0 }],
      algorithmName: 'SHA-512',
      hashValue: 'aA==',
      saltValue: 'cw==',
      spinCount: 100000,
      securityDescriptor: 'O:WDG:WD',
    }).status.ok,
  );
  const loaded = roundTrip(mod, wb);
  const ranges = readAll(
    () => loaded.getProtectedRangeCount(0),
    (i) => loaded.getProtectedRange(0, i),
  );
  assert.equal(ranges.length, 2);
  assert.equal(ranges[0].name, 'Inputs');
  assert.deepEqual(ranges[0].ranges, [
    { firstRow: 1, firstCol: 1, lastRow: 2, lastCol: 2 },
    { firstRow: 4, firstCol: 4, lastRow: 4, lastCol: 4 },
  ]);
  assert.equal(ranges[0].password, 'CC1A');
  assert.equal(ranges[1].spinCount, 100000);
  assert.equal(ranges[1].securityDescriptor, 'O:WDG:WD');

  assert.ok(loaded.insertRows(0, 0, 2).ok);
  assert.deepEqual(
    readAll(
      () => loaded.getProtectedRangeCount(0),
      (i) => loaded.getProtectedRange(0, i),
    )[0].ranges,
    [
      { firstRow: 3, firstCol: 1, lastRow: 4, lastCol: 2 },
      { firstRow: 6, firstCol: 4, lastRow: 6, lastCol: 4 },
    ],
  );
  assert.ok(
    loaded.updateProtectedRange(0, 0, {
      name: 'Renamed',
      ranges: [{ firstRow: 3, firstCol: 3, lastRow: 3, lastCol: 3 }],
    }).ok,
  );
  assert.equal(
    readAll(
      () => loaded.getProtectedRangeCount(0),
      (i) => loaded.getProtectedRange(0, i),
    )[0].password,
    '',
  );
  assert.ok(loaded.removeProtectedRange(0, 0).ok);
  assert.deepEqual(
    readAll(
      () => loaded.getProtectedRangeCount(0),
      (i) => loaded.getProtectedRange(0, i),
    ).map((r) => r.name),
    ['Totals'],
  );
  release(loaded);
  release(wb);
});

test('protected range arguments are validated', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault().value;
  assert.ok(
    wb.addProtectedRange(0, { name: 'Inputs', ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }] }).status
      .ok,
  );
  assert.equal(
    wb.addProtectedRange(0, { name: 'Inputs', ranges: [{ firstRow: 0, firstCol: 1, lastRow: 0, lastCol: 1 }] }).status
      .ok,
    false,
  );
  assert.equal(
    wb.addProtectedRange(0, { name: '', ranges: [{ firstRow: 0, firstCol: 1, lastRow: 0, lastCol: 1 }] }).status.ok,
    false,
  );
  assert.equal(wb.addProtectedRange(0, { name: 'Other', ranges: [] }).status.ok, false);
  assert.equal(
    wb.addProtectedRange(9, { name: 'Other', ranges: [{ firstRow: 0, firstCol: 1, lastRow: 0, lastCol: 1 }] }).status
      .ok,
    false,
  );
  assert.equal(
    wb.updateProtectedRange(0, 5, { name: 'Other', ranges: [{ firstRow: 0, firstCol: 1, lastRow: 0, lastCol: 1 }] }).ok,
    false,
  );
  assert.equal(wb.removeProtectedRange(0, 5).ok, false);
  assert.equal(
    readAll(
      () => wb.getProtectedRangeCount(0),
      (i) => wb.getProtectedRange(0, i),
    ).length,
    1,
  );
  release(wb);
});

test('note visibility reads back, persists and rejects a missing note', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault().value;
  assert.ok(wb.setComment(0, 1, 1, 'Alice', 'shown').ok);
  assert.ok(wb.setComment(0, 3, 3, 'Alice', 'hidden').ok);
  assert.equal(wb.getCommentVisible(0, 1, 1).value, 0);
  assert.ok(wb.setCommentVisible(0, 1, 1, true).ok);
  const loaded = roundTrip(mod, wb);
  assert.equal(loaded.getCommentVisible(0, 1, 1).value, 1);
  assert.equal(loaded.getCommentVisible(0, 3, 3).value, 0);
  assert.ok(loaded.setComment(0, 1, 1, 'Alice', 'edited').ok);
  assert.equal(loaded.getCommentVisible(0, 1, 1).value, 1);
  assert.equal(loaded.getCommentVisible(0, 7, 7).status.ok, false);
  assert.equal(loaded.setCommentVisible(0, 7, 7, true).ok, false);
  release(loaded);
  release(wb);
});

test('range-list inputs need at least one rectangle', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault().value;
  assert.throws(() => wb.addProtectedRange(0, { name: 'Inputs' }), TypeError);
  assert.throws(() => wb.addValidation(0, { type: 3, formula1: '"A,B"' }), TypeError);
  assert.throws(() => wb.addCfRule(0, { type: 0, formula1: 'TRUE' }), TypeError);
  assert.equal(
    readAll(
      () => wb.getCfRuleCount(0),
      (i) => wb.getCfRule(0, i),
    ).length,
    0,
  );
  assert.equal(wb.addValidation(0, { ranges: [], type: 3, formula1: '"A,B"' }).status.code, 2);
  assert.equal(
    readAll(
      () => wb.getValidationCount(0),
      (i) => wb.getValidation(0, i),
    ).length,
    0,
  );
  release(wb);
});
