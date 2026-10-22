import test from 'node:test';
import { assert, getModule } from './smoke_support.mjs';

// ---- Workbook protection, protected ranges and note visibility ---------------

const release = (wb) => wb.dispose();

const roundTrip = (mod, wb) => {
  const saved = wb.save();
  assert.ok(saved.status.ok, JSON.stringify(saved.status));
  return mod.Workbook.loadBytes(saved.bytes);
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
  const wb = mod.Workbook.createDefault();
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
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.addProtectedRange(0, { name: 'Inputs', sqref: 'B2:C3 E5', password: 'CC1A' }).ok);
  assert.ok(
    wb.addProtectedRange(0, {
      name: 'Totals',
      sqref: 'A10:A12',
      algorithmName: 'SHA-512',
      hashValue: 'aA==',
      saltValue: 'cw==',
      spinCount: 100000,
      securityDescriptor: 'O:WDG:WD',
    }).ok,
  );
  const loaded = roundTrip(mod, wb);
  const ranges = loaded.getProtectedRanges(0);
  assert.ok(ranges.status.ok);
  assert.equal(ranges.length, 2);
  assert.equal(ranges[0].name, 'Inputs');
  assert.equal(ranges[0].sqref, 'B2:C3 E5');
  assert.equal(ranges[0].password, 'CC1A');
  assert.equal(ranges[1].spinCount, 100000);
  assert.equal(ranges[1].securityDescriptor, 'O:WDG:WD');

  assert.ok(loaded.insertRows(0, 0, 2).ok);
  assert.equal(loaded.getProtectedRanges(0)[0].sqref, 'B4:C5 E7');
  assert.ok(loaded.updateProtectedRange(0, 0, { name: 'Renamed', sqref: 'D4' }).ok);
  assert.equal(loaded.getProtectedRanges(0)[0].password, '');
  assert.ok(loaded.removeProtectedRange(0, 0).ok);
  assert.deepEqual(
    loaded.getProtectedRanges(0).map((r) => r.name),
    ['Totals'],
  );
  release(loaded);
  release(wb);
});

test('protected range arguments are validated', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.addProtectedRange(0, { name: 'Inputs', sqref: 'A1' }).ok);
  assert.equal(wb.addProtectedRange(0, { name: 'Inputs', sqref: 'B1' }).ok, false);
  assert.equal(wb.addProtectedRange(0, { name: '', sqref: 'B1' }).ok, false);
  assert.equal(wb.addProtectedRange(0, { name: 'Other', sqref: 'not a range' }).ok, false);
  assert.equal(wb.addProtectedRange(9, { name: 'Other', sqref: 'B1' }).ok, false);
  assert.equal(wb.updateProtectedRange(0, 5, { name: 'Other', sqref: 'B1' }).ok, false);
  assert.equal(wb.removeProtectedRange(0, 5).ok, false);
  assert.equal(wb.getProtectedRanges(0).length, 1);
  release(wb);
});

test('note visibility reads back, persists and rejects a missing note', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
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
