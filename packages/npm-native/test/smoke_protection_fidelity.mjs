import test from 'node:test';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';

// ---- Workbook protection, protected ranges and note visibility ---------------

const release = (wb) => wb.dispose();

const roundTrip = (mod, wb) => {
  const saved = wb.save();
  return mod.Workbook.load(saved);
};

const fullProtection = {
  lockStructure: true,
  lockWindows: true,
  lockRevision: false,
  workbookLegacyPasswordHash: 'CF1A',
  workbookAlgorithmName: 'SHA-512',
  workbookHashValue: 'aGFzaA==',
  workbookSaltValue: 'c2FsdA==',
  workbookSpinCount: 100000,
  revisionsLegacyPasswordHash: '8D2B',
  revisionsAlgorithmName: 'SHA-256',
  revisionsHashValue: 'cmV2',
  revisionsSaltValue: 'c2Fs',
  revisionsSpinCount: 50000,
};

test('workbook protection is absent by default and every field round-trips', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.getWorkbookProtection(), null);
  wb.setWorkbookProtection(fullProtection);
  assert.deepEqual(wb.getWorkbookProtection(), fullProtection);
  const loaded = roundTrip(mod, wb);
  assert.deepEqual(loaded.getWorkbookProtection(), fullProtection);
  // A record with every lock off stays present; null removes the element.
  loaded.setWorkbookProtection({});
  assert.equal(roundTrip(mod, loaded).getWorkbookProtection().lockStructure, false);
  loaded.setWorkbookProtection(null);
  assert.equal(roundTrip(mod, loaded).getWorkbookProtection(), null);
  release(loaded);
  release(wb);
});

test('protected ranges add, list, update, remove and follow row edits', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).addProtectedRange({
    name: 'Inputs',
    ranges: [
      { firstRow: 1, firstCol: 1, lastRow: 2, lastCol: 2 },
      { firstRow: 4, firstCol: 4, lastRow: 4, lastCol: 4 },
    ],
    legacyPasswordHash: 'CC1A',
  });
  wb.sheet(0).addProtectedRange({
    name: 'Totals',
    ranges: [{ firstRow: 9, firstCol: 0, lastRow: 11, lastCol: 0 }],
    algorithmName: 'SHA-512',
    hashValue: 'aA==',
    saltValue: 'cw==',
    spinCount: 100000,
    securityDescriptor: 'O:WDG:WD',
  });
  const loaded = roundTrip(mod, wb);
  const ranges = loaded.sheet(0).listProtectedRanges();
  assert.equal(ranges.length, 2);
  assert.equal(ranges[0].name, 'Inputs');
  assert.deepEqual(ranges[0].ranges, [
    { firstRow: 1, firstCol: 1, lastRow: 2, lastCol: 2 },
    { firstRow: 4, firstCol: 4, lastRow: 4, lastCol: 4 },
  ]);
  assert.equal(ranges[0].legacyPasswordHash, 'CC1A');
  assert.equal(ranges[1].spinCount, 100000);
  assert.equal(ranges[1].securityDescriptor, 'O:WDG:WD');

  loaded.sheet(0).insertRows(0, 1);
  assert.deepEqual(loaded.sheet(0).listProtectedRanges()[0].ranges, [
    { firstRow: 3, firstCol: 1, lastRow: 4, lastCol: 2 },
    { firstRow: 6, firstCol: 4, lastRow: 6, lastCol: 4 },
  ]);
  loaded.sheet(0).updateProtectedRange(0, {
    name: 'Renamed',
    ranges: [{ firstRow: 3, firstCol: 3, lastRow: 3, lastCol: 3 }],
  });
  // A partial update leaves the keys it omits as they were.
  assert.equal(loaded.sheet(0).listProtectedRanges()[0].name, 'Renamed');
  assert.equal(loaded.sheet(0).listProtectedRanges()[0].legacyPasswordHash, 'CC1A');
  loaded.sheet(0).updateProtectedRange(0, { legacyPasswordHash: '' });
  assert.equal(loaded.sheet(0).listProtectedRanges()[0].legacyPasswordHash, null);
  loaded.sheet(0).updateProtectedRange(0, { legacyPasswordHash: 'CC1A' });
  loaded.sheet(0).updateProtectedRange(0, { legacyPasswordHash: null });
  assert.equal(loaded.sheet(0).listProtectedRanges()[0].legacyPasswordHash, null);
  assert.deepEqual(loaded.sheet(0).listProtectedRanges()[0].ranges, [
    { firstRow: 3, firstCol: 3, lastRow: 3, lastCol: 3 },
  ]);
  loaded.sheet(0).removeProtectedRange(0);
  assert.deepEqual(
    loaded
      .sheet(0)
      .listProtectedRanges()
      .map((r) => r.name),
    ['Totals'],
  );
  release(loaded);
  release(wb);
});

test('protected range arguments are validated', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).addProtectedRange({ name: 'Inputs', ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }] });
  assert.throws(
    () =>
      wb.sheet(0).addProtectedRange({ name: 'Inputs', ranges: [{ firstRow: 0, firstCol: 1, lastRow: 0, lastCol: 1 }] }),
    isFormulonError,
  );
  assert.throws(
    () => wb.sheet(0).addProtectedRange({ name: '', ranges: [{ firstRow: 0, firstCol: 1, lastRow: 0, lastCol: 1 }] }),
    isFormulonError,
  );
  assert.throws(() => wb.sheet(0).addProtectedRange({ name: 'Other', ranges: [] }), isFormulonError);
  assert.throws(
    () =>
      wb.sheet(9).addProtectedRange({ name: 'Other', ranges: [{ firstRow: 0, firstCol: 1, lastRow: 0, lastCol: 1 }] }),
    isFormulonError,
  );
  assert.throws(
    () =>
      wb
        .sheet(0)
        .updateProtectedRange(5, { name: 'Other', ranges: [{ firstRow: 0, firstCol: 1, lastRow: 0, lastCol: 1 }] }),
    isFormulonError,
  );
  assert.throws(() => wb.sheet(0).removeProtectedRange(5), isFormulonError);
  assert.equal(wb.sheet(0).listProtectedRanges().length, 1);
  release(wb);
});

test('note visibility reads back, persists and rejects a missing note', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setComment(1, 1, { author: 'Alice', text: 'shown' });
  wb.sheet(0).setComment(3, 3, { author: 'Alice', text: 'hidden' });
  assert.equal(wb.sheet(0).getCommentVisible(1, 1), false);
  wb.sheet(0).setCommentVisible(1, 1, true);
  const loaded = roundTrip(mod, wb);
  assert.equal(loaded.sheet(0).getCommentVisible(1, 1), true);
  assert.equal(loaded.sheet(0).getCommentVisible(3, 3), false);
  loaded.sheet(0).setComment(1, 1, { author: 'Alice', text: 'edited' });
  assert.equal(loaded.sheet(0).getCommentVisible(1, 1), true);
  const absent = loaded.sheet(0).getCommentVisible(7, 7);
  assert.equal(absent, null);
  assert.throws(() => loaded.sheet(99).getCommentVisible(1, 1), isFormulonError);
  assert.throws(() => loaded.sheet(0).setCommentVisible(7, 7, true), isFormulonError);
  release(loaded);
  release(wb);
});

test('range-list inputs need at least one rectangle', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.throws(() => wb.sheet(0).addProtectedRange({ name: 'Inputs' }), TypeError);
  assert.throws(() => wb.sheet(0).addValidation({ type: 3, formula1: '"A,B"' }), TypeError);
  assert.throws(() => wb.sheet(0).addCfRule({ type: 0, formula1: 'TRUE' }), TypeError);
  assert.equal(wb.sheet(0).listCfRules().length, 0);
  assert.throws(
    () => wb.sheet(0).addValidation({ ranges: [], type: 3, formula1: '"A,B"' }),
    (e) => isFormulonError(e, 2),
  );
  assert.equal(wb.sheet(0).listValidations().length, 0);
  release(wb);
});
