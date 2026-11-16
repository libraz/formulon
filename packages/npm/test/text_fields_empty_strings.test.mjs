// Empty-value rules of the comment, person, hyperlink, validation, defined-name, protection, table and
// formatReference text fields (api-conventions §5 rule 30): `""` is refused where it is not a legal value, means the
// same as the absent value where the file holds one state, and is a value where the file keeps it.

import assert from 'node:assert/strict';
import test from 'node:test';

import { getModule } from './smoke_support.mjs';

async function withBook(fn) {
  const f = await getModule();
  const wb = f.Workbook.createDefault();
  try {
    return await fn(wb, wb.sheet(0), f);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

const invalid = (f) => (e) => e.name === 'FormulonError' && e.code === f.StatusCode.InvalidArgument;

const CELL = { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 };

// ---- formatReference ----

test('formatReference: sheetName "" is refused, null and omitted write no qualifier', async () => {
  await withBook((wb, _ws, f) => {
    assert.throws(() => wb.formatReference('B2:C3', { sheetName: '' }), invalid(f));
    assert.throws(() => wb.formatReference({ sheetName: '', range: CELL }), invalid(f));
    assert.equal(wb.formatReference('B2:C3', { sheetName: null }), 'B2:C3');
    assert.equal(wb.formatReference('B2:C3'), 'B2:C3');
    assert.equal(wb.formatReference(wb.parseReference('B2:C3')), 'B2:C3');
  });
});

// ---- hyperlink ----

test('setHyperlink: target "" and location "" are refused and add nothing', async () => {
  await withBook((_wb, ws, f) => {
    assert.throws(() => ws.setHyperlink({ range: 'A1', target: '' }), invalid(f));
    assert.throws(() => ws.setHyperlink({ range: 'A1', target: 'https://example.com/', location: '' }), invalid(f));
    assert.throws(() => ws.setHyperlink({ range: 'A1', target: null, location: '' }), invalid(f));
    assert.throws(() => ws.insertHyperlinkAt(0, { range: 'A1', target: '' }), invalid(f));
    assert.equal(ws.listHyperlinks().length, 0);
  });
});

test('setHyperlink: null target or location is absent', async () => {
  await withBook((_wb, ws) => {
    ws.setHyperlink({ range: 'A1', target: 'https://example.com/', location: null });
    ws.setHyperlink({ range: 'B1', target: null, location: 'Sheet1!C3' });
    assert.equal(ws.getHyperlink('A1').location, null);
    assert.equal(ws.getHyperlink('B1').target, null);
    assert.equal(ws.getHyperlink('B1').location, 'Sheet1!C3');
  });
});

test('setHyperlink: display "" and tooltip "" are the absent value', async () => {
  await withBook((_wb, ws) => {
    ws.setHyperlink({ range: 'A1', target: 'https://example.com/', display: '', tooltip: '' });
    ws.setHyperlink({ range: 'B1', target: 'https://example.com/' });
    const empty = ws.getHyperlink('A1');
    const omitted = ws.getHyperlink('B1');
    assert.equal(empty.display, omitted.display);
    assert.equal(empty.tooltip, null);
    assert.equal(omitted.tooltip, null);
  });
});

// ---- data validation ----

const rule = (extra) => ({ ranges: 'D1:D3', type: 1, formula1: '1', formula2: '5', ...extra });

test('addValidation: formula1 "" and formula2 "" are refused and add nothing', async () => {
  await withBook((_wb, ws, f) => {
    assert.throws(() => ws.addValidation(rule({ formula1: '' })), invalid(f));
    assert.throws(() => ws.addValidation(rule({ formula2: '' })), invalid(f));
    assert.equal(ws.listValidations().length, 0);
  });
});

test('addValidation: null formula is absent', async () => {
  await withBook((_wb, ws) => {
    const id = ws.addValidation(rule({ formula2: null }));
    assert.equal(ws.getValidation(id).formula2, null);
    assert.notEqual(ws.getValidation(id).formula1, null);
  });
});

test('updateValidation: "" formula is refused, null clears it', async () => {
  await withBook((_wb, ws, f) => {
    const id = ws.addValidation(rule());
    assert.throws(() => ws.updateValidation(id, { formula1: '' }), invalid(f));
    assert.throws(() => ws.updateValidation(id, { formula2: '' }), invalid(f));
    assert.notEqual(ws.getValidation(id).formula1, null);
    assert.notEqual(ws.getValidation(id).formula2, null);
    ws.updateValidation(id, { formula2: null });
    assert.equal(ws.getValidation(id).formula2, null);
  });
});

test('addValidation: "" titles and messages are the absent value', async () => {
  await withBook((_wb, ws) => {
    const id = ws.addValidation(rule({ errorTitle: '', errorMessage: '', promptTitle: '', promptMessage: '' }));
    const got = ws.getValidation(id);
    assert.equal(got.errorTitle, null);
    assert.equal(got.errorMessage, null);
    assert.equal(got.promptTitle, null);
    assert.equal(got.promptMessage, null);
  });
});

// ---- defined name ----

test('setDefinedName: comment "" and omitted are the same state', async () => {
  await withBook((wb) => {
    wb.setDefinedName({ name: 'WithEmpty', formula: '=Sheet1!$A$1', comment: '' });
    wb.setDefinedName({ name: 'WithNone', formula: '=Sheet1!$A$1' });
    assert.equal(wb.getDefinedName('WithEmpty').comment, null);
    assert.equal(wb.getDefinedName('WithNone').comment, null);
  });
});

// ---- notes ----

test('setComment: empty text writes a note with no text, null removes it', async () => {
  await withBook((_wb, ws) => {
    ws.setComment(1, 1, { text: '' });
    const note = ws.getComment(1, 1);
    assert.notEqual(note, null);
    assert.equal(note.text, '');
    assert.equal(ws.listComments().length, 1);
    ws.setComment(1, 1, null);
    assert.equal(ws.getComment(1, 1), null);
    assert.equal(ws.listComments().length, 0);
  });
});

test('setComment: author "" and omitted are the same state', async () => {
  await withBook((_wb, ws) => {
    ws.setComment(1, 1, { text: 'a', author: '' });
    ws.setComment(2, 2, { text: 'a' });
    assert.equal(ws.getComment(1, 1).author, null);
    assert.equal(ws.getComment(2, 2).author, null);
  });
});

// ---- people and threaded comments ----

test('addPerson: id "" and providerId "" are refused, userId "" is absent', async () => {
  await withBook((wb, _ws, f) => {
    assert.throws(() => wb.addPerson({ id: '', displayName: 'Ann' }), invalid(f));
    assert.throws(() => wb.addPerson({ displayName: 'Ann', providerId: '' }), invalid(f));
    assert.equal(wb.listPeople().length, 0);
    const id = wb.addPerson({ displayName: 'Ann', userId: '' });
    assert.equal(wb.getPerson(id).userId, null);
    assert.equal(wb.getPerson(id).providerId, null);
  });
});

test('addThreadedComment: id "", mention personId "" and mentionId "" are refused; omitted generates', async () => {
  await withBook((wb, ws, f) => {
    const person = wb.addPerson({ displayName: 'Ann' });
    const base = { text: '@Ann hi', personId: person, cell: 'B2' };
    assert.throws(() => ws.addThreadedComment({ ...base, id: '' }), invalid(f));
    assert.throws(() => ws.addThreadedComment({ ...base, mentions: [{ personId: '', start: 0, end: 4 }] }), invalid(f));
    assert.throws(
      () => ws.addThreadedComment({ ...base, mentions: [{ personId: person, mentionId: '', start: 0, end: 4 }] }),
      invalid(f),
    );
    assert.equal(ws.listThreadedComments().length, 0);

    const id = ws.addThreadedComment({ ...base, mentions: [{ personId: person, start: 0, end: 4 }] });
    const [comment] = ws.listThreadedComments();
    assert.equal(comment.id, id);
    assert.match(comment.mentions[0].mentionId, /^\{[0-9A-F-]{36}\}$/);
  });
});

// ---- table ----

test('addTable: name "" and a "" column name are refused; omitted columns derive', async () => {
  await withBook((_wb, ws, f) => {
    assert.throws(() => ws.addTable({ range: 'A1:B3', name: '' }), invalid(f));
    assert.throws(() => ws.addTable({ range: 'A1:B3', name: 'Sales', columns: ['', ''] }), invalid(f));
    assert.throws(() => ws.addTable({ range: 'A1:B3', name: 'Sales', columns: ['A', ''] }), invalid(f));
    // Nothing was created, so the name is still free.
    const table = ws.addTable({ range: 'A1:B3', name: 'Sales' });
    assert.equal(table.name, 'Sales');
  });
});

// ---- protection ----

const RANGE = 'A1:B2';
const PROTECTED_RANGE_TEXT = ['legacyPasswordHash', 'algorithmName', 'hashValue', 'saltValue', 'securityDescriptor'];

for (const key of PROTECTED_RANGE_TEXT) {
  test(`protected range ${key}: "" is refused on add and update, null clears`, async () => {
    await withBook((_wb, ws, f) => {
      assert.throws(() => ws.addProtectedRange({ name: 'R', ranges: RANGE, [key]: '' }), invalid(f));
      assert.equal(ws.listProtectedRanges().length, 0);
      ws.addProtectedRange({ name: 'R', ranges: RANGE, [key]: 'x' });
      assert.equal(ws.listProtectedRanges()[0][key], 'x');
      assert.throws(() => ws.updateProtectedRange('R', { [key]: '' }), invalid(f));
      assert.equal(ws.listProtectedRanges()[0][key], 'x');
      ws.updateProtectedRange('R', { [key]: null });
      assert.equal(ws.listProtectedRanges()[0][key], null);
    });
  });
}

test('protected range: omitted and null strings are absent', async () => {
  await withBook((_wb, ws) => {
    ws.addProtectedRange({ name: 'R', ranges: RANGE });
    ws.addProtectedRange({ name: 'S', ranges: 'C1:D2', hashValue: null });
    for (const range of ws.listProtectedRanges()) {
      for (const key of PROTECTED_RANGE_TEXT) assert.equal(range[key], null);
    }
  });
});

for (const key of ['algorithmName', 'hashValue', 'saltValue', 'legacyPasswordHash']) {
  test(`sheet protection ${key}: "" is refused on set and update, null clears`, async () => {
    await withBook((_wb, ws, f) => {
      assert.throws(() => ws.setSheetProtection({ [key]: '' }), invalid(f));
      assert.equal(ws.getSheetProtection(), null);
      ws.setSheetProtection({ [key]: 'x' });
      assert.throws(() => ws.updateSheetProtection({ [key]: '' }), invalid(f));
      assert.equal(ws.getSheetProtection()[key], 'x');
      ws.updateSheetProtection({ [key]: null });
      assert.equal(ws.getSheetProtection()[key], null);
    });
  });
}

for (const key of [
  'workbookLegacyPasswordHash',
  'workbookAlgorithmName',
  'workbookHashValue',
  'workbookSaltValue',
  'revisionsLegacyPasswordHash',
  'revisionsAlgorithmName',
  'revisionsHashValue',
  'revisionsSaltValue',
]) {
  test(`workbook protection ${key}: "" is refused, null is absent`, async () => {
    await withBook((wb, _ws, f) => {
      assert.throws(() => wb.setWorkbookProtection({ lockStructure: true, [key]: '' }), invalid(f));
      assert.equal(wb.getWorkbookProtection(), null);
      wb.setWorkbookProtection({ lockStructure: true, [key]: null });
      assert.equal(wb.getWorkbookProtection()[key], null);
    });
  });
}

// ---- differential number format ----

test('addDxf: numberFormat "" is the absent value', async () => {
  await withBook((wb) => {
    const empty = wb.ooxml.addDxf({ font: { bold: true }, numberFormat: '' });
    const omitted = wb.ooxml.addDxf({ font: { bold: true } });
    assert.equal(wb.ooxml.getDxf(empty).numberFormat, wb.ooxml.getDxf(omitted).numberFormat);
    assert.equal(wb.ooxml.getDxf(empty).numberFormat, null);
  });
});
