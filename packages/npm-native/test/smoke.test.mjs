// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Smoke tests for the staged @libraz/formulon-native npm package.
//
// These tests intentionally import the *staged* dist/ artefacts (resolved
// via the package's own package.json "main" field), not the raw build/
// output. That way we catch staging mistakes -- missing files, wrong
// relative paths, broken "exports" entries -- before the tarball ships.
//
// Run via `make node-test` (which runs `make node-package` first).

import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const pkgRoot = path.resolve(__dirname, '..');
const pkgJsonPath = path.join(pkgRoot, 'package.json');

// Resolve the staged entry point through the package's own package.json
// "main" field. Fails (rather than skips) if the package isn't staged --
// that's the whole point of running these tests against dist/.
async function loadStagedModule() {
  const raw = await readFile(pkgJsonPath, 'utf8');
  const pkg = JSON.parse(raw);
  const main = pkg.main;
  if (!main) {
    throw new Error(`package.json is missing "main": ${pkgJsonPath}`);
  }
  const mainPath = path.resolve(pkgRoot, main);
  // Node ESM requires file:// URLs for absolute paths on Windows;
  // pathToFileURL is the portable form.
  return import(pathToFileURL(mainPath).href);
}

let modPromise;
function getModule() {
  if (!modPromise) {
    modPromise = loadStagedModule();
  }
  return modPromise;
}

test('default export exposes Workbook + evalFormula + version', async () => {
  const mod = await getModule();
  assert.equal(typeof mod.Workbook, 'function');
  assert.equal(typeof mod.evalFormula, 'function');
  assert.equal(typeof mod.version, 'function');
  assert.equal(typeof mod.statusString, 'function');
  assert.equal(typeof mod.ValueKind, 'object');
  assert.equal(typeof mod.PivotCellKind, 'object');
});

test('version() returns a non-empty string', async () => {
  const mod = await getModule();
  const v = mod.version();
  assert.equal(typeof v, 'string');
  assert.ok(v.length > 0, `expected non-empty version, got ${JSON.stringify(v)}`);
});

test("evalFormula('=1+2') returns NUMBER 3", async () => {
  const mod = await getModule();
  const r = mod.evalFormula('=1+2');
  assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
  assert.equal(r.value.kind, mod.ValueKind.Number);
  assert.equal(r.value.number, 3);
});

test("evalFormula('=#REF!') surfaces ERROR value (not failed status)", async () => {
  const mod = await getModule();
  const r = mod.evalFormula('=#REF!');
  assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
  assert.equal(r.value.kind, mod.ValueKind.Error);
});

test('Workbook.createDefault + setNumber + getValue round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.sheetCount(), 1);
  const setRc = wb.setNumber(0, 0, 0, 42);
  assert.ok(setRc.ok, `setNumber: ${JSON.stringify(setRc)}`);
  const r = wb.getValue(0, 0, 0);
  assert.ok(r.status.ok, `getValue: ${JSON.stringify(r.status)}`);
  assert.equal(r.value.kind, mod.ValueKind.Number);
  assert.equal(r.value.number, 42);
});

test('pivotCount + pivotLayout expose PivotTable projection status', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.pivotCount(0), 0);

  const missing = wb.pivotLayout(0, 0);
  assert.equal(missing.status.ok, false);
  assert.notEqual(missing.status.status, 0);
  assert.equal(missing.top, 0);
  assert.equal(missing.left, 0);
  assert.equal(missing.rows, 0);
  assert.equal(missing.cols, 0);
  assert.deepEqual(missing.cells, []);
});

test("Workbook.createDefault + setFormula '=1+2' + recalc -> 3", async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setFormula(0, 0, 0, '=1+2').ok);
  assert.ok(wb.recalc().ok);
  const r = wb.getValue(0, 0, 0);
  assert.ok(r.status.ok);
  assert.equal(r.value.kind, mod.ValueKind.Number);
  assert.equal(r.value.number, 3);
});

test("evaluateFormulaArray('=SEQUENCE(2,3)') returns a 2x3 grid", async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const r = wb.evaluateFormulaArray(0, 0, 0, '=SEQUENCE(2,3)');
  assert.ok(r.status.ok);
  assert.equal(r.rows, 2);
  assert.equal(r.cols, 3);
  assert.equal(r.cells.length, 2);
  assert.equal(r.cells[0].length, 3);
  // Row-major 1..6.
  assert.equal(r.cells[0][0].kind, mod.ValueKind.Number);
  assert.equal(r.cells[0][0].number, 1);
  assert.equal(r.cells[1][2].number, 6);
});

test("evaluateFormulaArray('=1+2') reports a 1x1 scalar array", async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const r = wb.evaluateFormulaArray(0, 0, 0, '=1+2');
  assert.ok(r.status.ok);
  assert.equal(r.rows, 1);
  assert.equal(r.cols, 1);
  assert.equal(r.cells[0][0].kind, mod.ValueKind.Number);
  assert.equal(r.cells[0][0].number, 3);
});

test('Workbook.createEmpty + addSheet -> sheetCount/sheetName work', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createEmpty();
  assert.ok(wb.addSheet('Sheet1').ok);
  assert.equal(wb.sheetCount(), 1);
  const sn = wb.sheetName(0);
  assert.ok(sn.status.ok);
  assert.equal(sn.value, 'Sheet1');
});

test('save() returns Uint8Array; loadBytes round-trips the value', async () => {
  const mod = await getModule();
  const wb1 = mod.Workbook.createDefault();
  assert.ok(wb1.setNumber(0, 0, 0, 42).ok);
  const sr = wb1.save();
  assert.ok(sr.status.ok, `save: ${JSON.stringify(sr.status)}`);
  assert.ok(sr.bytes instanceof Uint8Array, 'expected Uint8Array');
  assert.ok(sr.bytes.length > 0, 'expected non-empty save buffer');

  const wb2 = mod.Workbook.loadBytes(sr.bytes);
  const r = wb2.getValue(0, 0, 0);
  assert.ok(r.status.ok, `loaded getValue: ${JSON.stringify(r.status)}`);
  assert.equal(r.value.kind, mod.ValueKind.Number);
  assert.equal(r.value.number, 42);
});

// -- Expanded surface coverage -----------------------------------------
// These tests exercise the methods the addon mirrors from the embind
// binding. They are deliberately shallow: each call should round-trip
// some observable state without crashing the Node process.

test('isValid returns true for a live workbook', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.isValid(), true);
});

test('addMerge + getMerges round-trip a single range', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const range = { firstRow: 1, firstCol: 1, lastRow: 2, lastCol: 3 };
  const ar = wb.addMerge(0, range);
  assert.ok(ar.ok, `addMerge: ${JSON.stringify(ar)}`);
  const list = wb.getMerges(0);
  assert.ok(Array.isArray(list), `expected Array, got ${typeof list}`);
  assert.equal(list.length, 1);
  assert.deepEqual(
    {
      firstRow: list[0].firstRow,
      firstCol: list[0].firstCol,
      lastRow: list[0].lastRow,
      lastCol: list[0].lastCol,
    },
    range,
  );
});

test('addHyperlink + getHyperlinks round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.getHyperlinks(0).length, 0);
  // Three entries with progressively more optional fields populated.
  assert.ok(wb.addHyperlink(0, 1, 2, 'https://example.com/', '', '').ok);
  assert.ok(wb.addHyperlink(0, 3, 4, 'mailto:hello@example.com', 'Hello', '').ok);
  assert.ok(wb.addHyperlink(0, 5, 6, 'https://example.org/x', 'See X', 'External link').ok);
  const list = wb.getHyperlinks(0);
  assert.equal(list.length, 3);
  assert.equal(list[0].row, 1);
  assert.equal(list[0].col, 2);
  assert.equal(list[0].target, 'https://example.com/');
  assert.equal(list[0].display, '');
  assert.equal(list[0].tooltip, '');
  assert.equal(list[1].target, 'mailto:hello@example.com');
  assert.equal(list[1].display, 'Hello');
  assert.equal(list[2].display, 'See X');
  assert.equal(list[2].tooltip, 'External link');
  // Sheet-out-of-range is rejected.
  assert.ok(!wb.addHyperlink(999, 0, 0, 'https://x/', '', '').ok);
  // clearHyperlinks drops everything.
  assert.ok(wb.clearHyperlinks(0).ok);
  assert.equal(wb.getHyperlinks(0).length, 0);
});

test('removeHyperlink / removeHyperlinkAt / clearHyperlinks surface on an empty sheet', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // No-op variants on an empty list still return kOk.
  assert.ok(wb.removeHyperlink(0, 0, 0).ok);
  assert.ok(wb.clearHyperlinks(0).ok);
  // Out-of-range index is rejected.
  assert.ok(!wb.removeHyperlinkAt(0, 0).ok);
  // Sheet-index-out-of-range is rejected on every variant.
  assert.ok(!wb.removeHyperlink(999, 0, 0).ok);
  assert.ok(!wb.removeHyperlinkAt(999, 0).ok);
  assert.ok(!wb.clearHyperlinks(999).ok);
  // The hyperlink list is still empty afterwards.
  assert.equal(wb.getHyperlinks(0).length, 0);
});

test('removeMerge + removeMergeAt + clearMerges step-wise prune the merge list', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const a = { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 };
  const b = { firstRow: 4, firstCol: 4, lastRow: 5, lastCol: 5 };
  assert.ok(wb.addMerge(0, a).ok);
  assert.ok(wb.addMerge(0, b).ok);
  // removeMerge with an overlap that hits `a` only.
  const overlap = { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 };
  assert.ok(wb.removeMerge(0, overlap).ok);
  let list = wb.getMerges(0);
  assert.equal(list.length, 1);
  assert.equal(list[0].firstRow, 4);
  // removeMergeAt drops the survivor by index.
  assert.ok(wb.removeMergeAt(0, 0).ok);
  list = wb.getMerges(0);
  assert.equal(list.length, 0);
  // clearMerges is a no-op on an empty list and stays kOk.
  assert.ok(wb.addMerge(0, a).ok);
  assert.ok(wb.addMerge(0, b).ok);
  assert.ok(wb.clearMerges(0).ok);
  list = wb.getMerges(0);
  assert.equal(list.length, 0);
});

test('setSheetZoom + getSheetView surface the new zoom', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setSheetZoom(0, 175).ok);
  const v = wb.getSheetView(0);
  assert.ok(v.status.ok, `getSheetView: ${JSON.stringify(v.status)}`);
  assert.equal(v.view.zoomScale, 175);
  // Default freeze / tab-hidden state survives a zoom change.
  assert.equal(v.view.freezeRows, 0);
  assert.equal(v.view.freezeCols, 0);
  assert.equal(v.view.tabHidden, 0);
});

test('getSheetView surfaces display / orientation defaults', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const v = wb.getSheetView(0);
  assert.ok(v.status.ok, `getSheetView: ${JSON.stringify(v.status)}`);
  assert.equal(v.view.showGridLines, 1);
  assert.equal(v.view.showRowColHeaders, 1);
  assert.equal(v.view.showZeros, 1);
  assert.equal(v.view.rightToLeft, 0);
  assert.equal(v.view.tabSelected, 0);
  assert.equal(v.view.viewMode, '');
});

test('setSheetShowGridLines / setSheetRightToLeft / setSheetTabSelected / setSheetViewMode round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setSheetShowGridLines(0, false).ok);
  assert.ok(wb.setSheetShowRowColHeaders(0, false).ok);
  assert.ok(wb.setSheetShowZeros(0, false).ok);
  assert.ok(wb.setSheetRightToLeft(0, true).ok);
  assert.ok(wb.setSheetTabSelected(0, true).ok);
  assert.ok(wb.setSheetViewMode(0, 'pageBreakPreview').ok);
  const v = wb.getSheetView(0);
  assert.ok(v.status.ok, `getSheetView: ${JSON.stringify(v.status)}`);
  assert.equal(v.view.showGridLines, 0);
  assert.equal(v.view.showRowColHeaders, 0);
  assert.equal(v.view.showZeros, 0);
  assert.equal(v.view.rightToLeft, 1);
  assert.equal(v.view.tabSelected, 1);
  assert.equal(v.view.viewMode, 'pageBreakPreview');
});

test('insertRows shifts an existing literal forward', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Populate row 0; insert one row at row 0; the literal is now at row 1.
  assert.ok(wb.setNumber(0, 0, 0, 99).ok);
  assert.ok(wb.insertRows(0, 0, 1).ok);
  const r0 = wb.getValue(0, 0, 0);
  assert.ok(r0.status.ok);
  // The freshly-inserted row 0 is blank.
  assert.equal(r0.value.kind, mod.ValueKind.Blank);
  const r1 = wb.getValue(0, 1, 0);
  assert.ok(r1.status.ok);
  assert.equal(r1.value.kind, mod.ValueKind.Number);
  assert.equal(r1.value.number, 99);
});

test('setCellXfIndex + getCellXfIndex round-trip the xf id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Seed the cell so the xf attaches to a stored slot.
  assert.ok(wb.setNumber(0, 0, 0, 1).ok);
  // xfIndex 0 is always the default xf and is guaranteed to be valid.
  assert.ok(wb.setCellXfIndex(0, 0, 0, 0).ok);
  const r = wb.getCellXfIndex(0, 0, 0);
  assert.ok(r.status.ok, `getCellXfIndex: ${JSON.stringify(r.status)}`);
  assert.equal(typeof r.xfIndex, 'number');
  assert.equal(r.xfIndex, 0);
});

test('style building blocks: addFont -> addXf -> setCellXfIndex -> getXf', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // A fresh workbook starts with empty styles tables.
  assert.equal(wb.fontCount(), 0);
  assert.equal(wb.fillCount(), 0);
  assert.equal(wb.borderCount(), 0);
  assert.equal(wb.xfCount(), 0);

  const fontResult = wb.addFont({
    name: 'Arial',
    size: 12,
    bold: true,
    italic: false,
    strike: false,
    underline: 0,
    colorArgb: 0xff112233,
  });
  assert.ok(fontResult.status.ok, `addFont: ${JSON.stringify(fontResult.status)}`);
  // Linear-search dedup: the second call with the same payload returns the same index.
  const fontDup = wb.addFont({
    name: 'Arial',
    size: 12,
    bold: true,
    italic: false,
    strike: false,
    underline: 0,
    colorArgb: 0xff112233,
  });
  assert.equal(fontDup.index, fontResult.index);
  assert.equal(wb.fontCount(), 1);

  const fill = wb.addFill({ pattern: 1, fgArgb: 0xffff0000, bgArgb: 0xff000000 });
  assert.ok(fill.status.ok);
  const border = wb.addBorder({
    left: { style: 1, colorArgb: 0xff000000 },
    right: { style: 1, colorArgb: 0xff000000 },
    top: { style: 1, colorArgb: 0xff000000 },
    bottom: { style: 1, colorArgb: 0xff000000 },
    diagonal: { style: 0, colorArgb: 0 },
    diagonalUp: false,
    diagonalDown: false,
  });
  assert.ok(border.status.ok);

  const builtin = wb.addNumFmt('General');
  assert.ok(builtin.status.ok);
  assert.equal(builtin.numFmtId, 0);

  const custom = wb.addNumFmt('"USD" #,##0');
  assert.ok(custom.status.ok);
  assert.ok(custom.numFmtId >= 164, `expected custom id >= 164, got ${custom.numFmtId}`);

  const xf = wb.addXf({
    fontIndex: fontResult.index,
    fillIndex: fill.index,
    borderIndex: border.index,
    numFmtId: custom.numFmtId,
    horizontalAlign: 1,
    verticalAlign: 2,
    wrapText: true,
  });
  assert.ok(xf.status.ok, `addXf: ${JSON.stringify(xf.status)}`);

  assert.ok(wb.setNumber(0, 0, 0, 1).ok);
  assert.ok(wb.setCellXfIndex(0, 0, 0, xf.index).ok);

  const reread = wb.getCellXf(xf.index);
  assert.ok(reread.status.ok);
  assert.equal(reread.fontIndex, fontResult.index);
  assert.equal(reread.numFmtId, custom.numFmtId);
  assert.equal(reread.wrapText, true);

  const rfont = wb.getFont(fontResult.index);
  assert.ok(rfont.status.ok);
  assert.equal(rfont.name, 'Arial');
  assert.equal(rfont.bold, true);

  const rfill = wb.getFill(fill.index);
  assert.ok(rfill.status.ok);
  assert.equal(rfill.pattern, 1);

  const rborder = wb.getBorder(border.index);
  assert.ok(rborder.status.ok);
  assert.equal(rborder.left.style, 1);
  assert.equal(rborder.diagonalUp, false);

  const rfmt = wb.getNumFmt(custom.numFmtId);
  assert.ok(rfmt.status.ok);
  assert.equal(rfmt.formatCode, '"USD" #,##0');
});

test('addXf rejects out-of-range font_index', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const r = wb.addXf({
    fontIndex: 99,
    fillIndex: 0,
    borderIndex: 0,
    numFmtId: 0,
    horizontalAlign: 0,
    verticalAlign: 0,
    wrapText: false,
  });
  assert.ok(!r.status.ok, 'expected addXf to reject out-of-range font_index');
});

test('setIterativeProgress accepts a function and accepts null to clear', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Registration roundtrip; we don't assert the callback fires because
  // a non-iterative recalc never invokes the trampoline. The smoke
  // signal we want is "the addon does not crash on register / clear".
  const cb = () => true;
  const reg = wb.setIterativeProgress(cb);
  assert.ok(reg.ok, `setIterativeProgress(fn): ${JSON.stringify(reg)}`);
  const clr = wb.setIterativeProgress(null);
  assert.ok(clr.ok, `setIterativeProgress(null): ${JSON.stringify(clr)}`);
});

test('evaluateCfRange returns ok envelope with empty cells for a CF-less workbook', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // The default workbook has no CF rules, so the call should succeed
  // and return an empty cell list. NaN disables `TimePeriod` rules.
  const r = wb.evaluateCfRange(0, 0, 0, 4, 4, Number.NaN);
  assert.ok(r.status.ok, `evaluateCfRange: ${JSON.stringify(r.status)}`);
  assert.ok(Array.isArray(r.cells));
  assert.equal(r.cells.length, 0);
});

test('comments round-trip: setComment + getComment', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Pre-existing cell so the comment attaches to a stored slot.
  assert.ok(wb.setNumber(0, 1, 1, 7).ok);
  const sc = wb.setComment(0, 1, 1, 'libraz', 'hello world');
  assert.ok(sc.ok, `setComment: ${JSON.stringify(sc)}`);
  const c = wb.getComment(0, 1, 1);
  assert.ok(c !== null, 'expected non-null CommentEntry');
  assert.equal(c.author, 'libraz');
  assert.equal(c.text, 'hello world');
  // Removing surfaces null on the next read.
  assert.ok(wb.setComment(0, 1, 1, '', '').ok);
  assert.equal(wb.getComment(0, 1, 1), null);
});

test('cellCount + cellAt enumerate stored cells', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setNumber(0, 0, 0, 1).ok);
  assert.ok(wb.setNumber(0, 0, 1, 2).ok);
  assert.ok(wb.setFormula(0, 1, 0, '=A1+B1').ok);
  assert.ok(wb.recalc().ok);
  const n = wb.cellCount(0);
  assert.equal(typeof n, 'number');
  assert.ok(n >= 3, `expected >= 3 stored cells, got ${n}`);
  // Iterate; one of the entries should carry our formula.
  let foundFormula = false;
  for (let i = 0; i < n; ++i) {
    const e = wb.cellAt(0, i);
    assert.ok(e.status.ok, `cellAt(${i}): ${JSON.stringify(e.status)}`);
    assert.equal(typeof e.row, 'number');
    assert.equal(typeof e.col, 'number');
    if (e.formula === '=A1+B1') {
      foundFormula = true;
      assert.equal(e.value.kind, mod.ValueKind.Number);
      assert.equal(e.value.number, 3);
    }
  }
  assert.ok(foundFormula, 'expected to find the =A1+B1 cell during iteration');
});

test('default export exposes CfMatchKind ordinals', async () => {
  const mod = await getModule();
  assert.equal(typeof mod.CfMatchKind, 'object');
  assert.equal(mod.CfMatchKind.DifferentialFormat, 0);
  assert.equal(mod.CfMatchKind.ColorScale, 1);
  assert.equal(mod.CfMatchKind.DataBar, 2);
  assert.equal(mod.CfMatchKind.IconSet, 3);
});

test('default export exposes ErrorCode ordinals and versionString alias', async () => {
  const mod = await getModule();
  assert.equal(typeof mod.ErrorCode, 'object');
  assert.equal(mod.ErrorCode.Div0, 1);
  assert.equal(mod.ErrorCode.Ref, 3);
  assert.equal(mod.ErrorCode.NA, 6);
  assert.equal(typeof mod.versionString, 'function');
  assert.equal(mod.versionString(), mod.version());
});

test('setError() rejects a missing errorCode argument', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.throws(() => wb.setError(0, 0, 0), TypeError);
});

test('saveEx() writes xlsx and xlsb bytes; rejects a missing format argument', async () => {
  const mod = await getModule();
  assert.equal(typeof mod.WorkbookFormat, 'object');
  const wb = mod.Workbook.createDefault();
  const xlsx = wb.saveEx(mod.WorkbookFormat.Xlsx);
  assert.ok(xlsx.status.ok, JSON.stringify(xlsx.status));
  assert.ok(xlsx.bytes.length > 0);
  const xlsb = wb.saveEx(mod.WorkbookFormat.Xlsb);
  assert.ok(xlsb.status.ok, JSON.stringify(xlsb.status));
  assert.ok(xlsb.bytes.length > 0);
  assert.throws(() => wb.saveEx(), TypeError);
});

test('addValidation / getValidations / removeValidationAt / clearValidations round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Empty by default.
  assert.equal(wb.getValidations(0).length, 0);

  // List-type validation at A1:B3.
  const listRule = {
    ranges: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1 }],
    type: 3,
    op: 0,
    errorStyle: 1,
    allowBlank: true,
    showInputMessage: true,
    showErrorMessage: true,
    formula1: '"Yes,No,Maybe"',
    promptTitle: 'Choose',
    promptMessage: 'Pick one',
    errorTitle: 'Bad value',
    errorMessage: 'Pick from the list',
  };
  assert.ok(wb.addValidation(0, listRule).ok);

  // Decimal between [0, 100] across two rectangles.
  const decimalRule = {
    ranges: [
      { firstRow: 4, firstCol: 0, lastRow: 4, lastCol: 0 },
      { firstRow: 6, firstCol: 0, lastRow: 9, lastCol: 0 },
    ],
    type: 2,
    op: 0,
    formula1: '0',
    formula2: '100',
    allowBlank: false,
  };
  assert.ok(wb.addValidation(0, decimalRule).ok);

  const list = wb.getValidations(0);
  assert.equal(list.length, 2);
  // Booleans must arrive as JS booleans, not 0/1.
  assert.equal(typeof list[0].allowBlank, 'boolean');
  assert.equal(list[0].allowBlank, true);
  assert.equal(typeof list[0].showInputMessage, 'boolean');
  assert.equal(list[0].showInputMessage, true);
  assert.equal(typeof list[1].allowBlank, 'boolean');
  assert.equal(list[1].allowBlank, false);
  assert.equal(list[0].formula1, '"Yes,No,Maybe"');
  assert.equal(list[1].formula1, '0');
  assert.equal(list[1].formula2, '100');
  assert.equal(list[1].ranges.length, 2);
  assert.equal(list[1].ranges[1].firstRow, 6);

  // removeValidationAt removes the first rule.
  assert.ok(wb.removeValidationAt(0, 0).ok);
  const after = wb.getValidations(0);
  assert.equal(after.length, 1);
  assert.equal(after[0].type, 2);

  // Out-of-range index is rejected.
  assert.ok(!wb.removeValidationAt(0, 99).ok);
  // Sheet-out-of-range is rejected on every entry.
  assert.ok(!wb.addValidation(999, listRule).ok);
  assert.ok(!wb.removeValidationAt(999, 0).ok);
  assert.ok(!wb.clearValidations(999).ok);

  // clearValidations drops everything; safe to call again.
  assert.ok(wb.clearValidations(0).ok);
  assert.equal(wb.getValidations(0).length, 0);
  assert.ok(wb.clearValidations(0).ok);
});

// -- Newly wired surface (audit parity with the WASM binding) ----------
// These tests exercise the methods that were previously only reachable
// from the WASM package. They assert sensible return shapes rather than
// deep semantics; the point is "the addon wires them and does not throw".

test('addConditionalFormat + getConditionalFormats + clearConditionalFormats round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.getConditionalFormats(0).length, 0);
  // cellIs rule (type 1) comparing the cell to a literal.
  const add = wb.addConditionalFormat(0, {
    sqref: [{ firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 0 }],
    type: 1,
    op: 5,
    formula1: '10',
    dxfId: 0,
    stopIfTrue: false,
  });
  assert.ok(add.status.ok, `addConditionalFormat: ${JSON.stringify(add)}`);
  assert.equal(typeof add.index, 'number');
  const list = wb.getConditionalFormats(0);
  assert.equal(list.length, 1);
  assert.equal(list[0].type, 1);
  assert.equal(typeof list[0].id, 'string');
  assert.equal(list[0].sqref.length, 1);
  assert.equal(list[0].sqref[0].lastRow, 4);
  assert.equal(list[0].formula1, '10');
  // removeConditionalFormatAt drops the only rule.
  assert.ok(wb.removeConditionalFormatAt(0, 0).ok);
  assert.equal(wb.getConditionalFormats(0).length, 0);
  // clearConditionalFormats is a safe no-op on an empty sheet.
  assert.ok(wb.clearConditionalFormats(0).ok);
  // Sheet-out-of-range is rejected.
  assert.ok(!wb.clearConditionalFormats(999).ok);
});

test('named cell styles: cellStyleCount / getCellStyle / cellStyleXfCount', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // A fresh workbook may carry zero named styles; the accessors must
  // still return well-formed values rather than throwing.
  const count = wb.cellStyleCount();
  assert.equal(typeof count, 'number');
  assert.ok(count >= 0);
  assert.equal(typeof wb.cellStyleXfCount(), 'number');
  if (count > 0) {
    const cs = wb.getCellStyle(0);
    assert.ok(cs.status.ok, `getCellStyle: ${JSON.stringify(cs.status)}`);
    assert.equal(typeof cs.name, 'string');
    assert.equal(typeof cs.xfId, 'number');
  } else {
    // Out-of-range index surfaces an error status, not a throw.
    const cs = wb.getCellStyle(0);
    assert.equal(cs.status.ok, false);
  }
});

test('setCalcMode / calcMode round-trip the calc policy', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Default is automatic (0).
  assert.equal(wb.calcMode(), 0);
  assert.ok(wb.setCalcMode(1).ok);
  assert.equal(wb.calcMode(), 1);
  assert.ok(wb.setCalcMode(2).ok);
  assert.equal(wb.calcMode(), 2);
});

test('excelProfileId / setExcelProfileId round-trip the profile id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const def = wb.excelProfileId();
  assert.equal(typeof def, 'string');
  assert.ok(def.length > 0);
  assert.ok(wb.setExcelProfileId('mac-365-ja_JP').ok);
  assert.equal(wb.excelProfileId(), 'mac-365-ja_JP');
});

test('functionNames + functionMetadata expose the catalog', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const names = wb.functionNames();
  assert.ok(Array.isArray(names));
  assert.ok(names.length > 0, 'expected a non-empty function catalog');
  assert.ok(names.includes('SUM'), 'expected SUM in the catalog');
  const md = wb.functionMetadata('SUM', 0);
  assert.equal(md.ok, true);
  assert.equal(md.name, 'SUM');
  assert.equal(typeof md.minArity, 'number');
  // Unknown function returns { ok: false }.
  const miss = wb.functionMetadata('NOT_A_REAL_FUNCTION_XYZ', 0);
  assert.equal(miss.ok, false);
});

test('mergeFunctionMetadata overlays provider metadata; is identity without a provider', async () => {
  const mod = await getModule();
  assert.equal(typeof mod.mergeFunctionMetadata, 'function');
  const base = mod.Workbook.createDefault().functionMetadata('XLOOKUP', 0);
  assert.equal(base.ok, true);
  // The engine leaves display metadata empty.
  assert.equal(base.signatureTemplate, undefined);
  assert.equal(base.description, undefined);

  const entry = {
    signature: 'XLOOKUP(lookup_value, lookup_array, return_array)',
    description: 'Searches a range or an array.',
    aliases: { 'fr-FR': 'RECHERCHEX' },
    localized: {
      'fr-FR': { signature: 'RECHERCHEX(...)', description: 'Recherche.' },
    },
  };

  // Localized override wins for the matching locale.
  const fr = mod.mergeFunctionMetadata(base, entry, 'fr-FR');
  assert.equal(fr.signatureTemplate, 'RECHERCHEX(...)');
  assert.equal(fr.description, 'Recherche.');
  assert.equal(fr.localizedName, 'RECHERCHEX');
  // Structural fields survive the merge.
  assert.equal(fr.ok, true);
  assert.equal(fr.name, 'XLOOKUP');

  // A locale with no localized/alias entry falls back to the default
  // signature/description and the canonical display name.
  const de = mod.mergeFunctionMetadata(base, entry, 'de-DE');
  assert.equal(de.signatureTemplate, 'XLOOKUP(lookup_value, lookup_array, return_array)');
  assert.equal(de.description, 'Searches a range or an array.');
  assert.equal(de.localizedName, 'XLOOKUP');

  // No provider entry -> base returned verbatim, display metadata stays NULL.
  const none = mod.mergeFunctionMetadata(base, undefined, 'fr-FR');
  assert.equal(none, base);
  assert.equal(none.signatureTemplate, undefined);
  assert.equal(none.description, undefined);
});

test('localizeFunctionName / canonicalizeFunctionName round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // en-US locale (0): the localized name is the canonical name itself.
  assert.equal(wb.localizeFunctionName('SUM', 0), 'SUM');
  assert.equal(wb.canonicalizeFunctionName('SUM', 0), 'SUM');
  // An unknown name returns the empty string.
  assert.equal(wb.canonicalizeFunctionName('NOPE_XYZ', 0), '');
});

test('precedents / dependents return arrays for a small formula graph', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setNumber(0, 0, 0, 1).ok); // A1
  assert.ok(wb.setNumber(0, 0, 1, 2).ok); // B1
  assert.ok(wb.setFormula(0, 1, 0, '=A1+B1').ok); // A2 = A1 + B1
  assert.ok(wb.recalc().ok);
  // A2 reads A1 and B1.
  const prec = wb.precedents(0, 1, 0, 1);
  assert.ok(Array.isArray(prec));
  assert.equal(prec.length, 2);
  for (const node of prec) {
    assert.equal(typeof node.sheet, 'number');
    assert.equal(typeof node.row, 'number');
    assert.equal(typeof node.col, 'number');
  }
  // A1 is read by A2.
  const dep = wb.dependents(0, 0, 0, 1);
  assert.ok(Array.isArray(dep));
  assert.equal(dep.length, 1);
  assert.equal(dep[0].row, 1);
  assert.equal(dep[0].col, 0);
});

test('spillInfo reports a dynamic-array region', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Non-spill cell: engaged is false, fields zeroed.
  const none = wb.spillInfo(0, 0, 0);
  assert.equal(typeof none.engaged, 'boolean');
  assert.equal(none.engaged, false);
  // A spilling dynamic array anchored at A1.
  assert.ok(wb.setFormula(0, 0, 0, '=SEQUENCE(3,1)').ok);
  assert.ok(wb.recalc().ok);
  const info = wb.spillInfo(0, 0, 0);
  assert.equal(info.engaged, true);
  assert.equal(info.anchorRow, 0);
  assert.equal(info.anchorCol, 0);
  assert.equal(info.rows, 3);
  assert.equal(info.cols, 1);
});

test('getLambdaText is wired and reports the documented contract', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // A bare top-level LAMBDA renders as #CALC! in Excel 365 and does not
  // retain a renderable closure; the engine mirrors that, so the lambda
  // text surface returns an error status (not a throw). A plain number
  // cell is likewise not a lambda. Both prove the method marshals
  // arguments and propagates the Expected failure as a JS error status.
  assert.ok(wb.setFormula(0, 0, 0, '=LAMBDA(x,x+1)').ok);
  assert.ok(wb.recalc().ok);
  const bare = wb.getLambdaText(0, 0, 0);
  assert.equal(bare.status.ok, false);
  assert.equal(typeof bare.text, 'string');
  assert.equal(bare.text, '');

  assert.ok(wb.setNumber(0, 1, 0, 5).ok);
  const num = wb.getLambdaText(0, 1, 0);
  assert.equal(num.status.ok, false);
});

test('getExternalLinks returns an array (empty for a fresh workbook)', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const links = wb.getExternalLinks();
  assert.ok(Array.isArray(links));
  assert.equal(links.length, 0);
});

test('getSheetProtection / setSheetProtection round-trip the protection flags', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const before = wb.getSheetProtection(0);
  assert.ok(before.status.ok, `getSheetProtection: ${JSON.stringify(before.status)}`);
  assert.equal(before.protection.enabled, 0);
  // Enable protection and lock cell selection.
  const next = {
    ...before.protection,
    enabled: 1,
    selectLockedCells: 1,
    sort: 1,
  };
  assert.ok(wb.setSheetProtection(0, next).ok);
  const after = wb.getSheetProtection(0);
  assert.ok(after.status.ok);
  assert.equal(after.protection.enabled, 1);
  assert.equal(after.protection.selectLockedCells, 1);
  assert.equal(after.protection.sort, 1);
  // Sheet-out-of-range is rejected.
  assert.ok(!wb.getSheetProtection(999).status.ok);
  assert.ok(!wb.setSheetProtection(999, next).ok);
});
