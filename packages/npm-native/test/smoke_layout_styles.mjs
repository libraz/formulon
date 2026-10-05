import test from 'node:test';
import { assert, getModule } from './smoke_support.mjs';

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
  assert.ok(wb.addHyperlink(0, 1, 2, 'https://example.com/', '', '', '').ok);
  assert.ok(wb.addHyperlink(0, 3, 4, 'mailto:hello@example.com', 'Hello', '', '').ok);
  assert.ok(wb.addHyperlink(0, 5, 6, '', 'See X', 'Internal link', 'Sheet1!A1').ok);
  assert.ok(wb.addHyperlinkRange(0, 7, 8, 9, 10, 'https://example.com/range', 'Range', 'Range tip', '').ok);
  const list = wb.getHyperlinks(0);
  assert.equal(list.length, 4);
  assert.equal(list[0].row, 1);
  assert.equal(list[0].col, 2);
  assert.equal(list[0].lastRow, 1);
  assert.equal(list[0].lastCol, 2);
  assert.equal(list[0].target, 'https://example.com/');
  assert.equal(list[0].display, '');
  assert.equal(list[0].tooltip, '');
  assert.equal(list[1].target, 'mailto:hello@example.com');
  assert.equal(list[1].display, 'Hello');
  assert.equal(list[2].display, 'See X');
  assert.equal(list[2].tooltip, 'Internal link');
  assert.equal(list[2].location, 'Sheet1!A1');
  assert.deepEqual(
    {
      row: list[3].row,
      col: list[3].col,
      lastRow: list[3].lastRow,
      lastCol: list[3].lastCol,
      target: list[3].target,
      display: list[3].display,
      tooltip: list[3].tooltip,
    },
    {
      row: 7,
      col: 8,
      lastRow: 9,
      lastCol: 10,
      target: 'https://example.com/range',
      display: 'Range',
      tooltip: 'Range tip',
    },
  );
  // Sheet-out-of-range is rejected.
  assert.ok(!wb.addHyperlink(999, 0, 0, 'https://x/', '', '', '').ok);
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

test('sheet layout getters preserve explicit width presence flags', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setColumnWidth(0, 0, 0, 0).ok);
  assert.ok(wb.setColumnHidden(0, 1, 1, true).ok);
  assert.ok(wb.setRowHeight(0, 3, 30).ok);

  const columns = wb.getSheetColumns(0);
  assert.ok(columns.status.ok, `getSheetColumns: ${JSON.stringify(columns.status)}`);
  const widthZero = columns.columns.find((column) => column.first === 0 && column.last === 0);
  assert.ok(widthZero);
  assert.equal(widthZero.width, 0);
  assert.equal(widthZero.hasWidth, 1);
  assert.equal(widthZero.hasStyle, 0);
  const hidden = columns.columns.find((column) => column.first === 1 && column.last === 1);
  assert.ok(hidden);
  assert.equal(hidden.hidden, 1);
  assert.equal(hidden.hasWidth, 0);

  const rows = wb.getSheetRowOverrides(0);
  assert.ok(rows.status.ok, `getSheetRowOverrides: ${JSON.stringify(rows.status)}`);
  const row = rows.rows.find((entry) => entry.row === 3);
  assert.ok(row);
  assert.equal(row.height, 30);
  assert.equal(row.hasStyle, 0);
  assert.equal(row.styleXf, 0);
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
  // A fresh workbook seeds Excel's minimum style table: one default font,
  // one empty border, one default xf, and the two fills Excel reserves at
  // the front of every file (`none`, `gray125`). A caller's first record
  // therefore lands after them instead of displacing `none`.
  assert.equal(wb.fontCount().value, 1);
  assert.equal(wb.fillCount().value, 2);
  assert.equal(wb.borderCount().value, 1);
  assert.equal(wb.xfCount().value, 1);

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
  assert.equal(wb.fontCount().value, 2);

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
  assert.equal(reread.hasAlignment, true);
  assert.equal(reread.hasHorizontalAlign, true);
  assert.equal(reread.hasVerticalAlign, true);
  assert.equal(reread.hasWrapText, true);
  assert.equal(reread.hasJustifyLastLine, false);

  const optional = wb.addXf({
    fontIndex: fontResult.index,
    fillIndex: fill.index,
    borderIndex: border.index,
    numFmtId: custom.numFmtId,
    textRotation: 255,
    indent: 0,
    relativeIndent: -3,
    shrinkToFit: false,
    readingOrder: 0,
    justifyLastLine: true,
  });
  assert.ok(optional.status.ok, `optional alignment: ${JSON.stringify(optional.status)}`);
  const optionalRead = wb.getCellXf(optional.index);
  assert.ok(optionalRead.status.ok);
  assert.equal(optionalRead.textRotation, 255);
  assert.equal(optionalRead.indent, 0);
  assert.equal(optionalRead.relativeIndent, -3);
  assert.equal(optionalRead.shrinkToFit, false);
  assert.equal(optionalRead.readingOrder, 0);
  assert.equal(optionalRead.justifyLastLine, true);
  assert.equal(optionalRead.hasAlignment, true);
  assert.equal(optionalRead.hasJustifyLastLine, true);
  assert.equal(wb.addXf(optionalRead).index, optional.index);

  const omitted = wb.addXf({
    fontIndex: fontResult.index,
    fillIndex: fill.index,
    borderIndex: border.index,
    numFmtId: custom.numFmtId,
  });
  assert.ok(omitted.status.ok);
  const omittedRead = wb.getCellXf(omitted.index);
  assert.ok(omittedRead.status.ok);
  assert.equal(omittedRead.hasAlignment, false);

  const explicitEmpty = wb.addXf({
    fontIndex: fontResult.index,
    fillIndex: fill.index,
    borderIndex: border.index,
    numFmtId: custom.numFmtId,
    hasAlignment: true,
  });
  assert.ok(explicitEmpty.status.ok);
  assert.notEqual(explicitEmpty.index, omitted.index);
  const explicitEmptyRead = wb.getCellXf(explicitEmpty.index);
  assert.ok(explicitEmptyRead.status.ok);
  assert.equal(explicitEmptyRead.hasAlignment, true);

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

test('selector colours survive get/add identity and save/load', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const theme = { kind: 2, rgb: 0, theme: 3, tint: 0.5, indexed: 0 };
  const indexed = { kind: 3, rgb: 0, theme: 0, tint: 0, indexed: 9 };
  const automatic = { kind: 4, rgb: 0, theme: 0, tint: 0, indexed: 0 };

  const font = wb.addFont({ name: 'SelectorFont', size: 11, colorArgb: 0x01020304, color: theme });
  assert.ok(font.status.ok);
  const gotFont = wb.getFont(font.index);
  assert.ok(gotFont.status.ok);
  assert.equal(gotFont.color.kind, 2);
  assert.equal(gotFont.color.theme, 3);
  assert.equal(gotFont.color.tint, 0.5);
  assert.equal(gotFont.colorArgb, 0x01020304);
  const fontAgain = wb.addFont(gotFont);
  assert.ok(fontAgain.status.ok);
  assert.equal(fontAgain.index, font.index);

  const fill = wb.addFill({
    pattern: 1,
    fgArgb: 0x05060708,
    bgArgb: 0x090a0b0c,
    fg: indexed,
    bg: automatic,
  });
  assert.ok(fill.status.ok);
  const gotFill = wb.getFill(fill.index);
  assert.ok(gotFill.status.ok);
  assert.equal(gotFill.fg.kind, 3);
  assert.equal(gotFill.fg.indexed, 9);
  assert.equal(gotFill.bg.kind, 4);
  const fillAgain = wb.addFill(gotFill);
  assert.ok(fillAgain.status.ok);
  assert.equal(fillAgain.index, fill.index);

  const border = wb.addBorder({
    left: { style: 1, colorArgb: 0x01020304, color: theme },
    right: { style: 1, colorArgb: 0x05060708, color: indexed },
    top: { style: 1, colorArgb: 0x090a0b0c, color: automatic },
    bottom: { style: 0, colorArgb: 0 },
    diagonal: { style: 0, colorArgb: 0 },
  });
  assert.ok(border.status.ok);
  const gotBorder = wb.getBorder(border.index);
  assert.ok(gotBorder.status.ok);
  assert.equal(gotBorder.left.color.kind, 2);
  assert.equal(gotBorder.left.color.theme, 3);
  assert.equal(gotBorder.right.color.kind, 3);
  assert.equal(gotBorder.right.color.indexed, 9);
  assert.equal(gotBorder.top.color.kind, 4);
  const borderAgain = wb.addBorder(gotBorder);
  assert.ok(borderAgain.status.ok);
  assert.equal(borderAgain.index, border.index);

  const dxf = wb.addDxf({
    font: { name: 'DxfSelector', size: 9, colorArgb: 0x11121314, color: automatic },
    fill: { pattern: 1, fgArgb: 0x15161718, fg: indexed },
    border: { left: { style: 1, colorArgb: 0x191a1b1c, color: theme } },
  });
  assert.ok(dxf.status.ok);
  const gotDxf = wb.getDxf(dxf.index);
  assert.ok(gotDxf.status.ok);
  assert.equal(gotDxf.font.color.kind, 4);
  assert.equal(gotDxf.fill.fg.kind, 3);
  assert.equal(gotDxf.border.left.color.kind, 2);
  const dxfAgain = wb.addDxf(gotDxf);
  assert.ok(dxfAgain.status.ok);
  assert.equal(dxfAgain.index, dxf.index);

  const saved = wb.save();
  assert.ok(saved.status.ok);
  const loaded = mod.Workbook.loadBytes(saved.bytes);
  const loadedFont = loaded.getFont(font.index);
  const loadedFill = loaded.getFill(fill.index);
  const loadedBorder = loaded.getBorder(border.index);
  const loadedDxf = loaded.getDxf(dxf.index);
  assert.ok(loadedFont.status.ok);
  assert.ok(loadedFill.status.ok);
  assert.ok(loadedBorder.status.ok);
  assert.ok(loadedDxf.status.ok);
  assert.equal(loadedFont.color.kind, 2);
  assert.equal(loadedFill.fg.kind, 3);
  assert.equal(loadedBorder.left.color.kind, 2);
  assert.equal(loadedDxf.font.color.kind, 4);
  loaded.dispose();
  wb.dispose();
});

test('dxf alignment and protection XML survive get/add identity and save/load', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const alignmentXml = '<alignment horizontal="center" wrapText="1"/>';
  const protectionXml = '<protection locked="0" hidden="1"/>';
  const alignment = wb.addDxf({ alignmentXml });
  assert.ok(alignment.status.ok);
  const protection = wb.addDxf({ protectionXml });
  assert.ok(protection.status.ok);
  assert.notEqual(alignment.index, protection.index);

  const gotAlignment = wb.getDxf(alignment.index);
  assert.ok(gotAlignment.status.ok);
  assert.equal(gotAlignment.alignmentXml, alignmentXml);
  assert.equal(gotAlignment.protectionXml, undefined);
  const gotProtection = wb.getDxf(protection.index);
  assert.ok(gotProtection.status.ok);
  assert.equal(gotProtection.alignmentXml, undefined);
  assert.equal(gotProtection.protectionXml, protectionXml);

  const alignmentAgain = wb.addDxf(gotAlignment);
  assert.ok(alignmentAgain.status.ok);
  assert.equal(alignmentAgain.index, alignment.index);
  const protectionAgain = wb.addDxf(gotProtection);
  assert.ok(protectionAgain.status.ok);
  assert.equal(protectionAgain.index, protection.index);

  const saved = wb.save();
  assert.ok(saved.status.ok);
  const loaded = mod.Workbook.loadBytes(saved.bytes);
  assert.ok(loaded.isValid());
  const loadedAlignment = loaded.getDxf(alignment.index);
  assert.ok(loadedAlignment.status.ok);
  assert.equal(loadedAlignment.alignmentXml, alignmentXml);
  assert.equal(loadedAlignment.protectionXml, undefined);
  const loadedProtection = loaded.getDxf(protection.index);
  assert.ok(loadedProtection.status.ok);
  assert.equal(loadedProtection.alignmentXml, undefined);
  assert.equal(loadedProtection.protectionXml, protectionXml);
  const loadedAlignmentAgain = loaded.addDxf(loadedAlignment);
  assert.ok(loadedAlignmentAgain.status.ok);
  assert.equal(loadedAlignmentAgain.index, alignment.index);
  const loadedProtectionAgain = loaded.addDxf(loadedProtection);
  assert.ok(loadedProtectionAgain.status.ok);
  assert.equal(loadedProtectionAgain.index, protection.index);
  loaded.dispose();
  wb.dispose();
});

test('addFont / getFont preserve superscript and round-trip to the same index', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const superscript = wb.addFont({ name: 'Arial', size: 12, vertAlign: 1, colorArgb: 0xff112233 });
  assert.ok(superscript.status.ok, `addFont: ${JSON.stringify(superscript.status)}`);
  assert.equal(wb.getFont(superscript.index).vertAlign, 1);

  const before = wb.fontCount().value;
  const again = wb.addFont(wb.getFont(superscript.index));
  assert.ok(again.status.ok);
  assert.equal(again.index, superscript.index);
  assert.equal(wb.fontCount().value, before);

  const edited = wb.getFont(superscript.index);
  edited.colorArgb = 0xff00ff00;
  const recolored = wb.addFont(edited);
  assert.ok(recolored.status.ok);
  assert.notEqual(recolored.index, superscript.index);
  const reread = wb.getFont(recolored.index);
  assert.equal(reread.vertAlign, 1);
  assert.equal(reread.colorArgb, 0xff00ff00);
});

test('addFont/setFont with a non-numeric field throws and does not commit a coerced value', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();

  const before = wb.fontCount().value;
  assert.throws(() => wb.addFont({ name: 'Arial', size: 'not a number' }), TypeError);
  assert.equal(wb.fontCount().value, before, 'no font must have been added for the rejected call');

  const font = wb.addFont({ name: 'Arial', size: 12 });
  assert.ok(font.status.ok);
  assert.throws(() => wb.setFont(font.index, { name: 'Arial', size: 'not a number' }), TypeError);
  // The font slot must still hold its original size, not a coerced 0.
  assert.equal(wb.getFont(font.index).size, 12);
});

test('phonetic runs keep their spans through a save/load round trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setText(0, 0, 0, '東京都');
  assert.deepEqual(wb.getCellPhoneticRuns(0, 0, 0).runs, []);

  const runs = [
    { sb: 0, eb: 2, text: 'トウキョウ' },
    { sb: 2, eb: 3, text: 'ト' },
  ];
  const stored = wb.setCellPhoneticRuns(0, 0, 0, runs);
  assert.ok(stored.ok, `setCellPhoneticRuns: ${JSON.stringify(stored)}`);
  assert.deepEqual(wb.getCellPhoneticRuns(0, 0, 0).runs, runs);
  // The flattening getter still reports the concatenation.
  assert.equal(wb.getCellPhonetic(0, 0, 0).value, 'トウキョウト');

  const saved = wb.save();
  assert.ok(saved.status.ok, `save: ${JSON.stringify(saved.status)}`);
  const loaded = mod.Workbook.loadBytes(saved.bytes);
  assert.deepEqual(loaded.getCellPhoneticRuns(0, 0, 0).runs, runs);

  // Writing the flattened reading back is the collapse the run API avoids.
  wb.setCellPhonetic(0, 0, 0, 'トウキョウト');
  assert.deepEqual(wb.getCellPhoneticRuns(0, 0, 0).runs, [{ sb: 0, eb: 3, text: 'トウキョウト' }]);

  const rejected = wb.setCellPhoneticRuns(0, 0, 0, [
    { sb: 2, eb: 3, text: 'ト' },
    { sb: 0, eb: 2, text: 'トウ' },
  ]);
  assert.equal(rejected.ok, false);

  loaded.dispose();
  wb.dispose();
});

test('phonetic properties are independent of the runs and survive a round trip', async () => {
  const mod = await getModule();
  // The result carries a `status` alongside the triple; compare the triple
  // alone so a status-shape change does not read as a value change.
  const props = (book) => {
    const { fontId, type, alignment } = book.getCellPhoneticProperties(0, 0, 0);
    return { fontId, type, alignment };
  };
  const wb = mod.Workbook.createDefault();
  wb.setText(0, 0, 0, '大阪');
  assert.deepEqual(props(wb), { fontId: 0, type: 0, alignment: 0 });

  assert.ok(wb.setCellPhoneticProperties(0, 0, 0, { fontId: 3, type: 2, alignment: 2 }).ok);
  // Setting the readings must not reset the rendering.
  assert.ok(wb.setCellPhoneticRuns(0, 0, 0, [{ sb: 0, eb: 2, text: 'おおさか' }]).ok);
  assert.deepEqual(props(wb), { fontId: 3, type: 2, alignment: 2 });

  const saved = wb.save();
  assert.ok(saved.status.ok, `save: ${JSON.stringify(saved.status)}`);
  const loaded = mod.Workbook.loadBytes(saved.bytes);
  assert.deepEqual(props(loaded), { fontId: 3, type: 2, alignment: 2 });

  // Two bits each on the binary side, so a wider ordinal is refused.
  assert.equal(wb.setCellPhoneticProperties(0, 0, 0, { fontId: 0, type: 4, alignment: 0 }).ok, false);
  assert.deepEqual(props(wb), { fontId: 3, type: 2, alignment: 2 });

  loaded.dispose();
  wb.dispose();
});

test('setDefaultFont declares what an unstyled cell is saved as', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.getFont(0).name, 'Calibri');

  // addFont can only ever append beside the seeded default.
  const appended = wb.addFont({ name: '游ゴシック', size: 11 });
  assert.ok(appended.status.ok);
  assert.ok(appended.index > 0);
  assert.equal(wb.getFont(0).name, 'Calibri');

  const declared = wb.setDefaultFont({ name: '游ゴシック', size: 11, hasCharset: true, charset: 128 });
  assert.ok(declared.ok, `setDefaultFont: ${JSON.stringify(declared)}`);
  assert.equal(wb.getFont(0).name, '游ゴシック');
  assert.equal(wb.getFont(0).charset, 128);

  wb.dispose();
});

test('setFont overwrites an existing slot and refuses an absent index', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const added = wb.addFont({ name: 'Meiryo', size: 12 });
  assert.ok(added.status.ok);

  const replaced = wb.setFont(added.index, { name: 'MS Gothic', size: 9 });
  assert.ok(replaced.ok, `setFont: ${JSON.stringify(replaced)}`);
  assert.equal(wb.getFont(added.index).name, 'MS Gothic');

  const before = wb.fontCount().value;
  assert.equal(wb.setFont(before, { name: 'MS Gothic', size: 9 }).ok, false);
  assert.equal(wb.fontCount().value, before);

  wb.dispose();
});

test('addDxf / getDxf round-trip a superscript differential font', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const added = wb.addDxf({ font: { name: 'Calibri', size: 9, vertAlign: 1 } });
  assert.ok(added.status.ok, `addDxf: ${JSON.stringify(added.status)}`);
  const readBack = wb.getDxf(added.index);
  assert.ok(readBack.status.ok);
  assert.equal(readBack.font.vertAlign, 1);

  const before = wb.dxfCount().value;
  const again = wb.addDxf(readBack);
  assert.ok(again.status.ok);
  assert.equal(again.index, added.index);
  assert.equal(wb.dxfCount().value, before);
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
