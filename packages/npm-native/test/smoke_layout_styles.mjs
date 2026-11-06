import test from 'node:test';
import { assert, getModule, isFormulonError, readAll } from './smoke_support.mjs';

test('addMerge + getMerges round-trip a single range', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const range = { firstRow: 1, firstCol: 1, lastRow: 2, lastCol: 3 };
  wb.addMerge(0, range);
  const list = wb.listMerges(0);
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
  assert.equal(wb.listHyperlinks(0).length, 0);
  // Three entries with progressively more optional fields populated.
  wb.addHyperlink(0, 1, 2, 'https://example.com/', '', '', '');
  wb.addHyperlink(0, 3, 4, 'mailto:hello@example.com', 'Hello', '', '');
  wb.addHyperlink(0, 5, 6, '', 'See X', 'Internal link', 'Sheet1!A1');
  wb.addHyperlinkRange(
    0,
    { firstRow: 7, firstCol: 8, lastRow: 9, lastCol: 10 },
    'https://example.com/range',
    'Range',
    'Range tip',
    '',
  );
  const list = wb.listHyperlinks(0);
  assert.equal(list.length, 4);
  assert.deepEqual(list[0].rect, { firstRow: 1, firstCol: 2, lastRow: 1, lastCol: 2 });
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
      rect: list[3].rect,
      target: list[3].target,
      display: list[3].display,
      tooltip: list[3].tooltip,
    },
    {
      rect: { firstRow: 7, firstCol: 8, lastRow: 9, lastCol: 10 },
      target: 'https://example.com/range',
      display: 'Range',
      tooltip: 'Range tip',
    },
  );
  // Sheet-out-of-range is rejected.
  assert.throws(() => wb.addHyperlink(999, 0, 0, 'https://x/', '', '', ''), isFormulonError);
  // clearHyperlinks drops everything.
  wb.clearHyperlinks(0);
  assert.equal(wb.listHyperlinks(0).length, 0);
});

test('removeHyperlink / removeHyperlinkAt / clearHyperlinks surface on an empty sheet', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Removing by cell reports NotFound when no hyperlink sits there; clearing is a no-op.
  assert.throws(
    () => wb.removeHyperlink(0, 0, 0),
    (e) => isFormulonError(e, mod.StatusCode.NotFound),
  );
  wb.clearHyperlinks(0);
  // Out-of-range index is rejected.
  assert.throws(
    () => wb.removeHyperlinkAt(0, 0),
    (e) => isFormulonError(e, mod.StatusCode.IndexOutOfRange),
  );
  // Sheet-index-out-of-range is rejected on every variant.
  const sheetRange = (e) => isFormulonError(e, mod.StatusCode.SheetIndexOutOfRange);
  assert.throws(() => wb.removeHyperlink(999, 0, 0), sheetRange);
  assert.throws(() => wb.removeHyperlinkAt(999, 0), sheetRange);
  assert.throws(() => wb.clearHyperlinks(999), sheetRange);
  // The hyperlink list is still empty afterwards.
  assert.equal(wb.listHyperlinks(0).length, 0);
});

test('removeMerge + removeMergeAt + clearMerges step-wise prune the merge list', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const a = { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 };
  const b = { firstRow: 4, firstCol: 4, lastRow: 5, lastCol: 5 };
  wb.addMerge(0, a);
  wb.addMerge(0, b);
  // removeMerge with an overlap that hits `a` only.
  const overlap = { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 };
  wb.removeMerge(0, overlap);
  let list = wb.listMerges(0);
  assert.equal(list.length, 1);
  assert.equal(list[0].firstRow, 4);
  // removeMergeAt drops the survivor by index.
  wb.removeMergeAt(0, 0);
  list = wb.listMerges(0);
  assert.equal(list.length, 0);
  // clearMerges is a no-op on an empty list.
  wb.addMerge(0, a);
  wb.addMerge(0, b);
  wb.clearMerges(0);
  list = wb.listMerges(0);
  assert.equal(list.length, 0);
});

test('setSheetZoom + getSheetView surface the new zoom', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setZoom(0, 175);
  const v = wb.getView(0);
  assert.equal(v.zoomScale, 175);
  // Default freeze / tab-hidden state survives a zoom change.
  assert.equal(v.freezeRows, 0);
  assert.equal(v.freezeCols, 0);
  assert.equal(v.tabHidden, 0);
});

test('sheet layout getters preserve explicit width presence flags', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setColumnWidth(0, 0, 0, 0);
  wb.setColumnHidden(0, 1, 1, true);
  wb.setRowHeight(0, 3, 3, 30);

  const columns = wb.listColumnOverrides(0);
  const widthZero = columns.find((column) => column.first === 0 && column.last === 0);
  assert.ok(widthZero);
  assert.equal(widthZero.width, 0);
  assert.equal(widthZero.hasWidth, 1);
  assert.equal(widthZero.hasStyle, 0);
  const hidden = columns.find((column) => column.first === 1 && column.last === 1);
  assert.ok(hidden);
  assert.equal(hidden.hidden, 1);
  assert.equal(hidden.hasWidth, 0);

  const rows = wb.listRowOverrides(0);
  const row = rows.find((entry) => entry.row === 3);
  assert.ok(row);
  assert.equal(row.height, 30);
  assert.equal(row.hasStyle, 0);
  assert.equal(row.styleXf, 0);
});

test('getSheetView surfaces display / orientation defaults', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const v = wb.getView(0);
  assert.equal(v.showGridLines, 1);
  assert.equal(v.showRowColHeaders, 1);
  assert.equal(v.showZeros, 1);
  assert.equal(v.rightToLeft, 0);
  assert.equal(v.tabSelected, 0);
  assert.equal(v.viewMode, '');
});

test('setSheetShowGridLines / setSheetRightToLeft / setSheetTabSelected / setSheetViewMode round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setShowGridLines(0, false);
  wb.setShowRowColHeaders(0, false);
  wb.setShowZeros(0, false);
  wb.setRightToLeft(0, true);
  wb.setTabSelected(0, true);
  wb.setViewMode(0, 'pageBreakPreview');
  const v = wb.getView(0);
  assert.equal(v.showGridLines, 0);
  assert.equal(v.showRowColHeaders, 0);
  assert.equal(v.showZeros, 0);
  assert.equal(v.rightToLeft, 1);
  assert.equal(v.tabSelected, 1);
  assert.equal(v.viewMode, 'pageBreakPreview');
});

test('insertRows shifts an existing literal forward', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Populate row 0; insert one row at row 0; the literal is now at row 1.
  wb.setNumber(0, 0, 0, 99);
  wb.insertRows(0, 0, 1);
  const r0 = wb.getValue(0, 0, 0);
  // The freshly-inserted row 0 is blank.
  assert.equal(r0.kind, mod.ValueKind.Blank);
  const r1 = wb.getValue(0, 1, 0);
  assert.equal(r1.kind, mod.ValueKind.Number);
  assert.equal(r1.number, 99);
});

test('setCellXfIndex + getCellXfIndex round-trip the xf id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Seed the cell so the xf attaches to a stored slot.
  wb.setNumber(0, 0, 0, 1);
  // xfIndex 0 is always the default xf and is guaranteed to be valid.
  wb.setCellXfIndex(0, 0, 0, 0);
  const r = wb.getCellXfIndex(0, 0, 0);
  assert.equal(typeof r, 'number');
  assert.equal(r, 0);
});

test('style building blocks: addFont -> addXf -> setCellXfIndex -> getXf', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // A fresh workbook seeds Excel's minimum style table: one default font,
  // one empty border, one default xf, and the two fills Excel reserves at
  // the front of every file (`none`, `gray125`). A caller's first record
  // therefore lands after them instead of displacing `none`.
  assert.equal(wb.getFontCount(), 1);
  assert.equal(wb.getFillCount(), 2);
  assert.equal(wb.getBorderCount(), 1);
  assert.equal(wb.getCellXfCount(), 1);

  const fontResult = wb.addFont({
    name: 'Arial',
    size: 12,
    bold: true,
    italic: false,
    strike: false,
    underline: 0,
    colorArgb: 0xff112233,
  });
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
  assert.equal(fontDup, fontResult);
  assert.equal(wb.getFontCount(), 2);

  const fill = wb.addFill({ pattern: 1, fgArgb: 0xffff0000, bgArgb: 0xff000000 });
  const border = wb.addBorder({
    left: { style: 1, colorArgb: 0xff000000 },
    right: { style: 1, colorArgb: 0xff000000 },
    top: { style: 1, colorArgb: 0xff000000 },
    bottom: { style: 1, colorArgb: 0xff000000 },
    diagonal: { style: 0, colorArgb: 0 },
    diagonalUp: false,
    diagonalDown: false,
  });

  const builtin = wb.addNumFmt('General');
  assert.equal(builtin, 0);

  const custom = wb.addNumFmt('"USD" #,##0');
  assert.ok(custom >= 164, `expected custom id >= 164, got ${custom}`);

  const xf = wb.addCellXf({
    fontIndex: fontResult,
    fillIndex: fill,
    borderIndex: border,
    numFmtId: custom,
    horizontalAlign: 1,
    verticalAlign: 2,
    wrapText: true,
  });

  wb.setNumber(0, 0, 0, 1);
  wb.setCellXfIndex(0, 0, 0, xf);

  const reread = wb.getCellXf(xf);
  assert.equal(reread.fontIndex, fontResult);
  assert.equal(reread.numFmtId, custom);
  assert.equal(reread.wrapText, true);
  assert.equal(reread.hasAlignment, true);
  assert.equal(reread.hasHorizontalAlign, true);
  assert.equal(reread.hasVerticalAlign, true);
  assert.equal(reread.hasWrapText, true);
  assert.equal(reread.hasJustifyLastLine, false);

  const optional = wb.addCellXf({
    fontIndex: fontResult,
    fillIndex: fill,
    borderIndex: border,
    numFmtId: custom,
    textRotation: 255,
    indent: 0,
    relativeIndent: -3,
    shrinkToFit: false,
    readingOrder: 0,
    justifyLastLine: true,
  });
  const optionalRead = wb.getCellXf(optional);
  assert.equal(optionalRead.textRotation, 255);
  assert.equal(optionalRead.indent, 0);
  assert.equal(optionalRead.relativeIndent, -3);
  assert.equal(optionalRead.shrinkToFit, false);
  assert.equal(optionalRead.readingOrder, 0);
  assert.equal(optionalRead.justifyLastLine, true);
  assert.equal(optionalRead.hasAlignment, true);
  assert.equal(optionalRead.hasJustifyLastLine, true);
  assert.equal(wb.addCellXf(optionalRead), optional);

  const omitted = wb.addCellXf({
    fontIndex: fontResult,
    fillIndex: fill,
    borderIndex: border,
    numFmtId: custom,
  });
  const omittedRead = wb.getCellXf(omitted);
  assert.equal(omittedRead.hasAlignment, false);

  const explicitEmpty = wb.addCellXf({
    fontIndex: fontResult,
    fillIndex: fill,
    borderIndex: border,
    numFmtId: custom,
    hasAlignment: true,
  });
  assert.notEqual(explicitEmpty, omitted);
  const explicitEmptyRead = wb.getCellXf(explicitEmpty);
  assert.equal(explicitEmptyRead.hasAlignment, true);

  const rfont = wb.getFont(fontResult);
  assert.equal(rfont.name, 'Arial');
  assert.equal(rfont.bold, true);

  const rfill = wb.getFill(fill);
  assert.equal(rfill.pattern, 1);

  const rborder = wb.getBorder(border);
  assert.equal(rborder.left.style, 1);
  assert.equal(rborder.diagonalUp, false);

  const rfmt = wb.getNumFmtString(custom);
  assert.equal(rfmt, '"USD" #,##0');
});

test('selector colours survive get/add identity and save/load', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const theme = { kind: 2, rgb: 0, theme: 3, tint: 0.5, indexed: 0 };
  const indexed = { kind: 3, rgb: 0, theme: 0, tint: 0, indexed: 9 };
  const automatic = { kind: 4, rgb: 0, theme: 0, tint: 0, indexed: 0 };

  const font = wb.addFont({ name: 'SelectorFont', size: 11, colorArgb: 0x01020304, color: theme });
  const gotFont = wb.getFont(font);
  assert.equal(gotFont.color.kind, 2);
  assert.equal(gotFont.color.theme, 3);
  assert.equal(gotFont.color.tint, 0.5);
  assert.equal(gotFont.colorArgb, 0x01020304);
  const fontAgain = wb.addFont(gotFont);
  assert.equal(fontAgain, font);

  const fill = wb.addFill({
    pattern: 1,
    fgArgb: 0x05060708,
    bgArgb: 0x090a0b0c,
    fg: indexed,
    bg: automatic,
  });
  const gotFill = wb.getFill(fill);
  assert.equal(gotFill.fg.kind, 3);
  assert.equal(gotFill.fg.indexed, 9);
  assert.equal(gotFill.bg.kind, 4);
  const fillAgain = wb.addFill(gotFill);
  assert.equal(fillAgain, fill);

  const border = wb.addBorder({
    left: { style: 1, colorArgb: 0x01020304, color: theme },
    right: { style: 1, colorArgb: 0x05060708, color: indexed },
    top: { style: 1, colorArgb: 0x090a0b0c, color: automatic },
    bottom: { style: 0, colorArgb: 0 },
    diagonal: { style: 0, colorArgb: 0 },
  });
  const gotBorder = wb.getBorder(border);
  assert.equal(gotBorder.left.color.kind, 2);
  assert.equal(gotBorder.left.color.theme, 3);
  assert.equal(gotBorder.right.color.kind, 3);
  assert.equal(gotBorder.right.color.indexed, 9);
  assert.equal(gotBorder.top.color.kind, 4);
  const borderAgain = wb.addBorder(gotBorder);
  assert.equal(borderAgain, border);

  const dxf = wb.addDxf({
    font: { name: 'DxfSelector', size: 9, colorArgb: 0x11121314, color: automatic },
    fill: { pattern: 1, fgArgb: 0x15161718, fg: indexed },
    border: { left: { style: 1, colorArgb: 0x191a1b1c, color: theme } },
  });
  const gotDxf = wb.getDxf(dxf);
  assert.equal(gotDxf.font.color.kind, 4);
  assert.equal(gotDxf.fill.fg.kind, 3);
  assert.equal(gotDxf.border.left.color.kind, 2);
  const dxfAgain = wb.addDxf(gotDxf);
  assert.equal(gotDxf.font.hasSize, true);
  const bareDxf = wb.addDxf({ font: { bold: true, hasBold: true } });
  assert.equal(wb.getDxf(bareDxf).font.hasSize, false);
  assert.equal(dxfAgain, dxf);

  const saved = wb.save();
  const loaded = mod.Workbook.load(saved);
  const loadedFont = loaded.getFont(font);
  const loadedFill = loaded.getFill(fill);
  const loadedBorder = loaded.getBorder(border);
  const loadedDxf = loaded.getDxf(dxf);
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
  const protection = wb.addDxf({ protectionXml });
  assert.notEqual(alignment, protection);

  const gotAlignment = wb.getDxf(alignment);
  assert.equal(gotAlignment.alignmentXml, alignmentXml);
  assert.equal(gotAlignment.protectionXml, undefined);
  const gotProtection = wb.getDxf(protection);
  assert.equal(gotProtection.alignmentXml, undefined);
  assert.equal(gotProtection.protectionXml, protectionXml);

  const alignmentAgain = wb.addDxf(gotAlignment);
  assert.equal(alignmentAgain, alignment);
  const protectionAgain = wb.addDxf(gotProtection);
  assert.equal(protectionAgain, protection);

  const saved = wb.save();
  const loaded = mod.Workbook.load(saved);
  assert.equal(loaded.disposed, false);
  const loadedAlignment = loaded.getDxf(alignment);
  assert.equal(loadedAlignment.alignmentXml, alignmentXml);
  assert.equal(loadedAlignment.protectionXml, undefined);
  const loadedProtection = loaded.getDxf(protection);
  assert.equal(loadedProtection.alignmentXml, undefined);
  assert.equal(loadedProtection.protectionXml, protectionXml);
  const loadedAlignmentAgain = loaded.addDxf(loadedAlignment);
  assert.equal(loadedAlignmentAgain, alignment);
  const loadedProtectionAgain = loaded.addDxf(loadedProtection);
  assert.equal(loadedProtectionAgain, protection);
  loaded.dispose();
  wb.dispose();
});

test('addFont / getFont preserve superscript and round-trip to the same index', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const superscript = wb.addFont({ name: 'Arial', size: 12, vertAlign: 1, colorArgb: 0xff112233 });
  assert.equal(wb.getFont(superscript).vertAlign, 1);

  const before = wb.getFontCount();
  const again = wb.addFont(wb.getFont(superscript));
  assert.equal(again, superscript);
  assert.equal(wb.getFontCount(), before);

  const edited = wb.getFont(superscript);
  edited.colorArgb = 0xff00ff00;
  const recolored = wb.addFont(edited);
  assert.notEqual(recolored, superscript);
  const reread = wb.getFont(recolored);
  assert.equal(reread.vertAlign, 1);
  assert.equal(reread.colorArgb, 0xff00ff00);
});

test('addFont/setFont with a non-numeric field throws and does not commit a coerced value', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();

  const before = wb.getFontCount();
  assert.throws(() => wb.addFont({ name: 'Arial', size: 'not a number' }), TypeError);
  assert.equal(wb.getFontCount(), before, 'no font must have been added for the rejected call');

  const font = wb.addFont({ name: 'Arial', size: 12 });
  assert.throws(() => wb.setFont(font, { name: 'Arial', size: 'not a number' }), TypeError);
  // The font slot must still hold its original size, not a coerced 0.
  assert.equal(wb.getFont(font).size, 12);
});

test('phonetic runs keep their spans through a save/load round trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setText(0, 0, 0, '東京都');
  assert.deepEqual(
    readAll(
      () => wb.getCellPhoneticRunCount(0, 0, 0),
      (i) => wb.getCellPhoneticRun(0, 0, 0, i),
    ),
    [],
  );

  const runs = [
    { sb: 0, eb: 2, text: 'トウキョウ' },
    { sb: 2, eb: 3, text: 'ト' },
  ];
  wb.setCellPhoneticRuns(0, 0, 0, runs);
  assert.deepEqual(
    readAll(
      () => wb.getCellPhoneticRunCount(0, 0, 0),
      (i) => wb.getCellPhoneticRun(0, 0, 0, i),
    ),
    runs,
  );
  // The flattening getter still reports the concatenation.
  assert.equal(wb.getCellPhonetic(0, 0, 0), 'トウキョウト');

  const saved = wb.save();
  const loaded = mod.Workbook.load(saved);
  assert.deepEqual(
    readAll(
      () => loaded.getCellPhoneticRunCount(0, 0, 0),
      (i) => loaded.getCellPhoneticRun(0, 0, 0, i),
    ),
    runs,
  );

  // Writing the flattened reading back is the collapse the run API avoids.
  wb.setCellPhonetic(0, 0, 0, 'トウキョウト');
  assert.deepEqual(
    readAll(
      () => wb.getCellPhoneticRunCount(0, 0, 0),
      (i) => wb.getCellPhoneticRun(0, 0, 0, i),
    ),
    [{ sb: 0, eb: 3, text: 'トウキョウト' }],
  );

  assert.throws(
    () =>
      wb.setCellPhoneticRuns(0, 0, 0, [
        { sb: 2, eb: 3, text: 'ト' },
        { sb: 0, eb: 2, text: 'トウ' },
      ]),
    isFormulonError,
  );

  loaded.dispose();
  wb.dispose();
});

test('phonetic properties are independent of the runs and survive a round trip', async () => {
  const mod = await getModule();
  // Compare the triple alone: the record carries more fields than the ones under test.
  const props = (book) => {
    const { fontId, type, alignment } = book.getCellPhoneticProperties(0, 0, 0);
    return { fontId, type, alignment };
  };
  const wb = mod.Workbook.createDefault();
  wb.setText(0, 0, 0, '大阪');
  assert.deepEqual(props(wb), { fontId: 0, type: 0, alignment: 0 });

  wb.setCellPhoneticProperties(0, 0, 0, { fontId: 3, type: 2, alignment: 2 });
  // Setting the readings must not reset the rendering.
  wb.setCellPhoneticRuns(0, 0, 0, [{ sb: 0, eb: 2, text: 'おおさか' }]);
  assert.deepEqual(props(wb), { fontId: 3, type: 2, alignment: 2 });

  const saved = wb.save();
  const loaded = mod.Workbook.load(saved);
  assert.deepEqual(props(loaded), { fontId: 3, type: 2, alignment: 2 });

  // Two bits each on the binary side, so a wider ordinal is refused.
  assert.throws(() => wb.setCellPhoneticProperties(0, 0, 0, { fontId: 0, type: 4, alignment: 0 }), isFormulonError);
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
  assert.ok(appended > 0);
  assert.equal(wb.getFont(0).name, 'Calibri');

  wb.setDefaultFont({ name: '游ゴシック', size: 11, hasCharset: true, charset: 128 });
  assert.equal(wb.getFont(0).name, '游ゴシック');
  assert.equal(wb.getFont(0).charset, 128);

  wb.dispose();
});

test('setFont overwrites an existing slot and refuses an absent index', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const added = wb.addFont({ name: 'Meiryo', size: 12 });

  wb.setFont(added, { name: 'MS Gothic', size: 9 });
  assert.equal(wb.getFont(added).name, 'MS Gothic');

  const before = wb.getFontCount();
  assert.throws(() => wb.setFont(before, { name: 'MS Gothic', size: 9 }), isFormulonError);
  assert.equal(wb.getFontCount(), before);

  wb.dispose();
});

test('addDxf / getDxf round-trip a superscript differential font', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const added = wb.addDxf({ font: { name: 'Calibri', size: 9, vertAlign: 1 } });
  const readBack = wb.getDxf(added);
  assert.equal(readBack.font.vertAlign, 1);

  const before = wb.getDxfCount();
  const again = wb.addDxf(readBack);
  assert.equal(again, added);
  assert.equal(wb.getDxfCount(), before);
});

test('addXf rejects out-of-range font_index', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.throws(
    () =>
      wb.addCellXf({
        fontIndex: 99,
        fillIndex: 0,
        borderIndex: 0,
        numFmtId: 0,
        horizontalAlign: 0,
        verticalAlign: 0,
        wrapText: false,
      }),
    isFormulonError,
  );
});

test('active sheet is recorded in bookViews and survives a round trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.addSheet('B');
  wb.addSheet('C');
  assert.equal(wb.getActiveSheet(), 0);
  wb.setActiveSheet(2);
  const active = wb.getActiveSheet();
  assert.equal(active, 2);
  assert.throws(() => wb.setActiveSheet(9), isFormulonError);
  assert.equal(wb.getActiveSheet(), 2);
  const saved = wb.save();
  const loaded = mod.Workbook.load(saved);
  try {
    assert.equal(loaded.getActiveSheet(), 2);
  } finally {
    loaded.dispose();
    wb.dispose();
  }
});

test('the show-phonetic flag survives value writes and a round trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.getCellShowPhonetic(0, 0, 0), false);
  wb.setText(0, 0, 0, '漢字');
  wb.setCellShowPhonetic(0, 0, 0, true);
  wb.setText(0, 0, 0, '漢字');
  assert.equal(wb.getCellShowPhonetic(0, 0, 0), true);
  const saved = wb.save();
  const loaded = mod.Workbook.load(saved);
  try {
    assert.equal(loaded.getCellShowPhonetic(0, 0, 0), true);
    loaded.setCellShowPhonetic(0, 0, 0, false);
    assert.equal(loaded.getCellShowPhonetic(0, 0, 0), false);
    assert.throws(() => loaded.setCellShowPhonetic(99, 0, 0, true), isFormulonError);
  } finally {
    loaded.dispose();
    wb.dispose();
  }
});

test('an explicit all-zero phoneticPr is reported as present and survives a round trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setText(0, 0, 0, '漢字');
  wb.setCellPhoneticRuns(0, 0, 0, [{ sb: 0, eb: 2, text: 'かんじ' }]);
  assert.equal(wb.getCellPhoneticProperties(0, 0, 0).present, false);
  wb.setCellPhoneticProperties(0, 0, 0, { fontId: 0, type: 0, alignment: 0 });
  assert.equal(wb.getCellPhoneticProperties(0, 0, 0).present, true);
  const saved = wb.save();
  const loaded = mod.Workbook.load(saved);
  try {
    const got = loaded.getCellPhoneticProperties(0, 0, 0);
    assert.equal(got.present, true);
    assert.deepEqual([got.fontId, got.type, got.alignment], [0, 0, 0]);
  } finally {
    loaded.dispose();
    wb.dispose();
  }
});

test('scroll origin, pane origin and selection round-trip through save', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setFreeze(0, 2, 1);
  const view = wb.getView(0);
  assert.equal(view.paneTopLeftRow, 2);
  assert.equal(view.paneTopLeftCol, 1);
  wb.setTopLeft(0, 5, 0);
  wb.setPaneTopLeft(0, 40, 6);
  assert.throws(() => wb.setPaneTopLeft(0, 1, 6), isFormulonError);
  const ranges = [
    { firstRow: 41, firstCol: 7, lastRow: 43, lastCol: 8 },
    { firstRow: 50, firstCol: 9, lastRow: 50, lastCol: 9 },
  ];
  wb.setSelection(0, { activeRow: 41, activeCol: 7, ranges });
  assert.throws(
    () =>
      wb.setSelection(0, {
        activeRow: 0,
        activeCol: 0,
        ranges: [{ firstRow: 3, firstCol: 0, lastRow: 1, lastCol: 0 }],
      }),
    isFormulonError,
  );
  const saved = wb.save();
  const loaded = mod.Workbook.load(saved);
  try {
    const got = loaded.getView(0);
    assert.deepEqual([got.topLeftRow, got.topLeftCol, got.paneTopLeftRow, got.paneTopLeftCol], [5, 0, 40, 6]);
    const selection = loaded.getSelection(0);
    assert.equal(selection.activeRow, 41);
    assert.equal(selection.activeCol, 7);
    assert.deepEqual(
      selection.ranges.map((r) => [r.firstRow, r.firstCol, r.lastRow, r.lastCol]),
      [
        [41, 7, 43, 8],
        [50, 9, 50, 9],
      ],
    );
    assert.throws(() => loaded.getSelection(99), isFormulonError);
  } finally {
    loaded.dispose();
    wb.dispose();
  }
});

test('sheet split reads back, round-trips and exchanges with the freeze', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.deepEqual([wb.getView(0).splitX, wb.getView(0).splitY], [0, 0]);
    wb.setSplit(0, 2000, 1500);
    const view = wb.getView(0);
    assert.deepEqual([view.splitX, view.splitY, view.freezeRows, view.freezeCols], [2000, 1500, 0, 0]);
    const saved = wb.save();
    const loaded = mod.Workbook.load(saved);
    try {
      const got = loaded.getView(0);
      assert.deepEqual([got.splitX, got.splitY, got.freezeRows, got.freezeCols], [2000, 1500, 0, 0]);
    } finally {
      loaded.dispose();
    }
    wb.setFreeze(0, 2, 1);
    assert.deepEqual([wb.getView(0).splitX, wb.getView(0).splitY], [0, 0]);
    wb.setSplit(0, 0, 900);
    const split = wb.getView(0);
    assert.deepEqual([split.splitX, split.splitY, split.freezeRows, split.freezeCols], [0, 900, 0, 0]);
    assert.throws(() => wb.setSplit(0, -1, 0), isFormulonError);
    assert.throws(() => wb.setSplit(0, 0, Number.NaN), isFormulonError);
    assert.throws(() => wb.setSplit(0, Number.POSITIVE_INFINITY, 0), isFormulonError);
    assert.equal(wb.getView(0).splitY, 900);
    wb.setSplit(0, 0, 0);
    assert.equal(wb.getView(0).splitY, 0);
  } finally {
    wb.dispose();
  }
});
