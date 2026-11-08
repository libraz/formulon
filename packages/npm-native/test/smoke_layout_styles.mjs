import test from 'node:test';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';

const rgb = (value) => ({ kind: 1, argb: value, theme: 0, tint: 0, indexed: 0 });
const NO_COLOR = { kind: 0, argb: 0, theme: 0, tint: 0, indexed: 0 };

test('addMerge + getMerges round-trip a single range', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const range = { firstRow: 1, firstCol: 1, lastRow: 2, lastCol: 3 };
  wb.sheet(0).addMerge(range);
  const list = wb.sheet(0).listMerges();
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
  assert.equal(wb.sheet(0).listHyperlinks().length, 0);
  // Three entries with progressively more optional fields populated.
  wb.sheet(0).setHyperlink(1, 2, { target: 'https://example.com/' });
  wb.sheet(0).setHyperlink(3, 4, { target: 'mailto:hello@example.com', display: 'Hello' });
  wb.sheet(0).setHyperlink(5, 6, {
    target: null,
    location: 'Sheet1!A1',
    display: 'See X',
    tooltip: 'Internal link',
  });
  wb.sheet(0).setHyperlinkRange(
    { firstRow: 7, firstCol: 8, lastRow: 9, lastCol: 10 },
    { target: 'https://example.com/range', display: 'Range', tooltip: 'Range tip' },
  );
  const list = wb.sheet(0).listHyperlinks();
  assert.equal(list.length, 4);
  assert.deepEqual(list[0].rect, { firstRow: 1, firstCol: 2, lastRow: 1, lastCol: 2 });
  assert.equal(list[0].target, 'https://example.com/');
  assert.equal(list[0].display, null);
  assert.equal(list[0].tooltip, null);
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
  assert.throws(() => wb.sheet(999).setHyperlink(0, 0, { target: 'https://x/' }), isFormulonError);
  // clearHyperlinks drops everything.
  wb.sheet(0).clearHyperlinks();
  assert.equal(wb.sheet(0).listHyperlinks().length, 0);
});

test('setHyperlink upserts one link per cell and a null link removes it', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setHyperlink(1, 1, { target: 'https://example.com/a', display: 'A' });
  wb.sheet(0).setHyperlink(1, 1, { target: 'https://example.com/b' });
  const links = wb.sheet(0).listHyperlinks();
  assert.equal(links.length, 1);
  assert.equal(links[0].target, 'https://example.com/b');
  assert.equal(links[0].display, null);

  wb.sheet(0).setHyperlinkRange(
    { firstRow: 3, firstCol: 0, lastRow: 4, lastCol: 1 },
    { target: 'https://example.com/range' },
  );
  assert.equal(wb.sheet(0).listHyperlinks().length, 2);
  wb.sheet(0).setHyperlinkRange({ firstRow: 3, firstCol: 0, lastRow: 4, lastCol: 1 }, null);
  assert.equal(wb.sheet(0).listHyperlinks().length, 1);

  wb.sheet(0).setHyperlink(1, 1, null);
  assert.equal(wb.sheet(0).getHyperlink(1, 1), null);
  // Removing a link that is not there is a no-op.
  wb.sheet(0).setHyperlink(1, 1, null);
  assert.equal(wb.sheet(0).listHyperlinks().length, 0);
});

test('removeHyperlink / clearHyperlinks surface on an empty sheet', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Removing by cell reports NotFound when no hyperlink sits there; clearing is a no-op.
  assert.throws(
    () => wb.sheet(0).removeHyperlink(0, 0),
    (e) => isFormulonError(e, mod.StatusCode.NotFound),
  );
  wb.sheet(0).clearHyperlinks();
  // Sheet-index-out-of-range is rejected on every variant.
  const sheetRange = (e) => isFormulonError(e, mod.StatusCode.SheetIndexOutOfRange);
  assert.throws(() => wb.sheet(999).removeHyperlink(0, 0), sheetRange);
  assert.throws(() => wb.sheet(999).clearHyperlinks(), sheetRange);
  // The hyperlink list is still empty afterwards.
  assert.equal(wb.sheet(0).listHyperlinks().length, 0);
});

test('removeMerge + clearMerges step-wise prune the merge list', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const a = { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 };
  const b = { firstRow: 4, firstCol: 4, lastRow: 5, lastCol: 5 };
  wb.sheet(0).addMerge(a);
  wb.sheet(0).addMerge(b);
  // removeMerge with an overlap that hits `a` only.
  const overlap = { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 };
  wb.sheet(0).removeMerge(overlap);
  let list = wb.sheet(0).listMerges();
  assert.equal(list.length, 1);
  assert.equal(list[0].firstRow, 4);
  // A rect covering the survivor drops it; a rect with no overlap is NotFound.
  assert.throws(
    () => wb.sheet(0).removeMerge({ firstRow: 9, firstCol: 9, lastRow: 9, lastCol: 9 }),
    (e) => isFormulonError(e, mod.StatusCode.NotFound),
  );
  wb.sheet(0).removeMerge(b);
  list = wb.sheet(0).listMerges();
  assert.equal(list.length, 0);
  // clearMerges is a no-op on an empty list.
  wb.sheet(0).addMerge(a);
  wb.sheet(0).addMerge(b);
  wb.sheet(0).clearMerges();
  list = wb.sheet(0).listMerges();
  assert.equal(list.length, 0);
});

test('setSheetZoom + getSheetView surface the new zoom', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).updateView({ zoomPct: 175 });
  const v = wb.sheet(0).getView();
  assert.equal(v.zoomPct, 175);
  // Default freeze / visibility state survives a zoom change.
  assert.equal(v.freezeRows, 0);
  assert.equal(v.freezeCols, 0);
  assert.equal(v.visibility, mod.SheetVisibility.Visible);
});

test('sheet layout getters preserve explicit width presence flags', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setColumnWidth(0, 0, 0);
  wb.sheet(0).setColumnHidden(1, 1, true);
  wb.sheet(0).setRowHeight(3, 3, 30);

  const columns = wb.sheet(0).listColumnOverrides();
  const widthZero = columns.find((column) => column.first === 0 && column.last === 0);
  assert.ok(widthZero);
  assert.equal(widthZero.widthChars, 0);
  assert.equal(wb.sheet(0).ooxml.getColumnXfIndex(0), null);
  assert.equal(widthZero.hidden, false);
  const hidden = columns.find((column) => column.first === 1 && column.last === 1);
  assert.ok(hidden);
  assert.equal(hidden.hidden, true);
  assert.equal(hidden.widthChars, null);

  const rows = wb.sheet(0).listRowOverrides();
  const row = rows.find((entry) => entry.row === 3);
  assert.ok(row);
  assert.equal(row.heightPt, 30);
  assert.equal(wb.sheet(0).ooxml.getRowXfIndex(3), null);
});

test('getSheetView surfaces display / orientation defaults', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const v = wb.sheet(0).getView();
  assert.equal(v.showGridLines, true);
  assert.equal(v.showRowColHeaders, true);
  assert.equal(v.showZeros, true);
  assert.equal(v.rightToLeft, false);
  assert.equal(v.tabSelected, false);
  assert.equal(v.viewMode, mod.SheetViewMode.Normal);
});

test('setSheetShowGridLines / setSheetRightToLeft / setSheetTabSelected / setSheetViewMode round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).updateView({ showGridLines: false });
  wb.sheet(0).updateView({ showRowColHeaders: false });
  wb.sheet(0).updateView({ showZeros: false });
  wb.sheet(0).updateView({ rightToLeft: true });
  wb.sheet(0).updateView({ tabSelected: true });
  wb.sheet(0).updateView({ viewMode: mod.SheetViewMode.PageBreakPreview });
  const v = wb.sheet(0).getView();
  assert.equal(v.showGridLines, false);
  assert.equal(v.showRowColHeaders, false);
  assert.equal(v.showZeros, false);
  assert.equal(v.rightToLeft, true);
  assert.equal(v.tabSelected, true);
  assert.equal(v.viewMode, mod.SheetViewMode.PageBreakPreview);
});

test('insertRows shifts an existing literal forward', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Populate row 0; insert one row at row 0; the literal is now at row 1.
  wb.sheet(0).setNumber(0, 0, 99);
  wb.sheet(0).insertRows(0, 1);
  const r0 = wb.sheet(0).getTypedValue(0, 0);
  // The freshly-inserted row 0 is blank.
  assert.equal(r0.kind, mod.ValueKind.Blank);
  const r1 = wb.sheet(0).getTypedValue(1, 0);
  assert.equal(r1.kind, mod.ValueKind.Number);
  assert.equal(r1.number, 99);
});

test('setCellXfIndex + getCellXfIndex round-trip the xf id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Seed the cell so the xf attaches to a stored slot.
  wb.sheet(0).setNumber(0, 0, 1);
  // xfIndex 0 is always the default xf and is guaranteed to be valid.
  wb.sheet(0).ooxml.setCellXfIndex(0, 0, 0);
  const r = wb.sheet(0).ooxml.getCellXfIndex(0, 0);
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
  assert.equal(wb.ooxml.listFonts().length, 1);
  assert.equal(wb.ooxml.listFills().length, 2);
  assert.equal(wb.ooxml.listBorders().length, 1);
  assert.equal(wb.ooxml.listCellXfs().length, 1);

  const fontResult = wb.ooxml.addFont({
    name: 'Arial',
    sizePt: 12,
    bold: true,
    italic: false,
    strike: false,
    underline: 0,
    color: rgb(0xff112233),
  });
  // Linear-search dedup: the second call with the same payload returns the same index.
  const fontDup = wb.ooxml.addFont({
    name: 'Arial',
    sizePt: 12,
    bold: true,
    italic: false,
    strike: false,
    underline: 0,
    color: rgb(0xff112233),
  });
  assert.equal(fontDup, fontResult);
  assert.equal(wb.ooxml.listFonts().length, 2);

  const fill = wb.ooxml.addFill({ pattern: 1, foreground: rgb(0xffff0000), background: rgb(0xff000000) });
  const border = wb.ooxml.addBorder({
    left: { style: 1, color: rgb(0xff000000) },
    right: { style: 1, color: rgb(0xff000000) },
    top: { style: 1, color: rgb(0xff000000) },
    bottom: { style: 1, color: rgb(0xff000000) },
    diagonal: { style: 0, color: NO_COLOR },
    diagonalUp: false,
    diagonalDown: false,
  });

  const builtin = wb.ooxml.addNumberFormat('General');
  assert.equal(builtin, 0);

  const custom = wb.ooxml.addNumberFormat('"USD" #,##0');
  assert.ok(custom >= 164, `expected custom id >= 164, got ${custom}`);

  const xf = wb.ooxml.addCellXf({
    fontIndex: fontResult,
    fillIndex: fill,
    borderIndex: border,
    numFmtId: custom,
    horizontal: 1,
    vertical: 2,
    wrapText: true,
  });

  wb.sheet(0).setNumber(0, 0, 1);
  wb.sheet(0).ooxml.setCellXfIndex(0, 0, xf);

  const reread = wb.ooxml.getCellXf(xf);
  assert.equal(reread.fontIndex, fontResult);
  assert.equal(reread.numFmtId, custom);
  assert.equal(reread.wrapText, true);
  assert.equal(reread.horizontal, 1);
  assert.equal(reread.vertical, 2);
  assert.equal(reread.justifyLastLine, null);

  const optional = wb.ooxml.addCellXf({
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
  const optionalRead = wb.ooxml.getCellXf(optional);
  assert.equal(optionalRead.textRotation, 255);
  assert.equal(optionalRead.indent, 0);
  assert.equal(optionalRead.relativeIndent, -3);
  assert.equal(optionalRead.shrinkToFit, false);
  assert.equal(optionalRead.readingOrder, 0);
  assert.equal(optionalRead.justifyLastLine, true);
  assert.equal(optionalRead.horizontal, null);
  assert.equal(wb.ooxml.addCellXf(optionalRead), optional);

  const omitted = wb.ooxml.addCellXf({
    fontIndex: fontResult,
    fillIndex: fill,
    borderIndex: border,
    numFmtId: custom,
  });
  const omittedRead = wb.ooxml.getCellXf(omitted);
  assert.deepEqual([omittedRead.horizontal, omittedRead.wrapText, omittedRead.justifyLastLine], [null, null, null]);

  const explicitEmpty = wb.ooxml.addCellXf({
    fontIndex: fontResult,
    fillIndex: fill,
    borderIndex: border,
    numFmtId: custom,
    hasAlignment: true,
  });
  assert.notEqual(explicitEmpty, omitted);

  const rfont = wb.ooxml.getFont(fontResult);
  assert.equal(rfont.name, 'Arial');
  assert.equal(rfont.bold, true);

  const rfill = wb.ooxml.getFill(fill);
  assert.equal(rfill.pattern, 1);

  const rborder = wb.ooxml.getBorder(border);
  assert.equal(rborder.left.style, 1);
  assert.equal(rborder.diagonalUp, false);

  const rfmt = wb.ooxml.getNumberFormat(custom);
  assert.equal(rfmt, '"USD" #,##0');
});

test('selector colours survive get/add identity and save/load', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const theme = { kind: 2, argb: 0, theme: 3, tint: 0.5, indexed: 0 };
  const indexed = { kind: 3, argb: 0, theme: 0, tint: 0, indexed: 9 };
  const automatic = { kind: 4, argb: 0, theme: 0, tint: 0, indexed: 0 };

  const font = wb.ooxml.addFont({ name: 'SelectorFont', sizePt: 11, color: theme });
  const gotFont = wb.ooxml.getFont(font);
  assert.equal(gotFont.color.kind, 2);
  assert.equal(gotFont.color.theme, 3);
  assert.equal(gotFont.color.tint, 0.5);
  const fontAgain = wb.ooxml.addFont(gotFont);
  assert.equal(fontAgain, font);

  const fill = wb.ooxml.addFill({
    pattern: 1,
    foreground: indexed,
    background: automatic,
  });
  const gotFill = wb.ooxml.getFill(fill);
  assert.equal(gotFill.foreground.kind, 3);
  assert.equal(gotFill.foreground.indexed, 9);
  assert.equal(gotFill.background.kind, 4);
  const fillAgain = wb.ooxml.addFill(gotFill);
  assert.equal(fillAgain, fill);

  const border = wb.ooxml.addBorder({
    left: { style: 1, color: theme },
    right: { style: 1, color: indexed },
    top: { style: 1, color: automatic },
    bottom: { style: 0, color: NO_COLOR },
    diagonal: { style: 0, color: NO_COLOR },
  });
  const gotBorder = wb.ooxml.getBorder(border);
  assert.equal(gotBorder.left.color.kind, 2);
  assert.equal(gotBorder.left.color.theme, 3);
  assert.equal(gotBorder.right.color.kind, 3);
  assert.equal(gotBorder.right.color.indexed, 9);
  assert.equal(gotBorder.top.color.kind, 4);
  const borderAgain = wb.ooxml.addBorder(gotBorder);
  assert.equal(borderAgain, border);

  const dxf = wb.ooxml.addDxf({
    font: { name: 'DxfSelector', sizePt: 9, color: automatic },
    fill: { pattern: 1, foreground: indexed },
    border: { left: { style: 1, color: theme } },
  });
  const gotDxf = wb.ooxml.getDxf(dxf);
  assert.equal(gotDxf.font.color.kind, 4);
  assert.equal(gotDxf.fill.foreground.kind, 3);
  assert.equal(gotDxf.border.left.color.kind, 2);
  const dxfAgain = wb.ooxml.addDxf(gotDxf);
  assert.notEqual(gotDxf.font.sizePt, null);
  const bareDxf = wb.ooxml.addDxf({ font: { bold: true } });
  assert.equal(wb.ooxml.getDxf(bareDxf).font.sizePt, null);
  assert.equal(dxfAgain, dxf);

  const saved = wb.save();
  const loaded = mod.Workbook.load(saved);
  const loadedFont = loaded.ooxml.getFont(font);
  const loadedFill = loaded.ooxml.getFill(fill);
  const loadedBorder = loaded.ooxml.getBorder(border);
  const loadedDxf = loaded.ooxml.getDxf(dxf);
  assert.equal(loadedFont.color.kind, 2);
  assert.equal(loadedFill.foreground.kind, 3);
  assert.equal(loadedBorder.left.color.kind, 2);
  assert.equal(loadedDxf.font.color.kind, 4);
  loaded.dispose();
  wb.dispose();
});

test('dxf alignment and protection records survive get/add identity and save/load', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const alignmentRecord = {
    horizontal: mod.HorizontalAlign.Center,
    vertical: mod.VerticalAlign.Bottom,
    wrapText: true,
    shrinkToFit: false,
    indent: 0,
    textRotation: 0,
    readingOrder: mod.ReadingOrder.Context,
    justifyLastLine: false,
    relativeIndent: 0,
  };
  const protectionRecord = { locked: false, hidden: true };
  const alignment = wb.ooxml.addDxf({ alignment: alignmentRecord });
  const protection = wb.ooxml.addDxf({ protection: protectionRecord });
  assert.notEqual(alignment, protection);

  const gotAlignment = wb.ooxml.getDxf(alignment);
  assert.deepEqual(gotAlignment.alignment, alignmentRecord);
  assert.equal(gotAlignment.protection, undefined);
  const gotProtection = wb.ooxml.getDxf(protection);
  assert.equal(gotProtection.alignment, undefined);
  assert.deepEqual(gotProtection.protection, protectionRecord);
  for (const key of ['alignmentXml', 'protectionXml']) {
    assert.equal(key in gotAlignment, false, key);
  }

  const alignmentAgain = wb.ooxml.addDxf(gotAlignment);
  assert.equal(alignmentAgain, alignment);
  const protectionAgain = wb.ooxml.addDxf(gotProtection);
  assert.equal(protectionAgain, protection);

  const saved = wb.save();
  const loaded = mod.Workbook.load(saved);
  assert.equal(loaded.disposed, false);
  const loadedAlignment = loaded.ooxml.getDxf(alignment);
  assert.deepEqual(loadedAlignment.alignment, alignmentRecord);
  assert.equal(loadedAlignment.protection, undefined);
  const loadedProtection = loaded.ooxml.getDxf(protection);
  assert.equal(loadedProtection.alignment, undefined);
  assert.deepEqual(loadedProtection.protection, protectionRecord);
  const loadedAlignmentAgain = loaded.ooxml.addDxf(loadedAlignment);
  assert.equal(loadedAlignmentAgain, alignment);
  const loadedProtectionAgain = loaded.ooxml.addDxf(loadedProtection);
  assert.equal(loadedProtectionAgain, protection);
  loaded.dispose();
  wb.dispose();
});

test('a dxf carries its number format as a format code and reads null without one', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const dxf = wb.ooxml.addDxf({ numberFormat: '0.00%' });
  const read = wb.ooxml.getDxf(dxf);
  assert.equal(read.numberFormat, '0.00%');
  for (const key of ['numFmt', 'numFmtId', 'numFmtCode']) assert.equal(key in read, false, key);
  assert.equal(wb.ooxml.addDxf(read), dxf);
  const bare = wb.ooxml.addDxf({ font: { bold: true } });
  assert.equal(wb.ooxml.getDxf(bare).numberFormat, null);
  const loaded = mod.Workbook.load(wb.save());
  assert.equal(loaded.ooxml.getDxf(dxf).numberFormat, '0.00%');
  loaded.dispose();
  wb.dispose();
});

test('addFont / getFont preserve superscript and round-trip to the same index', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const superscript = wb.ooxml.addFont({ name: 'Arial', sizePt: 12, vertAlign: 1, color: rgb(0xff112233) });
  assert.equal(wb.ooxml.getFont(superscript).vertAlign, 1);

  const before = wb.ooxml.listFonts().length;
  const again = wb.ooxml.addFont(wb.ooxml.getFont(superscript));
  assert.equal(again, superscript);
  assert.equal(wb.ooxml.listFonts().length, before);

  const edited = wb.ooxml.getFont(superscript);
  edited.color = rgb(0xff00ff00);
  const recolored = wb.ooxml.addFont(edited);
  assert.notEqual(recolored, superscript);
  const reread = wb.ooxml.getFont(recolored);
  assert.equal(reread.vertAlign, 1);
  assert.equal(reread.color.argb, 0xff00ff00);
});

test('addFont/setFont with a non-numeric field throws and does not commit a coerced value', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();

  const before = wb.ooxml.listFonts().length;
  assert.throws(() => wb.ooxml.addFont({ name: 'Arial', sizePt: 'not a number' }), TypeError);
  assert.equal(wb.ooxml.listFonts().length, before, 'no font must have been added for the rejected call');

  const font = wb.ooxml.addFont({ name: 'Arial', sizePt: 12 });
  assert.throws(() => wb.ooxml.setFont(font, { name: 'Arial', sizePt: 'not a number' }), TypeError);
  // The font slot must still hold its original size, not a coerced 0.
  assert.equal(wb.ooxml.getFont(font).sizePt, 12);
});

test('phonetic runs keep their spans through a save/load round trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setText(0, 0, '東京都');
  assert.deepEqual(wb.sheet(0).getCellPhoneticRuns(0, 0), []);

  const runs = [
    { start: 0, end: 2, text: 'トウキョウ' },
    { start: 2, end: 3, text: 'ト' },
  ];
  wb.sheet(0).setCellPhoneticRuns(0, 0, runs);
  assert.deepEqual(wb.sheet(0).getCellPhoneticRuns(0, 0), runs);
  // The flattening getter still reports the concatenation.
  assert.equal(wb.sheet(0).getCellPhonetic(0, 0), 'トウキョウト');

  const saved = wb.save();
  const loaded = mod.Workbook.load(saved);
  assert.deepEqual(loaded.sheet(0).getCellPhoneticRuns(0, 0), runs);

  // Writing the flattened reading back is the collapse the run API avoids.
  wb.sheet(0).setCellPhonetic(0, 0, 'トウキョウト');
  assert.deepEqual(wb.sheet(0).getCellPhoneticRuns(0, 0), [{ start: 0, end: 3, text: 'トウキョウト' }]);

  assert.throws(
    () =>
      wb.sheet(0).setCellPhoneticRuns(0, 0, [
        { start: 2, end: 3, text: 'ト' },
        { start: 0, end: 2, text: 'トウ' },
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
    const { fontIndex, type, alignment } = book.sheet(0).getCellPhoneticProperties(0, 0);
    return { fontIndex, type, alignment };
  };
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setText(0, 0, '大阪');
  assert.deepEqual(props(wb), { fontIndex: 0, type: 0, alignment: 0 });

  wb.sheet(0).setCellPhoneticProperties(0, 0, { fontIndex: 3, type: 2, alignment: 2 });
  // Setting the readings must not reset the rendering.
  wb.sheet(0).setCellPhoneticRuns(0, 0, [{ start: 0, end: 2, text: 'おおさか' }]);
  assert.deepEqual(props(wb), { fontIndex: 3, type: 2, alignment: 2 });

  const saved = wb.save();
  const loaded = mod.Workbook.load(saved);
  assert.deepEqual(props(loaded), { fontIndex: 3, type: 2, alignment: 2 });

  // Two bits each on the binary side, so a wider ordinal is refused.
  assert.throws(
    () => wb.sheet(0).setCellPhoneticProperties(0, 0, { fontIndex: 0, type: 4, alignment: 0 }),
    isFormulonError,
  );
  assert.deepEqual(props(wb), { fontIndex: 3, type: 2, alignment: 2 });

  loaded.dispose();
  wb.dispose();
});

test('setDefaultFont declares what an unstyled cell is saved as', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.getDefaultFont().name, 'Calibri');

  // addFont can only ever append beside the seeded default.
  const appended = wb.ooxml.addFont({ name: '游ゴシック', sizePt: 11 });
  assert.ok(appended > 0);
  assert.equal(wb.getDefaultFont().name, 'Calibri');
  assert.deepEqual(wb.getDefaultFont(), wb.ooxml.getFont(0));

  wb.setDefaultFont({ name: '游ゴシック', sizePt: 11, charset: 128 });
  assert.equal(wb.getDefaultFont().name, '游ゴシック');
  assert.equal(wb.getDefaultFont().charset, 128);
  assert.deepEqual(wb.getDefaultFont(), wb.ooxml.getFont(0));

  wb.dispose();
});

test('setFont overwrites an existing slot and refuses an absent index', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const added = wb.ooxml.addFont({ name: 'Meiryo', sizePt: 12 });

  wb.ooxml.setFont(added, { name: 'MS Gothic', sizePt: 9 });
  assert.equal(wb.ooxml.getFont(added).name, 'MS Gothic');

  const before = wb.ooxml.listFonts().length;
  assert.throws(() => wb.ooxml.setFont(before, { name: 'MS Gothic', sizePt: 9 }), isFormulonError);
  assert.equal(wb.ooxml.listFonts().length, before);

  wb.dispose();
});

test('addDxf / getDxf round-trip a superscript differential font', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const added = wb.ooxml.addDxf({ font: { name: 'Calibri', sizePt: 9, vertAlign: 1 } });
  const readBack = wb.ooxml.getDxf(added);
  assert.equal(readBack.font.vertAlign, 1);

  const before = wb.ooxml.listDxfs().length;
  const again = wb.ooxml.addDxf(readBack);
  assert.equal(again, added);
  assert.equal(wb.ooxml.listDxfs().length, before);
});

test('addXf rejects out-of-range font_index', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.throws(
    () =>
      wb.ooxml.addCellXf({
        fontIndex: 99,
        fillIndex: 0,
        borderIndex: 0,
        numFmtId: 0,
        horizontal: 0,
        vertical: 0,
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
  assert.equal(wb.activeSheet().index, 0);
  wb.setActiveSheet(2);
  const active = wb.activeSheet();
  assert.equal(active.index, 2);
  assert.equal(active, wb.sheet(2));
  assert.throws(() => wb.setActiveSheet(9), isFormulonError);
  assert.equal(wb.activeSheet().index, 2);
  const saved = wb.save();
  const loaded = mod.Workbook.load(saved);
  try {
    assert.equal(loaded.activeSheet().index, 2);
  } finally {
    loaded.dispose();
    wb.dispose();
  }
});

test('the show-phonetic flag survives value writes and a round trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.sheet(0).getCellShowPhonetic(0, 0), false);
  wb.sheet(0).setText(0, 0, '漢字');
  wb.sheet(0).setCellShowPhonetic(0, 0, true);
  wb.sheet(0).setText(0, 0, '漢字');
  assert.equal(wb.sheet(0).getCellShowPhonetic(0, 0), true);
  const saved = wb.save();
  const loaded = mod.Workbook.load(saved);
  try {
    assert.equal(loaded.sheet(0).getCellShowPhonetic(0, 0), true);
    loaded.sheet(0).setCellShowPhonetic(0, 0, false);
    assert.equal(loaded.sheet(0).getCellShowPhonetic(0, 0), false);
    assert.throws(() => loaded.sheet(99).setCellShowPhonetic(0, 0, true), isFormulonError);
  } finally {
    loaded.dispose();
    wb.dispose();
  }
});

test('an explicit all-zero phoneticPr is reported as present and survives a round trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setText(0, 0, '漢字');
  wb.sheet(0).setCellPhoneticRuns(0, 0, [{ start: 0, end: 2, text: 'かんじ' }]);
  assert.equal(wb.sheet(0).getCellPhoneticProperties(0, 0).present, false);
  wb.sheet(0).setCellPhoneticProperties(0, 0, { fontIndex: 0, type: 0, alignment: 0 });
  assert.equal(wb.sheet(0).getCellPhoneticProperties(0, 0).present, true);
  const saved = wb.save();
  const loaded = mod.Workbook.load(saved);
  try {
    const got = loaded.sheet(0).getCellPhoneticProperties(0, 0);
    assert.equal(got.present, true);
    assert.deepEqual([got.fontIndex, got.type, got.alignment], [0, 0, 0]);
  } finally {
    loaded.dispose();
    wb.dispose();
  }
});

test('scroll origin, pane origin and selection round-trip through save', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).updateView({ freezeRows: 2, freezeCols: 1 });
  const view = wb.sheet(0).getView();
  assert.equal(view.paneTopLeftRow, 2);
  assert.equal(view.paneTopLeftCol, 1);
  wb.sheet(0).updateView({ topLeftRow: 5, topLeftCol: 0 });
  wb.sheet(0).updateView({ paneTopLeftRow: 40, paneTopLeftCol: 6 });
  assert.throws(() => wb.sheet(0).updateView({ paneTopLeftRow: 1, paneTopLeftCol: 6 }), isFormulonError);
  const ranges = [
    { firstRow: 41, firstCol: 7, lastRow: 43, lastCol: 8 },
    { firstRow: 50, firstCol: 9, lastRow: 50, lastCol: 9 },
  ];
  wb.sheet(0).setSelection({ active: 'H42', ranges });
  assert.throws(
    () =>
      wb.sheet(0).setSelection({
        active: 'A1',
        ranges: [{ firstRow: 3, firstCol: 0, lastRow: 1, lastCol: 0 }],
      }),
    isFormulonError,
  );
  const saved = wb.save();
  const loaded = mod.Workbook.load(saved);
  try {
    const got = loaded.sheet(0).getView();
    assert.deepEqual([got.topLeftRow, got.topLeftCol, got.paneTopLeftRow, got.paneTopLeftCol], [5, 0, 40, 6]);
    const selection = loaded.sheet(0).getSelection();
    assert.equal(selection.activeRow, 41);
    assert.equal(selection.activeCol, 7);
    assert.deepEqual(
      selection.ranges.map((r) => [r.firstRow, r.firstCol, r.lastRow, r.lastCol]),
      [
        [41, 7, 43, 8],
        [50, 9, 50, 9],
      ],
    );
    assert.throws(() => loaded.sheet(99).getSelection(), isFormulonError);
  } finally {
    loaded.dispose();
    wb.dispose();
  }
});

test('sheet split reads back, round-trips and exchanges with the freeze', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.deepEqual([wb.sheet(0).getView().splitXTwips, wb.sheet(0).getView().splitYTwips], [0, 0]);
    wb.sheet(0).updateView({ splitXTwips: 2000, splitYTwips: 1500 });
    const view = wb.sheet(0).getView();
    assert.deepEqual([view.splitXTwips, view.splitYTwips, view.freezeRows, view.freezeCols], [2000, 1500, 0, 0]);
    const saved = wb.save();
    const loaded = mod.Workbook.load(saved);
    try {
      const got = loaded.sheet(0).getView();
      assert.deepEqual([got.splitXTwips, got.splitYTwips, got.freezeRows, got.freezeCols], [2000, 1500, 0, 0]);
    } finally {
      loaded.dispose();
    }
    wb.sheet(0).updateView({ freezeRows: 2, freezeCols: 1 });
    assert.deepEqual([wb.sheet(0).getView().splitXTwips, wb.sheet(0).getView().splitYTwips], [0, 0]);
    wb.sheet(0).updateView({ splitXTwips: 0, splitYTwips: 900 });
    const split = wb.sheet(0).getView();
    assert.deepEqual([split.splitXTwips, split.splitYTwips, split.freezeRows, split.freezeCols], [0, 900, 0, 0]);
    assert.throws(() => wb.sheet(0).updateView({ splitXTwips: -1, splitYTwips: 0 }), isFormulonError);
    assert.throws(() => wb.sheet(0).updateView({ splitXTwips: 0, splitYTwips: Number.NaN }), isFormulonError);
    assert.throws(
      () => wb.sheet(0).updateView({ splitXTwips: Number.POSITIVE_INFINITY, splitYTwips: 0 }),
      isFormulonError,
    );
    assert.equal(wb.sheet(0).getView().splitYTwips, 900);
    wb.sheet(0).updateView({ splitXTwips: 0, splitYTwips: 0 });
    assert.equal(wb.sheet(0).getView().splitYTwips, 0);
  } finally {
    wb.dispose();
  }
});
