// The OOXML-level layer: `Workbook.ooxml` and `Sheet.ooxml` hold the style record tables, the XF indices, the
// raw worksheet XML and the range-style remap, and nothing of it stays on the workbook or the sheet.
// Colours there take a colour string and read back a `Color`. FORMULON_PACKAGE=native runs it against the
// staged @libraz/formulon-native; otherwise FORMULON_NPM_ENTRY selects the staged @libraz/formulon subpath.

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const testDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(testDir, '..', '..', '..');
const native = process.env.FORMULON_PACKAGE === 'native';

async function loadPackage() {
  if (native) {
    return import(pathToFileURL(path.join(repoRoot, 'packages', 'npm-native', 'dist', 'index.mjs')).href);
  }
  const pkgRoot = path.resolve(testDir, '..');
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  const target = pkg.exports?.[process.env.FORMULON_NPM_ENTRY ?? '.']?.import;
  if (!target) throw new Error('package.json "exports" has no import target for the selected entry');
  return import(pathToFileURL(path.resolve(pkgRoot, target)).href);
}

const mod = await loadPackage();
const formulon = await mod.createFormulon();
const { ColorKind, StatusCode, Workbook } = formulon;

const WORKBOOK_OOXML = [
  'getCellXf',
  'listCellXfs',
  'addCellXf',
  'getFont',
  'setFont',
  'addFont',
  'listFonts',
  'getFill',
  'addFill',
  'listFills',
  'getBorder',
  'addBorder',
  'listBorders',
  'getNumberFormat',
  'addNumberFormat',
  'getDxf',
  'addDxf',
  'listDxfs',
  'getCellStyleXf',
  'addCellStyleXf',
  'listCellStyleXfs',
  'listPassthroughs',
  'hasPassthrough',
];

const SHEET_OOXML = [
  'getAutoFilterXml',
  'setAutoFilterXml',
  'getPageSetupXml',
  'setPageSetupXml',
  'getPageMarginsXml',
  'setPageMarginsXml',
  'getPrintOptionsXml',
  'setPrintOptionsXml',
  'getHeaderFooterXml',
  'setHeaderFooterXml',
  'getSheetPrXml',
  'setSheetPrXml',
  'setColumnXfIndex',
  'resetColumnXfIndex',
  'setRowXfIndex',
  'resetRowXfIndex',
  'getColumnXfIndex',
  'getRowXfIndex',
  'getCellXfIndex',
  'setCellXfIndex',
  'resetCellXfIndex',
  'setRangeXfIndex',
  'getRangeXfBases',
  'getRangeXfBasesExcluding',
  'remapRangeXfs',
  'snapshotRangeXfs',
  'restoreRangeXfs',
  'getEffectiveStyle',
];

// The members that stay on the workbook and the sheet.
const WORKBOOK_DIRECT = [
  'setDefaultFont',
  'getDefaultFont',
  'getThemeFonts',
  'setThemeFonts',
  'getThemeColors',
  'setThemeColors',
  'listCellStyles',
  'getCellStyle',
  'setCellStyle',
  'removeCellStyle',
];
const SHEET_DIRECT = ['updateCellFormat'];

function withSheet(body) {
  const wb = Workbook.createDefault();
  try {
    return body(wb.sheet(0), wb);
  } finally {
    wb.dispose();
  }
}

const rgb = (argb) => ({ kind: ColorKind.Rgb, argb: argb, theme: 0, tint: 0, indexed: 0 });

test('ooxml is one object per workbook and per sheet, with no public constructor', () => {
  withSheet((ws, wb) => {
    assert.equal(wb.ooxml, wb.ooxml);
    assert.equal(ws.ooxml, ws.ooxml);
    assert.equal(wb.sheet(0).ooxml, ws.ooxml);
    assert.notEqual(wb.addSheet('Two').ooxml, ws.ooxml);
    assert.throws(() => new wb.ooxml.constructor(), TypeError);
    assert.throws(() => new ws.ooxml.constructor(), TypeError);
    assert.throws(() => {
      wb.ooxml = {};
    }, TypeError);
  });
});

test('the OOXML-level members live only under ooxml', () => {
  withSheet((ws, wb) => {
    for (const name of WORKBOOK_OOXML) {
      assert.equal(typeof wb.ooxml[name], 'function', `Workbook.ooxml.${name}`);
      assert.equal(wb[name], undefined, `Workbook.${name}`);
    }
    for (const name of SHEET_OOXML) {
      assert.equal(typeof ws.ooxml[name], 'function', `Sheet.ooxml.${name}`);
      assert.equal(ws[name], undefined, `Sheet.${name}`);
    }
    for (const name of WORKBOOK_DIRECT) assert.equal(typeof wb[name], 'function', `Workbook.${name}`);
    for (const name of SHEET_DIRECT) assert.equal(typeof ws[name], 'function', `Sheet.${name}`);
  });
});

test('the style records are read and added through Workbook.ooxml, colours as Color', () => {
  withSheet((_ws, wb) => {
    const { ooxml } = wb;
    const fonts = ooxml.listFonts();
    assert.deepEqual(fonts[0], ooxml.getFont(0));
    const index = ooxml.addFont({ ...fonts[0], bold: true, color: '#112233' });
    assert.deepEqual(ooxml.getFont(index).color, rgb(0xff112233));
    assert.equal(ooxml.getFont(index).bold, true);
    const fill = ooxml.addFill({
      pattern: formulon.FillPattern.Solid,
      foreground: '#80FF0000',
      background: rgb(0xff00ff00),
    });
    assert.deepEqual(ooxml.getFill(fill).foreground, rgb(0x80ff0000));
    assert.equal(ooxml.listFills().length, fill + 1);
    assert.equal(typeof ooxml.getNumberFormat(14), 'string');
    assert.ok(ooxml.addNumberFormat('0.000') >= 164);
    assert.deepEqual(ooxml.listCellXfs()[0], ooxml.getCellXf(0));
    assert.ok(Array.isArray(wb.listCellStyles()));
    assert.deepEqual(ooxml.listPassthroughs(), []);
    assert.equal(ooxml.hasPassthrough('xl/none.xml'), false);
  });
});

test('Sheet.ooxml reads and writes the XF indices with the A1 forms', () => {
  withSheet((ws, wb) => {
    ws.updateCellFormat('B2', { numberFormat: '0.00' });
    const xf = ws.ooxml.getCellXfIndex('B2');
    assert.equal(xf, ws.ooxml.getCellXfIndex(1, 1));
    assert.equal(wb.ooxml.getCellXf(xf).numFmtId, 2);
    ws.ooxml.setCellXfIndex('C3', xf);
    assert.equal(ws.ooxml.getCellXfIndex('C3'), xf);
    ws.ooxml.resetCellXfIndex('C3');
    assert.notEqual(ws.ooxml.getCellXfIndex('C3'), xf);
    ws.ooxml.setRangeXfIndex('D1:D2', xf);
    assert.equal(ws.ooxml.getCellXfIndex('D2'), xf);
    const bases = ws.ooxml.getRangeXfBases('A1:D4');
    assert.ok(bases.includes(xf));
    ws.ooxml.remapRangeXfs('D1:D2', new Map([[xf, 0]]));
    assert.equal(ws.ooxml.getCellXfIndex('D1'), 0);
  });
});

test('Sheet.ooxml reads the row and column XF the style setters store', () => {
  withSheet((ws) => {
    ws.updateCellFormat('Z1', { numberFormat: '0.00' });
    const xf = ws.ooxml.getEffectiveStyle('Z1').xfIndex;
    assert.equal(ws.ooxml.getColumnXfIndex(1), null);
    ws.ooxml.setColumnXfIndex('B:C', xf);
    assert.equal(ws.ooxml.getColumnXfIndex(1), xf);
    assert.equal(ws.ooxml.getColumnXfIndex(2), xf);
    assert.equal(ws.ooxml.getColumnXfIndex(3), null);
    ws.ooxml.setRowXfIndex('4', xf);
    assert.equal(ws.ooxml.getRowXfIndex(3), xf);
    assert.equal(ws.ooxml.getRowXfIndex(4), null);
    for (const layout of [...ws.listColumnOverrides(), ...ws.listRowOverrides()]) {
      assert.equal(Object.hasOwn(layout, 'styleXf'), false);
    }
    assert.ok(ws.listColumnOverrides().length > 0);
    ws.ooxml.resetColumnXfIndex('B:C');
    assert.equal(ws.ooxml.getColumnXfIndex(1), null);
    ws.ooxml.resetRowXfIndex(3, 3);
    assert.equal(ws.ooxml.getRowXfIndex(3), null);
    assert.throws(
      () => ws.ooxml.getColumnXfIndex(-1),
      (e) => e.name === 'FormulonError' && e.code === StatusCode.InvalidArgument,
    );
    assert.throws(() => ws.ooxml.getColumnXfIndex(1.5), RangeError);
  });
});

test('Sheet.ooxml replaces and clears the worksheet XML fragments', () => {
  withSheet((ws) => {
    ws.ooxml.setSheetPrXml('<sheetPr><outlinePr summaryBelow="0"/></sheetPr>');
    assert.match(ws.ooxml.getSheetPrXml(), /summaryBelow="0"/);
    ws.ooxml.setSheetPrXml(null);
    assert.equal(ws.ooxml.getSheetPrXml(), null);
    assert.equal(ws.ooxml.getAutoFilterXml(), null);
  });
});

test('ooxml follows the handle it belongs to', () => {
  const wb = Workbook.createDefault();
  const removed = wb.addSheet('Gone');
  const ooxml = removed.ooxml;
  wb.removeSheet(removed);
  assert.throws(
    () => ooxml.getSheetPrXml(),
    (e) => e.name === 'FormulonError' && e.code === StatusCode.NotFound,
  );
  assert.throws(
    () => removed.ooxml,
    (e) => e.name === 'FormulonError' && e.code === StatusCode.NotFound,
  );
  const wbOoxml = wb.ooxml;
  wb.dispose();
  assert.throws(
    () => wbOoxml.listFonts(),
    (e) => e.name === 'FormulonError' && e.code === StatusCode.BindingInvalidHandle,
  );
  assert.throws(
    () => wb.ooxml,
    (e) => e.name === 'FormulonError' && e.code === StatusCode.BindingInvalidHandle,
  );
});

test('colour inputs outside ooxml take a colour string as well', () => {
  withSheet((ws) => {
    ws.setValues('A1', [[1], [2], [3]]);
    ws.addCfRule({
      type: formulon.CfRuleType.ColorScale,
      ranges: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 0 }],
      colorScale: { thresholds: [{ type: 3 }, { type: 4 }], colors: ['#FF0000', rgb(0xff00ff00)] },
    });
    assert.deepEqual(ws.listCfRules().at(-1).colorScale.colors, [rgb(0xffff0000), rgb(0xff00ff00)]);
    ws.updateCellFormat('B1', { font: { color: '#0000FF' }, fill: { foreground: '#FFFF00' } });
    const format = ws.getCellFormat('B1');
    assert.deepEqual(format.font.color, rgb(0xff0000ff));
    assert.deepEqual(format.fill.foreground, rgb(0xffffff00));
    assert.equal(format.fill.pattern, formulon.FillPattern.Solid);
  });
});
