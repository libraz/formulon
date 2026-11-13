// Style tables, number formats, table style defaults, save formats and enum spellings through the public entry.

import assert from 'node:assert/strict';
import test from 'node:test';

import { getModule } from './smoke_support.mjs';

async function withWorkbook(fn) {
  const f = await getModule();
  const wb = f.Workbook.createDefault();
  try {
    return await fn(wb, f);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

test('ooxml.setCellStyleXf replaces a cell-style xf and keeps its index', async () => {
  await withWorkbook((wb, f) => {
    const { ooxml } = wb;
    const font = ooxml.addFont({ ...ooxml.getFont(0), bold: true });
    const before = ooxml.getCellStyleXf(0);
    ooxml.setCellStyleXf(0, { ...before, fontIndex: font, applyFont: true });
    assert.equal(ooxml.getCellStyleXf(0).fontIndex, font);
    assert.throws(
      () => ooxml.setCellStyleXf(1000000, before),
      (e) => e.name === 'FormulonError' && e.code === f.StatusCode.IndexOutOfRange,
    );
  });
});

test('ooxml.listNumberFormats lists declared formats in id order', async () => {
  await withWorkbook((wb) => {
    const a = wb.ooxml.addNumberFormat('0.000');
    const b = wb.ooxml.addNumberFormat('0.0000');
    const listed = wb.ooxml.listNumberFormats();
    const ids = listed.map((n) => n.numFmtId);
    assert.deepEqual(
      ids,
      [...ids].sort((x, y) => x - y),
    );
    assert.deepEqual(
      listed.find((n) => n.numFmtId === a),
      { numFmtId: a, numberFormat: '0.000' },
    );
    assert.deepEqual(
      listed.find((n) => n.numFmtId === b),
      { numFmtId: b, numberFormat: '0.0000' },
    );
  });
});

test('setCellStyle takes outlineLevel and defaults styleXfIndex', async () => {
  await withWorkbook((wb) => {
    // Only the RowLevel_n / ColLevel_n built-ins (builtinId 1 and 2) carry an outline level.
    wb.setCellStyle({ name: 'RowLevel_3', builtinId: 1, outlineLevel: 2 });
    const style = wb.getCellStyle('RowLevel_3');
    assert.equal(style.outlineLevel, 2);
    assert.equal(style.styleXfIndex, 0);
    assert.equal('iLevel' in style, false);
    wb.setCellStyle({ name: 'Plain' });
    assert.equal(wb.getCellStyle('Plain').outlineLevel, 0);
    wb.setCellStyle(style);
    assert.equal(wb.getCellStyle('RowLevel_3').outlineLevel, 2);
  });
});

test('WorkbookFormat is Xlsx 0 and Xlsb 1', async () => {
  await withWorkbook((wb, f) => {
    assert.deepEqual({ ...f.WorkbookFormat }, { Xlsx: 0, Xlsb: 1 });
    const xlsx = wb.save(f.WorkbookFormat.Xlsx);
    assert.equal(xlsx[0], 0x50);
    assert.deepEqual(wb.save(), xlsx);
    assert.ok(wb.save(f.WorkbookFormat.Xlsb).length > 0);
    assert.throws(
      () => wb.save(2),
      (e) => e.name === 'FormulonError' && e.code === f.StatusCode.InvalidArgument,
    );
  });
});

test('CF match icon reads iconSet and iconIndex', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setValues('A1', [[1], [2], [3]]);
    ws.addCfRule({
      type: f.CfRuleType.IconSet,
      ranges: 'A1:A3',
      iconSet: {
        name: f.IconSetName.ThreeArrows,
        thresholds: [
          { type: f.CfValueObjectType.Percent, value: '33' },
          { type: f.CfValueObjectType.Percent, value: '67' },
        ],
      },
    });
    const icons = ws
      .evaluateCfRange('A1:A3')
      .flatMap((cell) => cell.matches)
      .filter((m) => m.kind === f.CfMatchKind.IconSet);
    assert.ok(icons.length > 0);
    for (const m of icons) {
      assert.equal(m.icon.iconSet, f.IconSetName.ThreeArrows);
      assert.equal(typeof m.icon.iconIndex, 'number');
      assert.equal('setName' in m.icon, false);
      assert.equal('index' in m.icon, false);
    }
  });
});

test('table style defaults use tableStyleName and pivotStyleName', async () => {
  await withWorkbook((wb) => {
    wb.updateTableStyleDefaults({ tableStyleName: 'TableStyleMedium2', pivotStyleName: 'PivotStyleLight16' });
    const read = wb.getTableStyleDefaults();
    assert.equal(read.tableStyleName, 'TableStyleMedium2');
    assert.equal(read.pivotStyleName, 'PivotStyleLight16');
    assert.equal('tableStyle' in read, false);
    assert.equal('pivotStyle' in read, false);
    wb.updateTableStyleDefaults({ tableStyleName: null });
    assert.equal(wb.getTableStyleDefaults().tableStyleName, null);
    assert.equal(wb.getTableStyleDefaults().pivotStyleName, 'PivotStyleLight16');
  });
});

test('PivotCellKind spells Column out', async () => {
  await withWorkbook((wb, f) => {
    assert.equal(f.PivotCellKind.ColumnLabel, 2);
    assert.equal(f.PivotCellKind.ColumnSubtotal, 5);
    assert.equal('ColLabel' in f.PivotCellKind, false);
    assert.equal('ColSubtotal' in f.PivotCellKind, false);
  });
});

test('a cell font cleared of its size or name takes the Normal font, and of its colour has none', async () => {
  await withWorkbook((wb, f) => {
    wb.setDefaultFont({ name: 'Arial', sizePt: 14 });
    assert.equal(wb.ooxml.getFont(0).sizePt, 14);
    const ws = wb.sheet(0);
    ws.updateCellFormat('B2', { font: { name: 'Courier New', sizePt: 20, color: '#FF0000' } });
    assert.equal(ws.getCellFormat('B2').font.sizePt, 20);

    ws.updateCellFormat('B2', { font: { name: null, sizePt: null, color: null } });
    const font = ws.getCellFormat('B2').font;
    assert.equal(font.sizePt, 14);
    assert.equal(font.name, 'Arial');
    assert.equal(font.color.kind, f.ColorKind.None);

    const added = wb.ooxml.addFont({ name: 'Courier New' });
    assert.equal(wb.ooxml.getFont(added).sizePt, 14);
    const unnamed = wb.ooxml.addFont({ sizePt: 9 });
    assert.equal(wb.ooxml.getFont(unnamed).name, 'Arial');

    wb.setDefaultFont({ name: 'Calibri' });
    assert.equal(wb.ooxml.getFont(0).sizePt, 14);
  });
});

test('a font patch naming a different face unlinks it from the theme; size-only and same-name keep the link', async () => {
  await withWorkbook((wb0, f) => {
    const wb = f.Workbook.createForProfile('mac-365-ja_JP');
    try {
      const ws = wb.sheet(0);
      assert.equal(ws.getCellFormat('B2').font.scheme, f.FontScheme.Minor);
      ws.updateCellFormat('B2', { font: { sizePt: 14 } });
      assert.equal(ws.getCellFormat('B2').font.scheme, f.FontScheme.Minor);
      ws.updateCellFormat('B2', { font: { name: ws.getCellFormat('B2').font.name } });
      assert.equal(ws.getCellFormat('B2').font.scheme, f.FontScheme.Minor);
      ws.updateCellFormat('B2', { font: { name: 'Arial' } });
      const font = ws.getCellFormat('B2').font;
      assert.equal(font.name, 'Arial');
      assert.equal(font.scheme, f.FontScheme.None);
      assert.equal(font.charset, 128);
      ws.updateCellFormat('C3', { font: { name: 'Arial', scheme: f.FontScheme.Major } });
      assert.equal(ws.getCellFormat('C3').font.scheme, f.FontScheme.Major);
    } finally {
      wb.dispose();
    }
  });
});
