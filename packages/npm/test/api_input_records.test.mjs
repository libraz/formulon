// Input records and read records on the public surface: an unknown key on an input is a TypeError, a record a getter
// returns is accepted by the matching setter, A1 text stands for a rectangle, and the writable-only input shapes
// default their omitted keys. FORMULON_NPM_ENTRY selects the staged subpath: "." (default) or "./threads".

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const entry = process.env.FORMULON_NPM_ENTRY ?? '.';

async function loadPackage() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  const target = pkg.exports?.[entry]?.import;
  if (!target) throw new Error(`package.json "exports" has no import target for ${entry}`);
  return import(pathToFileURL(path.resolve(pkgRoot, target)).href);
}

const f = await (await loadPackage()).createFormulon();

/** Runs `body(ws, wb)` on a fresh default workbook and disposes it. */
function withSheet(body) {
  const wb = f.Workbook.createDefault();
  try {
    return body(wb.sheet(0), wb);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

function expectCode(fn, code) {
  assert.throws(fn, (e) => {
    assert.equal(e.name, 'FormulonError', `expected a FormulonError, got ${e}`);
    assert.equal(e.code, code, `code ${e.code}`);
    return true;
  });
}

const expectInvalid = (fn) => expectCode(fn, f.StatusCode.InvalidArgument);
const expectTypeError = (fn) => assert.throws(fn, (e) => e instanceof TypeError, 'expected a TypeError');

const PERSON = '{00000000-0000-4000-8000-0000000000A1}';
const PERSON2 = '{00000000-0000-4000-8000-0000000000A2}';
const rect = (firstRow, firstCol, lastRow, lastCol) => ({ firstRow, firstCol, lastRow, lastCol });
const plain = (value) => JSON.parse(JSON.stringify(value) ?? 'null');

function fillTable(ws) {
  ws.setText('A1', 'Region');
  ws.setText('B1', 'Amount');
  ws.setText('A2', 'East');
  ws.setNumber('B2', 10);
  ws.setText('A3', 'West');
  ws.setNumber('B3', 30);
}

describe('an unknown key on an input record is a TypeError', () => {
  it('names the misspelt validation key and adds nothing', () => {
    withSheet((ws) => {
      assert.throws(
        () => ws.addValidation({ ranges: 'A1', type: f.ValidationType.List, fromula1: '"a,b"' }),
        (e) => e instanceof TypeError && e.message.includes('fromula1'),
      );
      assert.equal(ws.listValidations().length, 0);
    });
  });

  // Each case runs once with a valid record (must succeed) and once with a stray key added (must be a TypeError).
  const CASES = {
    addTable: (ws, _wb, x) => ws.addTable({ range: 'A1:B3', name: 'T', ...x }),
    setComment: (ws, _wb, x) => ws.setComment('A1', { text: 'x', ...x }),
    setSelection: (ws, _wb, x) => ws.setSelection({ active: 'A1', ...x }),
    setPrintTitles: (ws, _wb, x) => ws.setPrintTitles({ rows: '1', columns: null, ...x }),
    setAutoFilter: (ws, _wb, x) => ws.setAutoFilter({ range: 'A1:B3', ...x }),
    addThreadedComment: (ws, wb, x) => {
      wb.addPerson({ id: PERSON, displayName: 'A' });
      ws.addThreadedComment({
        id: '{00000000-0000-4000-8000-000000000001}',
        text: 'Hi',
        cell: 'A1',
        personId: PERSON,
        ...x,
      });
    },
    setArrayFormula: (ws, _wb, x) => ws.setArrayFormula({ range: 'A5:A6', formula: '=SEQUENCE(2)', ...x }),
    addProtectedRange: (ws, _wb, x) => ws.addProtectedRange({ name: 'p', ranges: 'A1:B2', ...x }),
    addSparklineGroup: (ws, _wb, x) =>
      ws.addSparklineGroup({
        type: f.SparklineType.Column,
        sparklines: [{ source: 'Sheet1!A1:D1', location: 'E1', ...x }],
      }),
    addCfRule: (ws, _wb, x) => ws.addCfRule({ type: f.CfRuleType.Expression, ranges: 'A1', formula1: 'TRUE', ...x }),
    addChart: (ws, _wb, x) => ws.addChart({ type: f.ChartPlotType.Bar, range: 'Sheet1!A1:B3', ...x }),
    setDefinedName: (_ws, wb, x) => wb.setDefinedName({ name: 'X', formula: '1', ...x }),
    addPerson: (_ws, wb, x) => wb.addPerson({ id: PERSON, displayName: 'A', ...x }),
    setWorkbookProtection: (_ws, wb, x) => wb.setWorkbookProtection({ lockStructure: true, ...x }),
    addFont: (_ws, wb, x) => wb.ooxml.addFont({ name: 'Arial', ...x }),
    addCellXf: (_ws, wb, x) => wb.ooxml.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 0, ...x }),
  };
  for (const [name, call] of Object.entries(CASES)) {
    it(`${name} refuses a stray key and accepts the same record without it`, () => {
      withSheet((ws, wb) => {
        fillTable(ws);
        call(ws, wb, {});
      });
      withSheet((ws, wb) => {
        fillTable(ws);
        expectTypeError(() => call(ws, wb, { bogus: 1 }));
      });
    });
  }

  it('a nested key of a CF format is refused', () => {
    withSheet((ws) => {
      fillTable(ws);
      expectTypeError(() =>
        ws.addCfRule({ type: f.CfRuleType.Expression, ranges: 'A1', formula1: 'TRUE', format: { bogus: 1 } }),
      );
    });
  });

  it('removed keys of the reshaped records are unknown keys', () => {
    withSheet((ws, wb) => {
      fillTable(ws);
      expectTypeError(() =>
        wb.ooxml.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 0, hasAlignment: true }),
      );
      expectTypeError(() =>
        ws.setAutoFilter({
          range: 'A1:B3',
          columns: [{ columnOffset: 0, kind: f.FilterKind.Custom, customCount: 1 }],
        }),
      );
      expectTypeError(() => ws.setSelection({ activeRow: 0, activeCol: 0 }));
    });
  });

  it('accepts the read-only keys of a record a getter returned', () => {
    withSheet((ws, wb) => {
      ws.setComment('B1', { text: 'note', author: 'Ann' });
      ws.setComment('A1', ws.getComment('B1'));
      assert.equal(ws.getComment('A1').text, 'note');
      wb.setTableStyle({ name: 'S' });
      wb.setTableStyle(wb.getTableStyle('S'));
      wb.setCellStyle(wb.getCellStyle('Normal'));
    });
  });
});

describe('A1 text stands for a rectangle wherever a record holds one', () => {
  it('a protected range takes ranges as one A1 string and reads back rectangles', () => {
    withSheet((ws) => {
      ws.addProtectedRange({ name: 'p', ranges: 'A1:B2' });
      assert.deepEqual(plain(ws.listProtectedRanges()[0].ranges), [rect(0, 0, 1, 1)]);
      ws.updateProtectedRange('p', { ranges: ['B2', 'D1:D3'] });
      assert.deepEqual(plain(ws.listProtectedRanges()[0].ranges), [rect(1, 1, 1, 1), rect(0, 3, 2, 3)]);
    });
  });

  it('a protected range read back is accepted as input', () => {
    withSheet((ws) => {
      ws.addProtectedRange({ name: 'p', ranges: 'A1:B2' });
      const read = ws.listProtectedRanges()[0];
      ws.addProtectedRange({ ...read, name: 'q' });
      assert.deepEqual(plain(ws.listProtectedRanges()[1].ranges), plain(read.ranges));
    });
  });

  it('an array formula takes range as A1 and reads it back as range', () => {
    withSheet((ws) => {
      ws.setArrayFormula({ range: 'A5:A6', formula: '=SEQUENCE(2)' });
      const read = ws.getArrayFormula('A5');
      assert.deepEqual(plain(read.range), rect(4, 0, 5, 0));
      assert.equal(read.rect, undefined);
      ws.setArrayFormula({ ...read, range: 'C5:C6' });
      assert.deepEqual(plain(ws.getArrayFormula('C5').range), rect(4, 2, 5, 2));
    });
  });

  it('every cell of a legacy array block reads its formula and no spill', () => {
    withSheet((ws) => {
      ws.setArrayFormula({ range: 'E1:E3', formula: '=A1:A3*2' });
      for (const a1 of ['E1', 'E2', 'E3']) {
        assert.equal(ws.getFormula(a1), '=A1:A3*2', a1);
        assert.equal(ws.getSpill(a1), null, a1);
      }
      assert.deepEqual(
        ws.listCells('E1:E3').map((c) => c.formula),
        ['=A1:A3*2', '=A1:A3*2', '=A1:A3*2'],
      );
      ws.setFormula('G1', '=SEQUENCE(2)');
      assert.equal(ws.getFormula('G2'), null);
    });
  });

  it('an AutoFilter takes its range as A1', () => {
    withSheet((ws) => {
      fillTable(ws);
      ws.setAutoFilter({ range: 'A1:B3' });
      assert.deepEqual(plain(ws.getAutoFilter().range), rect(0, 0, 2, 1));
    });
  });

  it('print titles take span strings and read back spans', () => {
    withSheet((ws) => {
      ws.setPrintTitles({ rows: '1:2', columns: 'A:B' });
      assert.deepEqual(plain(ws.getPrintTitles()), { rows: { first: 0, last: 1 }, columns: { first: 0, last: 1 } });
      ws.setPrintTitles({ rows: '$3', columns: null });
      assert.deepEqual(plain(ws.getPrintTitles()), { rows: { first: 2, last: 2 }, columns: null });
      ws.setPrintTitles(ws.getPrintTitles());
      assert.deepEqual(plain(ws.getPrintTitles().rows), { first: 2, last: 2 });
    });
  });

  it('print titles refuse a span on the wrong axis or a malformed one', () => {
    withSheet((ws) => {
      expectInvalid(() => ws.setPrintTitles({ rows: 'A:B', columns: null }));
      expectInvalid(() => ws.setPrintTitles({ rows: null, columns: '1:2' }));
      expectInvalid(() => ws.setPrintTitles({ rows: '1:', columns: null }));
    });
  });

  it('a data table takes range and input cells as records and re-adds from its own read', () => {
    withSheet((ws) => {
      ws.addDataTable({ range: 'B2:B3', columnInput: { row: 0, col: 0 } });
      const read = ws.getDataTable('B2');
      assert.deepEqual(plain(read.range), rect(1, 1, 2, 1));
      assert.equal(read.rect, undefined);
      ws.removeDataTable('B2');
      ws.addDataTable(read);
      assert.deepEqual(plain(ws.getDataTable('B2').range), rect(1, 1, 2, 1));
    });
  });

  it('a hyperlink, a spill and a parsed reference carry range, not rect', () => {
    withSheet((ws, wb) => {
      ws.setHyperlink('A1', { target: 'https://example.com' });
      assert.deepEqual(plain(ws.getHyperlink('A1').range), rect(0, 0, 0, 0));
      assert.equal(ws.getHyperlink('A1').rect, undefined);
      ws.setFormula('C1', '=SEQUENCE(2)');
      assert.deepEqual(plain(ws.getSpill('C1').range), rect(0, 2, 1, 2));
      assert.equal(ws.getSpill('C1').rect, undefined);
      const parsed = wb.parseReference('B2:D4');
      assert.deepEqual(plain(parsed.range), rect(1, 1, 3, 3));
      assert.equal(parsed.rect, undefined);
    });
  });
});

describe('a record a getter returns is accepted by the matching setter', () => {
  it('fonts, fills and borders round-trip', () => {
    withSheet((_ws, wb) => {
      const ooxml = wb.ooxml;
      const font = ooxml.addFont(ooxml.getFont(0));
      assert.deepEqual(plain(ooxml.getFont(font)), plain(ooxml.getFont(0)));
      const fill = ooxml.addFill(ooxml.getFill(1));
      assert.deepEqual(plain(ooxml.getFill(fill)), plain(ooxml.getFill(1)));
      const border = ooxml.addBorder(ooxml.getBorder(0));
      assert.deepEqual(plain(ooxml.getBorder(border)), plain(ooxml.getBorder(0)));
    });
  });

  it('colour strings written to a font, fill and border read back as RGB', () => {
    withSheet((_ws, wb) => {
      const ooxml = wb.ooxml;
      const font = ooxml.getFont(ooxml.addFont({ name: 'Arial', color: '#FF0000' }));
      assert.equal(font.color.kind, f.ColorKind.Rgb);
      assert.equal(font.color.argb >>> 0, 0xffff0000);
      const fill = ooxml.getFill(ooxml.addFill({ pattern: f.FillPattern.Solid, foreground: '#FFC7CE' }));
      assert.equal(fill.foreground.kind, f.ColorKind.Rgb);
      assert.equal(fill.foreground.argb >>> 0, 0xffffc7ce);
      const border = ooxml.getBorder(ooxml.addBorder({ left: { style: f.BorderStyle.Thin, color: '#00FF00' } }));
      assert.equal(border.left.color.argb >>> 0, 0xff00ff00);
    });
  });

  it('a cell XF read back is found again, and a minimal one reads absent alignment', () => {
    withSheet((_ws, wb) => {
      const ooxml = wb.ooxml;
      const index = ooxml.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 0, horizontal: 2 });
      assert.equal(ooxml.addCellXf(ooxml.getCellXf(index)), index);
      const minimal = ooxml.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 0 });
      const xf = ooxml.getCellXf(minimal);
      assert.equal(xf.horizontal, null);
      assert.equal(xf.vertical, null);
    });
  });

  it('a person without userId and providerId reads them as null', () => {
    withSheet((_ws, wb) => {
      wb.addPerson({ id: PERSON, displayName: 'A' });
      const person = wb.getPerson(PERSON);
      assert.equal(person.userId, null);
      assert.equal(person.providerId, null);
      wb.addPerson({ ...person, id: PERSON2 });
      assert.equal(wb.getPerson(PERSON2).displayName, 'A');
    });
  });

  it('a comment author is optional and a read comment passes back', () => {
    withSheet((ws) => {
      ws.setComment('B2', { text: 'check this' });
      assert.equal(ws.getComment('B2').author, null);
      ws.setComment('C3', { text: 'by', author: 'Ann' });
      ws.setComment('D4', ws.getComment('C3'));
      assert.deepEqual(plain(ws.getComment('D4')), plain(ws.getComment('C3')));
      // A read with no author passes back too.
      ws.setComment('E5', ws.getComment('B2'));
      assert.equal(ws.getComment('E5').author, null);
    });
  });

  it('workbook protection defaults every omitted key', () => {
    withSheet((_ws, wb) => {
      wb.setWorkbookProtection({ lockStructure: true });
      const read = wb.getWorkbookProtection();
      assert.equal(read.lockStructure, true);
      assert.equal(read.lockWindows, false);
      assert.equal(read.workbookSpinCount, 0);
      assert.equal(read.workbookHashValue, null);
      wb.setWorkbookProtection(read);
      assert.equal(wb.getWorkbookProtection().lockStructure, true);
    });
  });

  it('a table style defaults pivot, table and elements and a read passes back', () => {
    withSheet((_ws, wb) => {
      wb.setTableStyle({ name: 'S' });
      const read = wb.getTableStyle('S');
      assert.equal(read.pivot, true);
      assert.equal(read.table, true);
      assert.deepEqual(plain(read.elements), []);
      wb.setTableStyle({
        name: 'E',
        elements: [{ type: f.TableStyleElementType.WholeTable }],
      });
      const element = wb.getTableStyle('E').elements[0];
      assert.equal(element.dxfIndex, null);
      assert.equal(element.size, 1);
      wb.setTableStyle(wb.getTableStyle('E'));
    });
  });

  it('a cell style read back is accepted', () => {
    withSheet((_ws, wb) => {
      const good = wb.getCellStyle('Normal');
      wb.setCellStyle(good);
      assert.deepEqual(plain(wb.getCellStyle('Normal')), plain(good));
    });
  });

  it('a defined name carries its sheet in scope and a read passes back', () => {
    withSheet((_ws, wb) => {
      wb.addSheet('Second');
      wb.setDefinedName({ name: 'X', formula: 'Sheet1!$A$1', scope: 1 });
      const read = wb.getDefinedName('X', 1);
      assert.equal(read.scope, 1);
      assert.equal(read.localSheetId, undefined);
      wb.setDefinedName(read);
      assert.equal(wb.getDefinedName('X', 1).scope, 1);
      wb.setDefinedName({ name: 'G', formula: '1' });
      assert.equal(wb.getDefinedName('G').scope, null);
    });
  });

  it('a selection reads active as a cell address and passes back', () => {
    withSheet((ws) => {
      ws.setSelection({ active: 'C3', ranges: 'A1:D5' });
      const read = ws.getSelection();
      assert.deepEqual(plain(read.active), { row: 2, col: 2 });
      assert.equal(read.activeRow, undefined);
      assert.equal(read.activeCol, undefined);
      ws.setSelection({ active: { row: 0, col: 0 } });
      ws.setSelection(read);
      assert.deepEqual(plain(ws.getSelection().ranges), [rect(0, 0, 4, 3)]);
    });
  });

  it('table column totals use the TableColumn key names', () => {
    withSheet((ws) => {
      fillTable(ws);
      const table = ws.addTable({ range: 'A1:B3', name: 'T' });
      table.setColumnTotals(1, { totalsRowFunction: f.TableTotalsRowFunction.Sum });
      const column = table.getProperties().columns[1];
      assert.equal(column.totalsRowFunction, f.TableTotalsRowFunction.Sum);
      table.setColumnTotals(0, { totalsRowLabel: 'Total' });
      table.setColumnTotals(1, table.getProperties().columns[1]);
      assert.equal(table.getProperties().columns[1].totalsRowFunction, f.TableTotalsRowFunction.Sum);
      expectTypeError(() => table.setColumnTotals(1, { function: f.TableTotalsRowFunction.Sum }));
    });
  });

  it('a conditional-format rule read back is accepted and a Min threshold reads value null', () => {
    withSheet((ws) => {
      ws.addCfRule({
        type: f.CfRuleType.ColorScale,
        ranges: 'A1:A5',
        colorScale: {
          thresholds: [{ type: f.CfValueObjectType.Min }, { type: f.CfValueObjectType.Max }],
          colors: ['#FF0000', '#00FF00'],
        },
      });
      const rule = ws.listCfRules()[0];
      assert.equal(rule.colorScale.thresholds[0].value, null);
      ws.addCfRule(rule);
      assert.equal(ws.listCfRules().length, 2);
    });
  });

  it('a pivot data field read back is accepted with a null numberFormat', () => {
    withSheet((ws, wb) => {
      const { pivot } = pivotOver(ws, wb);
      const read = pivot.listDataFields()[0];
      assert.equal(read.numberFormat, null);
      pivot.setDataField(0, { ...read, aggregation: f.PivotAggregation.Average });
      assert.equal(pivot.listDataFields()[0].aggregation, f.PivotAggregation.Average);
    });
  });
});

describe('required keys are required at runtime', () => {
  it('a pivot filter without a field is a TypeError', () => {
    withSheet((ws, wb) => {
      const { pivot } = pivotOver(ws, wb);
      expectTypeError(() => pivot.addFilter({ kind: f.PivotFilterKind.LabelEquals, value: 'East' }));
      expectTypeError(() => pivot.addFilter({ fieldIndex: 'Region', value: 'East' }));
    });
  });

  it('a slicer takes its position only inside the anchor and needs a box or a size', () => {
    withSheet((ws) => {
      fillTable(ws);
      const table = ws.addTable({ range: 'A1:B3', name: 'T' });
      expectTypeError(() => ws.addSlicer(table, 0, { fromRow: 1, fromCol: 3, toRow: 9, toCol: 5 }));
      expectInvalid(() => ws.addSlicer(table, 0, { anchor: { fromRow: 1 } }));
      expectInvalid(() => ws.addSlicer(table, 0, {}));
      const slicer = ws.addSlicer(table, 0, { anchor: { fromRow: 1, fromCol: 3, toRow: 9, toCol: 5 } });
      const read = slicer.getDrawingProperties().anchor;
      assert.deepEqual([read.fromRow, read.fromCol, read.toRow, read.toCol], [1, 3, 9, 5]);
    });
  });

  it('a slicer patch merges the anchor keys it names and refuses them outside the anchor', () => {
    withSheet((ws) => {
      fillTable(ws);
      const table = ws.addTable({ range: 'A1:B3', name: 'T' });
      const slicer = ws.addSlicer(table, 0, { anchor: { fromRow: 1, fromCol: 3, toRow: 9, toCol: 5 } });
      expectTypeError(() => slicer.update({ fromRow: 2 }));
      slicer.update({ anchor: { fromRow: 2 } });
      const read = slicer.getDrawingProperties().anchor;
      assert.deepEqual([read.fromRow, read.fromCol, read.toRow, read.toCol], [2, 3, 9, 5]);
    });
  });

  it('theme fonts need all four faces and a read passes back', () => {
    withSheet((_ws, wb) => {
      expectTypeError(() => wb.setThemeFonts({ majorLatin: 'X' }));
      const read = wb.getThemeFonts();
      wb.setThemeFonts(read);
      assert.deepEqual(plain(wb.getThemeFonts()), plain(read));
    });
  });
});

describe('the date system is an options object', () => {
  it('convertCivilToSerial reads { date1904 }', () => {
    assert.equal(f.convertCivilToSerial({ year: 1904, month: 1, day: 2 }, { date1904: true }), 1);
    assert.equal(f.convertCivilToSerial({ year: 1900, month: 1, day: 1 }), 1);
    assert.equal(f.convertCivilToSerial({ year: 1900, month: 1, day: 1 }, {}), 1);
  });

  it('convertSerialToCivil reads { date1904 }', () => {
    assert.deepEqual(plain(f.convertSerialToCivil(1, { date1904: true })), {
      year: 1904,
      month: 1,
      day: 2,
      hour: 0,
      minute: 0,
      second: 0,
    });
  });

  it('a bare boolean, a non-boolean flag and an unknown option are TypeErrors', () => {
    expectTypeError(() => f.convertSerialToCivil(1, true));
    expectTypeError(() => f.convertCivilToSerial({ year: 1904, month: 1, day: 2 }, true));
    expectTypeError(() => f.convertCivilToSerial({ year: 1904, month: 1, day: 2 }, { date1904: 1 }));
    expectTypeError(() => f.convertCivilToSerial({ year: 1904, month: 1, day: 2 }, { bogus: true }));
  });
});

describe('merged function metadata always has the merged shape', () => {
  it('a missing entry still yields localizedName equal to the canonical name', () => {
    const base = f.getFunctionMetadata('SUM');
    const merged = f.mergeFunctionMetadata(base, undefined, 'fr-FR');
    assert.equal(merged.localizedName, 'SUM');
    assert.equal(merged.name, base.name);
  });

  it('a null entry behaves as a missing one', () => {
    const merged = f.mergeFunctionMetadata(f.getFunctionMetadata('SUM'), null, 'fr-FR');
    assert.equal(typeof merged.localizedName, 'string');
  });
});

describe('sparklines carry a source rectangle and a location address', () => {
  it('reads source as a SheetRect and location as a CellAddress, and the group passes back', () => {
    const wb = f.Workbook.createDefault();
    try {
      wb.addSheet('Second');
      const ws = wb.sheet(0);
      ws.addSparklineGroup({
        type: f.SparklineType.Column,
        sparklines: [{ source: 'Sheet1!A1:D1', location: 'E1' }],
      });
      const group = ws.listSparklineGroups()[0];
      assert.deepEqual(plain(group.sparklines[0].source), { sheet: 0, rect: rect(0, 0, 0, 3) });
      assert.deepEqual(plain(group.sparklines[0].location), { row: 0, col: 4 });
      assert.equal(group.sparklines[0].formula, undefined);
      wb.sheet(1).addSparklineGroup(group);
      assert.equal(wb.sheet(1).listSparklineGroups().length, 1);
    } finally {
      wb.dispose();
    }
  });

  it('dateAxisRange is a SheetRect or null', () => {
    withSheet((ws) => {
      ws.addSparklineGroup({
        type: f.SparklineType.Line,
        sparklines: [{ source: 'Sheet1!A1:D1', location: 'E1' }],
        dateAxisRange: 'Sheet1!A2:D2',
      });
      const group = ws.listSparklineGroups()[0];
      assert.deepEqual(plain(group.dateAxisRange), { sheet: 0, rect: rect(1, 0, 1, 3) });
      assert.equal(group.dateAxisRef, undefined);
      ws.updateSparklineGroup(0, { dateAxisRange: null });
      assert.equal(ws.listSparklineGroups()[0].dateAxisRange, null);
    });
  });
});

describe('AutoFilter custom conditions are a list', () => {
  const column = (conditions) => ({ columnOffset: 0, kind: f.FilterKind.Custom, conditions });

  it('one condition reads back as a one-element list without the old fields', () => {
    withSheet((ws) => {
      fillTable(ws);
      const { ComparisonOperator, FilterKind } = f;
      ws.setAutoFilter({
        range: 'A1:B5',
        columns: [
          {
            columnOffset: 0,
            kind: FilterKind.Custom,
            conditions: [{ operator: ComparisonOperator.GreaterThan, value: '3' }],
          },
        ],
      });
      const read = ws.getAutoFilter().columns[0];
      assert.deepEqual(plain(read.conditions), [{ operator: ComparisonOperator.GreaterThan, value: '3' }]);
      for (const key of ['customCount', 'operator1', 'value1', 'operator2', 'value2']) {
        assert.equal(read[key], undefined, key);
      }
    });
  });

  it('two conditions keep their order and the read passes back', () => {
    withSheet((ws) => {
      fillTable(ws);
      const { ComparisonOperator } = f;
      ws.setAutoFilter({
        range: 'A1:B5',
        columns: [
          {
            ...column([
              { operator: ComparisonOperator.GreaterThan, value: '1' },
              { operator: ComparisonOperator.LessThan, value: '9' },
            ]),
            customAnd: true,
          },
        ],
      });
      const read = ws.getAutoFilter();
      assert.equal(read.columns[0].conditions.length, 2);
      assert.equal(read.columns[0].conditions[1].value, '9');
      ws.setAutoFilter(read);
      assert.equal(ws.getAutoFilter().columns[0].conditions.length, 2);
    });
  });

  it('a custom column with no or three conditions is AutoFilterInvalid and a Between operator InvalidArgument', () => {
    withSheet((ws) => {
      fillTable(ws);
      const { ComparisonOperator } = f;
      const eq = { operator: ComparisonOperator.Equal, value: '1' };
      expectCode(() => ws.setAutoFilter({ range: 'A1:B5', columns: [column([])] }), f.StatusCode.AutoFilterInvalid);
      expectCode(
        () => ws.setAutoFilter({ range: 'A1:B5', columns: [column([eq, eq, eq])] }),
        f.StatusCode.AutoFilterInvalid,
      );
      expectInvalid(() =>
        ws.setAutoFilter({
          range: 'A1:B5',
          columns: [column([{ operator: ComparisonOperator.Between, value: '1' }])],
        }),
      );
    });
  });
});

describe('a conditional-format match is tagged by its kind', () => {
  function dataBarMatch(ws, border) {
    for (let r = 0; r < 4; r += 1) ws.setNumber(r, 0, r + 1);
    ws.addCfRule({
      type: f.CfRuleType.DataBar,
      ranges: 'A1:A4',
      dataBar: {
        min: { type: f.CfValueObjectType.Min },
        max: { type: f.CfValueObjectType.Max },
        fill: '#638EC6',
        ...(border === undefined ? {} : { border }),
      },
    });
    return ws.evaluateCfRange('A1:A4').flatMap((cell) => cell.matches)[0];
  }

  it('a data-bar match carries a bar record with a finite length', () => {
    withSheet((ws) => {
      const match = dataBarMatch(ws);
      assert.equal(match.kind, f.CfMatchKind.DataBar);
      assert.ok(Number.isFinite(match.bar.lengthPct));
      assert.equal(match.icon, undefined);
      assert.equal(match.dxfIndexEngaged, undefined);
    });
  });

  it('a border-less bar reads border null and a bordered one a colour', () => {
    withSheet((ws) => {
      assert.equal(dataBarMatch(ws).bar.border, null);
    });
    withSheet((ws) => {
      assert.equal(dataBarMatch(ws, '#FF0000').bar.border.kind, f.ColorKind.Rgb);
    });
  });

  it('a colour-scale match has no bar', () => {
    withSheet((ws) => {
      for (let r = 0; r < 4; r += 1) ws.setNumber(r, 0, r + 1);
      ws.addCfRule({
        type: f.CfRuleType.ColorScale,
        ranges: 'A1:A4',
        colorScale: {
          thresholds: [{ type: f.CfValueObjectType.Min }, { type: f.CfValueObjectType.Max }],
          colors: ['#FF0000', '#00FF00'],
        },
      });
      const match = ws.evaluateCfRange('A1:A4').flatMap((cell) => cell.matches)[0];
      assert.equal(match.kind, f.CfMatchKind.ColorScale);
      assert.equal(match.bar, undefined);
      assert.equal(match.color.kind, f.ColorKind.Rgb);
    });
  });
});

/** A cache over A1:B3 and a pivot with Region on rows and a Sum of Amount, built through the public surface. */
function pivotOver(ws, wb) {
  fillTable(ws);
  const cache = wb.addPivotCache();
  cache.addField('Region');
  cache.addField('Amount');
  for (const [region, amount] of [
    ['East', 10],
    ['West', 30],
  ]) {
    const record = cache.addRecord();
    cache.setRecordValue(record, 0, region);
    cache.setRecordValue(record, 1, amount);
  }
  const pivot = ws.addPivot({ name: 'P', cacheId: cache.cacheId, anchor: 'E1' });
  pivot.setFieldAxis('Region', f.PivotAxis.Row);
  pivot.setFieldAxis('Amount', f.PivotAxis.Value);
  const amount = 1;
  pivot.addDataField({ name: 'Sum of Amount', fieldIndex: amount, aggregation: f.PivotAggregation.Sum });
  return { cache, pivot };
}
