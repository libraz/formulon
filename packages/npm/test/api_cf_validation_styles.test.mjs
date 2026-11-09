// Validation and sparkline ids, validation / conditional-format updates, the ooxml style-table setters and theme
// script font lists on the public entry.
// FORMULON_NPM_ENTRY selects the staged subpath: "." (default) or "./threads".

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

const { StatusCode, ValidationType, CfRuleType, FontScheme, FillPattern, BorderStyle } = f;

/** Runs `body(ws, wb)` on a default workbook, then disposes it. */
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

const expectTypeError = (fn) => assert.throws(fn, (e) => e instanceof TypeError, 'expected a TypeError');

const whole = (ranges, extra = {}) => ({ type: ValidationType.Whole, ranges, formula1: '1', formula2: '5', ...extra });
const group = (source, location) => ({ sparklines: [{ source, location }] });

describe('a data-validation rule has an id', () => {
  it('addValidation returns the id the listed rule carries, and ids differ per rule', () => {
    withSheet((ws) => {
      const first = ws.addValidation(whole('A1:A3'));
      const second = ws.addValidation(whole('B1:B3'));
      assert.equal(typeof first, 'string');
      assert.notEqual(first, second);
      assert.deepEqual(
        ws.listValidations().map((v) => v.id),
        [first, second],
      );
    });
  });

  it('an id on the input is ignored and a re-added record gets a fresh id', () => {
    withSheet((ws) => {
      const first = ws.addValidation(whole('A1:A3'));
      const again = ws.addValidation(ws.listValidations()[0]);
      const withId = ws.addValidation({ ...whole('C1:C3'), id: first });
      assert.equal(new Set([first, again, withId]).size, 3);
      assert.equal(ws.listValidations().length, 3);
    });
  });

  it('a rule keeps its id across row edits and a removal does not reuse it', () => {
    withSheet((ws) => {
      const first = ws.addValidation(whole('A1:A3'));
      const second = ws.addValidation(whole('B1:B3'));
      ws.insertRows(0, 1);
      const moved = ws.listValidations();
      assert.deepEqual(
        moved.map((v) => v.id),
        [first, second],
      );
      assert.equal(moved[0].ranges[0].firstRow, 2);
      ws.removeValidation(first);
      const third = ws.addValidation(whole('D1:D3'));
      assert.notEqual(third, first);
      assert.notEqual(third, second);
    });
  });

  it('remove, update and the read accept the id or the position', () => {
    withSheet((ws) => {
      const first = ws.addValidation(whole('A1:A3'));
      const second = ws.addValidation(whole('B1:B3'));
      ws.updateValidation(second, { errorMessage: 'by id' });
      ws.updateValidation(0, { errorMessage: 'by position' });
      assert.equal(ws.listValidations()[1].errorMessage, 'by id');
      assert.equal(ws.listValidations()[0].errorMessage, 'by position');
      ws.removeValidation(first);
      assert.deepEqual(
        ws.listValidations().map((v) => v.id),
        [second],
      );
      ws.removeValidation(0);
      assert.equal(ws.listValidations().length, 0);
    });
  });

  it('an id no rule has is NotFound, a position past the end or below zero is IndexOutOfRange', () => {
    withSheet((ws) => {
      const id = ws.addValidation(whole('A1:A3'));
      expectCode(() => ws.updateValidation('no-such-id', {}), StatusCode.NotFound);
      expectCode(() => ws.removeValidation('no-such-id'), StatusCode.NotFound);
      expectCode(() => ws.removeValidation(1), StatusCode.IndexOutOfRange);
      expectCode(() => ws.updateValidation(5, {}), StatusCode.IndexOutOfRange);
      expectCode(() => ws.removeValidation(-1), StatusCode.IndexOutOfRange);
      ws.removeValidation(id);
      expectCode(() => ws.removeValidation(id), StatusCode.NotFound);
    });
  });

  it('a string that looks like a position is an id', () => {
    withSheet((ws) => {
      const id = ws.addValidation(whole('A1:A3'));
      ws.removeValidation(id);
      ws.addValidation(whole('A1:A3'));
      expectCode(() => ws.removeValidation(id), StatusCode.NotFound);
      assert.equal(ws.listValidations().length, 1);
    });
  });

  it('validateValue names the applied rule by id and none without a rule', () => {
    withSheet((ws) => {
      const id = ws.addValidation(whole('B1:B3'));
      assert.equal(ws.validateValue('B1', 9).ruleId, id);
      assert.equal(ws.validateValue('B1', 9).valid, false);
      assert.equal(ws.validateValue('B1', 3).ruleId, id);
      assert.equal(ws.validateValue('D9', 3).ruleId, null);
      assert.equal(ws.validateValue('D9', 3).valid, true);
    });
  });

  it('every rule has an id again after save and load', () => {
    withSheet((ws, wb) => {
      ws.addValidation(whole('A1:A3'));
      ws.addValidation(whole('B1:B3'));
      const back = f.Workbook.load(wb.save());
      try {
        const ids = back
          .sheet(0)
          .listValidations()
          .map((v) => v.id);
        assert.equal(ids.length, 2);
        assert.equal(new Set(ids).size, 2);
        for (const id of ids) assert.equal(typeof id, 'string');
      } finally {
        back.dispose();
      }
    });
  });
});

describe('updateValidation changes the named keys and keeps the rest', () => {
  it('a patch with one key leaves every other key as it was', () => {
    withSheet((ws) => {
      const id = ws.addValidation(whole('A1:A3', { errorMessage: 'old', promptTitle: 'P', allowBlank: false }));
      const before = ws.listValidations()[0];
      ws.updateValidation(id, { errorMessage: 'new' });
      assert.deepEqual(ws.listValidations()[0], { ...before, errorMessage: 'new' });
    });
  });

  it('an empty patch changes nothing', () => {
    withSheet((ws) => {
      const id = ws.addValidation(whole('A1:A3', { errorMessage: 'old' }));
      const before = ws.listValidations()[0];
      ws.updateValidation(id, {});
      assert.deepEqual(ws.listValidations()[0], before);
    });
  });

  it('null clears the two formulas and the four texts and no other key', () => {
    withSheet((ws) => {
      const id = ws.addValidation(
        whole('A1:A3', { errorTitle: 'et', errorMessage: 'em', promptTitle: 'pt', promptMessage: 'pm' }),
      );
      ws.updateValidation(id, { formula2: null, errorTitle: null, errorMessage: null, promptTitle: null });
      const rule = ws.listValidations()[0];
      assert.equal(rule.formula2, null);
      assert.equal(rule.errorTitle, null);
      assert.equal(rule.errorMessage, null);
      assert.equal(rule.promptTitle, null);
      assert.equal(rule.promptMessage, 'pm');
      assert.equal(rule.formula1, '=1');
      ws.updateValidation(id, { formula1: null, promptMessage: null });
      assert.equal(ws.listValidations()[0].formula1, null);
      for (const key of [
        'type',
        'operator',
        'errorStyle',
        'allowBlank',
        'showInputMessage',
        'showErrorMessage',
        'showInCellDropdown',
        'ranges',
      ]) {
        expectTypeError(() => ws.updateValidation(id, { [key]: null }));
      }
    });
  });

  it('type, operator and ranges are replaced when given', () => {
    withSheet((ws) => {
      const id = ws.addValidation(whole('A1:A3'));
      ws.updateValidation(id, { operator: f.ComparisonOperator.GreaterThan, ranges: 'C1:C2' });
      const rule = ws.listValidations()[0];
      assert.equal(rule.operator, f.ComparisonOperator.GreaterThan);
      assert.deepEqual(rule.ranges, [{ firstRow: 0, firstCol: 2, lastRow: 1, lastCol: 2 }]);
      ws.updateValidation(id, { type: ValidationType.Decimal });
      assert.equal(ws.listValidations()[0].type, ValidationType.Decimal);
    });
  });

  it('an empty range list is InvalidArgument and an unknown key is a TypeError', () => {
    withSheet((ws) => {
      const id = ws.addValidation(whole('A1:A3'));
      const before = ws.listValidations()[0];
      expectCode(() => ws.updateValidation(id, { ranges: [] }), StatusCode.InvalidArgument);
      expectTypeError(() => ws.updateValidation(id, { bogus: 1 }));
      assert.deepEqual(ws.listValidations()[0], before);
    });
  });

  it('a formula that does not parse fails with its parser status and leaves the rule unchanged', () => {
    withSheet((ws) => {
      const id = ws.addValidation(whole('A1:A3'));
      const before = ws.listValidations()[0];
      assert.throws(
        () => ws.updateValidation(id, { formula1: '((' }),
        (e) => e.name === 'FormulonError' && e.code >= 1000 && e.code <= 1014,
      );
      assert.deepEqual(ws.listValidations()[0], before);
    });
  });

  it('the record the list returns passes back as it is', () => {
    withSheet((ws) => {
      ws.addValidation(whole('A1:A3', { errorMessage: 'em' }));
      const record = ws.listValidations()[0];
      ws.updateValidation(0, record);
      assert.deepEqual(ws.listValidations()[0], record);
    });
  });
});

describe('a sparkline group has an id', () => {
  it('addSparklineGroup returns the id the listed group carries', () => {
    withSheet((ws) => {
      const first = ws.addSparklineGroup(group('Sheet1!A1:D1', 'F1'));
      const second = ws.addSparklineGroup(group('Sheet1!A2:D2', 'F2'));
      assert.equal(typeof first, 'string');
      assert.notEqual(first, second);
      assert.deepEqual(
        ws.listSparklineGroups().map((g) => g.id),
        [first, second],
      );
    });
  });

  it('update and remove accept the id or the position, and an unknown id is NotFound', () => {
    withSheet((ws) => {
      const first = ws.addSparklineGroup(group('Sheet1!A1:D1', 'F1'));
      const second = ws.addSparklineGroup(group('Sheet1!A2:D2', 'F2'));
      ws.updateSparklineGroup(second, { markers: true });
      ws.updateSparklineGroup(0, { high: true });
      assert.equal(ws.listSparklineGroups()[1].markers, true);
      assert.equal(ws.listSparklineGroups()[0].high, true);
      expectCode(() => ws.updateSparklineGroup('no-such-id', { markers: true }), StatusCode.NotFound);
      expectCode(() => ws.removeSparklineGroup('no-such-id'), StatusCode.NotFound);
      expectCode(() => ws.removeSparklineGroup(2), StatusCode.IndexOutOfRange);
      expectCode(() => ws.removeSparklineGroup(-1), StatusCode.IndexOutOfRange);
      ws.removeSparklineGroup(first);
      assert.deepEqual(
        ws.listSparklineGroups().map((g) => g.id),
        [second],
      );
      expectCode(() => ws.removeSparklineGroup(first), StatusCode.NotFound);
    });
  });

  it('a group keeps its id across row edits and a removed id is not reused', () => {
    withSheet((ws) => {
      const first = ws.addSparklineGroup(group('Sheet1!A1:D1', 'F1'));
      ws.insertRows(0, 0);
      const [moved] = ws.listSparklineGroups();
      assert.equal(moved.id, first);
      assert.deepEqual(moved.sparklines[0].location, { row: 1, col: 5 });
      ws.removeSparklineGroup(first);
      const second = ws.addSparklineGroup(group('Sheet1!A1:D1', 'F1'));
      assert.notEqual(second, first);
    });
  });

  it('the read record passes back through updateSparklineGroup as it is', () => {
    withSheet((ws) => {
      const id = ws.addSparklineGroup({ ...group('Sheet1!A1:D1', 'F1'), markers: true });
      const record = ws.listSparklineGroups()[0];
      ws.updateSparklineGroup(id, record);
      assert.deepEqual(ws.listSparklineGroups()[0], record);
    });
  });
});

describe('updateCfRule changes the named keys of one rule', () => {
  const expression = { type: CfRuleType.Expression, ranges: 'A1:A5', formula1: 'A1>1' };

  it('sets and clears dxfIndex on a rule addressed by position or id', () => {
    withSheet((ws, wb) => {
      wb.ooxml.addDxf({ font: { bold: true } });
      ws.addCfRule(expression);
      ws.updateCfRule(0, { dxfIndex: 0 });
      assert.equal(ws.listCfRules()[0].dxfIndex, 0);
      ws.updateCfRule(ws.listCfRules()[0].id, { dxfIndex: null });
      assert.equal(ws.listCfRules()[0].dxfIndex, null);
    });
  });

  it('a patch keeps every key it does not name', () => {
    withSheet((ws) => {
      ws.addCfRule({ ...expression, stopIfTrue: true });
      const before = ws.listCfRules()[0];
      ws.updateCfRule(0, { formula1: 'A1>2' });
      assert.deepEqual(ws.listCfRules()[0], { ...before, formula1: '=A1>2' });
    });
  });

  it('formula1 and dxfIndex clear with null and ranges replace', () => {
    withSheet((ws, wb) => {
      wb.ooxml.addDxf({ font: { bold: true } });
      ws.addCfRule({ ...expression, dxfIndex: 0 });
      ws.updateCfRule(0, { ranges: 'B1:B2', formula1: null, dxfIndex: null });
      const rule = ws.listCfRules()[0];
      assert.deepEqual(rule.ranges, [{ firstRow: 0, firstCol: 1, lastRow: 1, lastCol: 1 }]);
      assert.equal(rule.formula1, null);
      assert.equal(rule.dxfIndex, null);
      expectTypeError(() => ws.updateCfRule(0, { ranges: null }));
      expectCode(() => ws.updateCfRule(0, { ranges: [] }), StatusCode.InvalidArgument);
    });
  });

  it('an empty patch changes nothing', () => {
    withSheet((ws) => {
      ws.addCfRule(expression);
      const before = ws.listCfRules()[0];
      ws.updateCfRule(0, {});
      assert.deepEqual(ws.listCfRules()[0], before);
    });
  });

  it('a key the rule type has no use for and an unknown key are TypeErrors', () => {
    withSheet((ws) => {
      ws.addCfRule(expression);
      expectTypeError(() => ws.updateCfRule(0, { operator: f.ComparisonOperator.GreaterThan }));
      expectTypeError(() => ws.updateCfRule(0, { bogus: 1 }));
    });
  });

  it('a type change is InvalidArgument and leaves the rule as it was', () => {
    withSheet((ws) => {
      ws.addCfRule(expression);
      const before = ws.listCfRules()[0];
      expectCode(() => ws.updateCfRule(0, { type: CfRuleType.CellIs }), StatusCode.InvalidArgument);
      assert.deepEqual(ws.listCfRules()[0], before);
    });
  });

  it('a dxf index past the table is IndexOutOfRange', () => {
    withSheet((ws) => {
      ws.addCfRule(expression);
      expectCode(() => ws.updateCfRule(0, { dxfIndex: 50 }), StatusCode.IndexOutOfRange);
    });
  });

  it('an id no rule has is NotFound and a position past the end is IndexOutOfRange', () => {
    withSheet((ws) => {
      ws.addCfRule(expression);
      expectCode(() => ws.updateCfRule('no-such-id', {}), StatusCode.NotFound);
      expectCode(() => ws.updateCfRule(5, {}), StatusCode.IndexOutOfRange);
      expectCode(() => ws.updateCfRule(-1, {}), StatusCode.IndexOutOfRange);
    });
  });

  it('the record the list returns passes back as it is', () => {
    withSheet((ws) => {
      ws.addCfRule({ ...expression, stopIfTrue: true });
      const record = ws.listCfRules()[0];
      ws.updateCfRule(0, record);
      assert.deepEqual(ws.listCfRules()[0], record);
    });
  });

  it('a value rule keeps its payload and updates its operator', () => {
    withSheet((ws) => {
      ws.addCfRule({
        type: CfRuleType.CellIs,
        ranges: 'A1:A5',
        operator: f.ComparisonOperator.GreaterThan,
        formula1: '5',
      });
      ws.updateCfRule(0, { operator: f.ComparisonOperator.LessThan, formula1: '7' });
      const rule = ws.listCfRules()[0];
      assert.equal(rule.operator, f.ComparisonOperator.LessThan);
      assert.equal(rule.formula1, '=7');
    });
  });
});

describe('the ooxml style tables overwrite a record in place', () => {
  it('setFill takes the record getFill returns, changed', () => {
    withSheet((_ws, wb) => {
      const index = wb.ooxml.addFill({ pattern: FillPattern.Solid, foreground: '#FF0000' });
      const record = wb.ooxml.getFill(index);
      wb.ooxml.setFill(index, record);
      assert.deepEqual(wb.ooxml.getFill(index), record);
      wb.ooxml.setFill(index, { ...record, foreground: '#0000FF' });
      const changed = wb.ooxml.getFill(index);
      assert.equal(changed.pattern, FillPattern.Solid);
      assert.equal(changed.foreground.argb, 0xff0000ff);
      assert.equal(wb.ooxml.listFills().length, 3);
    });
  });

  it('setBorder takes the record getBorder returns, changed', () => {
    withSheet((_ws, wb) => {
      const index = wb.ooxml.addBorder({ left: { style: BorderStyle.Thin } });
      const record = wb.ooxml.getBorder(index);
      wb.ooxml.setBorder(index, record);
      assert.deepEqual(wb.ooxml.getBorder(index), record);
      wb.ooxml.setBorder(index, { ...record, left: { style: BorderStyle.None }, bottom: { style: BorderStyle.Thick } });
      const changed = wb.ooxml.getBorder(index);
      assert.equal(changed.left.style, BorderStyle.None);
      assert.equal(changed.bottom.style, BorderStyle.Thick);
    });
  });

  it('setCellXf takes the record getCellXf returns, changed', () => {
    withSheet((_ws, wb) => {
      const fill = wb.ooxml.addFill({ pattern: FillPattern.Solid, foreground: '#FF0000' });
      const index = wb.ooxml.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 0 });
      const record = wb.ooxml.getCellXf(index);
      wb.ooxml.setCellXf(index, record);
      assert.deepEqual(wb.ooxml.getCellXf(index), record);
      wb.ooxml.setCellXf(index, { ...record, fillIndex: fill, applyFill: true, wrapText: true });
      const changed = wb.ooxml.getCellXf(index);
      assert.equal(changed.fillIndex, fill);
      assert.equal(changed.applyFill, true);
      assert.equal(changed.wrapText, true);
    });
  });

  it('setDxf takes the record getDxf returns, changed', () => {
    withSheet((_ws, wb) => {
      const index = wb.ooxml.addDxf({ font: { bold: true } });
      const record = wb.ooxml.getDxf(index);
      wb.ooxml.setDxf(index, record);
      assert.deepEqual(wb.ooxml.getDxf(index), record);
      wb.ooxml.setDxf(index, { font: { italic: true } });
      const changed = wb.ooxml.getDxf(index);
      assert.equal(changed.font.italic, true);
      assert.equal(changed.font.bold, null);
    });
  });

  it('a setter restyles every cell format that names the record', () => {
    withSheet((ws, wb) => {
      const fill = wb.ooxml.addFill({ pattern: FillPattern.Solid, foreground: '#FF0000' });
      const xf = wb.ooxml.addCellXf({ fontIndex: 0, fillIndex: fill, borderIndex: 0, numFmtId: 0, applyFill: true });
      ws.setNumber('A1', 1);
      ws.ooxml.setCellXfIndex('A1', xf);
      assert.equal(ws.ooxml.getEffectiveStyle('A1').fillForeground.argb, 0xffff0000);
      wb.ooxml.setFill(fill, { pattern: FillPattern.Solid, foreground: '#00FF00' });
      assert.equal(ws.ooxml.getEffectiveStyle('A1').fillForeground.argb, 0xff00ff00);
    });
  });

  it('an index past the end is IndexOutOfRange and a negative one too', () => {
    withSheet((_ws, wb) => {
      const { ooxml } = wb;
      expectCode(() => ooxml.setFill(999, ooxml.getFill(0)), StatusCode.IndexOutOfRange);
      expectCode(() => ooxml.setBorder(999, ooxml.getBorder(0)), StatusCode.IndexOutOfRange);
      expectCode(() => ooxml.setCellXf(9999, ooxml.getCellXf(0)), StatusCode.IndexOutOfRange);
      expectCode(() => ooxml.setDxf(99, { font: { bold: true } }), StatusCode.IndexOutOfRange);
      expectCode(() => ooxml.setFill(-1, ooxml.getFill(0)), StatusCode.IndexOutOfRange);
      expectCode(() => ooxml.setBorder(-1, ooxml.getBorder(0)), StatusCode.IndexOutOfRange);
      expectCode(() => ooxml.setCellXf(-1, ooxml.getCellXf(0)), StatusCode.IndexOutOfRange);
      expectCode(() => ooxml.setDxf(-1, { font: { bold: true } }), StatusCode.IndexOutOfRange);
    });
  });

  it('a cell format naming a font, fill or border the tables lack is IndexOutOfRange', () => {
    withSheet((_ws, wb) => {
      const index = wb.ooxml.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 0 });
      const record = wb.ooxml.getCellXf(index);
      expectCode(() => wb.ooxml.setCellXf(index, { ...record, fillIndex: 99 }), StatusCode.IndexOutOfRange);
      assert.deepEqual(wb.ooxml.getCellXf(index), record);
    });
  });
});

describe('theme script fonts are listed', () => {
  it('a font set written with setThemeScriptFont appears in its own list only', () => {
    withSheet((_ws, wb) => {
      wb.setThemeScriptFont({ scheme: FontScheme.Major, script: 'Zzzz', typeface: 'Test Face' });
      const major = wb.listThemeScriptFonts(FontScheme.Major);
      const minor = wb.listThemeScriptFonts(FontScheme.Minor);
      assert.deepEqual(
        major.filter((x) => x.script === 'Zzzz'),
        [{ scheme: FontScheme.Major, script: 'Zzzz', typeface: 'Test Face' }],
      );
      assert.equal(
        minor.some((x) => x.script === 'Zzzz'),
        false,
      );
    });
  });

  it('every listed element is what getThemeScriptFont reads for its key', () => {
    withSheet((_ws, wb) => {
      const list = wb.listThemeScriptFonts(FontScheme.Minor);
      assert.ok(list.length > 0);
      for (const entry of list) {
        assert.equal(entry.scheme, FontScheme.Minor);
        assert.deepEqual(wb.getThemeScriptFont({ scheme: entry.scheme, script: entry.script }), entry);
      }
    });
  });

  it('FontScheme.None is InvalidArgument', () => {
    withSheet((_ws, wb) => {
      expectCode(() => wb.listThemeScriptFonts(FontScheme.None), StatusCode.InvalidArgument);
    });
  });
});

describe('the update, setter and list members are present', () => {
  it('Sheet, WorkbookOoxml and Workbook carry them', () => {
    withSheet((ws, wb) => {
      for (const name of ['updateValidation', 'updateCfRule']) assert.equal(typeof ws[name], 'function', name);
      for (const name of ['setFill', 'setBorder', 'setCellXf', 'setDxf']) {
        assert.equal(typeof wb.ooxml[name], 'function', name);
      }
      assert.equal(typeof wb.listThemeScriptFonts, 'function');
    });
  });
});
