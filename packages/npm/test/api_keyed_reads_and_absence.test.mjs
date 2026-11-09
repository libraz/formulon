// Keyed reads and absence on the public entry: workbook protection reads null when absent, a keyed read of a rule,
// validation, protected range or sparkline group returns null for a missing key, and an empty string is never a clear.
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
const { CfRuleType, StatusCode, ValidationType } = f;

/** Runs `body(ws, wb)` on a default workbook holding a small grid, then disposes it. */
function withSheet(body) {
  const wb = f.Workbook.createDefault();
  try {
    const ws = wb.sheet(0);
    ws.setText('A1', 'Region');
    ws.setText('B1', 'Amount');
    ws.setText('A2', 'East');
    ws.setNumber('B2', 10);
    return body(ws, wb);
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

const ABSENT_ID = '{00000000-0000-4000-8000-0000000000FF}';
const PERSON = '{00000000-0000-4000-8000-0000000000A1}';

describe('a workbook without protection reads null', () => {
  it('getWorkbookProtection is null on a fresh workbook', () => {
    withSheet((_ws, wb) => {
      assert.equal(wb.getWorkbookProtection(), null);
    });
  });

  it('a protection with every lock off is present and null removes it', () => {
    withSheet((_ws, wb) => {
      wb.setWorkbookProtection({});
      const protection = wb.getWorkbookProtection();
      assert.notEqual(protection, null);
      assert.equal(protection.lockStructure, false);
      wb.setWorkbookProtection(protection);
      assert.notEqual(wb.getWorkbookProtection(), null);
      wb.setWorkbookProtection(null);
      assert.equal(wb.getWorkbookProtection(), null);
    });
  });
});

describe('a keyed read returns the record or null', () => {
  it('getCfRule by id returns the record and null when absent', () => {
    withSheet((ws) => {
      const id = ws.addCfRule({ type: CfRuleType.Expression, ranges: 'A1:B2', formula1: 'A1>1' });
      assert.deepEqual(ws.getCfRule(id), ws.listCfRules()[0]);
      assert.equal(ws.getCfRule(ABSENT_ID), null);
    });
  });

  it('getValidation by id returns the record and null when absent', () => {
    withSheet((ws) => {
      const id = ws.addValidation({ type: ValidationType.Whole, ranges: 'D1:D3', formula1: '1', formula2: '5' });
      assert.deepEqual(ws.getValidation(id), ws.listValidations()[0]);
      assert.equal(ws.getValidation(ABSENT_ID), null);
    });
  });

  it('getProtectedRange by name returns the record and null when absent', () => {
    withSheet((ws) => {
      ws.addProtectedRange({ name: 'Totals', ranges: ['A1:B2'] });
      assert.deepEqual(ws.getProtectedRange('Totals'), ws.listProtectedRanges()[0]);
      assert.equal(ws.getProtectedRange('Missing'), null);
    });
  });

  it('getSparklineGroup by id returns the record and null when absent', () => {
    withSheet((ws) => {
      const id = ws.addSparklineGroup({ sparklines: [{ source: 'Sheet1!B2:B2', location: 'D1' }] });
      assert.deepEqual(ws.getSparklineGroup(id), ws.listSparklineGroups()[0]);
      assert.equal(ws.getSparklineGroup(ABSENT_ID), null);
    });
  });
});

describe('an empty string is never a clear', () => {
  it('empty style name is InvalidArgument on addTable, updateStyleOptions, addSlicer, updateTableStyleDefaults', () => {
    withSheet((ws, wb) => {
      expectCode(() => ws.addTable({ range: 'A1:B2', name: 'Bad', styleName: '' }), StatusCode.InvalidArgument);
      const table = ws.addTable({ range: 'A1:B2', name: 'Tbl' });
      expectCode(() => table.updateStyleOptions({ styleName: '' }), StatusCode.InvalidArgument);
      expectCode(
        () => ws.addSlicer(table, 0, { anchor: { fromRow: 8, fromCol: 3, toRow: 16, toCol: 5 }, styleName: '' }),
        StatusCode.InvalidArgument,
      );
      expectCode(() => wb.updateTableStyleDefaults({ tableStyleName: '' }), StatusCode.InvalidArgument);
      expectCode(() => wb.updateTableStyleDefaults({ pivotStyleName: '' }), StatusCode.InvalidArgument);
      table.updateStyleOptions({ styleName: null });
      assert.equal(table.getStyleOptions().styleName, null);
    });
  });

  it('empty totals label and parentId are InvalidArgument', () => {
    withSheet((ws, wb) => {
      const table = ws.addTable({ range: 'A1:B2', name: 'Tbl' });
      expectCode(() => table.setColumnTotals(0, { totalsRowLabel: '' }), StatusCode.InvalidArgument);
      expectCode(() => table.setColumnTotals(1, { totalsRowFormula: '' }), StatusCode.InvalidArgument);
      wb.addPerson({ id: PERSON, displayName: 'Alice' });
      expectCode(
        () => ws.addThreadedComment({ cell: 'D4', personId: PERSON, text: 'Hi', parentId: '' }),
        StatusCode.InvalidArgument,
      );
      assert.equal(ws.listThreadedComments().length, 0);
    });
  });

  it('styled false creates an unstyled table', () => {
    withSheet((ws) => {
      const plain = ws.addTable({ range: 'A1:B2', name: 'Plain', styled: false });
      assert.equal(plain.getStyleOptions(), null);
      const styled = ws.addTable({ range: 'D1:E2', name: 'Styled' });
      assert.notEqual(styled.getStyleOptions(), null);
      expectCode(
        () => ws.addTable({ range: 'G1:H2', name: 'Both', styled: false, styleName: 'TableStyleLight1' }),
        StatusCode.InvalidArgument,
      );
    });
  });
});
