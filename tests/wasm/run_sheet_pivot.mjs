import assert from 'node:assert/strict';
import { buildPivotWorkbookBytes, findPivotCell, PIVOT, VAL } from './smoke_support.mjs';

export function registerSheetPivot(Module, test) {
  test('setIterative accepts knobs without crashing', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setIterative(true, 50, 0.001).ok);
      // Set a fixed-point formula (sqrt(2) iteration). The engine may
      // surface NUMBER (converged) or ERROR depending on the recalc
      // path; we only assert that the call sequence doesn't crash and
      // that getValue returns a well-formed envelope.
      assert.ok(wb.setNumber(0, 0, 0, 1.0).ok);
      assert.ok(wb.setFormula(0, 0, 0, '=0.5*(A1+2/A1)').ok);
      assert.ok(wb.recalc().ok);
      const v = wb.getValue(0, 0, 0);
      assert.ok(v.status.ok);
      // Known well-formed kinds.
      assert.ok([VAL.NUMBER, VAL.ERROR, VAL.BLANK].includes(v.value.kind));
    } finally {
      wb.delete();
    }
  });

  test('renameSheet updates the sheet name', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.equal(wb.sheetCount().value, 1);
      assert.ok(wb.renameSheet(0, 'Renamed').ok);
      const r = wb.sheetName(0);
      assert.ok(r.status.ok);
      assert.equal(r.value, 'Renamed');
    } finally {
      wb.delete();
    }
  });

  test('renameSheet rejects forbidden characters', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const r = wb.renameSheet(0, 'Bad/Name');
      assert.equal(r.ok, false);
      assert.notEqual(r.status, 0);
    } finally {
      wb.delete();
    }
  });

  test('removeSheet drops a sheet but rejects the last one', () => {
    const wb = Module.Workbook.createEmpty();
    try {
      assert.ok(wb.addSheet('A').ok);
      assert.ok(wb.addSheet('B').ok);
      assert.ok(wb.addSheet('C').ok);
      assert.equal(wb.sheetCount().value, 3);
      assert.ok(wb.removeSheet(1).ok);
      assert.equal(wb.sheetCount().value, 2);
      assert.equal(wb.sheetName(0).value, 'A');
      assert.equal(wb.sheetName(1).value, 'C');
    } finally {
      wb.delete();
    }

    // Workbook with a single sheet rejects removeSheet(0).
    const lone = Module.Workbook.createDefault();
    try {
      const r = lone.removeSheet(0);
      assert.equal(r.ok, false);
    } finally {
      lone.delete();
    }
  });

  test('moveSheet rearranges sheets (Excel UI semantics)', () => {
    const wb = Module.Workbook.createEmpty();
    try {
      assert.ok(wb.addSheet('Alpha').ok);
      assert.ok(wb.addSheet('Beta').ok);
      assert.ok(wb.addSheet('Gamma').ok);
      // Move Alpha (0) to the end. Excel semantics: to=2 (post-removal).
      assert.ok(wb.moveSheet(0, 2).ok);
      assert.equal(wb.sheetName(0).value, 'Beta');
      assert.equal(wb.sheetName(1).value, 'Gamma');
      assert.equal(wb.sheetName(2).value, 'Alpha');
    } finally {
      wb.delete();
    }
  });

  test('setDefinedName adds, updates, and removes', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.equal(wb.definedNameCount().value, 0);
      assert.ok(wb.setDefinedName('Pi', '=3.14').ok);
      assert.equal(wb.definedNameCount().value, 1);
      const a = wb.definedNameAt(0);
      assert.ok(a.status.ok);
      assert.equal(a.name, 'Pi');
      assert.equal(a.formula, '=3.14');

      assert.ok(wb.setDefinedName('PI', '=3.14159').ok);
      assert.equal(wb.definedNameCount().value, 1);
      const b = wb.definedNameAt(0);
      assert.equal(b.name, 'Pi'); // authored case preserved
      assert.equal(b.formula, '=3.14159');

      assert.ok(wb.setDefinedName('Pi', '').ok);
      assert.equal(wb.definedNameCount().value, 0);
    } finally {
      wb.delete();
    }
  });

  test('cellCount + cellAt iterate populated cells in order', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setNumber(0, 0, 0, 1).ok);
      assert.ok(wb.setNumber(0, 0, 1, 2).ok);
      assert.ok(wb.setFormula(0, 1, 0, '=A1+B1').ok);
      assert.ok(wb.recalc().ok);

      const count = wb.cellCount(0).value;
      assert.ok(count >= 3, `expected >=3 cells, got ${count}`);

      // The very first iteration entry should be A1 (row=0, col=0).
      const c0 = wb.cellAt(0, 0);
      assert.ok(c0.status.ok);
      assert.equal(c0.row, 0);
      assert.equal(c0.col, 0);
      assert.equal(c0.formula, null);
      assert.equal(c0.value.kind, VAL.NUMBER);
      assert.equal(c0.value.number, 1);
    } finally {
      wb.delete();
    }
  });

  test('failed entry lookups omit optional payload fields', () => {
    const wb = Module.Workbook.createDefault();
    try {
      for (const entry of [wb.cellAt(99, 0), wb.definedNameAt(0), wb.tableAt(0), wb.passthroughAt(0)]) {
        assert.equal(entry.status.ok, false);
      }
      assert.ok(!('value' in wb.cellAt(99, 0)));
      assert.ok(!('name' in wb.definedNameAt(0)));
      assert.ok(!('name' in wb.tableAt(0)));
      assert.ok(!('path' in wb.passthroughAt(0)));
    } finally {
      wb.delete();
    }
  });

  test('pivotCount + pivotLayout expose PivotTable projection status', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.equal(wb.pivotCount(0).value, 0);

      const missing = wb.pivotLayout(0, 0);
      assert.equal(missing.status.ok, false);
      assert.notEqual(missing.status.status, 0);
      assert.equal(missing.top, 0);
      assert.equal(missing.left, 0);
      assert.equal(missing.rows, 0);
      assert.equal(missing.cols, 0);
      assert.deepEqual(missing.cells, []);
    } finally {
      wb.delete();
    }
  });

  test('pivotLayout projects loaded PivotTable cells for grid rendering', () => {
    const wb = Module.Workbook.loadBytes(buildPivotWorkbookBytes());
    try {
      assert.ok(wb.isValid(), Module.lastErrorMessage());
      assert.equal(wb.pivotCount(0).value, 0);
      assert.equal(wb.pivotCount(1).value, 1);

      const layout = wb.pivotLayout(1, 0);
      assert.ok(layout.status.ok, `status=${JSON.stringify(layout.status)}`);
      assert.equal(layout.top, 0);
      assert.equal(layout.left, 3);
      assert.equal(layout.rows, 4);
      assert.equal(layout.cols, 2);
      assert.ok(layout.cells.length > 0);

      const rowLabels = findPivotCell(layout, 0, 3);
      assert.ok(rowLabels);
      assert.equal(rowLabels.kind, PIVOT.HEADER);
      assert.equal(rowLabels.value.kind, VAL.TEXT);
      assert.equal(rowLabels.value.text, '行ラベル');

      const northLabel = findPivotCell(layout, 1, 3);
      assert.ok(northLabel);
      assert.equal(northLabel.kind, PIVOT.ROW_LABEL);
      assert.equal(northLabel.value.text, 'North');

      const northSum = findPivotCell(layout, 1, 4);
      assert.ok(northSum);
      assert.equal(northSum.kind, PIVOT.DATA);
      assert.equal(northSum.value.kind, VAL.NUMBER);
      assert.equal(northSum.value.number, 400);
      assert.equal(northSum.fieldName, 'Sum of Amount');

      const grandTotal = findPivotCell(layout, 3, 4);
      assert.ok(grandTotal);
      assert.equal(grandTotal.kind, PIVOT.GRAND_TOTAL);
      assert.equal(grandTotal.value.kind, VAL.NUMBER);
      assert.equal(grandTotal.value.number, 600);
    } finally {
      wb.delete();
    }
  });
}
