import assert from 'node:assert/strict';
import { StatusCode } from '../../packages/npm/common.mjs';
import { buildPivotWorkbookBytes, expectFailure, findPivotCell, PIVOT, VAL } from './smoke_support.mjs';

export function registerSheetPivot(Module, test) {
  test('setIterative accepts knobs without crashing', () => {
    const wb = Module.Workbook.createDefault();
    try {
      wb.updateIterative({ enabled: true, maxIterations: 50, maxChange: 0.001 });
      // Set a fixed-point formula (sqrt(2) iteration). The engine may
      // surface NUMBER (converged) or ERROR depending on the recalc
      // path; we only assert that the call sequence doesn't crash and
      // that getValue returns a well-formed value.
      wb.setNumber(0, 0, 0, 1.0);
      wb.setFormula(0, 0, 0, '=0.5*(A1+2/A1)');
      wb.recalc();
      const v = wb.getValue(0, 0, 0);
      // Known well-formed kinds.
      assert.ok([VAL.NUMBER, VAL.ERROR, VAL.BLANK].includes(v.kind));
    } finally {
      wb.delete();
    }
  });

  test('renameSheet updates the sheet name', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.equal(wb.getSheetCount(), 1);
      wb.renameSheet(0, 'Renamed');
      const r = wb.getSheetName(0);
      assert.equal(r, 'Renamed');
    } finally {
      wb.delete();
    }
  });

  test('renameSheet rejects forbidden characters', () => {
    const wb = Module.Workbook.createDefault();
    try {
      expectFailure(() => wb.renameSheet(0, 'Bad/Name'));
    } finally {
      wb.delete();
    }
  });

  test('removeSheet drops a sheet but rejects the last one', () => {
    const wb = Module.Workbook.createEmpty();
    try {
      wb.addSheet('A', wb.getSheetCount());
      wb.addSheet('B', wb.getSheetCount());
      wb.addSheet('C', wb.getSheetCount());
      assert.equal(wb.getSheetCount(), 3);
      wb.removeSheet(1);
      assert.equal(wb.getSheetCount(), 2);
      assert.equal(wb.getSheetName(0), 'A');
      assert.equal(wb.getSheetName(1), 'C');
    } finally {
      wb.delete();
    }

    // Workbook with a single sheet rejects removeSheet(0).
    const lone = Module.Workbook.createDefault();
    try {
      expectFailure(() => lone.removeSheet(0));
    } finally {
      lone.delete();
    }
  });

  test('moveSheet places a sheet before the sheet at the given position', () => {
    const wb = Module.Workbook.createEmpty();
    try {
      wb.addSheet('Alpha', wb.getSheetCount());
      wb.addSheet('Beta', wb.getSheetCount());
      wb.addSheet('Gamma', wb.getSheetCount());
      // The position is a slot of the current list: 3 is past the last sheet.
      wb.moveSheet(0, 3);
      assert.equal(wb.getSheetName(0), 'Beta');
      assert.equal(wb.getSheetName(1), 'Gamma');
      assert.equal(wb.getSheetName(2), 'Alpha');
      // Gamma goes before Beta; moving a sheet before itself or its successor changes nothing.
      wb.moveSheet(1, 0);
      wb.moveSheet(0, 0);
      wb.moveSheet(0, 1);
      assert.deepEqual(
        [0, 1, 2].map((i) => wb.getSheetName(i)),
        ['Gamma', 'Beta', 'Alpha'],
      );
    } finally {
      wb.delete();
    }
  });

  test('addSheet inserts before a position and names a null name from the profile', () => {
    const wb = Module.Workbook.createForProfile('win-365-de_DE');
    try {
      assert.equal(wb.addSheet(null, wb.getSheetCount()), 1);
      assert.equal(wb.getSheetName(1), 'Tabelle2');
      // The default name is the smallest free number, compared case-insensitively.
      wb.addSheet('tabelle3', wb.getSheetCount());
      assert.equal(wb.addSheet(undefined, wb.getSheetCount()), 3);
      assert.equal(wb.getSheetName(3), 'Tabelle4');

      assert.equal(wb.addSheet('Front', 0), 0);
      assert.equal(wb.getSheetName(0), 'Front');
      assert.equal(wb.getSheetName(1), 'Tabelle1');
      assert.equal(wb.addSheet('Middle', 2), 2);
      assert.deepEqual(
        [0, 1, 2, 3].map((i) => wb.getSheetName(i)),
        ['Front', 'Tabelle1', 'Middle', 'Tabelle2'],
      );

      expectFailure(() => wb.addSheet('', wb.getSheetCount()), StatusCode.InvalidSheetName);
      expectFailure(() => wb.addSheet('MIDDLE', wb.getSheetCount()), StatusCode.AlreadyExists);
      expectFailure(() => wb.addSheet('Past', wb.getSheetCount() + 1), StatusCode.SheetIndexOutOfRange);
      assert.throws(() => wb.addSheet(1, wb.getSheetCount()), TypeError);
    } finally {
      wb.delete();
    }
  });

  test('copySheet places the copy before the given position', () => {
    const wb = Module.Workbook.createEmpty();
    try {
      wb.addSheet('A', wb.getSheetCount());
      wb.addSheet('B', wb.getSheetCount());
      assert.equal(wb.copySheet(1, 0, 'First'), 0);
      assert.equal(wb.copySheet(0, wb.getSheetCount(), 'Last'), 3);
      assert.deepEqual(
        [0, 1, 2, 3].map((i) => wb.getSheetName(i)),
        ['First', 'A', 'B', 'Last'],
      );
    } finally {
      wb.delete();
    }
  });

  test('insertRows / deleteRows / insertColumns / deleteColumns take an inclusive span', () => {
    const wb = Module.Workbook.createDefault();
    try {
      wb.setNumber(0, 2, 2, 7);
      wb.insertRows(0, 2, 4);
      assert.equal(wb.getValue(0, 5, 2).number, 7);
      wb.deleteRows(0, 2, 4);
      assert.equal(wb.getValue(0, 2, 2).number, 7);
      wb.insertColumns(0, 2, 3);
      assert.equal(wb.getValue(0, 2, 4).number, 7);
      wb.deleteColumns(0, 2, 3);
      assert.equal(wb.getValue(0, 2, 2).number, 7);

      // A single row or column is the span `[n, n]`.
      wb.insertRows(0, 2, 2);
      assert.equal(wb.getValue(0, 3, 2).number, 7);
      for (const edit of [wb.insertRows, wb.deleteRows, wb.insertColumns, wb.deleteColumns]) {
        expectFailure(() => edit.call(wb, 0, 5, 3), StatusCode.InvalidArgument);
      }
      expectFailure(() => wb.insertRows(0, 0, 1048576), StatusCode.InvalidArgument);
    } finally {
      wb.delete();
    }
  });

  test('setDefinedName adds and updates, and removeDefinedName removes', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.equal(wb.listDefinedNames().length, 0);
      wb.setDefinedName({ name: 'Pi', formula: '=3.14' });
      const listed = wb.listDefinedNames();
      assert.equal(listed.length, 1);
      const a = listed[0];
      assert.equal(a.name, 'Pi');
      // The formula reads back with one leading '=', whether or not the input had it.
      assert.equal(a.formula, '=3.14');

      wb.setDefinedName({ name: 'PI', formula: '3.14159' });
      assert.equal(wb.listDefinedNames().length, 1);
      const b = wb.listDefinedNames()[0];
      assert.equal(b.name, 'PI'); // the update takes the new spelling
      assert.equal(b.formula, '=3.14159');

      // An empty formula is refused; removal has its own call.
      expectFailure(() => wb.setDefinedName({ name: 'Pi', formula: '' }), StatusCode.InvalidArgument);
      assert.equal(wb.listDefinedNames().length, 1);
      wb.removeDefinedName('Pi', -1);
      assert.equal(wb.listDefinedNames().length, 0);
    } finally {
      wb.delete();
    }
  });

  test('cellCount + cellAt iterate populated cells in order', () => {
    const wb = Module.Workbook.createDefault();
    try {
      wb.setNumber(0, 0, 0, 1);
      wb.setNumber(0, 0, 1, 2);
      wb.setFormula(0, 1, 0, '=A1+B1');
      wb.recalc();

      const count = wb.getCellCount(0);
      assert.ok(count >= 3, `expected >=3 cells, got ${count}`);

      // The very first iteration entry should be A1 (row=0, col=0).
      const c0 = wb.getCellAt(0, 0);
      assert.equal(c0.row, 0);
      assert.equal(c0.col, 0);
      assert.equal(c0.formula, null);
      assert.equal(c0.value.kind, VAL.NUMBER);
      assert.equal(c0.value.number, 1);
    } finally {
      wb.delete();
    }
  });

  test('failed entry lookups throw', () => {
    const wb = Module.Workbook.createDefault();
    try {
      expectFailure(() => wb.getCellAt(99, 0), StatusCode.SheetIndexOutOfRange);
      expectFailure(() => wb.getTable(0, 0), StatusCode.IndexOutOfRange);
      expectFailure(() => wb.getTable(9, 0), StatusCode.SheetIndexOutOfRange);
    } finally {
      wb.delete();
    }
  });

  test('pivotCount + pivotLayout expose PivotTable projection status', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.equal(wb.getPivotCount(0), 0);

      expectFailure(() => wb.getPivotLayout(0, 0), StatusCode.IndexOutOfRange);
    } finally {
      wb.delete();
    }
  });

  test('pivotLayout projects loaded PivotTable cells for grid rendering', () => {
    const wb = Module.Workbook.load(buildPivotWorkbookBytes());
    try {
      assert.equal(wb.disposed, false);
      assert.equal(wb.getPivotCount(0), 0);
      assert.equal(wb.getPivotCount(1), 1);

      const layout = wb.getPivotLayout(1, 0);
      assert.deepEqual(layout.rect, { firstRow: 0, firstCol: 3, lastRow: 3, lastCol: 4 });
      assert.ok(layout.cells.length > 0);

      const rowLabels = findPivotCell(layout, 0, 3);
      assert.ok(rowLabels);
      assert.equal(rowLabels.kind, PIVOT.HEADER);
      assert.equal(rowLabels.value.kind, VAL.TEXT);
      assert.equal(rowLabels.value.text, 'Row Labels');

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
