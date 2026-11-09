import assert from 'node:assert/strict';
import { StatusCode } from '../../packages/npm/common.mjs';
import { expectFailure, PIVOT, VAL } from './smoke_support.mjs';

// The error a call fails with.
function caught(fn) {
  try {
    fn();
  } catch (e) {
    return e;
  }
  return assert.fail('expected the call to throw');
}

const nestedErrorClass = (value) => (typeof value === 'symbol' ? TypeError : RangeError);
// A required whole-number field: a value that is no whole number is a TypeError, one past uint32 a RangeError.
const requiredErrorClass = (value) => (typeof value === 'number' && Number.isInteger(value) ? RangeError : TypeError);

function makeCheckedPivot(Module) {
  const wb = Module.Workbook.createDefault();
  const cache = wb.addPivotCache(0);
  wb.setPivotCacheWorksheetSource(cache, { ref: 'A1:B3', sheet: 'Sheet1' });
  wb.addPivotCacheField(cache, 'Region');
  wb.addPivotCacheField(cache, 'Amount');
  for (const [region, amount] of [
    ['East', 10],
    ['West', 30],
  ]) {
    const record = wb.addPivotCacheRecord(cache);
    wb.setPivotCacheRecordText(cache, record, 0, region);
    wb.setPivotCacheRecordNumber(cache, record, 1, amount);
  }
  const pivot = wb.addPivot(0, 'Pivot1', cache, 0, 4);
  const region = 0;
  const amount = 1;
  wb.setPivotFieldAxis(0, pivot, region, 0);
  wb.setPivotFieldAxis(0, pivot, amount, 2);
  wb.setPivotRowFieldOrder(0, pivot, [region]);
  wb.addPivotDataField(0, pivot, { name: 'Sum of Amount', fieldIndex: amount, aggregation: 0 });
  return { wb, pivot, cache, region, amount };
}

function checkedPivotGrid(wb, pivot) {
  const layout = wb.getPivotLayout(0, pivot);
  const rows = new Map();
  for (const cell of layout.cells) {
    const value = cell.value;
    const text = value.kind === VAL.NUMBER ? String(value.number) : value.kind === VAL.TEXT ? value.text : '';
    if (!rows.has(cell.row)) rows.set(cell.row, new Array(layout.rect.lastCol - layout.rect.firstCol + 1).fill(''));
    rows.get(cell.row)[cell.col - layout.rect.firstCol] = text;
  }
  return [...rows.keys()].sort((a, b) => a - b).map((row) => rows.get(row).join('|'));
}

const CHECKED_PIVOT_GRID = ['Row Labels|Sum of Amount', 'East|10', 'West|30', 'Grand Total|40'];

export function registerPivotRegressions(Module, test) {
  test('page field order moves listed fields onto the page axis and refuses the Values level', () => {
    const { wb, pivot, region } = makeCheckedPivot(Module);
    try {
      assert.deepEqual(wb.getPivotPageFieldOrder(0, pivot), []);
      wb.setPivotPageFieldOrder(0, pivot, [region]);
      assert.deepEqual(wb.getPivotPageFieldOrder(0, pivot), [region]);
      assert.deepEqual(wb.getPivotRowFieldOrder(0, pivot), []);
      expectFailure(() => wb.setPivotPageFieldOrder(0, pivot, [-2]), StatusCode.InvalidArgument);
      expectFailure(() => wb.setPivotPageFieldOrder(0, pivot, [region, region]), StatusCode.InvalidArgument);
      expectFailure(() => wb.setPivotPageFieldOrder(0, pivot, [99]), StatusCode.IndexOutOfRange);
      assert.deepEqual(wb.getPivotPageFieldOrder(0, pivot), [region]);
    } finally {
      wb.delete();
    }
  });

  test('setPivotFieldDateGroup accepts null and leaves an ungrouped field alone', () => {
    const { wb, pivot, region } = makeCheckedPivot(Module);
    try {
      wb.setPivotFieldDateGroup(0, pivot, region, null);
      assert.equal(wb.getPivotField(0, pivot, region).datePeriods, 0);
      assert.equal('clearPivotFieldDateGroup' in wb, false);
    } finally {
      wb.delete();
    }
  });

  test('addPivotFieldItem by cache index hides the blank item an empty label cannot name', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const cache = wb.addPivotCache(0);
      const region = wb.addPivotCacheField(cache, 'Region');
      const amount = wb.addPivotCacheField(cache, 'Amount');
      // Shared item 0 renders as "North"; shared item 1 is the blank, which
      // carries no label of its own and can only be named by its index.
      const north = wb.addPivotCacheFieldSharedItemText(cache, region, 'North');
      const blank = wb.addPivotCacheFieldSharedItemBlank(cache, region);
      assert.equal(blank, north + 1);

      const first = wb.addPivotCacheRecord(cache);
      wb.setPivotCacheRecordNumber(cache, first, region, 0);
      wb.setPivotCacheRecordNumber(cache, first, amount, 100);
      const second = wb.addPivotCacheRecord(cache);
      wb.setPivotCacheRecordBlank(cache, second, region);
      wb.setPivotCacheRecordNumber(cache, second, amount, 200);

      const pivot = wb.addPivot(0, 'PT', cache, 0, 0);
      const regionField = 0;
      const amountField = 1;
      wb.setPivotFieldAxis(0, pivot, regionField, 0);
      wb.setPivotFieldAxis(0, pivot, amountField, 2);
      wb.setPivotRowFieldOrder(0, pivot, [regionField]);
      wb.addPivotDataField(0, pivot, {
        name: 'Sum of Amount',
        fieldIndex: amountField,
        aggregation: 0,
      });

      const sumDataCells = () => {
        const layout = wb.getPivotLayout(0, pivot);
        let total = 0;
        for (const cell of layout.cells) {
          if (cell.kind === PIVOT.DATA && cell.value.kind === VAL.NUMBER) {
            total += cell.value.number;
          }
        }
        return total;
      };
      assert.equal(sumDataCells(), 300);

      wb.addPivotFieldItem(0, pivot, regionField, 'North', true);
      wb.addPivotFieldItem(0, pivot, regionField, { cacheIndex: 1 }, false);
      assert.equal(sumDataCells(), 100);
    } finally {
      wb.delete();
    }
  });

  test('pivot nested integers throw on narrowing before C mutation', () => {
    const invalid = [2 ** 32, 1.5, Number.NaN, Number.POSITIVE_INFINITY, Symbol('pivot numeric')];
    for (const value of invalid) {
      const data = makeCheckedPivot(Module);
      try {
        const before = data.wb.getPivotDataFieldCount(0, data.pivot);
        assert.throws(
          () =>
            data.wb.addPivotDataField(0, data.pivot, {
              name: 'Bad',
              fieldIndex: value,
              aggregation: 0,
            }),
          requiredErrorClass(value),
          `pivotDataFieldAdd fieldIndex=${String(value)} accepted`,
        );
        assert.equal(data.wb.getPivotDataFieldCount(0, data.pivot), before);
      } finally {
        data.wb.delete();
      }
    }
  });

  test('pivot filters throw on invalid field index, value kind, and data field index before mutation', () => {
    const invalid = [{ fieldIndex: 2 ** 32 }, { value: Symbol('pivot filter') }, { dataFieldIndex: 2 ** 32 }];
    for (const override of invalid) {
      const { wb, pivot, region } = makeCheckedPivot(Module);
      try {
        const before = wb.getPivotFilterCount(0, pivot);
        assert.throws(
          () =>
            wb.addPivotFilter(0, pivot, {
              fieldIndex: region,
              kind: 1,
              value: 15,
              ...override,
            }),
          nestedErrorClass(Object.values(override)[0]),
          `pivotFilterAdd ${String(Object.values(override)[0])} accepted`,
        );
        assert.equal(wb.getPivotFilterCount(0, pivot), before);
        assert.deepEqual(checkedPivotGrid(wb, pivot), CHECKED_PIVOT_GRID);
      } finally {
        wb.delete();
      }
    }
  });

  test('pivot filter operands of the wrong type throw a TypeError before mutation', () => {
    for (const override of [{ value: {} }, { value: true }, { valueHigh: [] }]) {
      const { wb, pivot, region } = makeCheckedPivot(Module);
      try {
        const before = wb.getPivotFilterCount(0, pivot);
        assert.throws(
          () => wb.addPivotFilter(0, pivot, { fieldIndex: region, kind: 1, value: 15, ...override }),
          TypeError,
          `pivotFilterAdd ${JSON.stringify(override)} accepted`,
        );
        assert.equal(wb.getPivotFilterCount(0, pivot), before);
      } finally {
        wb.delete();
      }
    }
  });

  test('a pivot filter reads back its operands as a number, a string or null', () => {
    const { wb, pivot, region } = makeCheckedPivot(Module);
    try {
      const index = wb.addPivotFilter(0, pivot, { fieldIndex: region, kind: 1, value: 15.5 });
      const stored = wb.getPivotFilter(0, pivot, index);
      assert.equal(stored.fieldIndex, region);
      assert.equal(stored.value, 15.5);
      assert.equal(stored.valueHigh, null);
    } finally {
      wb.delete();
    }
  });

  test('getPivotCacheRecordValue reads each record cell back as a tagged value', () => {
    const { wb, cache, region, amount } = makeCheckedPivot(Module);
    try {
      assert.deepEqual({ ...wb.getPivotCacheRecordValue(cache, 0, region).value }, { kind: VAL.TEXT, text: 'East' });
      assert.deepEqual({ ...wb.getPivotCacheRecordValue(cache, 1, amount).value }, { kind: VAL.NUMBER, number: 30 });

      const extra = wb.addPivotCacheRecord(cache);
      wb.setPivotCacheRecordBool(cache, extra, region, true);
      wb.setPivotCacheRecordBlank(cache, extra, amount);
      assert.deepEqual(
        { ...wb.getPivotCacheRecordValue(cache, extra, region).value },
        { kind: VAL.BOOL, boolean: true },
      );
      assert.equal(wb.getPivotCacheRecordValue(cache, extra, amount).value.kind, VAL.BLANK);
      wb.setPivotCacheRecordDate(cache, extra, amount, 45000);
      assert.deepEqual(
        { ...wb.getPivotCacheRecordValue(cache, extra, amount).value },
        { kind: VAL.NUMBER, number: 45000 },
      );

      const count = wb.getPivotCacheRecordCount(cache);
      expectFailure(() => wb.getPivotCacheRecordValue(cache, count, region), StatusCode.IndexOutOfRange);
      expectFailure(() => wb.getPivotCacheRecordValue(cache, 0, 2), StatusCode.IndexOutOfRange);
    } finally {
      wb.delete();
    }
  });

  test('pivot order arrays throw on invalid elements before replacing the order', () => {
    const invalid = [2 ** 32 + 1, 1.5, Number.NaN, Number.POSITIVE_INFINITY, Symbol('pivot order')];
    for (const value of invalid) {
      const { wb, pivot } = makeCheckedPivot(Module);
      try {
        assert.throws(
          () => wb.setPivotRowFieldOrder(0, pivot, [value]),
          nestedErrorClass(value),
          `pivotSetRowFieldOrder element=${String(value)} accepted`,
        );
        assert.deepEqual(checkedPivotGrid(wb, pivot), CHECKED_PIVOT_GRID);
      } finally {
        wb.delete();
      }
    }
  });

  test('pivot valid nested controls and nullish defaults remain accepted', () => {
    const { wb, pivot, region, amount } = makeCheckedPivot(Module);
    try {
      wb.addPivotDataField(0, pivot, {
        name: 'Count of Amount',
        fieldIndex: amount,
        aggregation: 1,
        numberFormat: null,
        showAs: undefined,
        showAsBaseField: null,
        showAsBaseItem: undefined,
      });
      assert.equal(wb.getPivotDataFieldCount(0, pivot), 2);
      wb.addPivotFilter(0, pivot, {
        fieldIndex: region,
        kind: 1,
        value: 15,
      });
      wb.setPivotRowFieldOrder(0, pivot, [0]);
    } finally {
      wb.delete();
    }
  });

  test('pivot nested getter failures are rethrown before any C mutation', () => {
    const { wb, pivot } = makeCheckedPivot(Module);
    try {
      const fieldCount = wb.getPivotFieldCount(0, pivot);
      const fieldSpec = {};
      Object.defineProperty(fieldSpec, 'customName', {
        enumerable: true,
        get() {
          throw new Error('pivot field getter');
        },
      });
      assert.throws(() => wb.updatePivotField(0, pivot, 0, fieldSpec), /pivot field getter/);
      assert.equal(wb.getPivotFieldCount(0, pivot), fieldCount);

      const indices = [];
      Object.defineProperty(indices, 0, {
        enumerable: true,
        get() {
          throw new Error('pivot order getter');
        },
      });
      assert.throws(() => wb.setPivotRowFieldOrder(0, pivot, indices), /pivot order getter/);
      assert.deepEqual(checkedPivotGrid(wb, pivot), CHECKED_PIVOT_GRID);
    } finally {
      wb.delete();
    }
  });

  test('a rejected argument reports its own diagnostic, not the previous call', () => {
    const wb = Module.Workbook.createDefault();
    try {
      // Leave a failure of a different call behind.
      const stale = caught(() => wb.getValue(99, 0, 0));
      assert.equal(stale.name, 'FormulonError');
      assert.ok(stale.message.length > 0);

      const rejected = caught(() =>
        wb.addTable(0, {
          range: { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 },
          name: 'T',
        }),
      );
      assert.notEqual(rejected.message, stale.message);
      assert.ok(rejected instanceof TypeError);
      assert.match(rejected.message, /columns/);
    } finally {
      wb.delete();
    }
  });

  test('list getters throw for a bad sheet', () => {
    const wb = Module.Workbook.createDefault();
    try {
      for (const list of [() => wb.listMerges(99), () => wb.listComments(99), () => wb.listHyperlinks(99)]) {
        expectFailure(list, StatusCode.SheetIndexOutOfRange);
      }
    } finally {
      wb.delete();
    }
  });

  test('index getters throw for a bad sheet or index', () => {
    const wb = Module.Workbook.createDefault();
    try {
      for (const count of [
        () => wb.getValidationCount(99),
        () => wb.getCfRuleCount(99),
        () => wb.getProtectedRangeCount(99),
        () => wb.getSparklineGroupCount(99),
        () => wb.getCellPhoneticRunCount(99, 0, 0),
      ]) {
        expectFailure(count, StatusCode.SheetIndexOutOfRange);
      }
      assert.deepEqual([wb.getValidationCount(0), wb.getCfRuleCount(0), wb.getExternalLinkCount()], [0, 0, 0]);
      for (const entry of [
        () => wb.getValidation(0, 0),
        () => wb.getCfRule(0, 0),
        () => wb.getProtectedRange(0, 0),
        () => wb.getSparklineGroup(0, 0),
        () => wb.getExternalLink(0),
        () => wb.getCellPhoneticRun(0, 0, 0, 0),
      ]) {
        expectFailure(entry, StatusCode.IndexOutOfRange);
      }
    } finally {
      wb.delete();
    }
  });

  // Deliberately last. A shadow-stack overflow is a WebAssembly trap, not an
  // exception: it destroys the module instance, so nothing after it here
  // would be measuring the engine any more. Running it at the end means the
  // rest of the suite has already reported when it fires.
  test('nesting up to the parser depth cap stays inside the shadow stack', () => {
    // The engine bounds formula nesting with one constant, and under WASM the
    // stack that bound protects is a link-time size. When the two disagree the
    // failure is not an error value: the shadow stack runs into the data
    // segment below it, and past that the module faults in a way no
    // `Expected` can report and no `try` can recover from. So the assertion is
    // that the cap is reached by a diagnostic (status 1007).
    //
    // Nested calls are the shape that matters -- parentheses are elided by the
    // parser and cost a fraction as much per level. The cap itself is
    // discovered rather than written down, so raising it in the engine does
    // not quietly narrow what this covers.
    const wb = Module.Workbook.createDefault();
    try {
      wb.setNumber(0, 0, 0, 7);

      const searchCeiling = 1024;
      let rejectedAt = 0;
      for (let depth = 1; depth <= searchCeiling; depth += 1) {
        const formula = `=${'SUM('.repeat(depth)}A1${')'.repeat(depth)}`;
        let out;
        try {
          out = wb.evaluateFormulaArray(0, 1, 0, formula).cells[0][0];
        } catch (e) {
          // A formula past the nesting cap is refused at entry with a parser status.
          if (e?.code === 1007) {
            rejectedAt = depth;
            break;
          }
          // Past this point the instance is gone and its worker threads keep
          // the process alive, so reporting and continuing would hang rather
          // than fail. Exit on the spot with the depth that did it.
          console.error(`FAIL nesting depth ${depth} (${formula.length} chars) faulted instead of evaluating.`);
          console.error(`  ${e && e.stack ? e.stack : e}`);
          console.error(
            '  The WASM shadow stack is too small for parser::kMaxFormulaAstDepth. ' +
              'See _FM_WASM_STACK_SIZE in cmake/FormulonWasm.cmake and the stack_probe artifact.',
          );
          process.exit(1);
        }
        if (out.kind === VAL.ERROR) {
          assert.equal(out.errorCode, 4, `depth=${depth} should be rejected as #NAME?`);
          rejectedAt = depth;
          break;
        }
        // Nested SUM over one cell is the identity. A wrong answer here is what
        // a stack that overran its limit looks like from the outside.
        assert.equal(out.kind, VAL.NUMBER, `depth=${depth}`);
        assert.equal(out.number, 7, `depth=${depth}`);
      }
      assert.ok(rejectedAt > 0, `no nesting depth up to ${searchCeiling} was rejected; the depth cap did not engage`);
    } finally {
      wb.delete();
    }
  });

  test('pivot anchor, grand totals, style info, name and subtotal functions use the update vocabulary', () => {
    const { wb, pivot, region } = makeCheckedPivot(Module);
    try {
      // The anchor is one cell.
      wb.setPivotAnchor(0, pivot, { row: 12, col: 6 });
      const rect = wb.getPivotLayout(0, pivot).rect;
      assert.equal(rect.firstRow, 12);
      assert.equal(rect.firstCol, 6);
      assert.throws(() => wb.setPivotAnchor(0, pivot, { row: 1.5, col: 0 }), TypeError);

      wb.updatePivotGrandTotals(0, pivot, false, true);
      assert.deepEqual({ ...wb.getPivotGrandTotals(0, pivot) }, { rows: false, columns: true });

      const style = wb.getPivotStyleOptions(0, pivot);
      assert.equal(style.styleName, 'PivotStyleLight16');
      wb.updatePivotStyleOptions(0, pivot, {
        styleName: 'PivotStyleMedium9',
        showRowHeaders: true,
        showColumnHeaders: false,
        showRowStripes: true,
        showColumnStripes: true,
        showLastColumn: false,
      });
      assert.deepEqual(
        { ...wb.getPivotStyleOptions(0, pivot) },
        {
          styleName: 'PivotStyleMedium9',
          showRowHeaders: true,
          showColumnHeaders: false,
          showRowStripes: true,
          showColumnStripes: true,
          showLastColumn: false,
        },
      );

      wb.renamePivot(0, pivot, 'Renamed');
      assert.equal(wb.getPivotName(0, pivot), 'Renamed');

      wb.setPivotFieldSubtotalFunctions(0, pivot, region, [0, 1]);
      assert.deepEqual([...wb.getPivotField(0, pivot, region).subtotalFunctions], [0, 1]);
      wb.setPivotFieldSubtotalFunctions(0, pivot, region, [1]);
      assert.deepEqual([...wb.getPivotField(0, pivot, region).subtotalFunctions], [1]);
      expectFailure(() => wb.setPivotFieldSubtotalFunctions(0, pivot, region, [0, 0]), StatusCode.InvalidArgument);
      assert.deepEqual([...wb.getPivotField(0, pivot, region).subtotalFunctions], [1]);
      wb.setPivotFieldSubtotalFunctions(0, pivot, region, []);
      assert.deepEqual([...wb.getPivotField(0, pivot, region).subtotalFunctions], []);
    } finally {
      wb.delete();
    }
  });
}
