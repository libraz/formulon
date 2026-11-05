import assert from 'node:assert/strict';
import { PIVOT, VAL } from './smoke_support.mjs';

const nestedErrorClass = (value) => (typeof value === 'symbol' ? TypeError : RangeError);

function makeCheckedPivot(Module) {
  const wb = Module.Workbook.createDefault();
  const must = (result, what) => {
    const status = typeof result.ok === 'boolean' ? result : result.status;
    assert.ok(status.ok, `${what}: ${JSON.stringify(status)}`);
    return result;
  };
  const cache = must(wb.addPivotCache(0), 'addPivotCache').index;
  must(
    wb.setPivotCacheWorksheetSource(cache, { present: true, ref: 'A1:B3', sheet: 'Sheet1' }),
    'setPivotCacheWorksheetSource',
  );
  must(wb.addPivotCacheField(cache, 'Region'), 'pivotCacheFieldAdd Region');
  must(wb.addPivotCacheField(cache, 'Amount'), 'pivotCacheFieldAdd Amount');
  for (const [region, amount] of [
    ['East', 10],
    ['West', 30],
  ]) {
    const record = must(wb.addPivotCacheRecord(cache), 'addPivotCacheRecord').index;
    must(wb.setPivotCacheRecordText(cache, record, 0, region), 'setPivotCacheRecordText');
    must(wb.setPivotCacheRecordNumber(cache, record, 1, amount), 'setPivotCacheRecordNumber');
  }
  const pivot = must(wb.addPivot(0, 'Pivot1', cache, 0, 4), 'addPivot').index;
  const region = must(wb.addPivotField(0, pivot, { sourceName: 'Region', axis: 0 }), 'region field').index;
  const amount = must(wb.addPivotField(0, pivot, { sourceName: 'Amount', axis: 2 }), 'amount field').index;
  must(wb.setPivotRowFieldOrder(0, pivot, [region]), 'setPivotRowFieldOrder');
  must(
    wb.addPivotDataField(0, pivot, { name: 'Sum of Amount', fieldIndex: amount, aggregation: 0 }),
    'addPivotDataField',
  );
  return { wb, pivot, region, amount };
}

function checkedPivotGrid(wb, pivot) {
  const layout = wb.getPivotLayout(0, pivot);
  assert.ok(layout.status.ok, JSON.stringify(layout.status));
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
  test('pivotFieldAddItemAt hides the blank item an empty label cannot name', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const cache = wb.addPivotCache(0);
      assert.ok(cache.status.ok, `pivotCacheCreate: ${JSON.stringify(cache.status)}`);
      const region = wb.addPivotCacheField(cache.index, 'Region');
      assert.ok(region.status.ok);
      const amount = wb.addPivotCacheField(cache.index, 'Amount');
      assert.ok(amount.status.ok);
      // Shared item 0 renders as "North"; shared item 1 is the blank, which
      // carries no label of its own and can only be named by its index.
      assert.ok(wb.addPivotCacheFieldSharedItemText(cache.index, region.index, 'North').ok);
      assert.ok(wb.addPivotCacheFieldSharedItemBlank(cache.index, region.index).ok);

      const first = wb.addPivotCacheRecord(cache.index);
      assert.ok(first.status.ok);
      assert.ok(wb.setPivotCacheRecordNumber(cache.index, first.index, region.index, 0).ok);
      assert.ok(wb.setPivotCacheRecordNumber(cache.index, first.index, amount.index, 100).ok);
      const second = wb.addPivotCacheRecord(cache.index);
      assert.ok(second.status.ok);
      assert.ok(wb.setPivotCacheRecordBlank(cache.index, second.index, region.index).ok);
      assert.ok(wb.setPivotCacheRecordNumber(cache.index, second.index, amount.index, 200).ok);

      const pivot = wb.addPivot(0, 'PT', cache.index, 0, 0);
      assert.ok(pivot.status.ok, `pivotCreate: ${JSON.stringify(pivot.status)}`);
      const regionField = wb.addPivotField(0, pivot.index, { sourceName: 'Region', axis: 0 });
      assert.ok(regionField.status.ok);
      const amountField = wb.addPivotField(0, pivot.index, { sourceName: 'Amount', axis: 2 });
      assert.ok(amountField.status.ok);
      assert.ok(wb.setPivotRowFieldOrder(0, pivot.index, [regionField.index]).ok);
      assert.ok(
        wb.addPivotDataField(0, pivot.index, {
          name: 'Sum of Amount',
          fieldIndex: amountField.index,
          aggregation: 0,
        }).status.ok,
      );

      const sumDataCells = () => {
        const layout = wb.getPivotLayout(0, pivot.index);
        assert.ok(layout.status.ok, `pivotLayout: ${JSON.stringify(layout.status)}`);
        let total = 0;
        for (const cell of layout.cells) {
          if (cell.kind === PIVOT.DATA && cell.value.kind === VAL.NUMBER) {
            total += cell.value.number;
          }
        }
        return total;
      };
      assert.equal(sumDataCells(), 300);

      assert.ok(wb.addPivotFieldItem(0, pivot.index, regionField.index, 'North', true).status.ok);
      assert.ok(wb.addPivotFieldSharedItem(0, pivot.index, regionField.index, 1, false).status.ok);
      assert.equal(sumDataCells(), 100);
    } finally {
      wb.delete();
    }
  });

  test('pivot nested integers throw on narrowing before C mutation', () => {
    const invalid = [2 ** 32, 1.5, Number.NaN, Number.POSITIVE_INFINITY, Symbol('pivot numeric')];
    for (const value of invalid) {
      const { wb, pivot } = makeCheckedPivot(Module);
      try {
        const before = wb.getPivotFieldCount(0, pivot).value;
        assert.throws(
          () => wb.addPivotField(0, pivot, { sourceName: 'Region', axis: value }),
          nestedErrorClass(value),
          `pivotFieldAdd axis=${String(value)} accepted`,
        );
        assert.equal(wb.getPivotFieldCount(0, pivot).value, before);
      } finally {
        wb.delete();
      }

      const data = makeCheckedPivot(Module);
      try {
        const before = data.wb.getPivotDataFieldCount(0, data.pivot).value;
        assert.throws(
          () =>
            data.wb.addPivotDataField(0, data.pivot, {
              name: 'Bad',
              fieldIndex: value,
              aggregation: 0,
            }),
          nestedErrorClass(value),
          `pivotDataFieldAdd fieldIndex=${String(value)} accepted`,
        );
        assert.equal(data.wb.getPivotDataFieldCount(0, data.pivot).value, before);
      } finally {
        data.wb.delete();
      }
    }
  });

  test('pivot filters throw on invalid axis, value kind, and field index before mutation', () => {
    const invalid = [
      { axis: 2 ** 32 },
      { valueKind: 1.5 },
      { valueKind: Number.NaN },
      { valueKind: Symbol('pivot filter') },
      { dataFieldIndex: 2 ** 32 },
    ];
    for (const override of invalid) {
      const { wb, pivot } = makeCheckedPivot(Module);
      try {
        const before = wb.getPivotFilterCount(0, pivot).value;
        assert.throws(
          () =>
            wb.addPivotFilter(0, pivot, {
              axis: 0,
              fieldName: 'Region',
              type: 1,
              valueKind: 1,
              valueDouble: 15,
              ...override,
            }),
          nestedErrorClass(Object.values(override)[0]),
          `pivotFilterAdd ${String(Object.values(override)[0])} accepted`,
        );
        assert.equal(wb.getPivotFilterCount(0, pivot).value, before);
        assert.deepEqual(checkedPivotGrid(wb, pivot), CHECKED_PIVOT_GRID);
      } finally {
        wb.delete();
      }
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
    const { wb, pivot, amount } = makeCheckedPivot(Module);
    try {
      const data = wb.addPivotDataField(0, pivot, {
        name: 'Count of Amount',
        fieldIndex: amount,
        aggregation: 1,
        numberFormat: null,
        showAs: undefined,
        showAsBaseField: null,
        showAsBaseItem: undefined,
      });
      assert.ok(data.status.ok, JSON.stringify(data.status));
      assert.equal(wb.getPivotDataFieldCount(0, pivot).value, 2);
      const filter = wb.addPivotFilter(0, pivot, {
        axis: 0,
        fieldName: 'Region',
        type: 1,
        valueKind: 1,
        valueDouble: 15,
      });
      assert.ok(filter.status.ok, JSON.stringify(filter));
      assert.ok(wb.setPivotRowFieldOrder(0, pivot, [0]).ok);
    } finally {
      wb.delete();
    }
  });

  test('pivot nested getter failures are rethrown before any C mutation', () => {
    const { wb, pivot } = makeCheckedPivot(Module);
    try {
      const fieldCount = wb.getPivotFieldCount(0, pivot).value;
      const fieldSpec = { axis: 0 };
      Object.defineProperty(fieldSpec, 'sourceName', {
        enumerable: true,
        get() {
          throw new Error('pivot field getter');
        },
      });
      assert.throws(() => wb.addPivotField(0, pivot, fieldSpec), /pivot field getter/);
      assert.equal(wb.getPivotFieldCount(0, pivot).value, fieldCount);

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
      // Leave a failure of a different call in the thread-local diagnostics.
      const stale = wb.getValue(99, 0, 0);
      assert.equal(stale.status.ok, false);
      assert.ok(stale.status.message.length > 0);

      const rejected = wb.addTable({
        range: { sheet: 0, rect: { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 } },
        name: 'T',
      });
      assert.equal(rejected.status.ok, false);
      assert.notEqual(rejected.status.message, stale.status.message);
      assert.match(rejected.status.message, /columns/);
      assert.equal(rejected.status.context, '');
    } finally {
      wb.delete();
    }
  });

  test('list getters expose failures through their array status', () => {
    const wb = Module.Workbook.createDefault();
    try {
      for (const list of [wb.listMerges(99), wb.listComments(99), wb.listHyperlinks(99)]) {
        assert.ok(Array.isArray(list));
        assert.equal(list.length, 0);
        assert.equal(list.status.ok, false);
      }
    } finally {
      wb.delete();
    }
  });

  test('index getters report a bad sheet or index through their status', () => {
    const wb = Module.Workbook.createDefault();
    try {
      for (const count of [
        wb.getValidationCount(99),
        wb.getCfRuleCount(99),
        wb.getProtectedRangeCount(99),
        wb.getSparklineGroupCount(99),
        wb.getCellPhoneticRunCount(99, 0, 0),
      ]) {
        assert.equal(count.status.ok, false);
      }
      assert.deepEqual(
        [wb.getValidationCount(0), wb.getCfRuleCount(0), wb.getExternalLinkCount()].map((r) => [r.status.ok, r.value]),
        [
          [true, 0],
          [true, 0],
          [true, 0],
        ],
      );
      for (const entry of [
        wb.getValidation(0, 0),
        wb.getCfRule(0, 0),
        wb.getProtectedRange(0, 0),
        wb.getSparklineGroup(0, 0),
        wb.getExternalLink(0),
        wb.getCellPhoneticRun(0, 0, 0, 0),
      ]) {
        assert.equal(entry.status.ok, false);
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
    // that the cap is reached by a diagnostic.
    //
    // Nested calls are the shape that matters -- parentheses are elided by the
    // parser and cost a fraction as much per level. The cap itself is
    // discovered rather than written down, so raising it in the engine does
    // not quietly narrow what this covers.
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setNumber(0, 0, 0, 7).ok);

      const searchCeiling = 1024;
      let rejectedAt = 0;
      for (let depth = 1; depth <= searchCeiling; depth += 1) {
        const formula = `=${'SUM('.repeat(depth)}A1${')'.repeat(depth)}`;
        let out;
        try {
          out = wb.evaluateFormula(0, 1, 0, formula);
        } catch (e) {
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
        assert.ok(out.status.ok, `depth=${depth} status=${JSON.stringify(out.status)}`);
        if (out.value.kind === VAL.ERROR) {
          assert.equal(out.value.errorCode, 4, `depth=${depth} should be rejected as #NAME?`);
          rejectedAt = depth;
          break;
        }
        // Nested SUM over one cell is the identity. A wrong answer here is what
        // a stack that overran its limit looks like from the outside.
        assert.equal(out.value.kind, VAL.NUMBER, `depth=${depth}`);
        assert.equal(out.value.number, 7, `depth=${depth}`);
      }
      assert.ok(rejectedAt > 0, `no nesting depth up to ${searchCeiling} was rejected; the depth cap did not engage`);
    } finally {
      wb.delete();
    }
  });
}
