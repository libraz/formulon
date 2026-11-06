import test from 'node:test';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';

test('a nested recalc from the iterative progress callback throws on serial and parallel passes', async () => {
  const mod = await getModule();
  const runOuterRecalc = (recalcOuter) => {
    const wb = mod.Workbook.createDefault();
    try {
      wb.setIterative({ enabled: true, maxIterations: 10, maxChange: 0.001 });
      wb.sheet(0).setFormula(0, 0, '=(A1+10)/2');
      let calls = 0;
      let nested;
      wb.setIterativeProgress(() => {
        calls += 1;
        try {
          wb.recalc();
        } catch (err) {
          nested = err;
        }
        return false;
      });

      recalcOuter(wb);
      assert.equal(calls, 1);
      assert.equal(nested?.name, 'FormulonError', `the nested recalc must throw: ${String(nested)}`);
      assert.equal(nested.code, 4005);
      assert.match(nested.message, /kGraphRecalcReentrant, 4005/);

      return nested.message;
    } finally {
      wb.setIterativeProgress(null);
      wb.dispose();
    }
  };

  const serial = runOuterRecalc((wb) => wb.recalc());
  const parallel = runOuterRecalc((wb) => wb.recalcParallel(2));
  assert.equal(parallel, serial, 'serial and parallel outer recalcs refuse the nested call alike');
});

test('a throwing iterative progress callback aborts the solve and its exception reaches the caller', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    wb.setIterative({ enabled: true, maxIterations: 100, maxChange: 0.0001 });
    wb.sheet(0).setFormula(0, 0, '=(A1+10)/2');
    let calls = 0;
    const thrown = new Error('callback failure');
    wb.setIterativeProgress(() => {
      calls += 1;
      throw thrown;
    });

    // The callback's own exception is rethrown by the outer call.
    let caught;
    try {
      wb.recalc();
    } catch (err) {
      caught = err;
    }
    assert.equal(caught, thrown);
    assert.equal(calls, 1, 'the solve must stop at the first throw');

    wb.sheet(0).getValue(0, 0);

    wb.setIterativeProgress(null);
    wb.sheet(0).setFormula(0, 0, '=(A1+10)/2');
    wb.recalc();
    const converged = wb.sheet(0).getValue(0, 0);
    assert.ok(Math.abs(converged.number - 10) < 0.001, `A1=${converged.number}`);
  } finally {
    wb.setIterativeProgress(null);
    wb.dispose();
  }
});

// `CfValueObjectInput.type`: 0 num, 1 percent, 2 percentile, 3 min,
// 4 max, 5 formula, 6 autoMin, 7 autoMax. Every threshold below states a
// `value`, which is what puts a borrowed string behind each CFVO.
const COLOR_SCALE_RULE = Object.freeze({
  ranges: [{ firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 0 }],
  type: 2,
  colorScale: {
    thresholds: [
      { type: 0, value: '0' },
      { type: 1, value: '50' },
      { type: 0, value: '100' },
    ],
    colors: [
      { r: 248, g: 105, b: 107, a: 255 },
      { r: 255, g: 235, b: 132, a: 255 },
      { r: 99, g: 190, b: 123, a: 255 },
    ],
  },
});

const DATA_BAR_RULE = Object.freeze({
  ranges: [{ firstRow: 0, firstCol: 1, lastRow: 4, lastCol: 1 }],
  type: 3,
  dataBar: {
    min: { type: 0, value: '5' },
    max: { type: 0, value: '95' },
    fill: { r: 0, g: 112, b: 192, a: 255 },
  },
});

const ICON_SET_RULE = Object.freeze({
  ranges: [{ firstRow: 0, firstCol: 2, lastRow: 4, lastCol: 2 }],
  type: 4,
  iconSet: {
    // 4Arrows: a 4-icon set carries three thresholds.
    name: 8,
    thresholds: [
      { type: 1, value: '0' },
      { type: 1, value: '33' },
      { type: 1, value: '67' },
    ],
  },
});

const EXPECTED_CFVO_VALUES = Object.freeze(['0', '50', '100', '5', '95', '0', '33', '67']);

// The threshold value strings of the three visual rules, flattened.
function cfvoValues(rules) {
  const byType = new Map(rules.map((rule) => [rule.type, rule]));
  const colorScale = byType.get(2);
  const dataBar = byType.get(3);
  const iconSet = byType.get(4);
  assert.ok(colorScale?.colorScale, 'colorScale payload missing');
  assert.ok(dataBar?.dataBar, 'dataBar payload missing');
  assert.ok(iconSet?.iconSet, 'iconSet payload missing');
  return [
    ...colorScale.colorScale.thresholds.map((t) => t.value),
    dataBar.dataBar.min.value,
    dataBar.dataBar.max.value,
    ...iconSet.iconSet.thresholds.map((t) => t.value),
  ];
}

test('CFVO value strings survive getCfRule and a save/load cycle', async () => {
  // Each CFVO `value` crosses the C ABI as a borrowed `const char*`. A
  // store that relocates while the later thresholds are pulled publishes
  // a pointer into freed bytes, so the strings come back wrong or empty
  // even though the rule count and colors still look right. Rules whose
  // thresholds are all `min` / `max` never take that path, because those
  // types carry no value string.
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  let bytes;
  try {
    for (const rule of [COLOR_SCALE_RULE, DATA_BAR_RULE, ICON_SET_RULE]) {
      wb.sheet(0).addCfRule(rule);
    }

    const inSession = wb.sheet(0).listCfRules();
    assert.equal(inSession.length, 3);
    assert.deepEqual(cfvoValues(inSession), EXPECTED_CFVO_VALUES);
    const byType = new Map(inSession.map((rule) => [rule.type, rule]));
    // The type travels with the value: a percent threshold read back as
    // type 0 would mean an absolute number instead.
    assert.deepEqual(
      byType.get(2).colorScale.thresholds.map((t) => t.type),
      [0, 1, 0],
    );
    assert.deepEqual(
      byType.get(4).iconSet.thresholds.map((t) => t.type),
      [1, 1, 1],
    );
    assert.equal(byType.get(2).colorScale.colors.length, 3);
    assert.equal(byType.get(2).colorScale.colors[2].g, 190);
    assert.equal(byType.get(3).dataBar.fill.b, 192);

    const saved = wb.save();
    bytes = saved;
  } finally {
    wb.dispose();
  }

  const reloaded = mod.Workbook.load(bytes);
  try {
    const rules = reloaded.sheet(0).listCfRules();
    assert.equal(rules.length, 3);
    assert.deepEqual(cfvoValues(rules), EXPECTED_CFVO_VALUES);
  } finally {
    reloaded.dispose();
  }
});

test('getIterative reads back what setIterative stored', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const initial = wb.getIterative();
    assert.equal(typeof initial.enabled, 'boolean');
    assert.equal(typeof initial.maxIterations, 'number');
    assert.equal(typeof initial.maxChange, 'number');

    wb.setIterative({ enabled: true, maxIterations: 42, maxChange: 0.25 });
    const enabled = wb.getIterative();
    assert.equal(enabled.enabled, true);
    assert.equal(enabled.maxIterations, 42);
    assert.equal(enabled.maxChange, 0.25);

    // The cap and threshold survive switching iteration back off, which is
    // what lets a host render the dialog with the stored values.
    wb.setIterative({ enabled: false, maxIterations: 42, maxChange: 0.25 });
    const disabled = wb.getIterative();
    assert.equal(disabled.enabled, false);
    assert.equal(disabled.maxIterations, 42);
    assert.equal(disabled.maxChange, 0.25);
  } finally {
    wb.dispose();
  }
});

test('addValidation defaults allowBlank to false', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    wb.sheet(0).addValidation({
      ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
      type: 3,
      formula1: '"a,b"',
    });
    const list = wb.sheet(0).listValidations();
    assert.equal(list.length, 1);
    assert.equal(list[0].allowBlank, false);
  } finally {
    wb.dispose();
  }
});

test('getCellStyleXf reports xfId alongside the cell-format fields', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const read = wb.getCellStyleXf(0);
    // A named-style xf is its own parent, so it never inherits one.
    assert.equal(read.xfId, 0);
    // The record shape matches `getCellXf`, which the documented
    // read / edit / addCellStyleXf round-trip depends on.
    const cellXf = wb.getCellXf(0);
    for (const key of Object.keys(cellXf)) {
      assert.ok(key in read, `getCellStyleXf is missing '${key}'`);
    }
  } finally {
    wb.dispose();
  }
});

test('pivotFieldAddItemAt hides the blank item an empty label cannot name', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const cache = wb.addPivotCache(0);
    const region = wb.addPivotCacheField(cache, 'Region');
    const amount = wb.addPivotCacheField(cache, 'Amount');
    // Shared item 0 renders as "North"; shared item 1 is the blank, which
    // carries no label of its own and can only be named by its index.
    wb.addPivotCacheFieldSharedItemText(cache, region, 'North');
    wb.addPivotCacheFieldSharedItemBlank(cache, region);

    const first = wb.addPivotCacheRecord(cache);
    wb.setPivotCacheRecordNumber(cache, first, region, 0);
    wb.setPivotCacheRecordNumber(cache, first, amount, 100);
    const second = wb.addPivotCacheRecord(cache);
    wb.setPivotCacheRecordBlank(cache, second, region);
    wb.setPivotCacheRecordNumber(cache, second, amount, 200);

    const pivot = wb.sheet(0).addPivot('PT', cache, 0, 0);
    const regionField = wb.sheet(0).addPivotField(pivot, { sourceName: 'Region', axis: 0 });
    const amountField = wb.sheet(0).addPivotField(pivot, { sourceName: 'Amount', axis: 2 });
    wb.sheet(0).setPivotRowFieldOrder(pivot, [regionField]);
    wb.sheet(0).addPivotDataField(pivot, {
      name: 'Sum of Amount',
      fieldIndex: amountField,
      aggregation: 0,
    });

    const sumDataCells = () => {
      const layout = wb.sheet(0).getPivotLayout(pivot);
      let total = 0;
      for (const cell of layout.cells) {
        if (cell.kind === mod.PivotCellKind.Data && cell.value.kind === mod.ValueKind.Number) {
          total += cell.value.number;
        }
      }
      return total;
    };
    assert.equal(sumDataCells(), 300);

    wb.sheet(0).addPivotFieldItem(pivot, regionField, 'North', true);
    wb.sheet(0).addPivotFieldSharedItem(pivot, regionField, 1, false);
    assert.equal(sumDataCells(), 100);
  } finally {
    wb.dispose();
  }
});

test('a rejected argument reports its own diagnostic, not the previous call', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    // Leave a failure of a different call in the thread-local diagnostics.
    const failure = (call) => {
      try {
        call();
      } catch (err) {
        return err;
      }
      return assert.fail('expected the call to throw');
    };
    const stale = failure(() => wb.sheet(99).getValue(0, 0));
    assert.ok(isFormulonError(stale));
    assert.ok(stale.message.length > 0);

    // An argument the binding rejects before any C call carries its own
    // message, and a TypeError rather than the destroyed-handle code.
    const shortColors = failure(() => wb.setThemeColors({ colors: [1, 2] }));
    assert.ok(shortColors instanceof TypeError);
    assert.match(shortColors.message, /setThemeColors/);
    assert.notEqual(shortColors.message, stale.message);

    const badCallback = failure(() => wb.setIterativeProgress(42));
    assert.ok(badCallback instanceof TypeError);
    assert.match(badCallback.message, /callback/);
  } finally {
    wb.setIterativeProgress(null);
    wb.dispose();
  }
});

test('a call on a disposed workbook names the destroyed handle', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.dispose();
  // 7000 is reserved for this: an argument fault reports 2 instead, so
  // "make a new workbook" stays the correct recovery for exactly this code.
  assert.throws(
    () => wb.recalc(),
    (e) => isFormulonError(e, 7000) && e.message.length > 0,
  );
});

test('sheet-layout setters refuse coordinates and metrics Excel cannot open', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    // These six setters marshal straight to the C ABI, which is where the
    // validation lives. The assertions below are what keeps that true: a
    // binding that grew its own path to the model would still pass every
    // round-trip test while quietly authoring a file Excel refuses, since
    // NaN and Infinity serialise to an empty string -- `<col width=""/>`
    // is not a lexical xsd:double at all, and a repair prompt rather than
    // an out-of-range value.
    const kInvalidArgument = 2;
    // The message is pinned, not just the code. Both JS bindings route
    // these calls to one set of C ABI helpers, so both report that shared
    // wording; a binding that grew its own check would keep returning
    // `kInvalidArgument` while the text stopped matching. The status code
    // alone cannot tell a shared implementation from two that agree today.
    const kSpanRefused = 'column span out of range';
    const kRowSpanRefused = 'row span out of range';
    const kMetricRefused = 'value must be finite and non-negative';
    const refused = (call, expectedMessage, label) => {
      assert.throws(
        call,
        (e) => isFormulonError(e, kInvalidArgument) && e.message === expectedMessage,
        `${label} should have been refused`,
      );
    };

    // One past Excel's last column / row.
    refused(() => wb.sheet(0).setColumnWidth(0, 16384, 10), kSpanRefused, 'setColumnWidth past the grid');
    refused(() => wb.sheet(0).setColumnHidden(0, 16384, true), kSpanRefused, 'setColumnHidden past the grid');
    refused(() => wb.sheet(0).setColumnOutline(0, 16384, 1), kSpanRefused, 'setColumnOutline past the grid');
    refused(() => wb.sheet(0).setRowHeight(0, 1048576, 10), kRowSpanRefused, 'setRowHeight past the grid');
    refused(() => wb.sheet(0).setRowHidden(0, 1048576, true), kRowSpanRefused, 'setRowHidden past the grid');
    refused(() => wb.sheet(0).setRowOutline(0, 1048576, 1), kRowSpanRefused, 'setRowOutline past the grid');

    for (const bad of [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY, -1]) {
      refused(() => wb.sheet(0).setColumnWidth(0, 2, bad), kMetricRefused, `setColumnWidth(${bad})`);
      refused(() => wb.sheet(0).setRowHeight(1, 1, bad), kMetricRefused, `setRowHeight(${bad})`);
    }

    // The context carries the values the engine saw, which is what makes a rejection actionable.
    assert.throws(
      () => wb.sheet(0).setColumnWidth(0, 16384, 10),
      (e) => /first=0 last=16384/.test(e.context),
    );
    assert.throws(
      () => wb.sheet(0).setRowHeight(1, 1, Number.NaN),
      (e) => /height=/.test(e.context),
    );

    // Zero is a real width and a real height, not a rejected one.
    wb.sheet(0).setColumnWidth(0, 2, 0);
    wb.sheet(0).setRowHeight(1, 1, 0);

    // The last coordinates Excel does address stay accepted -- an
    // off-by-one in either bound would refuse them.
    wb.sheet(0).setColumnWidth(16383, 16383, 10);
    wb.sheet(0).setColumnHidden(16383, 16383, true);
    wb.sheet(0).setColumnOutline(16383, 16383, 1);
    wb.sheet(0).setRowHeight(1048575, 1048575, 10);
    wb.sheet(0).setRowHidden(1048575, 1048575, true);
    wb.sheet(0).setRowOutline(1048575, 1048575, 1);
  } finally {
    wb.dispose();
  }
});

test('setVisibility states every tab visibility, veryHidden included', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const read = () => wb.sheet(0).getView();
    assert.equal(read().visibility, mod.SheetVisibility.Visible);
    // The only visible sheet cannot be hidden: kInvalidArgument, sheet left alone.
    assert.throws(
      () => wb.sheet(0).setVisibility(mod.SheetVisibility.Hidden),
      (e) => isFormulonError(e, 2),
    );
    assert.equal(read().visibility, mod.SheetVisibility.Visible);
    // The last sheet keeps its own refusal (kCannotRemoveLastSheet).
    assert.throws(
      () => wb.removeSheet(0),
      (e) => isFormulonError(e, 5052),
    );
    assert.equal(wb.listSheets().length, 1);
    // A second sheet lets the first one change state, but removing the first
    // while the second is hidden would leave no visible sheet.
    wb.addSheet('Second');
    wb.sheet(1).setVisibility(mod.SheetVisibility.Hidden);
    assert.throws(
      () => wb.removeSheet(0),
      (e) => isFormulonError(e, 2),
    );
    assert.equal(wb.listSheets().length, 2);
    wb.sheet(1).setVisibility(mod.SheetVisibility.Visible);

    wb.sheet(0).setVisibility(mod.SheetVisibility.VeryHidden);
    assert.equal(read().visibility, mod.SheetVisibility.VeryHidden);

    // Demotion and showing are explicit states too.
    wb.sheet(0).setVisibility(mod.SheetVisibility.Hidden);
    assert.equal(read().visibility, mod.SheetVisibility.Hidden);
    wb.sheet(0).setVisibility(mod.SheetVisibility.Visible);
    assert.equal(read().visibility, mod.SheetVisibility.Visible);

    // An unknown ordinal is refused and leaves the sheet alone.
    wb.sheet(0).setVisibility(mod.SheetVisibility.VeryHidden);
    assert.throws(() => wb.sheet(0).setVisibility(3), isFormulonError);
    assert.equal(read().visibility, mod.SheetVisibility.VeryHidden);
  } finally {
    wb.dispose();
  }
});
