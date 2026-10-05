import test from 'node:test';
import { assert, getModule } from './smoke_support.mjs';

test('a throwing iterative progress callback aborts the solve and reports a status', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.ok(wb.setIterative(true, 100, 0.0001).ok);
    assert.ok(wb.setFormula(0, 0, 0, '=(A1+10)/2').ok);
    let calls = 0;
    assert.ok(
      wb.setIterativeProgress(() => {
        calls += 1;
        throw new Error('callback failure');
      }).ok,
    );

    // Reported as a status rather than propagated, and with the same
    // ordinal the WASM binding uses (`kBindingCallbackException`).
    const aborted = wb.recalc();
    assert.equal(aborted.ok, false, `expected a failure envelope: ${JSON.stringify(aborted)}`);
    assert.equal(aborted.status, 7003);
    // The engine reported success (it only saw a cancellation), so there
    // is no thread-local diagnostic to read here -- this must be the
    // binding's own message, matching the WASM package's, not the residue
    // of an earlier call.
    assert.equal(aborted.message, 'the iterative progress callback threw; the solve was aborted');
    assert.equal(aborted.context, 'Workbook.setIterativeProgress');
    assert.equal(calls, 1, 'the solve must stop at the first throw');

    const mid = wb.getValue(0, 0, 0);
    assert.ok(mid.status.ok, JSON.stringify(mid.status));

    assert.ok(wb.setIterativeProgress(null).ok);
    assert.ok(wb.setFormula(0, 0, 0, '=(A1+10)/2').ok);
    const recovered = wb.recalc();
    assert.ok(recovered.ok, `recalc after the throwing callback: ${JSON.stringify(recovered)}`);
    const converged = wb.getValue(0, 0, 0);
    assert.ok(Math.abs(converged.value.number - 10) < 0.001, `A1=${converged.value.number}`);
  } finally {
    wb.setIterativeProgress(null);
    wb.dispose();
  }
});

// `CfValueObjectInput.type`: 0 num, 1 percent, 2 percentile, 3 min,
// 4 max, 5 formula, 6 autoMin, 7 autoMax. Every threshold below states a
// `value`, which is what puts a borrowed string behind each CFVO.
const COLOR_SCALE_RULE = Object.freeze({
  sqref: [{ firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 0 }],
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
  sqref: [{ firstRow: 0, firstCol: 1, lastRow: 4, lastCol: 1 }],
  type: 3,
  dataBar: {
    min: { type: 0, value: '5' },
    max: { type: 0, value: '95' },
    fill: { r: 0, g: 112, b: 192, a: 255 },
  },
});

const ICON_SET_RULE = Object.freeze({
  sqref: [{ firstRow: 0, firstCol: 2, lastRow: 4, lastCol: 2 }],
  type: 4,
  iconSet: {
    name: 0,
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

test('CFVO value strings survive getConditionalFormats and a save/load cycle', async () => {
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
      const added = wb.addConditionalFormat(0, rule);
      assert.ok(added.status.ok, `addConditionalFormat type=${rule.type}: ${JSON.stringify(added.status)}`);
    }

    const inSession = wb.getConditionalFormats(0);
    assert.ok(inSession.status.ok, JSON.stringify(inSession.status));
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
    assert.ok(saved.status.ok, JSON.stringify(saved.status));
    bytes = saved.bytes;
  } finally {
    wb.dispose();
  }

  const reloaded = mod.Workbook.loadBytes(bytes);
  try {
    const rules = reloaded.getConditionalFormats(0);
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
    assert.ok(initial.status.ok, `getIterative: ${JSON.stringify(initial.status)}`);
    assert.equal(typeof initial.enabled, 'boolean');
    assert.equal(typeof initial.maxIterations, 'number');
    assert.equal(typeof initial.maxChange, 'number');

    assert.ok(wb.setIterative(true, 42, 0.25).ok);
    const enabled = wb.getIterative();
    assert.ok(enabled.status.ok);
    assert.equal(enabled.enabled, true);
    assert.equal(enabled.maxIterations, 42);
    assert.equal(enabled.maxChange, 0.25);

    // The cap and threshold survive switching iteration back off, which is
    // what lets a host render the dialog with the stored values.
    assert.ok(wb.setIterative(false, 42, 0.25).ok);
    const disabled = wb.getIterative();
    assert.ok(disabled.status.ok);
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
    assert.ok(
      wb.addValidation(0, {
        ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
        type: 3,
        formula1: '"a,b"',
      }).ok,
    );
    const list = wb.getValidations(0);
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
    assert.ok(read.status.ok, `getCellStyleXf: ${JSON.stringify(read.status)}`);
    // A named-style xf is its own parent, so it never inherits one.
    assert.equal(read.xfId, 0);
    // The record shape matches `getCellXf`, which the documented
    // read / edit / addCellStyleXf round-trip depends on.
    const cellXf = wb.getCellXf(0);
    assert.ok(cellXf.status.ok);
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
    const cache = wb.pivotCacheCreate(0);
    assert.ok(cache.status.ok, `pivotCacheCreate: ${JSON.stringify(cache.status)}`);
    const region = wb.pivotCacheFieldAdd(cache.index, 'Region');
    assert.ok(region.status.ok);
    const amount = wb.pivotCacheFieldAdd(cache.index, 'Amount');
    assert.ok(amount.status.ok);
    // Shared item 0 renders as "North"; shared item 1 is the blank, which
    // carries no label of its own and can only be named by its index.
    assert.ok(wb.pivotCacheFieldAddSharedItemText(cache.index, region.index, 'North').ok);
    assert.ok(wb.pivotCacheFieldAddSharedItemBlank(cache.index, region.index).ok);

    const first = wb.pivotCacheRecordAdd(cache.index);
    assert.ok(first.status.ok);
    assert.ok(wb.pivotCacheRecordSetNumber(cache.index, first.index, region.index, 0).ok);
    assert.ok(wb.pivotCacheRecordSetNumber(cache.index, first.index, amount.index, 100).ok);
    const second = wb.pivotCacheRecordAdd(cache.index);
    assert.ok(second.status.ok);
    assert.ok(wb.pivotCacheRecordSetBlank(cache.index, second.index, region.index).ok);
    assert.ok(wb.pivotCacheRecordSetNumber(cache.index, second.index, amount.index, 200).ok);

    const pivot = wb.pivotCreate(0, 'PT', cache.index, 0, 0);
    assert.ok(pivot.status.ok, `pivotCreate: ${JSON.stringify(pivot.status)}`);
    const regionField = wb.pivotFieldAdd(0, pivot.index, { sourceName: 'Region', axis: 0 });
    assert.ok(regionField.status.ok);
    const amountField = wb.pivotFieldAdd(0, pivot.index, { sourceName: 'Amount', axis: 2 });
    assert.ok(amountField.status.ok);
    assert.ok(wb.pivotSetRowFieldOrder(0, pivot.index, [regionField.index]).ok);
    assert.ok(
      wb.pivotDataFieldAdd(0, pivot.index, {
        name: 'Sum of Amount',
        fieldIndex: amountField.index,
        aggregation: 0,
      }).status.ok,
    );

    const sumDataCells = () => {
      const layout = wb.pivotLayout(0, pivot.index);
      assert.ok(layout.status.ok, `pivotLayout: ${JSON.stringify(layout.status)}`);
      let total = 0;
      for (const cell of layout.cells) {
        if (cell.kind === mod.PivotCellKind.Data && cell.value.kind === mod.ValueKind.Number) {
          total += cell.value.number;
        }
      }
      return total;
    };
    assert.equal(sumDataCells(), 300);

    assert.ok(wb.pivotFieldAddItem(0, pivot.index, regionField.index, 'North', true).ok);
    assert.ok(wb.pivotFieldAddItemAt(0, pivot.index, regionField.index, 1, false).ok);
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
    const stale = wb.getValue(99, 0, 0);
    assert.equal(stale.status.ok, false);
    assert.ok(stale.status.message.length > 0);

    // An argument the binding rejects before any C call carries its own
    // message, and a code that names an argument fault rather than the
    // destroyed-handle code.
    const missingSpec = wb.pivotFieldAdd(0, 0);
    assert.equal(missingSpec.status.ok, false);
    assert.equal(missingSpec.status.status, 7001);
    assert.match(missingSpec.status.message, /pivotFieldAdd/);
    assert.notEqual(missingSpec.status.message, stale.status.message);
    assert.equal(missingSpec.status.context, '');

    const badCallback = wb.setIterativeProgress(42);
    assert.equal(badCallback.ok, false);
    assert.equal(badCallback.status, 7001);
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
  // 7000 is reserved for this: an argument fault reports 7001 instead, so
  // "make a new workbook" stays the correct recovery for exactly this code.
  const after = wb.recalc();
  assert.equal(after.ok, false);
  assert.equal(after.status, 7000);
  assert.ok(after.message.length > 0, 'the destroyed-handle status must carry a message');
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
    const kRowRefused = 'row index out of range';
    const kMetricRefused = 'value must be finite and non-negative';
    // These six setters answer with a bare Status, not a `{status}`
    // envelope, so the diagnostics are read off the returned object.
    const refused = (status, expectedMessage, label) => {
      assert.equal(status.ok, false, `${label} should have been refused`);
      assert.equal(status.status, kInvalidArgument, `${label}: ${JSON.stringify(status)}`);
      assert.equal(status.message, expectedMessage, `${label}: ${JSON.stringify(status)}`);
    };

    // One past Excel's last column / row.
    refused(wb.setColumnWidth(0, 0, 16384, 10), kSpanRefused, 'setColumnWidth past the grid');
    refused(wb.setColumnHidden(0, 0, 16384, true), kSpanRefused, 'setColumnHidden past the grid');
    refused(wb.setColumnOutline(0, 0, 16384, 1), kSpanRefused, 'setColumnOutline past the grid');
    refused(wb.setRowHeight(0, 1048576, 10), kRowRefused, 'setRowHeight past the grid');
    refused(wb.setRowHidden(0, 1048576, true), kRowRefused, 'setRowHidden past the grid');
    refused(wb.setRowOutline(0, 1048576, 1), kRowRefused, 'setRowOutline past the grid');

    for (const bad of [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY, -1]) {
      refused(wb.setColumnWidth(0, 0, 2, bad), kMetricRefused, `setColumnWidth(${bad})`);
      refused(wb.setRowHeight(0, 1, bad), kMetricRefused, `setRowHeight(${bad})`);
    }

    // The context names the C ABI entry point that refused the call and
    // the values it saw, which is what makes a rejection actionable.
    const span = wb.setColumnWidth(0, 0, 16384, 10);
    assert.match(span.context, /fm_sheet_set_column_width: first=0 last=16384/);
    const metric = wb.setRowHeight(0, 1, Number.NaN);
    assert.match(metric.context, /fm_sheet_set_row_height: height=/);

    // Zero is a real width and a real height, not a rejected one.
    assert.ok(wb.setColumnWidth(0, 0, 2, 0).ok);
    assert.ok(wb.setRowHeight(0, 1, 0).ok);

    // The last coordinates Excel does address stay accepted -- an
    // off-by-one in either bound would refuse them.
    assert.ok(wb.setColumnWidth(0, 16383, 16383, 10).ok);
    assert.ok(wb.setColumnHidden(0, 16383, 16383, true).ok);
    assert.ok(wb.setColumnOutline(0, 16383, 16383, 1).ok);
    assert.ok(wb.setRowHeight(0, 1048575, 10).ok);
    assert.ok(wb.setRowHidden(0, 1048575, true).ok);
    assert.ok(wb.setRowOutline(0, 1048575, 1).ok);
  } finally {
    wb.dispose();
  }
});

test('setSheetVisibility states veryHidden, which the bool setter cannot', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const read = () => {
      const r = wb.getSheetView(0);
      assert.ok(r.status.ok, `getSheetView: ${JSON.stringify(r.status)}`);
      return r.view;
    };
    assert.equal(read().visibility, mod.SheetVisibility.Visible);
    assert.equal(read().tabHidden, 0);

    assert.ok(wb.setSheetVisibility(0, mod.SheetVisibility.VeryHidden).ok);
    // The two-state view stays consistent: a very-hidden sheet reads as
    // hidden to a caller that knows only the bool, never as visible.
    assert.equal(read().visibility, mod.SheetVisibility.VeryHidden);
    assert.equal(read().tabHidden, 1);

    // Asking for plain "hidden" must not weaken the stronger state.
    assert.ok(wb.setSheetTabHidden(0, true).ok);
    assert.equal(read().visibility, mod.SheetVisibility.VeryHidden);

    // Demotion is the other direction the bool cannot express.
    assert.ok(wb.setSheetVisibility(0, mod.SheetVisibility.Hidden).ok);
    assert.equal(read().visibility, mod.SheetVisibility.Hidden);
    assert.equal(read().tabHidden, 1);

    // Showing the sheet clears either hidden state.
    assert.ok(wb.setSheetTabHidden(0, false).ok);
    assert.equal(read().visibility, mod.SheetVisibility.Visible);
    assert.equal(read().tabHidden, 0);

    // An unknown ordinal is refused and leaves the sheet alone.
    assert.ok(wb.setSheetVisibility(0, mod.SheetVisibility.VeryHidden).ok);
    assert.equal(wb.setSheetVisibility(0, 3).ok, false);
    assert.equal(read().visibility, mod.SheetVisibility.VeryHidden);
  } finally {
    wb.dispose();
  }
});
