import assert from 'node:assert/strict';
import { PIVOT, VAL } from './smoke_support.mjs';

export function registerPivotRegressions(Module, test) {
  test('pivotFieldAddItemAt hides the blank item an empty label cannot name', () => {
    const wb = Module.Workbook.createDefault();
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
          if (cell.kind === PIVOT.DATA && cell.value.kind === VAL.NUMBER) {
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

      const rejected = wb.createTable({ sheetIndex: 0, ref: 'A1:B2', name: 'T' });
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
      for (const list of [
        wb.getMerges(99),
        wb.getComments(99),
        wb.getHyperlinks(99),
        wb.getValidations(99),
        wb.getConditionalFormats(99),
      ]) {
        assert.ok(Array.isArray(list));
        assert.equal(list.length, 0);
        assert.equal(list.status.ok, false);
      }
      const links = wb.getExternalLinks();
      assert.ok(Array.isArray(links));
      assert.equal(links.status.ok, true);
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
          out = wb.evaluateFormulaText(0, 1, 0, formula);
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
