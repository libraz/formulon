import test from 'node:test';
import { assert, expectFailure, getModule, loadStagedModule, makeIterativeWorkbook } from './smoke_support.mjs';

test('[Symbol.dispose] is the same function as dispose()', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    assert.equal(typeof Symbol.dispose, 'symbol');
    assert.equal(wb[Symbol.dispose], wb.dispose);
  } finally {
    wb.dispose();
  }
});

test('[Symbol.dispose]() releases the handle, repeats safely and leaves the wrapper callable', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  assert.equal(wb.disposed, false);
  wb[Symbol.dispose]();
  wb[Symbol.dispose]();
  wb.dispose();
  assert.equal(wb.disposed, true);
  expectFailure(() => wb.getSheetCount(), Module.StatusCode.BindingInvalidHandle);
});

test('dispose() and [Symbol.dispose]() are refused inside the iterative progress callback', async () => {
  const Module = await getModule();
  for (const release of [(wb) => wb.dispose(), (wb) => wb[Symbol.dispose]()]) {
    const wb = makeIterativeWorkbook(Module);
    let caught = null;
    wb.setIterativeProgress(() => {
      try {
        release(wb);
      } catch (e) {
        caught = e;
      }
      return false;
    });
    wb.recalc();
    assert.ok(caught instanceof TypeError, 'release inside the callback must throw');
    assert.match(String(caught.message), /cannot dispose Workbook during an active method call/);
    assert.equal(wb.disposed, false);
    wb.setIterativeProgress(null);
    wb.dispose();
  }
});

test('the module object carries every table and helper exported by name', async () => {
  const Module = await getModule();
  const named = await loadStagedModule();
  const keys = Object.keys(named).filter((key) => key !== 'default' && key !== 'createFormulon');
  assert.ok(keys.includes('StatusCode') && keys.includes('ErrorCode') && keys.includes('PIVOT_VALUES_FIELD'));
  for (const key of keys) {
    assert.equal(Module[key], named[key], `${key} on the module object`);
  }
});
