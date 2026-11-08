import test from 'node:test';
import { assert, getModule, loadStagedModule } from './smoke_support.mjs';

test('[Symbol.dispose] is the same function as dispose()', async () => {
  const { Workbook } = await getModule();
  const wb = Workbook.createDefault();
  try {
    assert.equal(typeof Symbol.dispose, 'symbol');
    assert.equal(wb[Symbol.dispose], wb.dispose);
  } finally {
    wb.dispose();
  }
});

test('[Symbol.dispose]() releases the handle, repeats safely and leaves the wrapper callable', async () => {
  const { StatusCode, Workbook } = await getModule();
  const wb = Workbook.createDefault();
  assert.equal(wb.disposed, false);
  wb[Symbol.dispose]();
  wb[Symbol.dispose]();
  wb.dispose();
  assert.equal(wb.disposed, true);
  assert.throws(
    () => wb.listSheets(),
    (e) => e.name === 'FormulonError' && e.code === StatusCode.BindingInvalidHandle,
  );
});

test('dispose() and [Symbol.dispose]() are refused inside the iterative progress callback', async () => {
  const { Workbook } = await getModule();
  for (const release of [(wb) => wb.dispose(), (wb) => wb[Symbol.dispose]()]) {
    const wb = Workbook.createDefault();
    wb.updateIterative({ enabled: true, maxIterations: 100, maxChange: 0.0001 });
    wb.sheet(0).setFormula(0, 0, '=(A1+10)/2');
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
    assert.equal(caught?.name, 'FormulonError', 'release inside the callback must throw');
    assert.equal(caught.code, 4005);
    assert.match(String(caught.message), /cannot dispose Workbook during an active method call/);
    assert.equal(wb.disposed, false, 'the refused release leaves the handle intact');
    wb.setIterativeProgress(null);
    wb.dispose();
  }
});

test('the Formulon object carries every table and helper exported by name', async () => {
  const mod = await loadStagedModule();
  const formulon = await mod.createFormulon();
  const keys = Object.keys(mod).filter((key) => key !== 'default' && key !== 'createFormulon');
  for (const key of keys) {
    assert.equal(formulon[key], mod[key], `${key} on the Formulon object`);
  }
});
