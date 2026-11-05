import test from 'node:test';
import { assert, getModule } from './smoke_support.mjs';

test('[Symbol.dispose] is the same function as dispose()', async () => {
  const { Workbook } = await getModule();
  const wb = Workbook.createDefault().value;
  try {
    assert.equal(typeof Symbol.dispose, 'symbol');
    assert.equal(wb[Symbol.dispose], wb.dispose);
  } finally {
    wb.dispose();
  }
});

test('[Symbol.dispose]() releases the handle, repeats safely and leaves the wrapper callable', async () => {
  const { StatusCode, Workbook } = await getModule();
  const wb = Workbook.createDefault().value;
  assert.equal(wb.isValid(), true);
  wb[Symbol.dispose]();
  wb[Symbol.dispose]();
  wb.dispose();
  assert.equal(wb.isValid(), false);
  const after = wb.getSheetCount();
  assert.equal(after.status.ok, false);
  assert.equal(after.status.code, StatusCode.BindingInvalidHandle);
});

test('dispose() and [Symbol.dispose]() are refused inside the iterative progress callback', async () => {
  const { Workbook } = await getModule();
  for (const release of [(wb) => wb.dispose(), (wb) => wb[Symbol.dispose]()]) {
    const wb = Workbook.createDefault().value;
    assert.ok(wb.setIterative(true, 100, 0.0001).ok);
    assert.ok(wb.setFormula(0, 0, 0, '=(A1+10)/2').ok);
    let caught = null;
    assert.ok(
      wb.setIterativeProgress(() => {
        try {
          release(wb);
        } catch (e) {
          caught = e;
        }
        return false;
      }).ok,
    );
    wb.recalc();
    assert.ok(caught instanceof TypeError, 'release inside the callback must throw');
    assert.match(String(caught.message), /cannot dispose Workbook during an active method call/);
    assert.ok(wb.isValid(), 'the refused release leaves the handle intact');
    wb.setIterativeProgress(null);
    wb.dispose();
  }
});

test('the default export carries every table and helper exported by name', async () => {
  const mod = await getModule();
  const keys = Object.keys(mod).filter((key) => key !== 'default');
  for (const key of keys) {
    assert.equal(mod.default[key], mod[key], `${key} on the default export`);
  }
});
