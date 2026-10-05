import test from 'node:test';
import { assert, getModule } from './smoke_support.mjs';

test('setIterativeProgress accepts a function and accepts null to clear', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Registration roundtrip; we don't assert the callback fires because
  // a non-iterative recalc never invokes the trampoline. The smoke
  // signal we want is "the addon does not crash on register / clear".
  const cb = () => true;
  const reg = wb.setIterativeProgress(cb);
  assert.ok(reg.ok, `setIterativeProgress(fn): ${JSON.stringify(reg)}`);
  const clr = wb.setIterativeProgress(null);
  assert.ok(clr.ok, `setIterativeProgress(null): ${JSON.stringify(clr)}`);
});

test('setIterativeProgress keeps callbacks isolated per Workbook', async () => {
  const mod = await getModule();
  const first = mod.Workbook.createDefault();
  const second = mod.Workbook.createDefault();
  let firstCalls = 0;
  let secondCalls = 0;

  for (const wb of [first, second]) {
    assert.ok(wb.setIterative(true, 10, 0.001).ok);
    assert.ok(wb.setFormula(0, 0, 0, '=(A1+10)/2').ok);
  }
  assert.ok(
    first.setIterativeProgress(() => {
      firstCalls += 1;
      return true;
    }).ok,
  );
  assert.ok(
    second.setIterativeProgress(() => {
      secondCalls += 1;
      return true;
    }).ok,
  );

  assert.ok(first.recalc().ok);
  assert.ok(firstCalls > 0);
  assert.equal(secondCalls, 0);
});

test('dispose is rejected while this workbook is executing its progress callback', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setIterative(true, 10, 0.001).ok);
  assert.ok(wb.setFormula(0, 0, 0, '=(A1+10)/2').ok);
  let calls = 0;
  assert.ok(
    wb.setIterativeProgress(() => {
      calls += 1;
      assert.throws(() => wb.dispose(), /cannot dispose a Workbook/);
      return true;
    }).ok,
  );
  assert.ok(wb.recalc().ok);
  assert.ok(calls > 0);
  assert.equal(wb.isValid(), true);
  wb.dispose();
});

test('evaluateCfRange returns ok envelope with empty cells for a CF-less workbook', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // The default workbook has no CF rules, so the call should succeed
  // and return an empty cell list. NaN disables `TimePeriod` rules.
  const r = wb.evaluateCfRange(0, 0, 0, 4, 4, Number.NaN);
  assert.ok(r.status.ok, `evaluateCfRange: ${JSON.stringify(r.status)}`);
  assert.ok(Array.isArray(r.cells));
  assert.equal(r.cells.length, 0);
});

test('evaluateCfRange with todaySerial omitted disables timePeriod rules', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // timePeriod rule (type 15) for "today" (period 0) over A1. The cell
  // holds the serial for 1899-12-30, which a 0.0 basis would match.
  assert.ok(wb.setNumber(0, 0, 0, 0).ok);
  const dxf = wb.addDxf({ font: { bold: true } });
  assert.ok(dxf.status.ok, JSON.stringify(dxf.status));
  const add = wb.addConditionalFormat(0, {
    sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
    type: 15,
    timePeriod: 0,
    dxfId: dxf.index,
  });
  assert.ok(add.status.ok, `addConditionalFormat: ${JSON.stringify(add)}`);
  assert.ok(wb.recalc().ok);

  // Omitting todaySerial must behave like passing NaN (the documented
  // disabling value), not like passing 0 -- which is the valid serial for
  // 1899-12-30 and would make the rule match.
  const omitted = wb.evaluateCfRange(0, 0, 0, 0, 0);
  assert.ok(omitted.status.ok, `evaluateCfRange: ${JSON.stringify(omitted.status)}`);
  const explicitNaN = wb.evaluateCfRange(0, 0, 0, 0, 0, Number.NaN);
  assert.ok(explicitNaN.status.ok, `evaluateCfRange: ${JSON.stringify(explicitNaN.status)}`);
  assert.equal(omitted.cells.length, explicitNaN.cells.length);
  assert.equal(omitted.cells.length, 0);

  // A concrete basis of 0 does engage the rule, proving the default is
  // not merely a no-op for this workbook.
  const zeroBasis = wb.evaluateCfRange(0, 0, 0, 0, 0, 0);
  assert.ok(zeroBasis.status.ok, `evaluateCfRange: ${JSON.stringify(zeroBasis.status)}`);
  assert.equal(zeroBasis.cells.length, 1);
});
