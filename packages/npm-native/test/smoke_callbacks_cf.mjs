import test from 'node:test';
import { assert, getModule } from './smoke_support.mjs';

test('setIterativeProgress accepts a function and accepts null to clear', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Registration roundtrip; we don't assert the callback fires because
  // a non-iterative recalc never invokes the trampoline. The smoke
  // signal we want is "the addon does not crash on register / clear".
  const cb = () => true;
  wb.setIterativeProgress(cb);
  wb.setIterativeProgress(null);
});

test('setIterativeProgress keeps callbacks isolated per Workbook', async () => {
  const mod = await getModule();
  const first = mod.Workbook.createDefault();
  const second = mod.Workbook.createDefault();
  let firstCalls = 0;
  let secondCalls = 0;

  for (const wb of [first, second]) {
    wb.setIterative({ enabled: true, maxIterations: 10, maxChange: 0.001 });
    wb.sheet(0).setFormula(0, 0, '=(A1+10)/2');
  }
  first.setIterativeProgress(() => {
    firstCalls += 1;
    return true;
  });
  second.setIterativeProgress(() => {
    secondCalls += 1;
    return true;
  });

  first.recalc();
  assert.ok(firstCalls > 0);
  assert.equal(secondCalls, 0);
});

test('dispose is rejected while this workbook is executing its progress callback', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setIterative({ enabled: true, maxIterations: 10, maxChange: 0.001 });
  wb.sheet(0).setFormula(0, 0, '=(A1+10)/2');
  let calls = 0;
  wb.setIterativeProgress(() => {
    calls += 1;
    assert.throws(() => wb.dispose(), /cannot dispose Workbook during an active method call/);
    return true;
  });
  wb.recalc();
  assert.ok(calls > 0);
  assert.equal(wb.disposed, false);
  wb.dispose();
});

test('evaluateCfRange returns an empty list for a CF-less workbook', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // The default workbook has no CF rules, so the call should succeed
  // and return an empty cell list. NaN disables `TimePeriod` rules.
  const r = wb.sheet(0).evaluateCfRange({ firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 4 }, Number.NaN);
  assert.ok(Array.isArray(r));
  assert.equal(r.length, 0);
});

test('evaluateCfRange requires todaySerial; NaN disables timePeriod rules', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // timePeriod rule (type 15) for "today" (period 0) over A1. The cell
  // holds the serial for 1899-12-30, which a 0.0 basis would match.
  wb.sheet(0).setNumber(0, 0, 0);
  const dxf = wb.addDxf({ font: { bold: true } });
  wb.sheet(0).addCfRule({
    ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
    type: 15,
    timePeriod: 0,
    dxfId: dxf,
  });
  wb.recalc();

  // todaySerial is required: an omitted one is rejected rather than read as
  // 0, the valid serial for 1899-12-30 that would make the rule match. NaN
  // is the documented disabling value.
  assert.throws(() => wb.sheet(0).evaluateCfRange({ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }), TypeError);
  const explicitNaN = wb.sheet(0).evaluateCfRange({ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }, Number.NaN);
  assert.equal(explicitNaN.length, 0);

  // A concrete basis of 0 does engage the rule, proving the default is
  // not merely a no-op for this workbook.
  const zeroBasis = wb.sheet(0).evaluateCfRange({ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }, 0);
  assert.equal(zeroBasis.length, 1);
});
