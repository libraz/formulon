import test from 'node:test';
import { assert, expectFailure, getModule, makeIterativeWorkbook, threadsEntry, VAL } from './smoke_support.mjs';

test('Workbook.createDefault produces a valid single-sheet workbook', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    assert.equal(wb.disposed, false);
    assert.equal(wb.listSheets().length, 1);
  } finally {
    wb.dispose();
  }
});

test('print settings: raw XML round-trips and rejects malformed fragments', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    // An absent element reads back as null.
    assert.equal(wb.sheet(0).getPageSetupXml(), null);

    const fragment = '<pageSetup paperSize="9" orientation="portrait" scale="85"/>';
    wb.sheet(0).setPageSetupXml(fragment);
    assert.equal(wb.sheet(0).getPageSetupXml(), fragment);

    for (const bad of [
      '<pageSetup/><pageSetup/>',
      '<pageMargins left="1"/>',
      '<pageSetup orientation="portrait"',
      '<x:pageSetup/>',
    ]) {
      expectFailure(() => wb.sheet(0).setPageSetupXml(bad));
    }
    // A rejected set leaves the stored fragment untouched.
    assert.equal(wb.sheet(0).getPageSetupXml(), fragment);

    wb.sheet(0).setPageSetupXml(null);
    assert.equal(wb.sheet(0).getPageSetupXml(), null);
    const cleared = wb.sheet(0).getPageSetup();
    assert.equal(cleared.scale, 100);
    assert.equal(cleared.scaleStated, false);
  } finally {
    wb.dispose();
  }
});

test('print settings: typed patch leaves unstated attributes alone', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    wb.sheet(0).setPageSetupXml('<pageSetup paperSize="9" orientation="portrait" horizontalDpi="600" copies="3"/>');
    wb.sheet(0).setPageSetup({ orientation: Module.PageOrientation.Landscape });

    const xml = wb.sheet(0).getPageSetupXml();
    assert.match(xml, /orientation="landscape"/);
    // Attributes the engine does not model must survive the patch, or the
    // "open a template and change one thing" route loses data.
    assert.match(xml, /horizontalDpi="600"/);
    assert.match(xml, /copies="3"/);
    assert.match(xml, /paperSize="9"/);

    // A print scale outside Excel's range is rejected, not clamped.
    expectFailure(() => wb.sheet(0).setPageSetup({ scale: 9 }));
    expectFailure(() => wb.sheet(0).setPageSetup({ scale: 401 }));
    wb.sheet(0).setPageSetup({ scale: 400 });
  } finally {
    wb.dispose();
  }
});

test('print settings: a change reaches paginate with no save cycle', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    for (let row = 0; row < 200; row += 1) {
      for (let col = 0; col < 20; col += 1) {
        wb.sheet(0).setNumber(row, col, 1);
      }
    }
    const portrait = wb.sheet(0).getPagination().pageCount;
    wb.sheet(0).setPageSetup({ orientation: Module.PageOrientation.Landscape });
    assert.notEqual(wb.sheet(0).getPagination().pageCount, portrait);
  } finally {
    wb.dispose();
  }
});

test('print settings: an authored report survives save and load', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  let bytes;
  try {
    wb.sheet(0).setText(0, 0, '売上');
    wb.sheet(0).setPageSetup({
      paperSize: 9,
      orientation: Module.PageOrientation.Portrait,
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
    });
    wb.sheet(0).setPageMargins({ left: 0.5, right: 0.5, top: 0.8, bottom: 0.8 });
    wb.sheet(0).setPrintOptions({ horizontalCentered: true });
    wb.sheet(0).setHeaderFooter({ oddHeader: '&C月次報告', oddFooter: '&R&P / &N' });
    wb.sheet(0).setPrintArea([{ firstRow: 0, firstCol: 0, lastRow: 79, lastCol: 5 }]);
    wb.sheet(0).setPrintTitles({ rows: { first: 0, last: 1 }, columns: null });
    wb.sheet(0).addRowBreak(39, true);

    const saved = wb.save();
    bytes = saved;
  } finally {
    wb.dispose();
  }

  const reloaded = Module.Workbook.load(bytes);
  try {
    const setup = reloaded.sheet(0).getPageSetup();
    assert.equal(setup.paperSize, 9);
    assert.equal(setup.orientation, Module.PageOrientation.Portrait);
    assert.equal(setup.fitToPage, true);
    assert.equal(setup.fitToHeight, 0);
    assert.equal(reloaded.sheet(0).getPageMargins().left, 0.5);
    assert.deepEqual(reloaded.sheet(0).getPrintArea(), [{ firstRow: 0, firstCol: 0, lastRow: 79, lastCol: 5 }]);
    assert.deepEqual(reloaded.sheet(0).getPrintTitles(), { rows: { first: 0, last: 1 }, columns: null });
    // The header code is stored escaped and decodes back to `&C...`.
    assert.match(reloaded.sheet(0).getHeaderFooterXml(), /<oddHeader>&amp;C月次報告<\/oddHeader>/);
    const breaks = reloaded.sheet(0).listRowBreaks();
    assert.deepEqual(
      breaks.map((b) => b.id),
      [39],
    );
    assert.equal(breaks[0].manual, true);
  } finally {
    reloaded.dispose();
  }
});

test('setRangeXfIndex applies one xf across a rectangle', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    const xf = wb.addCellXf({
      fontIndex: 0,
      fillIndex: 0,
      borderIndex: 0,
      numFmtId: 0,
      wrapText: true,
      hasWrapText: true,
    });
    wb.sheet(0).setRangeXfIndex({ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 2 }, xf);
    // Cells that held nothing are materialised so the ruled box renders.
    for (const [row, col] of [
      [0, 0],
      [1, 1],
      [2, 2],
    ]) {
      assert.equal(wb.sheet(0).getCellXfIndex(row, col), xf);
    }
    assert.equal(wb.sheet(0).getCellXfIndex(3, 3), 0);
  } finally {
    wb.dispose();
  }
});

test('sheet layout getters report an absent width, height or style as null', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    wb.sheet(0).setColumnWidth(0, 0, 0);
    wb.sheet(0).setColumnHidden(1, 1, true);
    wb.sheet(0).setRowHeight(3, 3, 30);

    const columns = wb.sheet(0).listColumnOverrides();
    const widthZero = columns.find((column) => column.first === 0 && column.last === 0);
    assert.ok(widthZero);
    assert.equal(widthZero.width, 0);
    assert.equal(widthZero.styleXf, null);
    assert.equal(widthZero.hidden, false);
    const hidden = columns.find((column) => column.first === 1 && column.last === 1);
    assert.ok(hidden);
    assert.equal(hidden.hidden, true);
    assert.equal(hidden.width, null);

    const rows = wb.sheet(0).listRowOverrides();
    const row = rows.find((entry) => entry.row === 3);
    assert.ok(row);
    assert.equal(row.height, 30);
    assert.equal(row.styleXf, null);
    assert.equal(row.collapsed, false);
  } finally {
    wb.dispose();
  }
});

test('Unicode sheet identity rejects folded duplicate and permits casing rename', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createEmpty();
  try {
    wb.addSheet('Ä');
    wb.addSheet('Ö');
    expectFailure(() => wb.addSheet('ä'));
    expectFailure(() => wb.renameSheet(1, 'ä'));
    assert.equal(wb.sheet(0).name, 'Ä');
    assert.equal(wb.sheet(1).name, 'Ö');
    wb.renameSheet(0, 'ä');
    assert.equal(wb.sheet(0).name, 'ä');
    assert.equal(wb.sheet(1).name, 'Ö');
  } finally {
    wb.dispose();
  }
});

test('setNumber + getValue round-trips bit-exactly', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    const x = 3.141592653589793;
    wb.sheet(0).setNumber(0, 0, x);
    wb.recalc();
    const r = wb.sheet(0).getValue(0, 0);
    assert.equal(r.kind, VAL.NUMBER);
    assert.equal(r.number, x);
  } finally {
    wb.dispose();
  }
});

test('setFormula + recalc computes 41 + 1 == 42 in B1', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    wb.sheet(0).setNumber(0, 0, 41);
    wb.sheet(0).setFormula(0, 1, '=A1+1');
    wb.recalc();
    const b1 = wb.sheet(0).getValue(0, 1);
    assert.equal(b1.kind, VAL.NUMBER);
    assert.equal(b1.number, 42);
  } finally {
    wb.dispose();
  }
});

test('Workbook.recalcParallel evaluates a wide DAG and reports bounded telemetry', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  const branchCount = 32;
  try {
    wb.sheet(0).setNumber(0, 0, 1);
    for (let i = 0; i < branchCount; i += 1) {
      const row = i + 1;
      wb.sheet(0).setFormula(row, 1, `=A1+${i + 2}`);
      wb.sheet(0).setFormula(row, 2, `=B${row + 1}*2`);
    }

    const parallel = wb.recalcParallel(4);
    assert.equal(typeof parallel.cellsEvaluated, 'number');
    assert.equal(typeof parallel.sccsProcessed, 'number');
    assert.equal(typeof parallel.parallelSteps, 'number');
    assert.ok(parallel.cellsEvaluated > 0);
    assert.ok(parallel.sccsProcessed > 0);
    if (!threadsEntry) {
      // Without pthreads every launch is refused, so the pass runs serially.
      assert.equal(parallel.workerThreadsStarted, 0);
    }
    if (parallel.workerThreadsStarted <= 1) {
      // OS launch refusal or a partial launch of one worker is a documented
      // successful serial degradation.
      assert.equal(parallel.workerThreadsUsed, 0);
      assert.equal(parallel.parallelSteps, 0);
      assert.ok(parallel.serialFallbackSteps > 0, `stats=${JSON.stringify(parallel)}`);
    } else {
      assert.ok(parallel.parallelSteps > 0, `stats=${JSON.stringify(parallel)}`);
      assert.ok(parallel.workerThreadsStarted >= 2);
      assert.ok(parallel.workerThreadsStarted <= 4);
      assert.ok(parallel.workerThreadsUsed > 0);
      assert.ok(parallel.workerThreadsUsed <= parallel.workerThreadsStarted);
    }

    const first = wb.sheet(0).getValue(1, 1);
    const last = wb.sheet(0).getValue(branchCount, 2);
    assert.equal(first.kind, VAL.NUMBER);
    assert.equal(last.kind, VAL.NUMBER);
    assert.equal(first.number, 3);
    assert.equal(last.number, (1 + branchCount + 1) * 2);

    wb.sheet(0).setNumber(0, 0, 5);
    const callerOnly = wb.recalcParallel(1);
    assert.equal(callerOnly.parallelSteps, 0);
    assert.ok(callerOnly.serialFallbackSteps > 0);
    assert.equal(callerOnly.workerThreadsStarted, 0);
    assert.equal(callerOnly.workerThreadsUsed, 0);
    assert.equal(wb.sheet(0).getValue(branchCount, 2).number, (5 + branchCount + 1) * 2);

    expectFailure(() => wb.recalcParallel(9), Module.StatusCode.InvalidArgument);

    for (const invalidThreadCount of [
      -0.5,
      1.5,
      Number.NaN,
      Number.POSITIVE_INFINITY,
      Number.NEGATIVE_INFINITY,
      2 ** 32 + 1,
      null,
      undefined,
    ]) {
      const expected = invalidThreadCount === null || invalidThreadCount === undefined ? TypeError : RangeError;
      assert.throws(() => wb.recalcParallel(invalidThreadCount), expected, String(invalidThreadCount));
    }
  } finally {
    wb.dispose();
  }
});

// Status ordinals the callback contract is stated in terms of; see
// `src/utils/error.h`. Spelled out here because the binding does not
// export a status enum.
const STATUS = Object.freeze({
  RECALC_REENTRANT: 4005,
});

test('a throwing iterative progress callback aborts the solve and surfaces its own exception', async () => {
  const Module = await getModule();
  const wb = makeIterativeWorkbook(Module);
  try {
    let calls = 0;
    const thrown = new Error('callback failure');
    wb.setIterativeProgress(() => {
      calls += 1;
      throw thrown;
    });

    // The binding never lets an exception cross back into WASM, where no
    // landing pad exists to restore the engine's RAII state: the throw is
    // re-thrown from the outer call once the native frames have returned.
    assert.throws(
      () => wb.recalc(),
      (e) => e === thrown,
    );
    assert.equal(calls, 1, 'the solve must stop at the first throw');

    // The workbook is still readable, and the recalc re-entry guard was
    // released on the way out.
    const mid = wb.sheet(0).getValue(0, 0);
    assert.equal(mid.kind, VAL.NUMBER);

    try {
      wb.recalc();
    } catch (e) {
      assert.notEqual(e.code, STATUS.RECALC_REENTRANT, 'the re-entry flag stayed set');
    }

    wb.setIterativeProgress(null);
    wb.sheet(0).setFormula(0, 0, '=(A1+10)/2');
    wb.recalc();
    const converged = wb.sheet(0).getValue(0, 0);
    assert.equal(converged.kind, VAL.NUMBER);
    assert.ok(Math.abs(converged.number - 10) < 0.001, `A1=${converged.number}`);
  } finally {
    wb.setIterativeProgress(null);
    wb.dispose();
  }
});

test('delete() from inside its own progress callback throws instead of crashing', async () => {
  const Module = await getModule();
  const wb = makeIterativeWorkbook(Module);
  let calls = 0;
  let caught = null;
  wb.setIterativeProgress(() => {
    calls += 1;
    try {
      wb.dispose();
    } catch (e) {
      caught = e;
    }
    return false;
  });

  wb.recalc();
  assert.equal(calls, 1, 'the callback must run exactly once');
  assert.ok(caught instanceof Error, 'delete() inside the callback must throw, not free the handle');
  assert.match(String(caught.message), /cannot dispose Workbook during an active method call/);
  // The handle survived: delete() never reached fm_workbook_destroy, so
  // the workbook is still usable after the (deliberately cancelled) solve.
  assert.equal(wb.disposed, false);

  wb.setIterativeProgress(null);
  wb.dispose();
});

test('an uncaught delete() inside the progress callback aborts the solve like any other throw', async () => {
  const Module = await getModule();
  const wb = makeIterativeWorkbook(Module);
  let calls = 0;
  wb.setIterativeProgress(() => {
    calls += 1;
    wb.dispose(); // not caught here -- propagates the same as a plain `throw`
  });

  assert.throws(
    () => wb.recalc(),
    (e) => e instanceof TypeError && /cannot dispose Workbook during an active method call/.test(e.message),
  );
  assert.equal(calls, 1);
  assert.equal(wb.disposed, false);

  wb.setIterativeProgress(null);
  wb.dispose();
});

test('an iterative progress callback returning a non-boolean is read by truthiness', async () => {
  const Module = await getModule();
  for (const [returned, expectedSweeps] of [
    ['keep going', 0],
    [0, 1],
    [undefined, 0],
  ]) {
    const wb = makeIterativeWorkbook(Module);
    try {
      let calls = 0;
      wb.setIterativeProgress(() => {
        calls += 1;
        return returned;
      });
      wb.recalc();
      // A falsy return aborts after the first sweep; anything else lets
      // the solve run to convergence.
      if (expectedSweeps === 1) {
        assert.equal(calls, 1, `returned=${String(returned)}`);
      } else {
        assert.ok(calls > 1, `returned=${String(returned)}, calls=${calls}`);
      }
    } finally {
      wb.setIterativeProgress(null);
      wb.dispose();
    }
  }
});
