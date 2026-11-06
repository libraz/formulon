import test from 'node:test';
import { assert, expectFailure, getModule, makeIterativeWorkbook, threadsEntry, VAL } from './smoke_support.mjs';

test('Workbook.createDefault produces a valid single-sheet workbook', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    assert.equal(wb.disposed, false);
    assert.equal(wb.getSheetCount(), 1);
  } finally {
    wb.delete();
  }
});

test('print settings: raw XML round-trips and rejects malformed fragments', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    // An absent element reads back as the empty string, never null.
    assert.equal(wb.getPageSetupXml(0), '');

    const fragment = '<pageSetup paperSize="9" orientation="portrait" scale="85"/>';
    wb.setPageSetupXml(0, fragment);
    assert.equal(wb.getPageSetupXml(0), fragment);

    for (const bad of [
      '<pageSetup/><pageSetup/>',
      '<pageMargins left="1"/>',
      '<pageSetup orientation="portrait"',
      '<x:pageSetup/>',
    ]) {
      expectFailure(() => wb.setPageSetupXml(0, bad));
    }
    // A rejected set leaves the stored fragment untouched.
    assert.equal(wb.getPageSetupXml(0), fragment);

    wb.setPageSetupXml(0, '');
    assert.equal(wb.getPageSetupXml(0), '');
    const cleared = wb.getPageSetup(0);
    assert.equal(cleared.scale, 100);
    assert.equal(cleared.scaleStated, false);
  } finally {
    wb.delete();
  }
});

test('print settings: typed patch leaves unstated attributes alone', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    wb.setPageSetupXml(0, '<pageSetup paperSize="9" orientation="portrait" horizontalDpi="600" copies="3"/>');
    wb.setPageSetup(0, { orientation: 2 });

    const xml = wb.getPageSetupXml(0);
    assert.match(xml, /orientation="landscape"/);
    // Attributes the engine does not model must survive the patch, or the
    // "open a template and change one thing" route loses data.
    assert.match(xml, /horizontalDpi="600"/);
    assert.match(xml, /copies="3"/);
    assert.match(xml, /paperSize="9"/);

    // A print scale outside Excel's range is rejected, not clamped.
    expectFailure(() => wb.setPageSetup(0, { scale: 9 }));
    expectFailure(() => wb.setPageSetup(0, { scale: 401 }));
    wb.setPageSetup(0, { scale: 400 });
  } finally {
    wb.delete();
  }
});

test('print settings: a change reaches paginate with no save cycle', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    for (let row = 0; row < 200; row += 1) {
      for (let col = 0; col < 20; col += 1) {
        wb.setNumber(0, row, col, 1);
      }
    }
    const portrait = wb.getPagination(0).pageCount;
    wb.setPageSetup(0, { orientation: 2 });
    assert.notEqual(wb.getPagination(0).pageCount, portrait);
  } finally {
    wb.delete();
  }
});

test('print settings: an authored report survives save and load', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  let bytes;
  try {
    wb.setText(0, 0, 0, '売上');
    wb.setPageSetup(0, {
      paperSize: 9,
      orientation: 1,
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
    });
    wb.setPageMargins(0, { left: 0.5, right: 0.5, top: 0.8, bottom: 0.8 });
    wb.setPrintOptions(0, { horizontalCentered: true });
    wb.setHeaderFooter(0, { oddHeader: '&C月次報告', oddFooter: '&R&P / &N' });
    wb.setPrintArea(0, [{ firstRow: 0, firstCol: 0, lastRow: 79, lastCol: 5 }]);
    wb.setPrintTitles(0, '1:2', '');
    wb.addRowBreak(0, 39, true);

    const saved = wb.save();
    bytes = saved;
  } finally {
    wb.delete();
  }

  const reloaded = Module.Workbook.load(bytes);
  try {
    const setup = reloaded.getPageSetup(0);
    assert.equal(setup.paperSize, 9);
    assert.equal(setup.orientation, 1);
    assert.equal(setup.fitToPage, true);
    assert.equal(setup.fitToHeight, 0);
    assert.equal(reloaded.getPageMargins(0).left, 0.5);
    assert.deepEqual(reloaded.getPrintArea(0), [{ firstRow: 0, firstCol: 0, lastRow: 79, lastCol: 5 }]);
    const titles = reloaded.getPrintTitles(0);
    assert.equal(titles.repeatRows, '1:2');
    assert.equal(titles.repeatCols, '');
    // The header code is stored escaped and decodes back to `&C...`.
    assert.match(reloaded.getHeaderFooterXml(0), /<oddHeader>&amp;C月次報告<\/oddHeader>/);
    const breaks = reloaded.listRowBreaks(0);
    assert.deepEqual(
      breaks.map((b) => b.id),
      [39],
    );
    assert.equal(breaks[0].manual, true);
  } finally {
    reloaded.delete();
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
    wb.setRangeXfIndex(0, { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 2 }, xf);
    // Cells that held nothing are materialised so the ruled box renders.
    for (const [row, col] of [
      [0, 0],
      [1, 1],
      [2, 2],
    ]) {
      assert.equal(wb.getCellXfIndex(0, row, col), xf);
    }
    assert.equal(wb.getCellXfIndex(0, 3, 3), 0);
  } finally {
    wb.delete();
  }
});

test('sheet layout getters preserve explicit width presence flags', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    wb.setColumnWidth(0, 0, 0, 0);
    wb.setColumnHidden(0, 1, 1, true);
    wb.setRowHeight(0, 3, 3, 30);

    const columns = wb.listColumnOverrides(0);
    const widthZero = columns.find((column) => column.first === 0 && column.last === 0);
    assert.ok(widthZero);
    assert.equal(widthZero.width, 0);
    assert.equal(typeof widthZero.hasWidth, 'number');
    assert.equal(widthZero.hasWidth, 1);
    assert.equal(typeof widthZero.hasStyle, 'number');
    assert.equal(widthZero.hasStyle, 0);
    const hidden = columns.find((column) => column.first === 1 && column.last === 1);
    assert.ok(hidden);
    assert.equal(hidden.hidden, 1);
    assert.equal(typeof hidden.hasWidth, 'number');
    assert.equal(hidden.hasWidth, 0);

    const rows = wb.listRowOverrides(0);
    const row = rows.find((entry) => entry.row === 3);
    assert.ok(row);
    assert.equal(row.height, 30);
    assert.equal(row.hasStyle, 0);
    assert.equal(row.styleXf, 0);
  } finally {
    wb.delete();
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
    assert.equal(wb.getSheetName(0), 'Ä');
    assert.equal(wb.getSheetName(1), 'Ö');
    wb.renameSheet(0, 'ä');
    const sn = wb.getSheetName(0);
    assert.equal(sn, 'ä');
    assert.equal(wb.getSheetName(1), 'Ö');
  } finally {
    wb.delete();
  }
});

test('setNumber + getValue round-trips bit-exactly', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    const x = 3.141592653589793;
    wb.setNumber(0, 0, 0, x);
    wb.recalc();
    const r = wb.getValue(0, 0, 0);
    assert.equal(r.kind, VAL.NUMBER);
    assert.equal(r.number, x);
  } finally {
    wb.delete();
  }
});

test('setFormula + recalc computes 41 + 1 == 42 in B1', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    wb.setNumber(0, 0, 0, 41);
    wb.setFormula(0, 0, 1, '=A1+1');
    wb.recalc();
    const b1 = wb.getValue(0, 0, 1);
    assert.equal(b1.kind, VAL.NUMBER);
    assert.equal(b1.number, 42);
  } finally {
    wb.delete();
  }
});

test('Workbook.recalcParallel evaluates a wide DAG and reports bounded telemetry', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  const branchCount = 32;
  try {
    wb.setNumber(0, 0, 0, 1);
    for (let i = 0; i < branchCount; i += 1) {
      const row = i + 1;
      wb.setFormula(0, row, 1, `=A1+${i + 2}`);
      wb.setFormula(0, row, 2, `=B${row + 1}*2`);
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

    const first = wb.getValue(0, 1, 1);
    const last = wb.getValue(0, branchCount, 2);
    assert.equal(first.kind, VAL.NUMBER);
    assert.equal(last.kind, VAL.NUMBER);
    assert.equal(first.number, 3);
    assert.equal(last.number, (1 + branchCount + 1) * 2);

    wb.setNumber(0, 0, 0, 5);
    const callerOnly = wb.recalcParallel(1);
    assert.equal(callerOnly.parallelSteps, 0);
    assert.ok(callerOnly.serialFallbackSteps > 0);
    assert.equal(callerOnly.workerThreadsStarted, 0);
    assert.equal(callerOnly.workerThreadsUsed, 0);
    assert.equal(wb.getValue(0, branchCount, 2).number, (5 + branchCount + 1) * 2);

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
    wb.delete();
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
    const mid = wb.getValue(0, 0, 0);
    assert.equal(mid.kind, VAL.NUMBER);

    try {
      wb.recalc();
    } catch (e) {
      assert.notEqual(e.code, STATUS.RECALC_REENTRANT, 'the re-entry flag stayed set');
    }

    wb.setIterativeProgress(null);
    wb.setFormula(0, 0, 0, '=(A1+10)/2');
    wb.recalc();
    const converged = wb.getValue(0, 0, 0);
    assert.equal(converged.kind, VAL.NUMBER);
    assert.ok(Math.abs(converged.number - 10) < 0.001, `A1=${converged.number}`);
  } finally {
    wb.setIterativeProgress(null);
    wb.delete();
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
      wb.delete();
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
  wb.delete();
});

test('an uncaught delete() inside the progress callback aborts the solve like any other throw', async () => {
  const Module = await getModule();
  const wb = makeIterativeWorkbook(Module);
  let calls = 0;
  wb.setIterativeProgress(() => {
    calls += 1;
    wb.delete(); // not caught here -- propagates the same as a plain `throw`
  });

  assert.throws(
    () => wb.recalc(),
    (e) => e instanceof TypeError && /cannot dispose Workbook during an active method call/.test(e.message),
  );
  assert.equal(calls, 1);
  assert.equal(wb.disposed, false);

  wb.setIterativeProgress(null);
  wb.delete();
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
      wb.delete();
    }
  }
});
