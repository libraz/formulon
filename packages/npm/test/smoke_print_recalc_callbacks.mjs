import test from 'node:test';
import { assert, getModule, makeIterativeWorkbook, threadsEntry, VAL } from './smoke_support.mjs';

test('Workbook.createDefault produces a valid single-sheet workbook', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    assert.ok(wb.isValid());
    assert.equal(wb.sheetCount().value, 1);
  } finally {
    wb.delete();
  }
});

test('print settings: raw XML round-trips and rejects malformed fragments', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    // An absent element reads back as the empty string, never null.
    assert.equal(wb.getSheetPageSetupXml(0).xml, '');

    const fragment = '<pageSetup paperSize="9" orientation="portrait" scale="85"/>';
    assert.ok(wb.setSheetPageSetupXml(0, fragment).ok);
    assert.equal(wb.getSheetPageSetupXml(0).xml, fragment);

    for (const bad of [
      '<pageSetup/><pageSetup/>',
      '<pageMargins left="1"/>',
      '<pageSetup orientation="portrait"',
      '<x:pageSetup/>',
    ]) {
      assert.equal(wb.setSheetPageSetupXml(0, bad).ok, false, `expected rejection for ${bad}`);
    }
    // A rejected set leaves the stored fragment untouched.
    assert.equal(wb.getSheetPageSetupXml(0).xml, fragment);

    assert.ok(wb.setSheetPageSetupXml(0, '').ok);
    assert.equal(wb.getSheetPageSetupXml(0).xml, '');
    const cleared = wb.getSheetPageSetup(0);
    assert.ok(cleared.status.ok);
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
    assert.ok(
      wb.setSheetPageSetupXml(0, '<pageSetup paperSize="9" orientation="portrait" horizontalDpi="600" copies="3"/>').ok,
    );
    assert.ok(wb.setSheetPageSetup(0, { orientation: 2 }).ok);

    const xml = wb.getSheetPageSetupXml(0).xml;
    assert.match(xml, /orientation="landscape"/);
    // Attributes the engine does not model must survive the patch, or the
    // "open a template and change one thing" route loses data.
    assert.match(xml, /horizontalDpi="600"/);
    assert.match(xml, /copies="3"/);
    assert.match(xml, /paperSize="9"/);

    // A print scale outside Excel's range is rejected, not clamped.
    assert.equal(wb.setSheetPageSetup(0, { scale: 9 }).ok, false);
    assert.equal(wb.setSheetPageSetup(0, { scale: 401 }).ok, false);
    assert.ok(wb.setSheetPageSetup(0, { scale: 400 }).ok);
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
        assert.ok(wb.setNumber(0, row, col, 1).ok);
      }
    }
    const portrait = wb.paginate(0).pageCount;
    assert.ok(wb.setSheetPageSetup(0, { orientation: 2 }).ok);
    assert.notEqual(wb.paginate(0).pageCount, portrait);
  } finally {
    wb.delete();
  }
});

test('print settings: an authored report survives save and load', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  let bytes;
  try {
    assert.ok(wb.setText(0, 0, 0, '売上').ok);
    assert.ok(
      wb.setSheetPageSetup(0, {
        paperSize: 9,
        orientation: 1,
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
      }).ok,
    );
    assert.ok(wb.setSheetPageMargins(0, { left: 0.5, right: 0.5, top: 0.8, bottom: 0.8 }).ok);
    assert.ok(wb.setSheetPrintOptions(0, { horizontalCentered: true }).ok);
    assert.ok(wb.setSheetHeaderFooter(0, { oddHeader: '&C月次報告', oddFooter: '&R&P / &N' }).ok);
    assert.ok(wb.setSheetPrintArea(0, 'A1:F80').ok);
    assert.ok(wb.setSheetPrintTitles(0, '1:2', '').ok);
    assert.ok(wb.addSheetRowBreak(0, 39, true).ok);

    const saved = wb.save();
    assert.ok(saved.status.ok, JSON.stringify(saved.status));
    bytes = saved.bytes;
  } finally {
    wb.delete();
  }

  const reloaded = Module.Workbook.loadBytes(bytes);
  try {
    const setup = reloaded.getSheetPageSetup(0);
    assert.ok(setup.status.ok);
    assert.equal(setup.paperSize, 9);
    assert.equal(setup.orientation, 1);
    assert.equal(setup.fitToPage, true);
    assert.equal(setup.fitToHeight, 0);
    assert.equal(reloaded.getSheetPageMargins(0).left, 0.5);
    assert.equal(reloaded.getSheetPrintArea(0).ranges, 'A1:F80');
    const titles = reloaded.getSheetPrintTitles(0);
    assert.equal(titles.repeatRows, '1:2');
    assert.equal(titles.repeatCols, '');
    // The header code is stored escaped and decodes back to `&C...`.
    assert.match(reloaded.getSheetHeaderFooterXml(0).xml, /<oddHeader>&amp;C月次報告<\/oddHeader>/);
    const breaks = reloaded.getSheetRowBreaks(0);
    assert.deepEqual(
      breaks.breaks.map((b) => b.id),
      [39],
    );
    assert.equal(breaks.breaks[0].manual, true);
  } finally {
    reloaded.delete();
  }
});

test('setRangeXfIndex applies one xf across a rectangle', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    const xf = wb.addXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 0, wrapText: true, hasWrapText: true });
    assert.ok(xf.status.ok, JSON.stringify(xf.status));
    assert.ok(wb.setRangeXfIndex(0, 0, 0, 2, 2, xf.index).ok);
    // Cells that held nothing are materialised so the ruled box renders.
    for (const [row, col] of [
      [0, 0],
      [1, 1],
      [2, 2],
    ]) {
      assert.equal(wb.getCellXfIndex(0, row, col).xfIndex, xf.index);
    }
    assert.equal(wb.getCellXfIndex(0, 3, 3).xfIndex, 0);
  } finally {
    wb.delete();
  }
});

test('sheet layout getters preserve explicit width presence flags', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    assert.ok(wb.setColumnWidth(0, 0, 0, 0).ok);
    assert.ok(wb.setColumnHidden(0, 1, 1, true).ok);
    assert.ok(wb.setRowHeight(0, 3, 30).ok);

    const columns = wb.getSheetColumns(0);
    assert.ok(columns.status.ok, `getSheetColumns: ${JSON.stringify(columns.status)}`);
    const widthZero = columns.columns.find((column) => column.first === 0 && column.last === 0);
    assert.ok(widthZero);
    assert.equal(widthZero.width, 0);
    assert.equal(typeof widthZero.hasWidth, 'number');
    assert.equal(widthZero.hasWidth, 1);
    assert.equal(typeof widthZero.hasStyle, 'number');
    assert.equal(widthZero.hasStyle, 0);
    const hidden = columns.columns.find((column) => column.first === 1 && column.last === 1);
    assert.ok(hidden);
    assert.equal(hidden.hidden, 1);
    assert.equal(typeof hidden.hasWidth, 'number');
    assert.equal(hidden.hasWidth, 0);

    const rows = wb.getSheetRowOverrides(0);
    assert.ok(rows.status.ok, `getSheetRowOverrides: ${JSON.stringify(rows.status)}`);
    const row = rows.rows.find((entry) => entry.row === 3);
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
    assert.ok(wb.addSheet('Ä').ok);
    assert.ok(wb.addSheet('Ö').ok);
    assert.equal(wb.addSheet('ä').ok, false);
    const collision = wb.renameSheet(1, 'ä');
    assert.equal(collision.ok, false);
    assert.notEqual(collision.status, 0);
    assert.equal(wb.sheetName(0).value, 'Ä');
    assert.equal(wb.sheetName(1).value, 'Ö');
    assert.ok(wb.renameSheet(0, 'ä').ok);
    const sn = wb.sheetName(0);
    assert.ok(sn.status.ok);
    assert.equal(sn.value, 'ä');
    assert.equal(wb.sheetName(1).value, 'Ö');
  } finally {
    wb.delete();
  }
});

test('setNumber + getValue round-trips bit-exactly', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    const x = 3.141592653589793;
    assert.ok(wb.setNumber(0, 0, 0, x).ok);
    assert.ok(wb.recalc().ok);
    const r = wb.getValue(0, 0, 0);
    assert.ok(r.status.ok);
    assert.equal(r.value.kind, VAL.NUMBER);
    assert.equal(r.value.number, x);
  } finally {
    wb.delete();
  }
});

test('setFormula + recalc computes 41 + 1 == 42 in B1', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    assert.ok(wb.setNumber(0, 0, 0, 41).ok);
    assert.ok(wb.setFormula(0, 0, 1, '=A1+1').ok);
    assert.ok(wb.recalc().ok);
    const b1 = wb.getValue(0, 0, 1);
    assert.ok(b1.status.ok);
    assert.equal(b1.value.kind, VAL.NUMBER);
    assert.equal(b1.value.number, 42);
  } finally {
    wb.delete();
  }
});

test('Workbook.recalcParallel evaluates a wide DAG and reports bounded telemetry', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  const branchCount = 32;
  try {
    assert.ok(wb.setNumber(0, 0, 0, 1).ok);
    for (let i = 0; i < branchCount; i += 1) {
      const row = i + 1;
      assert.ok(wb.setFormula(0, row, 1, `=A1+${i + 2}`).ok);
      assert.ok(wb.setFormula(0, row, 2, `=B${row + 1}*2`).ok);
    }

    const parallel = wb.recalcParallel(4);
    assert.ok(parallel.status.ok, `status=${JSON.stringify(parallel.status)}`);
    assert.equal(typeof parallel.stats.cellsEvaluated, 'number');
    assert.equal(typeof parallel.stats.sccsProcessed, 'number');
    assert.equal(typeof parallel.stats.parallelSteps, 'number');
    assert.ok(parallel.stats.cellsEvaluated > 0);
    assert.ok(parallel.stats.sccsProcessed > 0);
    if (!threadsEntry) {
      // Without pthreads every launch is refused, so the pass runs serially.
      assert.equal(parallel.stats.workerThreadsStarted, 0);
    }
    if (parallel.stats.workerThreadsStarted <= 1) {
      // OS launch refusal or a partial launch of one worker is a documented
      // successful serial degradation.
      assert.equal(parallel.stats.workerThreadsUsed, 0);
      assert.equal(parallel.stats.parallelSteps, 0);
      assert.ok(parallel.stats.serialFallbackSteps > 0, `stats=${JSON.stringify(parallel.stats)}`);
    } else {
      assert.ok(parallel.stats.parallelSteps > 0, `stats=${JSON.stringify(parallel.stats)}`);
      assert.ok(parallel.stats.workerThreadsStarted >= 2);
      assert.ok(parallel.stats.workerThreadsStarted <= 4);
      assert.ok(parallel.stats.workerThreadsUsed > 0);
      assert.ok(parallel.stats.workerThreadsUsed <= parallel.stats.workerThreadsStarted);
    }

    const first = wb.getValue(0, 1, 1);
    const last = wb.getValue(0, branchCount, 2);
    assert.ok(first.status.ok);
    assert.ok(last.status.ok);
    assert.equal(first.value.kind, VAL.NUMBER);
    assert.equal(last.value.kind, VAL.NUMBER);
    assert.equal(first.value.number, 3);
    assert.equal(last.value.number, (1 + branchCount + 1) * 2);

    assert.ok(wb.setNumber(0, 0, 0, 5).ok);
    const callerOnly = wb.recalcParallel(1);
    assert.ok(callerOnly.status.ok, `status=${JSON.stringify(callerOnly.status)}`);
    assert.equal(callerOnly.stats.parallelSteps, 0);
    assert.ok(callerOnly.stats.serialFallbackSteps > 0);
    assert.equal(callerOnly.stats.workerThreadsStarted, 0);
    assert.equal(callerOnly.stats.workerThreadsUsed, 0);
    assert.equal(wb.getValue(0, branchCount, 2).value.number, (5 + branchCount + 1) * 2);

    const invalid = wb.recalcParallel(9);
    assert.equal(invalid.status.ok, false);
    assert.notEqual(invalid.status.status, 0);
    // Rejected by the binding itself, before any C ABI call: the context
    // is empty, never a value built from the threadCount the caller
    // passed (the Status contract's promise for a binding-raised
    // failure).
    assert.equal(invalid.status.context, '');
    assert.equal(invalid.stats.cellsEvaluated, 0);
    assert.equal(invalid.stats.sccsProcessed, 0);
    assert.equal(invalid.stats.parallelSteps, 0);
    assert.equal(invalid.stats.serialFallbackSteps, 0);
    assert.equal(invalid.stats.cycleRecoveries, 0);
    assert.equal(invalid.stats.workerThreadsStarted, 0);
    assert.equal(invalid.stats.workerThreadsUsed, 0);

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
      const invalidShape = wb.recalcParallel(invalidThreadCount);
      assert.equal(invalidShape.status.ok, false, `threadCount=${String(invalidThreadCount)}`);
      assert.notEqual(invalidShape.status.status, 0);
      assert.equal(invalidShape.stats.cellsEvaluated, 0);
      assert.equal(invalidShape.stats.sccsProcessed, 0);
      assert.equal(invalidShape.stats.parallelSteps, 0);
      assert.equal(invalidShape.stats.serialFallbackSteps, 0);
      assert.equal(invalidShape.stats.cycleRecoveries, 0);
      assert.equal(invalidShape.stats.workerThreadsStarted, 0);
      assert.equal(invalidShape.stats.workerThreadsUsed, 0);
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
  CALLBACK_EXCEPTION: 7003,
});

test('a throwing iterative progress callback aborts the solve and reports a status', async () => {
  const Module = await getModule();
  const wb = makeIterativeWorkbook(Module);
  try {
    let calls = 0;
    assert.ok(
      wb.setIterativeProgress(() => {
        calls += 1;
        throw new Error('callback failure');
      }).ok,
    );

    // The binding never lets an exception cross back into WASM, where no
    // landing pad exists to restore the engine's RAII state: the throw
    // comes back as a status on the same envelope every other call uses.
    const aborted = wb.recalc();
    assert.equal(aborted.ok, false, `expected a failure envelope: ${JSON.stringify(aborted)}`);
    assert.equal(aborted.status, STATUS.CALLBACK_EXCEPTION);
    assert.equal(calls, 1, 'the solve must stop at the first throw');

    // The workbook is still readable, and the recalc re-entry guard was
    // released on the way out.
    const mid = wb.getValue(0, 0, 0);
    assert.ok(mid.status.ok, JSON.stringify(mid.status));
    assert.equal(mid.value.kind, VAL.NUMBER);

    const second = wb.recalc();
    assert.notEqual(second.status, STATUS.RECALC_REENTRANT, 'the re-entry flag stayed set');

    assert.ok(wb.setIterativeProgress(null).ok);
    assert.ok(wb.setFormula(0, 0, 0, '=(A1+10)/2').ok);
    const recovered = wb.recalc();
    assert.ok(recovered.ok, `recalc after the throwing callback: ${JSON.stringify(recovered)}`);
    const converged = wb.getValue(0, 0, 0);
    assert.ok(converged.status.ok, JSON.stringify(converged.status));
    assert.equal(converged.value.kind, VAL.NUMBER);
    assert.ok(Math.abs(converged.value.number - 10) < 0.001, `A1=${converged.value.number}`);
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
  assert.ok(
    wb.setIterativeProgress(() => {
      calls += 1;
      try {
        wb.delete();
      } catch (e) {
        caught = e;
      }
      return false;
    }).ok,
  );

  const result = wb.recalc();
  assert.equal(calls, 1, 'the callback must run exactly once');
  assert.ok(caught instanceof Error, 'delete() inside the callback must throw, not free the handle');
  assert.match(String(caught.message), /iterative progress callback/);
  // The handle survived: delete() never reached fm_workbook_destroy, so
  // the workbook is still usable after the (deliberately cancelled) solve.
  assert.ok(result.ok, `recalc after a cancelled solve: ${JSON.stringify(result)}`);
  assert.ok(wb.isValid());

  wb.setIterativeProgress(null);
  wb.delete();
});

test('an uncaught delete() inside the progress callback aborts the solve like any other throw', async () => {
  const Module = await getModule();
  const wb = makeIterativeWorkbook(Module);
  let calls = 0;
  assert.ok(
    wb.setIterativeProgress(() => {
      calls += 1;
      wb.delete(); // not caught here -- propagates the same as a plain `throw`
    }).ok,
  );

  const result = wb.recalc();
  assert.equal(calls, 1);
  assert.equal(result.ok, false, `expected a failure envelope: ${JSON.stringify(result)}`);
  assert.equal(result.status, STATUS.CALLBACK_EXCEPTION);
  assert.ok(wb.isValid(), 'the handle must still be alive: delete() never completed');

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
      assert.ok(
        wb.setIterativeProgress(() => {
          calls += 1;
          return returned;
        }).ok,
      );
      assert.ok(wb.recalc().ok);
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
