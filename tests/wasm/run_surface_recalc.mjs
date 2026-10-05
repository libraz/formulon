import assert from 'node:assert/strict';
import { VAL } from './smoke_support.mjs';

export function registerSurfaceRecalc(Module, test, threadsBuild) {
  test('versionString is non-empty', () => {
    const v = Module.versionString();
    assert.equal(typeof v, 'string');
    assert.ok(v.length > 0, `expected non-empty version, got ${JSON.stringify(v)}`);
  });

  test('statusString covers known codes', () => {
    assert.equal(Module.statusString(0), 'kOk');
    // Unknown / out-of-band codes still return a non-null string.
    const unknown = Module.statusString(123456);
    assert.equal(typeof unknown, 'string');
    assert.ok(unknown.length > 0);
  });

  test('errorDisplayName returns Excel literals', () => {
    assert.equal(Module.errorDisplayName(1), '#DIV/0!');
    assert.equal(Module.errorDisplayName(999), '#UNKNOWN!');
  });

  test('evalFormula(=SUM(1,2,3)) returns NUMBER 6', () => {
    const r = Module.evalFormula('=SUM(1,2,3)');
    assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
    assert.equal(r.value.kind, VAL.NUMBER);
    assert.equal(r.value.number, 6);
  });

  test('evalFormula(=IF(TRUE,1,2)) returns NUMBER 1', () => {
    const r = Module.evalFormula('=IF(TRUE,1,2)');
    assert.ok(r.status.ok);
    assert.equal(r.value.kind, VAL.NUMBER);
    assert.equal(r.value.number, 1);
  });

  test('evalFormula evaluates intersection read-only and returns #NULL!', () => {
    const r = Module.evalFormula('=A1 B1');
    assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
    assert.equal(r.value.kind, VAL.ERROR);
    assert.equal(r.value.errorCode, 0); // ErrorCode::Null / #NULL!
  });

  test('evalFormula(=CONCAT) returns TEXT "hello world"', () => {
    const r = Module.evalFormula('=CONCAT("hello"," ","world")');
    assert.ok(r.status.ok);
    assert.equal(r.value.kind, VAL.TEXT);
    assert.equal(r.value.text, 'hello world');
  });

  test('numeric comparisons share the 15-digit bucket without fuzzy exact lookups', () => {
    const operators = [
      ['=', 1],
      ['<>', 0],
      ['<', 0],
      ['<=', 1],
      ['>', 0],
      ['>=', 1],
    ];
    for (const [operator, expected] of operators) {
      const r = Module.evalFormula(`=7.1${operator}7.1000000000000005`);
      assert.ok(r.status.ok, `operator=${operator} status=${JSON.stringify(r.status)}`);
      assert.equal(r.value.kind, VAL.BOOL, `operator=${operator}`);
      assert.equal(r.value.boolean, expected, `operator=${operator}`);
    }

    // MATCH and XLOOKUP exact modes intentionally retain literal IEEE equality;
    // the formula comparison bucket must not make their exact guards fuzzy.
    // Populate the two worksheet values through setNumber so the parser's
    // decimal-literal canonicalisation cannot collapse the adjacent doubles.
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setNumber(0, 0, 0, 7.1).ok);
      assert.ok(wb.setNumber(0, 1, 0, 7.1000000000000005).ok);
      assert.ok(wb.setFormula(0, 2, 0, '=MATCH(A2,A1:A1,0)').ok);
      assert.ok(wb.setFormula(0, 3, 0, '=XLOOKUP(A2,A1:A1,A1:A1,"MISS",0)').ok);
      assert.ok(wb.recalc().ok);

      const match = wb.getValue(0, 2, 0);
      assert.ok(match.status.ok, `MATCH status=${JSON.stringify(match.status)}`);
      assert.equal(match.value.kind, VAL.ERROR);
      assert.equal(match.value.errorCode, 6); // ErrorCode::NA / #N/A.

      const xlookup = wb.getValue(0, 3, 0);
      assert.ok(xlookup.status.ok, `XLOOKUP status=${JSON.stringify(xlookup.status)}`);
      assert.equal(xlookup.value.kind, VAL.TEXT);
      assert.equal(xlookup.value.text, 'MISS');
    } finally {
      wb.delete();
    }
  });

  test('Workbook construction + setNumber + recalc + getValue', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.isValid());
      assert.equal(wb.sheetCount().value, 1);
      const nameRes = wb.sheetName(0);
      assert.ok(nameRes.status.ok);
      assert.equal(nameRes.value, 'Sheet1');

      assert.ok(wb.setNumber(0, 0, 0, 42).ok);
      assert.ok(wb.setFormula(0, 1, 0, '=A1*2').ok);
      assert.ok(wb.recalc().ok);

      const a1 = wb.getValue(0, 0, 0);
      assert.ok(a1.status.ok);
      assert.equal(a1.value.kind, VAL.NUMBER);
      assert.equal(a1.value.number, 42);

      const b1 = wb.getValue(0, 1, 0);
      assert.ok(b1.status.ok);
      assert.equal(b1.value.kind, VAL.NUMBER);
      assert.equal(b1.value.number, 84);

      // addSheet should grow sheetCount.
      assert.ok(wb.addSheet('Second').ok);
      assert.equal(wb.sheetCount().value, 2);
    } finally {
      wb.delete();
    }
  });

  test('Workbook.recalcParallel evaluates a wide DAG and reports bounded telemetry', () => {
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
      if (!threadsBuild) {
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

  test('paginate exposes a status envelope and page geometry', () => {
    const wb = Module.Workbook.createDefault();
    try {
      // Asserts the envelope's shape, not the page geometry: how many rows
      // fit a page is the print engine's business and is pinned by the
      // native print suite against Excel. Row 400 is far enough down to
      // force a break under any plausible body height.
      assert.ok(wb.setNumber(0, 0, 0, 1).ok);
      assert.ok(wb.setNumber(0, 400, 0, 2).ok);
      const result = wb.paginate(0);
      assert.ok(result.status.ok, `status=${JSON.stringify(result.status)}`);
      assert.ok(result.pageCount >= 2, `pageCount=${result.pageCount}`);
      assert.deepEqual(result.printArea, []);
      assert.ok(result.horizontalBreaks.length > 0);
      assert.deepEqual(
        result.horizontalBreaks,
        [...new Set(result.horizontalBreaks)].sort((a, b) => a - b),
      );
      assert.deepEqual(result.verticalBreaks, []);
    } finally {
      wb.delete();
    }
  });

  test('setPinnedNow / pinnedNow / clearPinnedNow drive the clock seam', () => {
    const wb = Module.Workbook.createDefault();
    try {
      // Unpinned by default: the workbook follows the host clock.
      assert.equal(wb.pinnedNow().now, null);

      assert.ok(wb.setPinnedNow(2026, 4, 23, 15, 30, 45).ok);
      assert.deepEqual(wb.pinnedNow().now, {
        year: 2026,
        month: 4,
        day: 23,
        hour: 15,
        minute: 30,
        second: 45,
      });

      // 2026-04-23 is serial 46135 under the 1900 date system.
      assert.ok(wb.setFormula(0, 0, 0, '=TODAY()').ok);
      assert.ok(wb.recalc().ok);
      const a1 = wb.getValue(0, 0, 0);
      assert.equal(a1.value.kind, VAL.NUMBER);
      assert.equal(a1.value.number, 46135);

      // The pin is a calendar instant, not a normalising constructor.
      assert.equal(wb.setPinnedNow(2026, 13, 1, 0, 0, 0).ok, false);
      assert.equal(wb.setPinnedNow(2025, 2, 29, 0, 0, 0).ok, false);

      assert.ok(wb.clearPinnedNow().ok);
      assert.equal(wb.pinnedNow().now, null);
    } finally {
      wb.delete();
    }
  });
}
