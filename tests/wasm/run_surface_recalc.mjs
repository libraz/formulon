import assert from 'node:assert/strict';
import { StatusCode } from '../../packages/npm/common.mjs';
import { expectFailure, VAL } from './smoke_support.mjs';

export function registerSurfaceRecalc(Module, test, threadsBuild) {
  test('getVersion is non-empty', () => {
    const v = Module.getVersion();
    assert.equal(typeof v, 'string');
    assert.ok(v.length > 0, `expected non-empty version, got ${JSON.stringify(v)}`);
  });

  test('getStatusName covers known codes', () => {
    assert.equal(Module.getStatusName(0), 'OK');
    // Unknown / out-of-band codes still return a non-null string.
    const unknown = Module.getStatusName(123456);
    assert.equal(typeof unknown, 'string');
    assert.ok(unknown.length > 0);
  });

  test('getErrorDisplayName returns Excel literals', () => {
    assert.equal(Module.getErrorDisplayName(1), '#DIV/0!');
    assert.equal(Module.getErrorDisplayName(999), '#UNKNOWN!');
  });

  test('evalFormula(=SUM(1,2,3)) returns NUMBER 6', () => {
    const r = Module.evalFormula('=SUM(1,2,3)');
    assert.equal(r.kind, VAL.NUMBER);
    assert.equal(r.number, 6);
  });

  test('evalFormula(=IF(TRUE,1,2)) returns NUMBER 1', () => {
    const r = Module.evalFormula('=IF(TRUE,1,2)');
    assert.equal(r.kind, VAL.NUMBER);
    assert.equal(r.number, 1);
  });

  test('evalFormula evaluates intersection read-only and returns #NULL!', () => {
    const r = Module.evalFormula('=A1 B1');
    assert.equal(r.kind, VAL.ERROR);
    assert.equal(r.errorCode, 0); // ErrorCode::Null / #NULL!
  });

  test('evalFormula(=CONCAT) returns TEXT "hello world"', () => {
    const r = Module.evalFormula('=CONCAT("hello"," ","world")');
    assert.equal(r.kind, VAL.TEXT);
    assert.equal(r.text, 'hello world');
  });

  test('numeric comparisons share the 15-digit bucket without fuzzy exact lookups', () => {
    const operators = [
      ['=', true],
      ['<>', false],
      ['<', false],
      ['<=', true],
      ['>', false],
      ['>=', true],
    ];
    for (const [operator, expected] of operators) {
      const r = Module.evalFormula(`=7.1${operator}7.1000000000000005`);
      assert.equal(r.kind, VAL.BOOL, `operator=${operator}`);
      assert.equal(r.boolean, expected, `operator=${operator}`);
    }

    // MATCH and XLOOKUP exact modes intentionally retain literal IEEE equality;
    // the formula comparison bucket must not make their exact guards fuzzy.
    // Populate the two worksheet values through setNumber so the parser's
    // decimal-literal canonicalisation cannot collapse the adjacent doubles.
    const wb = Module.Workbook.createDefault();
    try {
      wb.setNumber(0, 0, 0, 7.1);
      wb.setNumber(0, 1, 0, 7.1000000000000005);
      wb.setFormula(0, 2, 0, '=MATCH(A2,A1:A1,0)');
      wb.setFormula(0, 3, 0, '=XLOOKUP(A2,A1:A1,A1:A1,"MISS",0)');
      wb.recalc();

      const match = wb.getValue(0, 2, 0);
      assert.equal(match.kind, VAL.ERROR);
      assert.equal(match.errorCode, 6); // ErrorCode::NA / #N/A.

      const xlookup = wb.getValue(0, 3, 0);
      assert.equal(xlookup.kind, VAL.TEXT);
      assert.equal(xlookup.text, 'MISS');
    } finally {
      wb.delete();
    }
  });

  test('Workbook construction + setNumber + recalc + getValue', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.equal(wb.disposed, false);
      assert.equal(wb.getSheetCount(), 1);
      const nameRes = wb.getSheetName(0);
      assert.equal(nameRes, 'Sheet1');

      wb.setNumber(0, 0, 0, 42);
      wb.setFormula(0, 1, 0, '=A1*2');
      wb.recalc();

      const a1 = wb.getValue(0, 0, 0);
      assert.equal(a1.kind, VAL.NUMBER);
      assert.equal(a1.number, 42);

      const b1 = wb.getValue(0, 1, 0);
      assert.equal(b1.kind, VAL.NUMBER);
      assert.equal(b1.number, 84);

      // addSheet should grow sheetCount.
      wb.addSheet('Second');
      assert.equal(wb.getSheetCount(), 2);
    } finally {
      wb.delete();
    }
  });

  test('Workbook.recalcParallel evaluates a wide DAG and reports bounded telemetry', () => {
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
      if (!threadsBuild) {
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

      expectFailure(() => wb.recalcParallel(9), StatusCode.InvalidArgument);

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

  test('paginate exposes page geometry', () => {
    const wb = Module.Workbook.createDefault();
    try {
      // Asserts the result's shape, not the page geometry: how many rows
      // fit a page is the print engine's business and is pinned by the
      // native print suite against Excel. Row 400 is far enough down to
      // force a break under any plausible body height.
      wb.setNumber(0, 0, 0, 1);
      wb.setNumber(0, 400, 0, 2);
      const result = wb.getPagination(0);
      assert.ok(result.pageCount >= 2, `pageCount=${result.pageCount}`);
      assert.deepEqual(result.printArea, []);
      assert.ok(result.rowBreaks.length > 0);
      assert.deepEqual(
        result.rowBreaks,
        [...new Set(result.rowBreaks)].sort((a, b) => a - b),
      );
      assert.deepEqual(result.columnBreaks, []);
    } finally {
      wb.delete();
    }
  });

  test('setPinnedNow / pinnedNow / clearPinnedNow drive the clock seam', () => {
    const wb = Module.Workbook.createDefault();
    try {
      // Unpinned by default: the workbook follows the host clock.
      assert.equal(wb.getPinnedNow(), null);

      wb.setPinnedNow(2026, 4, 23, 15, 30, 45);
      assert.deepEqual(wb.getPinnedNow(), {
        year: 2026,
        month: 4,
        day: 23,
        hour: 15,
        minute: 30,
        second: 45,
      });

      // 2026-04-23 is serial 46135 under the 1900 date system.
      wb.setFormula(0, 0, 0, '=TODAY()');
      wb.recalc();
      const a1 = wb.getValue(0, 0, 0);
      assert.equal(a1.kind, VAL.NUMBER);
      assert.equal(a1.number, 46135);

      // The pin is a calendar instant, not a normalising constructor.
      expectFailure(() => wb.setPinnedNow(2026, 13, 1, 0, 0, 0));
      expectFailure(() => wb.setPinnedNow(2025, 2, 29, 0, 0, 0));

      wb.clearPinnedNow();
      assert.equal(wb.getPinnedNow(), null);
    } finally {
      wb.delete();
    }
  });
}
