import assert from 'node:assert/strict';

const RANGE = { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 };

function withWorkbook(Module, fn) {
  const wb = Module.Workbook.createDefault();
  try {
    fn(wb);
  } finally {
    wb.delete();
  }
}

function protection(wb) {
  const result = wb.getSheetProtection(0);
  assert.ok(result.status.ok, JSON.stringify(result.status));
  return result.protection;
}

export function registerLayoutValidation(Module, test) {
  test('layout nested integer fields throw before native mutation', () => {
    withWorkbook(Module, (wb) => {
      const before = protection(wb);
      assert.throws(() => wb.setSheetProtection(0, { ...before, enabled: 1.5 }), RangeError);
      assert.deepEqual(protection(wb), before);

      assert.throws(() => wb.setSheetProtection(0, { ...before, spinCount: 2 ** 32 }), RangeError);
      assert.deepEqual(protection(wb), before);
    });
  });

  test('layout nested getters are read once and leave state unchanged on failure', () => {
    withWorkbook(Module, (wb) => {
      const beforeProtection = protection(wb);
      const protectionInput = { ...beforeProtection, enabled: 1 };
      Object.defineProperty(protectionInput, 'formatRows', {
        enumerable: true,
        get() {
          throw new Error('wasm protection getter');
        },
      });
      assert.throws(() => wb.setSheetProtection(0, protectionInput), /wasm protection getter/);
      assert.deepEqual(protection(wb), beforeProtection);

      const beforeFormat = wb.getSheetFormatDefaults(0);
      const formatInput = { defaultColWidth: 12.5 };
      Object.defineProperty(formatInput, 'defaultRowHeight', {
        enumerable: true,
        get() {
          throw new Error('wasm format getter');
        },
      });
      assert.throws(() => wb.setSheetFormatDefaults(0, formatInput), /wasm format getter/);
      assert.deepEqual(wb.getSheetFormatDefaults(0), beforeFormat);
    });
  });

  test('geometry and partial recalc throw on invalid nested coordinates before C calls', () => {
    withWorkbook(Module, (wb) => {
      assert.throws(() => wb.getCellRectPt(0, { ...RANGE, firstRow: 1.5 }, 0), RangeError);

      const throwingRange = {};
      Object.defineProperty(throwingRange, 'firstRow', {
        enumerable: true,
        get() {
          throw new Error('wasm geometry getter');
        },
      });
      assert.throws(() => wb.getCellRectPt(0, throwingRange, 0), /wasm geometry getter/);

      assert.ok(wb.setNumber(0, 0, 0, 1).ok);
      assert.ok(wb.setFormula(0, 0, 1, '=A1+1').ok);
      assert.throws(() => wb.partialRecalc({ ...RANGE, sheet: 0, firstRow: 1.5 }), RangeError);

      const throwingViewport = { sheet: 0, firstRow: 0, lastRow: 0, firstCol: 0 };
      Object.defineProperty(throwingViewport, 'lastCol', {
        enumerable: true,
        get() {
          throw new Error('wasm viewport getter');
        },
      });
      assert.throws(() => wb.partialRecalc(throwingViewport), /wasm viewport getter/);
    });
  });
}
