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
  test('layout nested integer fields reject before native mutation', () => {
    withWorkbook(Module, (wb) => {
      const before = protection(wb);
      const invalid = wb.setSheetProtection(0, { ...before, enabled: 1.5 });
      assert.equal(invalid.status, 2, JSON.stringify(invalid));
      assert.equal(invalid.ok, false);
      assert.deepEqual(protection(wb), before);

      const spin = wb.setSheetProtection(0, { ...before, spinCount: 2 ** 32 });
      assert.equal(spin.status, 2, JSON.stringify(spin));
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
      const protectionResult = wb.setSheetProtection(0, protectionInput);
      assert.equal(protectionResult.status, 2, JSON.stringify(protectionResult));
      assert.deepEqual(protection(wb), beforeProtection);

      const beforeFormat = wb.getSheetFormatDefaults(0);
      const formatInput = { defaultColWidth: 12.5 };
      Object.defineProperty(formatInput, 'defaultRowHeight', {
        enumerable: true,
        get() {
          throw new Error('wasm format getter');
        },
      });
      const formatResult = wb.setSheetFormatDefaults(0, formatInput);
      assert.equal(formatResult.status, 2, JSON.stringify(formatResult));
      assert.deepEqual(wb.getSheetFormatDefaults(0), beforeFormat);
    });
  });

  test('geometry and partial recalc reject invalid nested coordinates before C calls', () => {
    withWorkbook(Module, (wb) => {
      const fractional = wb.getCellRectPt(0, { ...RANGE, firstRow: 1.5 }, 0);
      assert.equal(fractional.status.status, 2, JSON.stringify(fractional));
      assert.equal(fractional.status.ok, false);

      const throwingRange = {};
      Object.defineProperty(throwingRange, 'firstRow', {
        enumerable: true,
        get() {
          throw new Error('wasm geometry getter');
        },
      });
      const geometry = wb.getCellRectPt(0, throwingRange, 0);
      assert.equal(geometry.status.status, 2, JSON.stringify(geometry));
      assert.equal(geometry.status.ok, false);

      assert.ok(wb.setNumber(0, 0, 0, 1).ok);
      assert.ok(wb.setFormula(0, 0, 1, '=A1+1').ok);
      const recalc = wb.partialRecalc({ ...RANGE, sheet: 0, firstRow: 1.5 });
      assert.equal(recalc.status.status, 2, JSON.stringify(recalc));
      assert.equal(recalc.status.ok, false);
      assert.equal(recalc.recomputed, 0);

      const throwingViewport = { sheet: 0, firstRow: 0, lastRow: 0, firstCol: 0 };
      Object.defineProperty(throwingViewport, 'lastCol', {
        enumerable: true,
        get() {
          throw new Error('wasm viewport getter');
        },
      });
      const viewport = wb.partialRecalc(throwingViewport);
      assert.equal(viewport.status.status, 2, JSON.stringify(viewport));
      assert.equal(viewport.recomputed, 0);
    });
  });
}
