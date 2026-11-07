import assert from 'node:assert/strict';
import { StatusCode } from '../../packages/npm/common.mjs';
import { expectFailure } from './smoke_support.mjs';

const RANGE = { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 };

function expectRangeError(fn) {
  assert.throws(fn, RangeError);
}

function expectTypeError(fn) {
  assert.throws(fn, TypeError);
}

function deleteAfterLifecycleProbe(wb, lifecycle) {
  if (wb.isDeleted()) return;
  try {
    wb.delete();
  } catch (error) {
    // Older artifacts can leave deleteLater scheduled after the getter
    // returns; there is no synchronous flush API. The fixed guard leaves the
    // workbook normally deletable here.
    if (lifecycle !== 'deleteLater') throw error;
  }
}

export function registerPositional(Module, test, hookProbe) {
  if (hookProbe !== undefined) {
    test('WASM positional guards install before factory hooks', () => {
      // The runner supplies these observations from onRuntimeInitialized and
      // postRun. Keeping this assertion in the shared smoke suite protects
      // direct factories and npm shims from reordering the installation.
      assert.deepEqual(hookProbe.runtimeInitial, { threw: true, a1: 42 });
      assert.deepEqual(hookProbe.runtimeReplacement, { threw: true, a1: 42 });
      assert.deepEqual(hookProbe.postInitial, { threw: true, a1: 42 });
      assert.deepEqual(hookProbe.postReplacement, { threw: true, a1: 42 });
    });
  }

  test('WASM uint32 positions reject coercion and preserve the cell', () => {
    const wb = Module.Workbook.createDefault();
    try {
      wb.setNumber(0, 0, 0, 42);
      for (const value of [2 ** 32, -1, 1.5, Number.NaN, Number.POSITIVE_INFINITY]) {
        expectRangeError(() => wb.setNumber(value, 0, 0, 7));
        assert.equal(wb.getValue(0, 0, 0).number, 42);
      }
      expectTypeError(() => wb.setNumber(Symbol('sheet'), 0, 0, 7));
      assert.equal(wb.getValue(0, 0, 0).number, 42);

      // A valid but out-of-range sheet remains a native status failure. This
      // distinguishes the guard from the engine's own sheet lookup.
      expectFailure(() => wb.setNumber(1, 0, 0, 7));
      assert.equal(wb.getValue(0, 0, 0).number, 42);
    } finally {
      wb.delete();
    }
  });

  test('WASM signed and string positions reject symbols and out-of-range values', () => {
    const wb = Module.Workbook.createDefault();
    try {
      for (const value of [2 ** 31, -(2 ** 31) - 1, 1.5, Number.NaN, Number.POSITIVE_INFINITY]) {
        expectRangeError(() => wb.setError(0, 0, 0, value));
      }
      expectTypeError(() => wb.setError(0, 0, 0, Symbol('error')));
      expectTypeError(() => wb.setText(0, 0, 0, Symbol('text')));
      wb.setText(0, 0, 0, 'valid');

      expectFailure(() => wb.setError(1, 0, 0, -1), StatusCode.SheetIndexOutOfRange);
      assert.equal(typeof Module.getStatusName(-1), 'string');
      assert.equal(Module.getErrorDisplayName, undefined);
      expectFailure(() => Module.setLogMinLevel(-1), StatusCode.InvalidArgument);
      expectTypeError(() => Module.getStatusName(Symbol('status')));
      expectTypeError(() => Module.setLogMinLevel(Symbol('level')));
    } finally {
      wb.delete();
    }
  });

  test('WASM guard messages name the facade method and count its arguments', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.throws(() => wb.setNumber('x', 0, 0, 1), {
        name: 'TypeError',
        message: 'Sheet.setNumber sheet index must be a primitive number',
      });
      assert.throws(() => wb.setNumber(0, -1, 0, 1), {
        name: 'RangeError',
        message: 'Sheet.setNumber argument 1 is outside uint32 range',
      });
      assert.throws(() => wb.setNumber(0, 0, 0, 'x'), {
        name: 'TypeError',
        message: 'Sheet.setNumber argument 3 must be a primitive number',
      });
      assert.throws(() => wb.addSheet(1), {
        name: 'TypeError',
        message: 'Workbook.addSheet argument 1 must be a primitive string',
      });
      assert.throws(() => wb.setPrintTitles(0), {
        name: 'TypeError',
        message: 'Sheet.setPrintTitles argument 1 is required',
      });
    } finally {
      wb.delete();
    }
    assert.throws(() => Module.Workbook.createForProfile(1), {
      name: 'TypeError',
      message: 'Workbook.createForProfile argument 1 must be a primitive string',
    });
    assert.throws(() => Module.evaluateFormula(1), {
      name: 'TypeError',
      message: 'evaluateFormula argument 1 must be a primitive string',
    });
  });

  test('WASM string free function is primitive-only', () => {
    expectTypeError(() => Module.evaluateFormula(Symbol('formula')));
    expectTypeError(() =>
      Module.evaluateFormula({
        toString() {
          throw new Error('toString');
        },
      }),
    );
    const result = Module.evaluateFormula('=1+1');
    assert.equal(result.number, 2);
  });

  test('WASM paging positions reject invalid cursors and retain nullish defaults', () => {
    const wb = Module.Workbook.createDefault();
    try {
      wb.setNumber(0, 0, 0, 42);
      for (const value of [-1, 1.5, Number.NaN, Number.POSITIVE_INFINITY, 2 ** 53, 2 ** 64]) {
        expectRangeError(() => wb.getCellsInRange(0, RANGE, value, 1));
        expectRangeError(() => wb.getCellsInRange(0, RANGE, 0, value));
        expectRangeError(() => wb.getInvalidCells(0, value, 1));
        expectRangeError(() => wb.getInvalidCells(0, 0, value));
      }
      expectTypeError(() => wb.getCellsInRange(0, RANGE, Symbol('cursor'), 1));
      expectTypeError(() => wb.getInvalidCells(0, Symbol('cursor'), 1));

      wb.getCellsInRange(0, RANGE, null, null);
      wb.getCellsInRange(0, RANGE, undefined, undefined);
      wb.getInvalidCells(0, null, null);
      wb.getInvalidCells(0, undefined, undefined);
      assert.equal(wb.getValue(0, 0, 0).number, 42);
    } finally {
      wb.delete();
    }
  });

  test('WASM-only table and auto-filter positions are guarded', () => {
    const wb = Module.Workbook.createDefault();
    try {
      for (const method of ['removeTable', 'updateTable', 'getAutoFilterXml']) {
        expectRangeError(() => wb[method](2 ** 32, {}));
      }
      expectRangeError(() => wb.setAutoFilterXml(2 ** 32, ''));
      expectTypeError(() => wb.setAutoFilterXml(0, Symbol('xml')));

      assert.equal(wb.getAutoFilterXml(0), null);
      wb.setAutoFilterXml(0, '');
    } finally {
      wb.delete();
    }
  });

  test('WASM guard preserves embind descriptors and receiver errors', () => {
    const descriptor = Object.getOwnPropertyDescriptor(Module.Workbook.prototype, 'setNumber');
    assert.equal(descriptor.enumerable, true);
    assert.equal(descriptor.writable, true);
    assert.equal(descriptor.configurable, true);
    const wb = Module.Workbook.createDefault();
    try {
      const detached = wb.setNumber;
      assert.throws(() => detached(0, 0, 0, 7));
      const forged = Object.create(Module.Workbook.prototype);
      assert.throws(() => forged.setNumber(0, 0, 0, 7));
      wb.setNumber.call(wb, 0, 0, 0, 42);
      assert.equal(wb.getValue(0, 0, 0).number, 42);
    } finally {
      wb.delete();
    }
  });

  test('WASM overload leaves validate positions and track active Workbook calls', () => {
    const wb = Module.Workbook.createDefault();
    try {
      wb.setNumber(0, 0, 0, 42);
      const leaf = (name, argCount) => {
        const table = Module.Workbook.prototype[name].overloadTable;
        assert.ok(Array.isArray(table), `${name} overload table missing`);
        const entry = table.find((candidate) => candidate?.argCount === argCount);
        assert.equal(typeof entry, 'function', `${name} overload leaf missing`);
        return entry;
      };
      const getCellsLeaf = leaf('getCellsInRange', 4);
      const listInvalidLeaf = leaf('getInvalidCells', 3);
      const insertImageLeaf = leaf('addImage', 3);

      expectRangeError(() => getCellsLeaf.call(wb, 0, RANGE, 2 ** 64, 1));
      expectRangeError(() => listInvalidLeaf.call(wb, 0, -1, 1));
      expectRangeError(() => insertImageLeaf.call(wb, 2 ** 32, new Uint8Array(), {}));
      assert.equal(wb.getValue(0, 0, 0).number, 42);

      wb.getCellsInRange(0, RANGE, 0, 1);
      wb.getInvalidCells(0, 0, 1);
      expectFailure(() => wb.addImage(0, new Uint8Array(), {}), StatusCode.BindingNullPointer);

      const nestedRange = {};
      Object.defineProperties(nestedRange, {
        firstRow: {
          get() {
            wb.delete();
            return 0;
          },
        },
        firstCol: { value: 0 },
        lastRow: { value: 0 },
        lastCol: { value: 0 },
      });
      assert.throws(() => getCellsLeaf.call(wb, 0, nestedRange, 0, 1), {
        name: 'FormulonError',
        code: 4005,
        message: /cannot dispose Workbook during an active method call/,
      });
      assert.equal(wb.isDeleted(), false);
      assert.equal(wb.getValue(0, 0, 0).number, 42);
    } finally {
      if (!wb.isDeleted()) wb.delete();
    }
  });

  test('WASM guards reject reentrant Workbook deletion during nested reads', () => {
    for (const lifecycle of ['delete', 'deleteLater', Symbol.dispose]) {
      const wb = Module.Workbook.createDefault();
      try {
        const before = wb.getCellXfCount();
        const spec = {};
        Object.defineProperty(spec, 'horizontalAlign', {
          get() {
            wb[lifecycle]();
            return 1;
          },
        });
        assert.throws(
          () => wb.addCellXf(spec),
          /cannot dispose Workbook during an active method call/,
          `${String(lifecycle)} unexpectedly succeeded`,
        );
        assert.equal(wb.getCellXfCount(), before);
        assert.equal(wb.isDeleted(), false);
      } finally {
        deleteAfterLifecycleProbe(wb, lifecycle);
      }
    }
  });

  test('WASM lifecycle guards cover inherited descriptor calls during nested reads', () => {
    for (const lifecycle of ['delete', 'deleteLater', Symbol.dispose]) {
      const wb = Module.Workbook.createDefault();
      const inherited = Object.getPrototypeOf(Object.getPrototypeOf(wb));
      const descriptor = Object.getOwnPropertyDescriptor(inherited, lifecycle);
      assert.equal(typeof descriptor?.value, 'function', `${String(lifecycle)} descriptor missing`);
      try {
        const before = wb.getCellXfCount();
        const spec = {};
        Object.defineProperty(spec, 'horizontalAlign', {
          get() {
            descriptor.value.call(wb);
            return 1;
          },
        });
        assert.throws(
          () => wb.addCellXf(spec),
          /cannot dispose Workbook during an active method call/,
          `${String(lifecycle)} inherited call unexpectedly succeeded`,
        );
        assert.equal(wb.getCellXfCount(), before);
        assert.equal(wb.isDeleted(), false);
      } finally {
        deleteAfterLifecycleProbe(wb, lifecycle);
      }
    }
  });

  test('WASM Workbook [Symbol.dispose] is dispose(): the handle goes, the wrapper stays', () => {
    const wb = Module.Workbook.createDefault();
    assert.equal(wb[Symbol.dispose], wb.dispose);
    wb[Symbol.dispose]();
    wb[Symbol.dispose]();
    assert.equal(wb.disposed, true);
    assert.equal(wb.isDeleted(), false);
    expectFailure(() => wb.getSheetCount(), StatusCode.BindingInvalidHandle);
    wb.delete();
  });

  test('WASM lifecycle depth follows empty Proxy Workbook receivers', () => {
    const wb = Module.Workbook.createDefault();
    const proxy = new Proxy(wb, {});
    try {
      proxy.setNumber(0, 0, 0, 42);

      let originalDeleteError;
      const originalDeleteSpec = {};
      Object.defineProperty(originalDeleteSpec, 'horizontalAlign', {
        get() {
          try {
            wb.delete();
          } catch (error) {
            originalDeleteError = error;
          }
          return 1;
        },
      });
      proxy.addCellXf(originalDeleteSpec);
      assert.equal(originalDeleteError?.code, 4005);
      assert.equal(proxy.isDeleted(), false);

      let proxyDeleteError;
      const proxyDeleteSpec = {};
      Object.defineProperty(proxyDeleteSpec, 'horizontalAlign', {
        get() {
          try {
            proxy.delete();
          } catch (error) {
            proxyDeleteError = error;
          }
          return 1;
        },
      });
      proxy.addCellXf(proxyDeleteSpec);
      assert.equal(proxyDeleteError?.code, 4005);
      assert.equal(wb.isDeleted(), false);
    } finally {
      if (!wb.isDeleted()) wb.delete();
    }
  });

  test('WASM lifecycle guard permits a getter to catch its deletion error', () => {
    const wb = Module.Workbook.createDefault();
    try {
      let caught;
      const spec = {};
      Object.defineProperty(spec, 'horizontalAlign', {
        get() {
          try {
            wb.delete();
          } catch (error) {
            caught = error;
          }
          return 1;
        },
      });
      wb.addCellXf(spec);
      assert.equal(caught?.code, 4005);
      assert.equal(wb.isDeleted(), false);
    } finally {
      if (!wb.isDeleted()) wb.delete();
    }
  });
}
