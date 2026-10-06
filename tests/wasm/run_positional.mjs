import assert from 'node:assert/strict';

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
      assert.ok(wb.setNumber(0, 0, 0, 42).ok);
      for (const value of [2 ** 32, -1, 1.5, Number.NaN, Number.POSITIVE_INFINITY]) {
        expectRangeError(() => wb.setNumber(value, 0, 0, 7));
        assert.equal(wb.getValue(0, 0, 0).value.number, 42);
      }
      expectTypeError(() => wb.setNumber(Symbol('sheet'), 0, 0, 7));
      assert.equal(wb.getValue(0, 0, 0).value.number, 42);

      // A valid but out-of-range sheet remains a native status failure. This
      // distinguishes the guard from the engine's own sheet lookup.
      const native = wb.setNumber(1, 0, 0, 7);
      assert.equal(native.ok, false);
      assert.equal(wb.getValue(0, 0, 0).value.number, 42);
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
      assert.ok(wb.setText(0, 0, 0, 'valid').ok);

      const errorControl = wb.setError(1, 0, 0, -1);
      assert.equal(typeof errorControl.ok, 'boolean');
      assert.equal(typeof Module.statusString(-1), 'string');
      assert.equal(typeof Module.errorDisplayName(-1), 'string');
      assert.equal(typeof Module.setLogMinLevel(-1).ok, 'boolean');
      expectTypeError(() => Module.statusString(Symbol('status')));
      expectTypeError(() => Module.errorDisplayName(Symbol('error')));
      expectTypeError(() => Module.setLogMinLevel(Symbol('level')));
    } finally {
      wb.delete();
    }
  });

  test('WASM string free function is primitive-only', () => {
    expectTypeError(() => Module.evalFormula(Symbol('formula')));
    expectTypeError(() =>
      Module.evalFormula({
        toString() {
          throw new Error('toString');
        },
      }),
    );
    const result = Module.evalFormula('=1+1');
    assert.equal(result.status.ok, true);
    assert.equal(result.value.number, 2);
  });

  test('WASM paging positions reject invalid cursors and retain nullish defaults', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setNumber(0, 0, 0, 42).ok);
      for (const value of [-1, 1.5, Number.NaN, Number.POSITIVE_INFINITY, 2 ** 53, 2 ** 64]) {
        expectRangeError(() => wb.getCellsInRange(0, RANGE, value, 1));
        expectRangeError(() => wb.getCellsInRange(0, RANGE, 0, value));
        expectRangeError(() => wb.listInvalidCells(0, value, 1));
        expectRangeError(() => wb.listInvalidCells(0, 0, value));
      }
      expectTypeError(() => wb.getCellsInRange(0, RANGE, Symbol('cursor'), 1));
      expectTypeError(() => wb.listInvalidCells(0, Symbol('cursor'), 1));

      assert.equal(wb.getCellsInRange(0, RANGE, null, null).status.ok, true);
      assert.equal(wb.getCellsInRange(0, RANGE, undefined, undefined).status.ok, true);
      assert.equal(wb.listInvalidCells(0, null, null).status.ok, true);
      assert.equal(wb.listInvalidCells(0, undefined, undefined).status.ok, true);
      assert.equal(wb.getValue(0, 0, 0).value.number, 42);
    } finally {
      wb.delete();
    }
  });

  test('WASM-only table and auto-filter positions are guarded', () => {
    const wb = Module.Workbook.createDefault();
    try {
      for (const method of ['removeTable', 'updateTable', 'getSheetAutoFilterXml']) {
        expectRangeError(() => wb[method](2 ** 32, {}));
      }
      expectRangeError(() => wb.setSheetAutoFilterXml(2 ** 32, ''));
      expectTypeError(() => wb.setSheetAutoFilterXml(0, Symbol('xml')));

      const got = wb.getSheetAutoFilterXml(0);
      assert.equal(typeof got.status.ok, 'boolean');
      assert.equal(typeof got.xml, 'string');
      assert.equal(typeof wb.setSheetAutoFilterXml(0, '').ok, 'boolean');
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
      assert.equal(wb.setNumber.call(wb, 0, 0, 0, 42).ok, true);
      assert.equal(wb.getValue(0, 0, 0).value.number, 42);
    } finally {
      wb.delete();
    }
  });

  test('WASM overload leaves validate positions and track active Workbook calls', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setNumber(0, 0, 0, 42).ok);
      const leaf = (name, argCount) => {
        const table = Module.Workbook.prototype[name].overloadTable;
        assert.ok(Array.isArray(table), `${name} overload table missing`);
        const entry = table.find((candidate) => candidate?.argCount === argCount);
        assert.equal(typeof entry, 'function', `${name} overload leaf missing`);
        return entry;
      };
      const getCellsLeaf = leaf('getCellsInRange', 4);
      const listInvalidLeaf = leaf('listInvalidCells', 3);
      const insertImageLeaf = leaf('insertImage', 3);

      expectRangeError(() => getCellsLeaf.call(wb, 0, RANGE, 2 ** 64, 1));
      expectRangeError(() => listInvalidLeaf.call(wb, 0, -1, 1));
      expectRangeError(() => insertImageLeaf.call(wb, 2 ** 32, new Uint8Array(), {}));
      assert.equal(wb.getValue(0, 0, 0).value.number, 42);

      assert.equal(wb.getCellsInRange(0, RANGE, 0, 1).status.ok, true);
      assert.equal(wb.listInvalidCells(0, 0, 1).status.ok, true);
      assert.equal(typeof wb.insertImage(0, new Uint8Array(), {}).status.ok, 'boolean');

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
        name: 'TypeError',
        message: /cannot dispose Workbook during an active method call/,
      });
      assert.equal(wb.isDeleted(), false);
      assert.equal(wb.getValue(0, 0, 0).value.number, 42);
    } finally {
      if (!wb.isDeleted()) wb.delete();
    }
  });

  test('WASM guards reject reentrant Workbook deletion during nested reads', () => {
    for (const lifecycle of ['delete', 'deleteLater', Symbol.dispose]) {
      const wb = Module.Workbook.createDefault();
      try {
        const before = wb.xfCount().value;
        const spec = {};
        Object.defineProperty(spec, 'horizontalAlign', {
          get() {
            wb[lifecycle]();
            return 1;
          },
        });
        assert.throws(
          () => wb.addXf(spec),
          /cannot dispose Workbook during an active method call/,
          `${String(lifecycle)} unexpectedly succeeded`,
        );
        assert.equal(wb.xfCount().value, before);
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
        const before = wb.xfCount().value;
        const spec = {};
        Object.defineProperty(spec, 'horizontalAlign', {
          get() {
            descriptor.value.call(wb);
            return 1;
          },
        });
        assert.throws(
          () => wb.addXf(spec),
          /cannot dispose Workbook during an active method call/,
          `${String(lifecycle)} inherited call unexpectedly succeeded`,
        );
        assert.equal(wb.xfCount().value, before);
        assert.equal(wb.isDeleted(), false);
      } finally {
        deleteAfterLifecycleProbe(wb, lifecycle);
      }
    }
  });

  test('WASM lifecycle depth follows empty Proxy Workbook receivers', () => {
    const wb = Module.Workbook.createDefault();
    const proxy = new Proxy(wb, {});
    try {
      assert.ok(proxy.setNumber(0, 0, 0, 42).ok);

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
      const originalDeleteResult = proxy.addXf(originalDeleteSpec);
      assert.ok(originalDeleteResult.status.ok, JSON.stringify(originalDeleteResult.status));
      assert.ok(originalDeleteError instanceof TypeError);
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
      const proxyDeleteResult = proxy.addXf(proxyDeleteSpec);
      assert.ok(proxyDeleteResult.status.ok, JSON.stringify(proxyDeleteResult.status));
      assert.ok(proxyDeleteError instanceof TypeError);
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
      const result = wb.addXf(spec);
      assert.ok(result.status.ok, JSON.stringify(result.status));
      assert.ok(caught instanceof TypeError);
      assert.equal(wb.isDeleted(), false);
    } finally {
      if (!wb.isDeleted()) wb.delete();
    }
  });
}
