import { spawnSync } from 'node:child_process';
import test from 'node:test';
import { assert, getModule, pkgRoot } from './smoke_support.mjs';

const RANGE = { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 };

test('validation and merge range integers reject narrowing before mutation', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const before = wb.sheet(0).listValidations().length;
    for (const value of [256, Number.NaN, Number.POSITIVE_INFINITY, 1.5, -1]) {
      assert.throws(
        () => wb.sheet(0).addValidation({ ranges: [RANGE], type: value }),
        RangeError,
        `validation type=${String(value)}`,
      );
      assert.equal(wb.sheet(0).listValidations().length, before);
    }

    assert.throws(
      () => wb.sheet(0).addValidation({ ranges: [{ ...RANGE, firstRow: 2 ** 32 }] }),
      RangeError,
      'validation range firstRow',
    );
    assert.equal(wb.sheet(0).listValidations().length, before);
  } finally {
    wb.dispose();
  }
});

test('filter kind is validated without aborting and date-group bytes do not wrap', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const before = wb.sheet(0).getAutoFilter();
    assert.throws(
      () =>
        wb.sheet(0).setAutoFilter({
          range: RANGE,
          columns: [{ columnOffset: 0, kind: 1, dateGroups: [{ month: 257 }] }],
        }),
      RangeError,
    );
    assert.deepEqual(wb.sheet(0).getAutoFilter(), before);

    wb.sheet(0).setAutoFilter({ range: RANGE, columns: [{ columnOffset: 0, kind: 0 }] });
    assert.equal(wb.sheet(0).getAutoFilter().columns[0].kind, 0);
  } finally {
    wb.dispose();
  }
});

test('filter wrong primitive is caught in a child process instead of aborting Node', async () => {
  await getModule();
  const moduleUrl = new URL('../dist/index.mjs', import.meta.url).href;
  const script = `
    import { createFormulon } from ${JSON.stringify(moduleUrl)};
    const mod = await createFormulon();
    const wb = mod.Workbook.createDefault();
    try {
      wb.sheet(0).setAutoFilter({ range: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }, columns: [{ kind: 'bad' }] });
      process.exitCode = 2;
    } catch (error) {
      if (!(error instanceof TypeError)) process.exitCode = 3;
      else if (wb.sheet(0).getAutoFilter() !== null) process.exitCode = 4;
      else console.log('caught');
    } finally {
      wb.dispose();
    }
  `;
  const child = spawnSync(process.execPath, ['--input-type=module', '--eval', script], {
    cwd: pkgRoot,
    encoding: 'utf8',
  });
  assert.equal(child.signal, null, `child signal=${child.signal}\n${child.stderr}`);
  assert.equal(child.status, 0, `child status=${child.status}\n${child.stderr}`);
  assert.match(child.stdout, /caught/);
});

test('threaded mention and image integer fields reject before C mutation', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const beforeComments = wb.sheet(0).listThreadedComments().length;
    assert.throws(
      () =>
        wb.sheet(0).addThreadedComment({
          id: '{AAAAAAAA-0000-0000-0000-000000000001}',
          personId: '{11111111-1111-1111-1111-111111111111}',
          text: 'x',
          row: 2 ** 32,
          mentions: [{ personId: 'p', mentionId: 'm', start: 2 ** 32 }],
        }),
      RangeError,
    );
    assert.equal(wb.sheet(0).listThreadedComments().length, beforeComments);

    assert.throws(
      () => wb.sheet(0).addPicture(new Uint8Array([1, 2, 3]), { anchor: { rowOffsetEmu: Number.NaN } }),
      RangeError,
    );
    assert.equal(wb.sheet(0).listDrawings().length, 0);

    // The offsets fit the binding's int64 range; only the image bytes are refused, by the host.
    assert.throws(
      () =>
        wb.sheet(0).addPicture(new Uint8Array([1, 2, 3]), {
          anchor: { rowOffsetEmu: -(2 ** 63), colOffsetEmu: 2 ** 63 - 1024 },
        }),
      (e) => e.name === 'FormulonError',
    );

    assert.throws(
      () => wb.sheet(0).addPicture(new Uint8Array([1, 2, 3]), { anchor: { widthEmu: 2 ** 63 } }),
      RangeError,
    );
    assert.equal(wb.sheet(0).listDrawings().length, 0);
  } finally {
    wb.dispose();
  }
});

test('UI reader stops at a throwing getter and leaves the workbook unchanged', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const filter = { range: RANGE, columns: [{}] };
    Object.defineProperty(filter.columns[0], 'kind', {
      enumerable: true,
      get() {
        throw new Error('filter getter');
      },
    });
    assert.throws(() => wb.sheet(0).setAutoFilter(filter), /filter getter/);
    assert.equal(wb.sheet(0).getAutoFilter(), null);

    const validation = { ranges: [RANGE] };
    Object.defineProperty(validation, 'type', {
      enumerable: true,
      get() {
        throw new Error('validation getter');
      },
    });
    assert.throws(() => wb.sheet(0).addValidation(validation), /validation getter/);
    assert.equal(wb.sheet(0).listValidations().length, 0);
  } finally {
    wb.dispose();
  }
});

test('nullish UI fields retain their defaults', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    wb.sheet(0).setAutoFilter({
      range: { ...RANGE, lastCol: 1 },
      columns: [
        { columnOffset: null, kind: undefined },
        { columnOffset: 1, kind: 1, dateGroups: [{ month: null }] },
      ],
      sort: null,
    });
    const stored = wb.sheet(0).getAutoFilter();
    assert.equal(stored.columns[0].columnOffset, 0);
    assert.equal(stored.columns[0].kind, 0);
    assert.equal(stored.columns[1].dateGroups[0].month, 0);
  } finally {
    wb.dispose();
  }
});
