import { spawnSync } from 'node:child_process';
import test from 'node:test';
import { assert, getModule, pkgRoot } from './smoke_support.mjs';

const RANGE = { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 };

function statusOf(result) {
  return typeof result.ok === 'boolean' ? result : result.status;
}

test('validation and merge range integers reject narrowing before mutation', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const before = wb.getValidations(0).length;
    for (const value of [256, Number.NaN, Number.POSITIVE_INFINITY, 1.5, -1]) {
      assert.throws(
        () => wb.addValidation(0, { ranges: [RANGE], type: value }),
        RangeError,
        `validation type=${String(value)}`,
      );
      assert.equal(wb.getValidations(0).length, before);
    }

    assert.throws(
      () => wb.addValidation(0, { ranges: [{ ...RANGE, firstRow: 2 ** 32 }] }),
      RangeError,
      'validation range firstRow',
    );
    assert.equal(wb.getValidations(0).length, before);
  } finally {
    wb.dispose();
  }
});

test('filter kind is validated without aborting and date-group bytes do not wrap', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const before = wb.getAutoFilter(0).autoFilter;
    assert.throws(
      () =>
        wb.setAutoFilter(0, {
          range: RANGE,
          columns: [{ colId: 0, kind: 1, dateGroups: [{ month: 257 }] }],
        }),
      RangeError,
    );
    assert.deepEqual(wb.getAutoFilter(0).autoFilter, before);

    const keep = wb.setAutoFilter(0, { range: RANGE, columns: [{ colId: 0, kind: 0 }] });
    assert.ok(statusOf(keep).ok, JSON.stringify(keep));
    assert.equal(wb.getAutoFilter(0).autoFilter.columns[0].kind, 0);
  } finally {
    wb.dispose();
  }
});

test('filter wrong primitive is caught in a child process instead of aborting Node', async () => {
  await getModule();
  const moduleUrl = new URL('../dist/index.mjs', import.meta.url).href;
  const script = `
    import mod from ${JSON.stringify(moduleUrl)};
    const wb = mod.Workbook.createDefault();
    try {
      wb.setAutoFilter(0, { range: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }, columns: [{ kind: 'bad' }] });
      process.exitCode = 2;
    } catch (error) {
      if (!(error instanceof TypeError)) process.exitCode = 3;
      else if (wb.getAutoFilter(0).autoFilter !== null) process.exitCode = 4;
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
    const beforeComments = wb.getThreadedComments(0).comments.length;
    assert.throws(
      () =>
        wb.addThreadedComment(0, {
          id: '{AAAAAAAA-0000-0000-0000-000000000001}',
          personId: '{11111111-1111-1111-1111-111111111111}',
          row: 2 ** 32,
          mentions: [{ personId: 'p', mentionId: 'm', start: 2 ** 32 }],
        }),
      RangeError,
    );
    assert.equal(wb.getThreadedComments(0).comments.length, beforeComments);

    assert.throws(() => wb.insertImage(0, new Uint8Array([1, 2, 3]), { rowOffEmu: Number.NaN }), RangeError);
    assert.equal(wb.listDrawingObjects(0).length, 0);

    assert.doesNotThrow(() =>
      wb.insertImage(0, new Uint8Array([1, 2, 3]), {
        rowOffEmu: -(2 ** 63),
        colOffEmu: 2 ** 63 - 1024,
      }),
    );

    assert.throws(() => wb.insertImage(0, new Uint8Array([1, 2, 3]), { widthEmu: 2 ** 63 }), RangeError);
    assert.equal(wb.listDrawingObjects(0).length, 0);
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
    assert.throws(() => wb.setAutoFilter(0, filter), /filter getter/);
    assert.equal(wb.getAutoFilter(0).autoFilter, null);

    const validation = { ranges: [RANGE] };
    Object.defineProperty(validation, 'type', {
      enumerable: true,
      get() {
        throw new Error('validation getter');
      },
    });
    assert.throws(() => wb.addValidation(0, validation), /validation getter/);
    assert.equal(wb.getValidations(0).length, 0);
  } finally {
    wb.dispose();
  }
});

test('nullish UI fields retain their defaults', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const result = wb.setAutoFilter(0, {
      range: { ...RANGE, lastCol: 1 },
      columns: [
        { colId: null, kind: undefined },
        { colId: 1, kind: 1, dateGroups: [{ month: null }] },
      ],
      sort: null,
    });
    assert.ok(statusOf(result).ok, JSON.stringify(result));
    const stored = wb.getAutoFilter(0).autoFilter;
    assert.equal(stored.columns[0].colId, 0);
    assert.equal(stored.columns[0].kind, 0);
    assert.equal(stored.columns[1].dateGroups[0].month, 0);
  } finally {
    wb.dispose();
  }
});
