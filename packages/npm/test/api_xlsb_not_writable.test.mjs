// Features the xlsb writer cannot store fail the save with IoFeatureNotWritable; protected-range and sparkline
// inputs outside Excel's limits fail with InvalidArgument.

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

import { getModule, isFormulonError } from './smoke_support.mjs';

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..', '..');
const fixtureDir = path.join(repoRoot, 'tests', 'fixtures', 'excel', 'xlsb_parity');

function fixture(...parts) {
  return new Uint8Array(readFileSync(path.join(fixtureDir, ...parts)));
}

function cells(count) {
  return Array.from({ length: count }, (_, i) => ({ firstRow: i, firstCol: 0, lastRow: i, lastCol: 0 }));
}

function invalidArgument(f, context) {
  return (e) => {
    assert.ok(isFormulonError(e, f.StatusCode.InvalidArgument), String(e));
    if (context) assert.ok(e.context.includes(context), e.context);
    return true;
  };
}

async function withWorkbook(fn) {
  const f = await getModule();
  const wb = f.Workbook.createDefault();
  try {
    return await fn(wb, f);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

test('an xlsb save of a protected range with a security descriptor fails and names the feature and sheet', async () => {
  await withWorkbook((wb, f) => {
    wb.sheet(0).addProtectedRange({ name: 'guarded', ranges: 'A1:B2', securityDescriptor: 'O:WDG:WDD:(A;;CC;;;WD)' });
    assert.throws(
      () => wb.save(f.WorkbookFormat.Xlsb),
      (e) => {
        assert.ok(isFormulonError(e, f.StatusCode.IoFeatureNotWritable), String(e));
        assert.match(e.context, /kind=protected_range_descriptor/);
        assert.match(e.context, /sheet_index=0/);
        return true;
      },
    );
    assert.ok(wb.save(f.WorkbookFormat.Xlsx).length > 0);
  });
});

test('a protected range name of 255 UTF-16 units is accepted and 256 is InvalidArgument', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    assert.equal(ws.addProtectedRange({ name: 'a'.repeat(255), ranges: 'A1:B2' }), 'a'.repeat(255));
    assert.throws(
      () => ws.addProtectedRange({ name: 'b'.repeat(256), ranges: 'A1:B2' }),
      (e) => isFormulonError(e, f.StatusCode.InvalidArgument),
    );
    // An astral character counts as two units: 128 of them are 256.
    assert.throws(
      () => ws.addProtectedRange({ name: '\u{1F600}'.repeat(128), ranges: 'A1:B2' }),
      (e) => isFormulonError(e, f.StatusCode.InvalidArgument),
    );
  });
});

test('a sparkline source that is a union fails with InvalidArgument and a 3-D source is accepted', async () => {
  await withWorkbook((wb, f) => {
    wb.addSheet('Sheet2');
    const ws = wb.sheet(0);
    assert.throws(
      () => ws.addSparklineGroup({ sparklines: [{ source: 'Sheet1!A1:A3,Sheet1!C1:C3', location: 'E1' }] }),
      (e) => isFormulonError(e, f.StatusCode.InvalidArgument),
    );
    ws.addSparklineGroup({ sparklines: [{ source: 'Sheet1:Sheet2!A1:A3', location: 'E1' }] });
    assert.equal(ws.listSparklineGroups().length, 1);
  });
});

test('an xlsb save of a sparkline with a union source fails and names the feature and sheet', async () => {
  const f = await getModule();
  const wb = f.Workbook.load(fixture('sparklines', 'union_source_src.xlsx'));
  try {
    assert.throws(
      () => wb.save(f.WorkbookFormat.Xlsb),
      (e) => {
        assert.ok(isFormulonError(e, f.StatusCode.IoFeatureNotWritable), String(e));
        assert.match(e.context, /kind=sparklines/);
        assert.match(e.context, /sheet_index=0/);
        return true;
      },
    );
    assert.ok(wb.save(f.WorkbookFormat.Xlsx).length > 0);
  } finally {
    wb.dispose();
  }
});

test('an xlsb save of a protected range with a 256-unit name fails and names the feature', async () => {
  const f = await getModule();
  const wb = f.Workbook.load(fixture('protection', 'title_256_src.xlsx'));
  try {
    assert.equal(wb.sheet(0).getProtectedRange(0).name.length, 256);
    assert.throws(
      () => wb.save(f.WorkbookFormat.Xlsb),
      (e) => {
        assert.ok(isFormulonError(e, f.StatusCode.IoFeatureNotWritable), String(e));
        assert.match(e.context, /kind=protected_range_title/);
        return true;
      },
    );
    assert.ok(wb.save(f.WorkbookFormat.Xlsx).length > 0);
  } finally {
    wb.dispose();
  }
});

test('sparkline inputs Excel cannot store are InvalidArgument', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    assert.throws(
      () => ws.addSparklineGroup({ sparklines: [{ source: 'Sheet1!E1:E6', location: 'A1:A2' }] }),
      invalidArgument(f),
    );
    assert.throws(
      () =>
        ws.addSparklineGroup({
          dateAxisRange: 'Sheet1!A1:A3,Sheet1!C1:C3',
          sparklines: [{ source: 'Sheet1!E1:E6', location: 'G1' }],
        }),
      invalidArgument(f, 'date_axis_ref='),
    );
    assert.throws(
      () => ws.addSparklineGroup({ sparklines: [{ source: 'Sheet1!#REF!', location: 'G1' }] }),
      invalidArgument(f, 'formula='),
    );
    assert.equal(ws.listSparklineGroups().length, 0);
  });
});

test('a sparkline update with a union source is InvalidArgument and keeps the group', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const id = ws.addSparklineGroup({ sparklines: [{ source: 'Sheet1!E1:E6', location: 'G1' }] });
    assert.throws(
      () => ws.updateSparklineGroup(id, { sparklines: [{ source: 'Sheet1!A1:A3,Sheet1!C1:C3', location: 'G1' }] }),
      invalidArgument(f, 'formula='),
    );
    assert.ok(ws.getSparklineGroup(id).sparklines[0].source);
  });
});

test('a protected range holds 8191 rectangles and not 8192, and its name is not empty', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    assert.equal(ws.addProtectedRange({ name: 'ok', ranges: cells(8191) }), 'ok');
    assert.ok(wb.save(f.WorkbookFormat.Xlsb).length > 0);
    assert.throws(
      () => ws.addProtectedRange({ name: 'big', ranges: cells(8192) }),
      invalidArgument(f, 'part=ranges requested=8192 limit=8191'),
    );
    assert.throws(() => ws.addProtectedRange({ name: '', ranges: 'A1:B2' }), invalidArgument(f));
    assert.throws(
      () => ws.updateProtectedRange('ok', { ranges: cells(8192) }),
      invalidArgument(f, 'part=ranges requested=8192 limit=8191'),
    );
    assert.equal(ws.listProtectedRanges().length, 1);
  });
});

test('a protected range name past 255 units is refused with its count, on add and on update', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    assert.throws(
      () => ws.addProtectedRange({ name: 'c'.repeat(256), ranges: 'A1:B2' }),
      invalidArgument(f, 'part=name requested=256 limit=255'),
    );
    ws.addProtectedRange({ name: 'ok', ranges: 'A1:B2' });
    assert.throws(
      () => ws.updateProtectedRange('ok', { name: 'd'.repeat(256) }),
      invalidArgument(f, 'part=name requested=256 limit=255'),
    );
    assert.equal(ws.getProtectedRange('ok').name, 'ok');
  });
});

test('a sparkline whose source is #REF! reads back without one and the read-back value is accepted', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    const id = ws.addSparklineGroup({ sparklines: [{ source: '#REF!', location: 'G1' }] });
    const read = ws.getSparklineGroup(id).sparklines[0];
    assert.ok(read.source === null || read.source === '');
    ws.addSparklineGroup({ sparklines: [{ source: read.source, location: 'G2' }] });
    ws.addSparklineGroup({ sparklines: [{ source: null, location: 'G3' }] });
    ws.updateSparklineGroup(id, { sparklines: [read] });
    assert.equal(ws.listSparklineGroups().length, 3);
  });
});

for (const [kind, parts] of [
  ['format_defaults', ['format', 'base_256_src.xlsx']],
  ['format_defaults', ['format', 'width_256_src.xlsx']],
  ['outline_level', ['format', 'outline_row_8_src.xlsx']],
  ['sheet_view', ['views', 'bad_window_id_src.xlsx']],
]) {
  test(`an xlsb save of ${parts.join('/')} fails with kind=${kind} and the xlsx save succeeds`, async () => {
    const f = await getModule();
    const wb = f.Workbook.load(fixture(...parts));
    try {
      assert.throws(
        () => wb.save(f.WorkbookFormat.Xlsb),
        (e) => {
          assert.ok(isFormulonError(e, f.StatusCode.IoFeatureNotWritable), String(e));
          assert.ok(e.context.includes(`kind=${kind}`), e.context);
          assert.match(e.context, /sheet_index=0/);
          return true;
        },
      );
      assert.ok(wb.save(f.WorkbookFormat.Xlsx).length > 0);
    } finally {
      wb.dispose();
    }
  });
}

test('format defaults outside Excel limits are InvalidArgument and write nothing', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.updateFormatDefaults({ defaultColumnWidthChars: 10, defaultRowHeightPt: 20, baseColumnWidthChars: 9 });
    const before = ws.getFormatDefaults();
    for (const patch of [
      { defaultColumnWidthChars: 255.5 },
      { defaultColumnWidthChars: 256 },
      { defaultRowHeightPt: 409.7 },
      { baseColumnWidthChars: 8.5 },
      { baseColumnWidthChars: 256 },
    ]) {
      assert.throws(() => ws.updateFormatDefaults(patch), invalidArgument(f), JSON.stringify(patch));
      const after = ws.getFormatDefaults();
      assert.equal(after.defaultColumnWidthChars, before.defaultColumnWidthChars);
      assert.equal(after.defaultRowHeightPt, before.defaultRowHeightPt);
      assert.equal(after.baseColumnWidthChars, before.baseColumnWidthChars);
    }
  });
});

test('format defaults at Excel limits are accepted', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.updateFormatDefaults({ defaultColumnWidthChars: 255, defaultRowHeightPt: 409.6, baseColumnWidthChars: 255 });
    const d = ws.getFormatDefaults();
    assert.equal(d.defaultColumnWidthChars, 255);
    assert.equal(d.defaultRowHeightPt, 409.6);
    assert.equal(d.baseColumnWidthChars, 255);
    ws.updateFormatDefaults({ baseColumnWidthChars: 0 });
    assert.equal(ws.getFormatDefaults().baseColumnWidthChars, 0);
  });
});
