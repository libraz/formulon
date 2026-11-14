// getCellPicture / setCellPicture: a picture placed in a cell reads back with its format, size and alt text; null
// removes it; a refused picture leaves the cell as it was.

import assert from 'node:assert/strict';
import test from 'node:test';

import { getModule } from './smoke_support.mjs';

async function withWorkbook(fn) {
  const f = await getModule();
  const wb = f.Workbook.createDefault();
  try {
    return await fn(wb, f);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

function png(width, height) {
  const bytes = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52];
  for (const v of [width, height]) bytes.push((v >>> 24) & 0xff, (v >>> 16) & 0xff, (v >>> 8) & 0xff, v & 0xff);
  bytes.push(8, 6, 0, 0, 0, 0x1f, 0x15, 0xc4, 0x89);
  return new Uint8Array(bytes);
}

const JPEG_COMPONENTS = [0x03, 0x01, 0x11, 0x00, 0x02, 0x11, 0x01, 0x03, 0x11, 0x01];

function jpeg(width, height) {
  const frame = [0xff, 0xc0, 0x00, 0x11, 0x08, height >> 8, height & 0xff, width >> 8, width & 0xff];
  return new Uint8Array([0xff, 0xd8, 0xff, ...frame.slice(1), ...JPEG_COMPONENTS, 0xff, 0xd9]);
}

const GIF = new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 4, 0, 4, 0, 0, 0, 0, 0x3b]);

const isStatus = (code) => (e) => e.name === 'FormulonError' && e.code === code;

test('setCellPicture places a PNG with alt text and getCellPicture reads it back', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const bytes = png(16, 8);
    ws.setCellPicture(1, 2, { bytes, altText: 'A red square' });
    const picture = ws.getCellPicture(1, 2);
    assert.deepEqual([...picture.bytes], [...bytes]);
    assert.equal(picture.altText, 'A red square');
    assert.equal(picture.format, f.ImageFormat.Png);
    assert.equal(picture.widthPx, 16);
    assert.equal(picture.heightPx, 8);
    assert.deepEqual(ws.getCellPicture('C2').altText, 'A red square');
  });
});

test('a JPEG without alt text is stored as given and reads altText null', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const bytes = jpeg(40, 24);
    ws.setCellPicture('A1', { bytes });
    ws.setCellPicture('A2', { bytes, altText: null });
    for (const a1 of ['A1', 'A2']) {
      const picture = ws.getCellPicture(a1);
      assert.deepEqual([...picture.bytes], [...bytes]);
      assert.equal(picture.altText, null);
      assert.equal(picture.format, f.ImageFormat.Jpeg);
      assert.equal(picture.widthPx, 40);
      assert.equal(picture.heightPx, 24);
    }
  });
});

test('the row and column form and the A1 form address the same cell', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setCellPicture('D5', { bytes: png(2, 2), altText: 'by a1' });
    assert.equal(ws.getCellPicture(4, 3).altText, 'by a1');
    ws.setCellPicture(5, 3, { bytes: png(3, 3), altText: 'by row and col' });
    assert.equal(ws.getCellPicture('D6').altText, 'by row and col');
  });
});

test('bytes may be an ArrayBuffer or a view and the result is a copy', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    const source = png(5, 5);
    ws.setCellPicture('A1', { bytes: source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength) });
    const padded = new Uint8Array(source.length + 4);
    padded.set(source, 2);
    ws.setCellPicture('A2', { bytes: padded.subarray(2, 2 + source.length) });
    for (const a1 of ['A1', 'A2']) assert.deepEqual([...ws.getCellPicture(a1).bytes], [...source]);
    source.fill(0);
    assert.notDeepEqual([...ws.getCellPicture('A1').bytes], [...source]);
  });
});

test('a CellPicture read back is accepted and format, widthPx and heightPx on input are ignored', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setCellPicture('A1', { bytes: png(7, 3), altText: 'copy me' });
    ws.setCellPicture('B1', ws.getCellPicture('A1'));
    assert.deepEqual(ws.getCellPicture('B1'), ws.getCellPicture('A1'));

    ws.setCellPicture('C1', { bytes: png(9, 4), format: f.ImageFormat.Gif, widthPx: 999, heightPx: 999 });
    const picture = ws.getCellPicture('C1');
    assert.equal(picture.format, f.ImageFormat.Png);
    assert.equal(picture.widthPx, 9);
    assert.equal(picture.heightPx, 4);
  });
});

test('a cell without a picture reads null, also one holding an ordinary value', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    assert.equal(ws.getCellPicture('A1'), null);
    ws.setValue('A1', 12);
    assert.equal(ws.getCellPicture(0, 0), null);
  });
});

test('setCellPicture null removes the picture, empties the cell and is a no-op when absent', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setCellPicture('B2', { bytes: png(2, 2), altText: 'x' });
    ws.setCellPicture('B2', null);
    assert.equal(ws.getCellPicture('B2'), null);
    assert.equal(ws.getValue('B2'), null);
    ws.setCellPicture('B2', null);
    ws.setValue('C3', 5);
    ws.setCellPicture('C3', null);
    assert.equal(ws.getValue('C3'), 5);
  });
});

test('a value written over a picture drops it and a picture replaces a formula', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setCellPicture('A1', { bytes: png(2, 2) });
    ws.setValue('A1', 3);
    assert.equal(ws.getCellPicture('A1'), null);
    ws.setFormula('A2', '=1+1');
    ws.setCellPicture('A2', { bytes: png(2, 2) });
    assert.notEqual(ws.getCellPicture('A2'), null);
  });
});

test('a formula reading a picture cell evaluates it as #VALUE!', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setCellPicture('A1', { bytes: png(2, 2), altText: 'alt' });
    ws.setFormula('B1', '=A1');
    ws.setFormula('C1', '=ISERROR(A1)');
    wb.recalc();
    for (const a1 of ['A1', 'B1']) {
      const value = ws.getValue(a1);
      assert.ok(value instanceof f.ExcelError);
      assert.equal(value.toString(), '#VALUE!');
    }
    assert.equal(ws.getValue('C1'), true);
  });
});

test('empty data and any other image format are refused with IoImageUnsupported and change nothing', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const original = png(2, 2);
    ws.setCellPicture('A1', { bytes: original, altText: 'keep' });
    for (const bytes of [new Uint8Array(0), GIF, new TextEncoder().encode('not an image')]) {
      assert.throws(() => ws.setCellPicture('A1', { bytes }), isStatus(f.StatusCode.IoImageUnsupported));
      assert.throws(() => ws.setCellPicture('B1', { bytes }), isStatus(f.StatusCode.IoImageUnsupported));
    }
    assert.equal(ws.getCellPicture('B1'), null);
    assert.equal(ws.getCellPicture('A1').altText, 'keep');
    assert.deepEqual([...ws.getCellPicture('A1').bytes], [...original]);
  });
});

test('data over 32 MiB is refused with SecurityResourceLimit', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const bytes = new Uint8Array(32 * 1024 * 1024 + 1);
    bytes.set(png(2, 2));
    assert.throws(() => ws.setCellPicture('A1', { bytes }), isStatus(f.StatusCode.SecurityResourceLimit));
    assert.equal(ws.getCellPicture('A1'), null);
  });
});

test('an empty altText is InvalidArgument and a coordinate past the grid is refused', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    assert.throws(
      () => ws.setCellPicture('A1', { bytes: png(2, 2), altText: '' }),
      isStatus(f.StatusCode.InvalidArgument),
    );
    assert.equal(ws.getCellPicture('A1'), null);
    assert.throws(() => ws.getCellPicture(1048576, 0), isStatus(f.StatusCode.InvalidArgument));
    assert.throws(() => ws.setCellPicture(0, 16384, null), isStatus(f.StatusCode.InvalidArgument));
    assert.throws(() => ws.getCellPicture('Sheet1!A1'), isStatus(f.StatusCode.InvalidArgument));
  });
});

test('a malformed picture argument is a TypeError', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    assert.throws(() => ws.setCellPicture('A1', undefined), TypeError);
    assert.throws(() => ws.setCellPicture('A1', 'picture'), TypeError);
    assert.throws(() => ws.setCellPicture('A1', {}), TypeError);
    assert.throws(() => ws.setCellPicture('A1', { bytes: 'abc' }), TypeError);
    assert.throws(() => ws.setCellPicture('A1', { bytes: null }), TypeError);
    assert.throws(() => ws.setCellPicture('A1', { bytes: png(2, 2), altText: 5 }), TypeError);
    assert.throws(() => ws.setCellPicture('A1', { bytes: png(2, 2), description: 'x' }), TypeError);
    assert.equal(ws.getCellPicture('A1'), null);
  });
});

test('a picture survives a save and load as .xlsx and as .xlsb', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const first = png(12, 9);
    const second = jpeg(30, 20);
    ws.setCellPicture('B2', { bytes: first, altText: 'a chart' });
    ws.setCellPicture('D3', { bytes: second });
    for (const format of [f.WorkbookFormat.Xlsx, f.WorkbookFormat.Xlsb]) {
      const loaded = f.Workbook.load(wb.save(format));
      try {
        const sheet = loaded.sheet(0);
        const a = sheet.getCellPicture('B2');
        assert.deepEqual([...a.bytes], [...first], `format ${format}`);
        assert.equal(a.altText, 'a chart');
        assert.equal(a.format, f.ImageFormat.Png);
        const b = sheet.getCellPicture('D3');
        assert.deepEqual([...b.bytes], [...second], `format ${format}`);
        assert.equal(b.altText, null);
        assert.equal(b.format, f.ImageFormat.Jpeg);
        assert.equal(sheet.getCellPicture('A1'), null);
      } finally {
        loaded.dispose();
      }
    }
  });
});
