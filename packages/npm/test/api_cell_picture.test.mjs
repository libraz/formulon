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

test('a formula over a picture cell returns the picture, ISERROR is false and the picture is reported', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const bytes = png(2, 2);
    ws.setCellPicture('A1', { bytes, altText: 'alt' });
    ws.setCellPicture('A2', { bytes: jpeg(3, 3) });
    ws.setFormula('B1', '=A1');
    ws.setFormula('C1', '=ISERROR(A1)');
    ws.setFormula('D3', '=A1:A2');
    wb.recalc();
    const ids = { A1: ws.getTypedValue('A1').pictureId, A2: ws.getTypedValue('A2').pictureId };
    const expected = { B1: ids.A1, D3: ids.A1, D4: ids.A2 };
    for (const [a1, id] of Object.entries(expected)) {
      assert.deepEqual(ws.getTypedValue(a1), { kind: f.ValueKind.Picture, pictureId: id }, a1);
      const value = ws.getValue(a1);
      assert.ok(value instanceof f.ExcelPicture, a1);
      assert.equal(value.pictureId, id, a1);
      assert.equal(value.toString(), '[picture]');
    }
    assert.equal(ws.getValue('C1'), false);
    const formulaPicture = ws.getCellPicture('B1');
    assert.deepEqual([...formulaPicture.bytes], [...bytes]);
    assert.equal(formulaPicture.altText, 'alt');
    assert.equal(ws.getCellPicture('D3').format, f.ImageFormat.Png);
    assert.equal(ws.getCellPicture('D4').format, f.ImageFormat.Jpeg);
    assert.equal(ws.getCellPicture('C1'), null);
    assert.deepEqual(wb.getCellPicture(ids.A1), ws.getCellPicture('A1'));
  });
});

test('a picture value read from one cell places the same picture in another', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setCellPicture('A1', { bytes: png(2, 2), altText: 'alt' });
    ws.setFormula('B2', '=1+1');
    const read = ws.getValue('A1');
    ws.setValue('B2', read);
    assert.equal(ws.getFormula('B2'), null);
    assert.deepEqual(ws.getCellPicture('B2'), ws.getCellPicture('A1'));
    assert.equal(ws.getValue('B2').pictureId, read.pictureId);
    ws.setValue('B3', { kind: f.ValueKind.Picture, pictureId: read.pictureId });
    assert.equal(ws.getValue('B3').pictureId, read.pictureId);
    ws.setFormula('C1', '=B3');
    wb.recalc();
    assert.equal(ws.getValue('C1').pictureId, read.pictureId);
  });
});

test('a picture id nothing names is NotFound on a write and null on a read', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    assert.equal(wb.getCellPicture(0), null);
    assert.equal(wb.getCellPicture(2 ** 31), null);
    assert.throws(() => ws.setValue('A1', new f.ExcelPicture(999)), isStatus(f.StatusCode.NotFound));
    assert.throws(
      () => ws.setValue('A1', { kind: f.ValueKind.Picture, pictureId: 2 ** 31 }),
      isStatus(f.StatusCode.NotFound),
    );
    assert.equal(ws.getValue('A1'), null);
  });
});

test('an id whose last cell was cleared still names its picture and can be placed again', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setCellPicture('A1', { bytes: png(2, 2), altText: 'kept' });
    const { pictureId } = ws.getValue('A1');
    ws.setCellPicture('A1', null);
    assert.equal(wb.getCellPicture(pictureId).altText, 'kept');
    ws.setValue('B1', new f.ExcelPicture(pictureId));
    assert.equal(ws.getCellPicture('B1').altText, 'kept');
  });
});

test('ExcelPicture and the picture id arguments follow the boundary table', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    assert.equal(f.ValueKind.Picture, 8);
    assert.ok(Object.is(new f.ExcelPicture(-0).pictureId, 0));
    assert.equal(new f.ExcelPicture(0).pictureId, 0);
    assert.equal(new f.ExcelPicture(2 ** 32 - 1).pictureId, 2 ** 32 - 1);
    assert.equal(String(new f.ExcelPicture(7)), '[picture]');
    assert.throws(() => {
      new f.ExcelPicture(1).pictureId = 2;
    }, TypeError);
    assert.equal(wb.getCellPicture(-0), wb.getCellPicture(0));
    const rangeErrors = [Number.NaN, Infinity, -Infinity, 1.5, 2 ** 32];
    const typeErrors = [null, undefined, true, '1', 'ä', {}];
    for (const id of rangeErrors) {
      assert.throws(() => new f.ExcelPicture(id), RangeError, String(id));
      assert.throws(() => wb.getCellPicture(id), RangeError, String(id));
      assert.throws(() => ws.setValue('A1', { kind: f.ValueKind.Picture, pictureId: id }), RangeError, String(id));
    }
    for (const id of typeErrors) {
      assert.throws(() => new f.ExcelPicture(id), TypeError, String(id));
      assert.throws(() => wb.getCellPicture(id), TypeError, String(id));
      assert.throws(() => ws.setValue('A1', { kind: f.ValueKind.Picture, pictureId: id }), TypeError, String(id));
    }
    assert.throws(() => wb.getCellPicture(), TypeError);
    assert.throws(() => ws.setValue('A1', { kind: f.ValueKind.Picture }), TypeError);
    assert.throws(() => ws.setValues(0, 0, [[undefined]]), TypeError);
    assert.throws(() => new f.ExcelPicture(-1), isStatus(f.StatusCode.InvalidArgument));
    assert.throws(() => wb.getCellPicture(-1), isStatus(f.StatusCode.InvalidArgument));
    assert.throws(
      () => ws.setValue('A1', { kind: f.ValueKind.Picture, pictureId: -1 }),
      isStatus(f.StatusCode.InvalidArgument),
    );
    assert.equal(ws.getValue('A1'), null);
  });
});

test('removing a picture leaves a formula cell and a spill member producing one', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setCellPicture('A1', { bytes: png(2, 2) });
    ws.setCellPicture('A2', { bytes: jpeg(3, 3) });
    ws.setFormula('B1', '=A1');
    ws.setFormula('D3', '=A1:A2');
    wb.recalc();
    ws.setCellPicture('B1', null);
    ws.setCellPicture('D4', null);
    assert.equal(ws.getFormula('B1'), '=A1');
    assert.notEqual(ws.getCellPicture('B1'), null);
    assert.notEqual(ws.getCellPicture('D4'), null);
  });
});

test('alt text loses trailing spaces and an alt left empty reads null', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setCellPicture('A1', { bytes: png(2, 2), altText: 'x  ' });
    ws.setCellPicture('A2', { bytes: png(2, 2), altText: '   ' });
    assert.equal(ws.getCellPicture('A1').altText, 'x');
    assert.equal(ws.getCellPicture('A2').altText, null);
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

test('a coordinate past the grid is refused', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
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

test('formatValue refuses a picture value', async () => {
  await withWorkbook((wb, f) => {
    wb.sheet(0).setCellPicture('A1', { bytes: png(2, 2) });
    const picture = wb.sheet(0).getValue('A1');
    assert.throws(() => wb.formatValue(picture, 'General'), isStatus(f.StatusCode.InvalidArgument));
  });
});
