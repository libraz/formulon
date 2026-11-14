// applyTheme / saveTheme: a theme file (.thmx) round-trips between workbooks; a refused file leaves the theme
// unchanged and fails with the status of its defect.

import assert from 'node:assert/strict';
import test from 'node:test';

import { getModule, isFormulonError } from './smoke_support.mjs';

const FONTS = {
  majorLatin: 'Major Source',
  majorEastAsian: 'Major EA',
  minorLatin: 'Minor Source',
  minorEastAsian: 'Minor EA',
};

async function withWorkbooks(fn) {
  const f = await getModule();
  const source = f.Workbook.createDefault();
  const target = f.Workbook.createDefault();
  try {
    return await fn(source, target, f);
  } finally {
    if (!source.disposed) source.dispose();
    if (!target.disposed) target.dispose();
  }
}

test('saveTheme of a fresh workbook is a zip carrying the theme part', async () => {
  await withWorkbooks((wb) => {
    const bytes = wb.saveTheme();
    assert.ok(bytes instanceof Uint8Array);
    assert.deepEqual([...bytes.subarray(0, 2)], [0x50, 0x4b]);
    assert.ok(Buffer.from(bytes).includes('theme/theme/theme1.xml'));
  });
});

test('saveTheme then applyTheme carries the theme fonts to another workbook', async () => {
  await withWorkbooks((source, target) => {
    source.setThemeFonts(FONTS);
    const expected = source.getThemeFonts();
    target.applyTheme(source.saveTheme());
    assert.deepEqual(target.getThemeFonts(), expected);
    assert.equal(target.getThemeFonts().minorLatin, 'Minor Source');
  });
});

test('applyTheme takes an ArrayBuffer as well as a Uint8Array', async () => {
  await withWorkbooks((source, target) => {
    source.setThemeFonts(FONTS);
    const bytes = source.saveTheme();
    target.applyTheme(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
    assert.equal(target.getThemeFonts().majorLatin, 'Major Source');
  });
});

test('applyTheme refuses empty bytes with InvalidArgument and leaves the theme unchanged', async () => {
  await withWorkbooks((_source, target, f) => {
    target.setThemeFonts(FONTS);
    const before = target.getThemeFonts();
    assert.throws(
      () => target.applyTheme(new Uint8Array(0)),
      (e) => isFormulonError(e, f.StatusCode.InvalidArgument),
    );
    assert.deepEqual(target.getThemeFonts(), before);
  });
});

test('applyTheme refuses bytes that are not a zip archive with IoZipCorrupt and leaves the theme unchanged', async () => {
  await withWorkbooks((_source, target, f) => {
    target.setThemeFonts(FONTS);
    const before = target.getThemeFonts();
    assert.throws(
      () => target.applyTheme(new TextEncoder().encode('not a zip archive')),
      (e) => isFormulonError(e, f.StatusCode.IoZipCorrupt),
    );
    assert.deepEqual(target.getThemeFonts(), before);
  });
});
