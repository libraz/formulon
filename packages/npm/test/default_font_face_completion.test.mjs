// A bare rename of the default font saves with the family and charset Excel completes for the face; a stated
// family or charset is written as given, and so is an unchanged name.
// FORMULON_NPM_ENTRY selects the staged subpath: "." (default) or "./threads".

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const entry = process.env.FORMULON_NPM_ENTRY ?? '.';

async function loadPackage() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  const target = pkg.exports?.[entry]?.import;
  if (!target) throw new Error(`package.json "exports" has no import target for ${entry}`);
  return import(pathToFileURL(path.resolve(pkgRoot, target)).href);
}

const f = await (await loadPackage()).createFormulon();

/** Saves a ja-JP workbook after `edit(wb)` and reads the default font back from the reloaded package. */
function savedDefaultFont(edit) {
  const wb = f.Workbook.createForProfile('mac-365-ja_JP');
  try {
    edit(wb);
    const back = f.Workbook.load(wb.save(f.WorkbookFormat.Xlsx));
    try {
      return back.getDefaultFont();
    } finally {
      back.dispose();
    }
  } finally {
    wb.dispose();
  }
}

// Mac Excel 16.115: [face, family, charset], null when the font states none.
const MEASURED = [
  ['Calibri', 2, null],
  ['Aptos', null, null],
  ['游明朝', 1, 128],
  ['等线', 4, 134],
  ['맑은 고딕', 2, 129],
  ['新細明體', 1, 136],
  ['ＭＳ Ｐゴシック', 2, 128],
  ['Meiryo', 2, 128],
  ['Arial', 2, null],
  ['Cambria', 1, null],
  ['Courier New', 1, null],
];

describe('setDefaultFont face completion', () => {
  for (const [face, family, charset] of MEASURED) {
    it(`a bare rename to ${face} saves family ${family} and charset ${charset}`, () => {
      const font = savedDefaultFont((wb) => wb.setDefaultFont({ name: face }));
      assert.equal(font.name, face);
      assert.equal(font.family, family);
      assert.equal(font.charset, charset);
    });
  }

  it('a rename to 游ゴシック from another face saves family 3', () => {
    const font = savedDefaultFont((wb) => {
      wb.setDefaultFont({ name: 'Arial' });
      wb.setDefaultFont({ name: '游ゴシック' });
    });
    assert.equal(font.family, 3);
    assert.equal(font.charset, 128);
  });

  it('a stated charset or family is written as given', () => {
    const charset = savedDefaultFont((wb) => wb.setDefaultFont({ name: 'Courier New', charset: 128 }));
    assert.equal(charset.family, null);
    assert.equal(charset.charset, 128);
    const family = savedDefaultFont((wb) => wb.setDefaultFont({ name: 'Arial', family: 1 }));
    assert.equal(family.family, 1);
    assert.equal(family.charset, null);
  });

  it('the stock default keeps its own record', () => {
    const font = savedDefaultFont(() => {});
    assert.equal(font.name, '游ゴシック');
    assert.equal(font.family, 2);
    assert.equal(font.charset, 128);
  });
});
