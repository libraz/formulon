// Empty-value rules of the raw sheet-XML setters, setCellPhonetic and the header/footer section patch:
// `""` is refused where it is not one element, and means "no value" where Excel stores one state for both.

import assert from 'node:assert/strict';
import test from 'node:test';

import { getModule } from './smoke_support.mjs';

async function withSheet(fn) {
  const f = await getModule();
  const wb = f.Workbook.createDefault();
  try {
    return await fn(wb.sheet(0), f);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

const XML_SETTERS = [
  ['AutoFilterXml', '<autoFilter ref="A1:C10"/>'],
  ['PageSetupXml', '<pageSetup paperSize="9" orientation="landscape"/>'],
  ['PageMarginsXml', '<pageMargins left="1" right="1" top="1" bottom="1" header="0.5" footer="0.5"/>'],
  ['PrintOptionsXml', '<printOptions gridLines="1"/>'],
  ['HeaderFooterXml', '<headerFooter><oddHeader>&amp;CTitle</oddHeader></headerFooter>'],
  ['SheetPrXml', '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'],
];

for (const [name, fragment] of XML_SETTERS) {
  test(`set${name}: "" is refused and changes nothing`, async () => {
    await withSheet((ws, { StatusCode }) => {
      ws.ooxml[`set${name}`](fragment);
      const before = ws.ooxml[`get${name}`]();
      assert.notEqual(before, null);
      assert.throws(
        () => ws.ooxml[`set${name}`](''),
        (e) => e.name === 'FormulonError' && e.code === StatusCode.InvalidArgument,
      );
      assert.equal(ws.ooxml[`get${name}`](), before);
    });
  });

  test(`set${name}: "" on an absent element is refused and it stays absent`, async () => {
    await withSheet((ws, { StatusCode }) => {
      const before = ws.ooxml[`get${name}`]();
      assert.throws(
        () => ws.ooxml[`set${name}`](''),
        (e) => e.name === 'FormulonError' && e.code === StatusCode.InvalidArgument,
      );
      assert.equal(ws.ooxml[`get${name}`](), before);
    });
  });

  test(`set${name}: null removes the element`, async () => {
    await withSheet((ws) => {
      ws.ooxml[`set${name}`](fragment);
      assert.notEqual(ws.ooxml[`get${name}`](), null);
      ws.ooxml[`set${name}`](null);
      assert.equal(ws.ooxml[`get${name}`](), null);
      // Removing what is not there is not an error.
      ws.ooxml[`set${name}`](null);
      assert.equal(ws.ooxml[`get${name}`](), null);
    });
  });

  test(`set${name}: a value that is neither a string nor null is a TypeError`, async () => {
    await withSheet((ws) => {
      assert.throws(() => ws.ooxml[`set${name}`](123), TypeError);
    });
  });
}

test('setPageSetupXml: null resets the structured view, "" leaves it', async () => {
  await withSheet((ws, { StatusCode }) => {
    ws.ooxml.setPageSetupXml('<pageSetup orientation="landscape" scale="50"/>');
    assert.throws(
      () => ws.ooxml.setPageSetupXml(''),
      (e) => e.code === StatusCode.InvalidArgument,
    );
    assert.equal(ws.getPageSetup().scalePct, 50);
    ws.ooxml.setPageSetupXml(null);
    assert.equal(ws.getPageSetup().scalePct, 100);
    assert.equal(ws.getPageSetup().scalePctStated, false);
  });
});

test('setCellPhonetic: "" and null both clear the reading', async () => {
  await withSheet((ws) => {
    ws.setText('A1', 'kanji');
    for (const empty of ['', null]) {
      ws.setCellPhonetic('A1', 'kana');
      assert.equal(ws.getCellPhonetic('A1'), 'kana');
      ws.setCellPhonetic('A1', empty);
      assert.equal(ws.getCellPhonetic('A1'), null, `cleared by ${JSON.stringify(empty)}`);
    }
    ws.setCellPhonetic(0, 0, 'kana');
    ws.setCellPhonetic(0, 0, null);
    assert.equal(ws.getCellPhonetic(0, 0), null);
  });
});

test('updateHeaderFooter: a section given "" or null is removed and the others stay', async () => {
  await withSheet((ws) => {
    for (const empty of ['', null]) {
      ws.updateHeaderFooter({ oddHeader: '&CHead', oddFooter: '&CFoot' });
      assert.equal(ws.getHeaderFooter().oddHeader, '&CHead');
      ws.updateHeaderFooter({ oddHeader: empty });
      const hf = ws.getHeaderFooter();
      assert.equal(hf.oddHeader, null, `header cleared by ${JSON.stringify(empty)}`);
      assert.equal(hf.oddFooter, '&CFoot');

      ws.updateHeaderFooter({ oddFooter: empty });
      assert.equal(ws.getHeaderFooter().oddFooter, null, `footer cleared by ${JSON.stringify(empty)}`);
    }
  });
});
