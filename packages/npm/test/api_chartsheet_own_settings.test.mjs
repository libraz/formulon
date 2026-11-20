// A chartsheet's protection, margins, header and footer, paper size and orientation set through the facade.

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

import { getModule, isFormulonError } from './smoke_support.mjs';

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..', '..');
const fixtureDir = path.join(repoRoot, 'tests', 'fixtures', 'excel', 'chartsheet_twins');

// `Sheet1` (a worksheet) then the chartsheet, from the Excel-saved twin `name`.
const fixture = (name, ext) => new Uint8Array(readFileSync(path.join(fixtureDir, `${name}.${ext}`)));

async function withBook(name, ext, fn) {
  const f = await getModule();
  const wb = f.Workbook.load(fixture(name, ext));
  try {
    return await fn(wb.sheet(1), f, wb);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

const refused = (f, message) => (e) => isFormulonError(e, f.StatusCode.InvalidArgument) && e.message.includes(message);

for (const ext of ['xlsx', 'xlsb']) {
  test(`setChartsheetProtection takes back what getChartsheetProtection reads (${ext})`, async () => {
    await withBook('cs_protect_hash', ext, (cs, f) => {
      const read = cs.getChartsheetProtection();
      assert.equal(read.algorithmName, 'SHA-512');
      cs.setChartsheetProtection(read);
      assert.deepEqual(cs.getChartsheetProtection(), read);
      cs.setChartsheetProtection({ allowEditObjects: true });
      assert.deepEqual(cs.getChartsheetProtection(), {
        algorithmName: null,
        hashValue: null,
        saltValue: null,
        spinCount: 0,
        legacyPasswordHash: null,
        allowEditContent: false,
        allowEditObjects: true,
      });
      assert.throws(() => cs.setChartsheetProtection({ hashValue: '' }), refused(f, 'hashValue'));
    });
  });

  test(`a chartsheet protection is never removed (${ext})`, async () => {
    await withBook('cs_control', ext, (cs, f, wb) => {
      assert.equal(cs.getChartsheetProtection(), null);
      cs.setChartsheetProtection(null);
      cs.setChartsheetProtection({ allowEditContent: true, allowEditObjects: true });
      assert.equal(cs.getChartsheetProtection(), null);
      cs.setChartsheetProtection({});
      assert.equal(cs.getChartsheetProtection().allowEditContent, false);
      const removed = refused(f, 'Chartsheet protection cannot be removed');
      assert.throws(() => cs.setChartsheetProtection(null), removed);
      assert.throws(() => cs.setChartsheetProtection({ allowEditContent: true, allowEditObjects: true }), removed);
      assert.throws(
        () => wb.sheet(0).setChartsheetProtection({}),
        refused(f, 'sheet_index is a worksheet, but this call needs a chartsheet'),
      );
    });
  });

  test(`paper size, margins and header / footer apply to a chartsheet (${ext})`, async () => {
    await withBook('cs_control', ext, async (cs, f, wb) => {
      cs.updatePageSetup({ paperSize: 5 });
      cs.updatePageMargins({ leftIn: 0.5, rightIn: 0.75, topIn: 1, bottomIn: 1.25, headerIn: 0.25, footerIn: 0.375 });
      cs.updateHeaderFooter({ oddHeader: '&CHDR' });
      assert.throws(() => cs.updatePageSetup({ scalePct: 50 }), refused(f, 'patch.scalePct states the scale'));
      cs.updatePageSetup({ scalePct: cs.getPageSetup().scalePct });

      const format = ext === 'xlsb' ? f.WorkbookFormat.Xlsb : f.WorkbookFormat.Xlsx;
      const reloaded = f.Workbook.load(wb.save(format));
      try {
        const again = reloaded.sheet(1);
        assert.equal(again.getPageSetup().paperSize, 5);
        assert.equal(again.getPageSetup().orientation, f.PageOrientation.Portrait);
        assert.equal(again.getPageMargins().leftIn, 0.5);
        assert.equal(again.getHeaderFooter().oddHeader, '&CHDR');
      } finally {
        reloaded.dispose();
      }
    });
  });
}
