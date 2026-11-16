// A print-titles span string that does not parse fails with the parser's status and diagnostic, as every A1
// argument does.

import assert from 'node:assert/strict';
import test from 'node:test';

import { getModule, isFormulonError } from './smoke_support.mjs';

test('setPrintTitles passes a span parse failure through unchanged', async () => {
  const f = await getModule();
  const wb = f.Workbook.createDefault();
  try {
    const ws = wb.sheet(0);
    for (const titles of [
      { rows: '1:!', columns: null },
      { rows: null, columns: 'A:!' },
    ]) {
      assert.throws(
        () => ws.setPrintTitles(titles),
        (e) => isFormulonError(e, f.StatusCode.ParserInvalidReference) && e.parseError !== undefined,
      );
    }
    assert.throws(
      () => ws.setPrintTitles({ rows: 'A:B', columns: null }),
      (e) => isFormulonError(e, f.StatusCode.InvalidArgument),
    );
    assert.equal(ws.getPrintTitles(), null);
  } finally {
    wb.dispose();
  }
});
