// The style XF reads of one row or column refuse a coordinate the grid cannot hold, as every engine call does.

import assert from 'node:assert/strict';
import test from 'node:test';

import { getModule, isFormulonError } from './smoke_support.mjs';

test('getRowXfIndex and getColumnXfIndex refuse a coordinate beyond the grid', async () => {
  const f = await getModule();
  const wb = f.Workbook.createDefault();
  try {
    const ooxml = wb.sheet(0).ooxml;
    assert.equal(ooxml.getRowXfIndex(1048575), null);
    assert.equal(ooxml.getColumnXfIndex(16383), null);
    assert.throws(
      () => ooxml.getRowXfIndex(1048576),
      (e) => isFormulonError(e, f.StatusCode.InvalidArgument),
    );
    assert.throws(
      () => ooxml.getColumnXfIndex(16384),
      (e) => isFormulonError(e, f.StatusCode.InvalidArgument),
    );
  } finally {
    wb.dispose();
  }
});
