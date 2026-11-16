// A table range naming a sheet the workbook does not have is NotFound, as an unknown sheet name is everywhere.

import assert from 'node:assert/strict';
import test from 'node:test';

import { getModule, isFormulonError } from './smoke_support.mjs';

test('Table.update with a range on an unknown sheet is NotFound', async () => {
  const f = await getModule();
  const wb = f.Workbook.createDefault();
  try {
    const ws = wb.sheet(0);
    const table = ws.addTable({ range: 'A1:B3', name: 'Sales' });
    assert.throws(
      () => table.update({ range: 'Nope!A1:B4' }),
      (e) => isFormulonError(e, f.StatusCode.NotFound),
    );
    wb.addSheet('Other');
    assert.throws(
      () => table.update({ range: 'Other!A1:B4' }),
      (e) => isFormulonError(e, f.StatusCode.InvalidArgument),
    );
  } finally {
    wb.dispose();
  }
});
