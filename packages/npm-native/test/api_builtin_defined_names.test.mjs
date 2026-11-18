// A defined name spelled as the profile locale's built-in name is stored as that built-in.

import assert from 'node:assert/strict';
import test from 'node:test';

import { getModule } from './smoke_support.mjs';

async function withWorkbook(profileId, fn) {
  const f = await getModule();
  const wb = f.Workbook.createForProfile(profileId);
  try {
    return await fn(wb, f);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

test('Print_Area defines the print area of the sheet its formula names', async () => {
  await withWorkbook('mac-365-ja_JP', (wb) => {
    wb.addSheet('Sheet2');
    wb.setDefinedName({ name: 'print_area', formula: '=Sheet2!$A$1:$B$2', comment: 'note', hidden: true });
    assert.deepEqual(wb.listDefinedNames(), [
      { name: '_xlnm.Print_Area', formula: '=Sheet2!$A$1:$B$2', scope: 1, comment: 'note', hidden: true },
    ]);
    assert.deepEqual(wb.sheet(1).getPrintArea(), [{ firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 }]);
    assert.equal(wb.getDefinedName('Print_Area'), null);
    assert.equal(wb.getDefinedName('Print_Area', 1).name, '_xlnm.Print_Area');
    wb.sheet(1).setPrintArea('A1:C3');
    assert.equal(wb.listDefinedNames().length, 1);
    assert.equal(wb.getDefinedName('PRINT_AREA', 1).formula, '=Sheet2!$A$1:$C$3');
    wb.removeDefinedName('Print_Area', 1);
    assert.deepEqual(wb.listDefinedNames(), []);
  });
});

test('only the locale spelling is a built-in, and the stored spelling needs a sheet', async () => {
  await withWorkbook('mac-365-de_DE', (wb, f) => {
    const sheet = wb.sheet(0).name;
    wb.setDefinedName({ name: 'Print_Area', formula: `=${sheet}!$A$1:$B$2` });
    wb.setDefinedName({ name: 'Druckbereich', formula: `=${sheet}!$A$1:$B$2` });
    assert.deepEqual(
      wb.listDefinedNames().map((n) => [n.name, n.scope]),
      [
        ['Print_Area', null],
        ['_xlnm.Print_Area', 0],
      ],
    );
    assert.throws(
      () => wb.setDefinedName({ name: '_xlnm.Print_Titles', formula: `=${sheet}!$1:$1` }),
      (e) => e.code === f.StatusCode.InvalidArgument,
    );
  });
});
