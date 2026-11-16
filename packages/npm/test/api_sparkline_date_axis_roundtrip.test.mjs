// A sparkline group whose stored date-axis text is not one rectangle reads `dateAxisRange: null` and writes that
// text back unchanged, as a sparkline source does: passing back what a read returned changes nothing.

import assert from 'node:assert/strict';
import test from 'node:test';

import { getModule } from './smoke_support.mjs';
import { packagePart } from './zip_part.mjs';

const DATE_AXIS = 'Dates';

// The `<xm:f>` texts of the saved first sheet: each sparkline source and the group's date axis.
function savedRangeTexts(wb) {
  return [...packagePart(wb.save(), 'xl/worksheets/sheet1.xml').matchAll(/<xm:f>([^<]*)<\/xm:f>/g)].map((m) => m[1]);
}

test('a group read back and written again keeps a date axis that is not one rectangle', async () => {
  const f = await getModule();
  const wb = f.Workbook.createDefault();
  try {
    wb.setDefinedName({ name: DATE_AXIS, formula: 'Sheet1!$A$1:$D$1' });
    const ws = wb.sheet(0);
    const id = ws.addSparklineGroup({
      dateAxis: true,
      dateAxisRange: DATE_AXIS,
      sparklines: [{ source: 'Sheet1!A2:D2', location: 'F2' }],
    });
    const read = ws.getSparklineGroup(id);
    assert.equal(read.dateAxisRange, null);
    ws.updateSparklineGroup(id, read);
    assert.deepEqual(savedRangeTexts(wb), [DATE_AXIS, 'Sheet1!A2:D2']);
    ws.removeSparklineGroup(id);
    const added = ws.addSparklineGroup(read);
    assert.deepEqual(savedRangeTexts(wb), [DATE_AXIS, 'Sheet1!A2:D2']);
    ws.updateSparklineGroup(added, { dateAxisRange: 'Sheet1!A3:D3' });
    assert.deepEqual(ws.getSparklineGroup(added).dateAxisRange, {
      sheet: 0,
      rect: { firstRow: 2, firstCol: 0, lastRow: 2, lastCol: 3 },
    });
  } finally {
    wb.dispose();
  }
});
