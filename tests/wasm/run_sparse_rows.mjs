import assert from 'node:assert/strict';

const ROWS = 50000;
const LAST_COL = 16383;

export function registerSparseRows(Module, test) {
  test('rows with a wide gap between stored cells neither abort nor count the gap', () => {
    const wb = Module.Workbook.createDefault();
    try {
      for (let r = 0; r < ROWS; r += 1) {
        wb.setNumber(0, r, 0, r);
        wb.setNumber(0, r, LAST_COL, r);
      }
      const added = wb.addCellXf({
        fontIndex: 0,
        fillIndex: 0,
        borderIndex: 0,
        numFmtId: 0,
        horizontal: 0,
        vertical: 2,
        wrapText: true,
      });
      wb.setRangeXfIndex(0, { firstRow: 0, firstCol: 0, lastRow: ROWS - 1, lastCol: 0 }, added);
      wb.setRangeXfIndex(0, { firstRow: 0, firstCol: LAST_COL, lastRow: ROWS - 1, lastCol: LAST_COL }, added);

      // Gap columns are not stored, so the count is exactly the written cells.
      const count = wb.getCellCount(0);
      assert.equal(count, ROWS * 2);

      const first = wb.getCellAt(0, 0);
      assert.equal(first.row, 0);
      assert.equal(first.col, 0);

      const second = wb.getCellAt(0, 1);
      assert.equal(second.row, 0);
      assert.equal(second.col, LAST_COL);

      const last = wb.getCellAt(0, ROWS * 2 - 1);
      assert.equal(last.row, ROWS - 1);
      assert.equal(last.col, LAST_COL);
    } finally {
      wb.delete();
    }
  });
}
