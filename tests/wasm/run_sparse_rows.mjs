import assert from 'node:assert/strict';

const ROWS = 50000;
const LAST_COL = 16383;

export function registerSparseRows(Module, test) {
  test('rows with a wide gap between stored cells neither abort nor count the gap', () => {
    const wb = Module.Workbook.createDefault();
    try {
      for (let r = 0; r < ROWS; r += 1) {
        assert.ok(wb.setNumber(0, r, 0, r).ok);
        assert.ok(wb.setNumber(0, r, LAST_COL, r).ok);
      }
      const added = wb.addXf({
        fontIndex: 0,
        fillIndex: 0,
        borderIndex: 0,
        numFmtId: 0,
        horizontalAlign: 0,
        verticalAlign: 2,
        wrapText: true,
      });
      assert.ok(added.status.ok);
      assert.ok(wb.setRangeXfIndex(0, 0, 0, ROWS - 1, 0, added.index).ok);
      assert.ok(wb.setRangeXfIndex(0, 0, LAST_COL, ROWS - 1, LAST_COL, added.index).ok);

      // Gap columns are not stored, so the count is exactly the written cells.
      const count = wb.cellCount(0);
      assert.ok(count.status.ok);
      assert.equal(count.value, ROWS * 2);

      const first = wb.cellAt(0, 0);
      assert.ok(first.status.ok);
      assert.equal(first.row, 0);
      assert.equal(first.col, 0);

      const second = wb.cellAt(0, 1);
      assert.ok(second.status.ok);
      assert.equal(second.row, 0);
      assert.equal(second.col, LAST_COL);

      const last = wb.cellAt(0, ROWS * 2 - 1);
      assert.ok(last.status.ok);
      assert.equal(last.row, ROWS - 1);
      assert.equal(last.col, LAST_COL);
    } finally {
      wb.delete();
    }
  });
}
