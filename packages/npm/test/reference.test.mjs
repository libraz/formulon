import test from 'node:test';
import { assert, getModule } from './smoke_support.mjs';

const dispose = (wb) => wb.delete();

test('spillInfo reports a spill blocked by a value', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.ok(wb.setFormula(0, 1, 3, '=SEQUENCE(5)').ok);
    assert.ok(wb.setText(0, 3, 3, 'x').ok);
    assert.ok(wb.recalc().ok);
    const blocked = wb.getSpillInfo(0, 1, 3);
    assert.equal(blocked.engaged, false);
    assert.equal(blocked.blocked, true);
    assert.deepEqual(blocked.rect, { firstRow: 1, firstCol: 3, lastRow: 5, lastCol: 3 });
    // Only the anchor reports the attempt.
    assert.equal(wb.getSpillInfo(0, 2, 3).blocked, false);
    assert.ok(wb.setBlank(0, 3, 3).ok);
    assert.ok(wb.recalc().ok);
    const engaged = wb.getSpillInfo(0, 1, 3);
    assert.equal(engaged.engaged, true);
    assert.equal(engaged.blocked, false);
    assert.equal(wb.getSpillInfo(0, 0, 0).blocked, false);
  } finally {
    dispose(wb);
  }
});

test('parseReference reads the INDIRECT grammar', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const ref = (text, options) => wb.parseReference(text, options);
    assert.deepEqual(ref('A1'), { sheetName: null, row0: 0, col0: 0, row1: 0, col1: 0 });
    assert.deepEqual(ref('$B$2:C3'), { sheetName: null, row0: 1, col0: 1, row1: 2, col1: 2 });
    assert.deepEqual(ref("'My Sheet'!A1:B2"), { sheetName: 'My Sheet', row0: 0, col0: 0, row1: 1, col1: 1 });
    assert.equal(ref('売上!A1').sheetName, '売上');
    assert.equal(ref("'O''Brien'!A1").sheetName, "O'Brien");
    assert.deepEqual(ref('B:D'), { sheetName: null, row0: 0, col0: 1, row1: 1048575, col1: 3 });
    assert.deepEqual(ref('2:11'), { sheetName: null, row0: 1, col0: 0, row1: 10, col1: 16383 });
    assert.equal(ref('XFE1'), null);
    assert.equal(ref(''), null);
    assert.equal(ref(' A1'), null);
    assert.equal(ref('A1,B2'), null);
    assert.deepEqual(ref('R2C3', { r1c1: true }), { sheetName: null, row0: 1, col0: 2, row1: 1, col1: 2 });
    assert.deepEqual(ref('R[-1]C', { r1c1: true, baseRow: 5, baseCol: 5 }), {
      sheetName: null,
      row0: 4,
      col0: 5,
      row1: 4,
      col1: 5,
    });
    assert.equal(ref('R[-1]C', { r1c1: true }), null);
    assert.throws(() => ref('R[-1]C', { r1c1: true, baseRow: 5 }), TypeError);
    assert.throws(() => ref('R[-1]C', { r1c1: true, baseCol: 5 }), TypeError);
    assert.throws(() => ref('A1', { baseRow: 2000000, baseCol: 0 }), RangeError);
    assert.throws(() => ref('A1', 5), TypeError);
    assert.throws(() => ref(5), TypeError);
  } finally {
    dispose(wb);
  }
});

test('formatReference writes A1 reference text and round-trips', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const fmt = (range, options) => wb.formatReference(range, options);
    assert.equal(fmt({ row0: 0, col0: 1, row1: 1048575, col1: 3 }), 'B:D');
    assert.equal(fmt({ row0: 1, col0: 0, row1: 10, col1: 16383 }), '2:11');
    assert.equal(fmt({ row0: 0, col0: 0, row1: 0, col1: 0 }), 'A1');
    assert.equal(fmt({ sheetName: 'My Sheet', row0: 0, col0: 0, row1: 0, col1: 0 }), "'My Sheet'!A1");
    assert.equal(fmt({ sheetName: 'TRUE', row0: 0, col0: 0, row1: 0, col1: 0 }), "'TRUE'!A1");
    assert.equal(fmt({ sheetName: '売上', row0: 0, col0: 0, row1: 0, col1: 0 }), '売上!A1');
    assert.equal(fmt({ sheetName: null, row0: 0, col0: 0, row1: 0, col1: 0 }), 'A1');
    assert.equal(fmt({ sheetName: '', row0: 0, col0: 0, row1: 0, col1: 0 }), 'A1');
    assert.equal(fmt({ row0: 0, col0: 0, row1: 1, col1: 1 }, { absolute: true }), '$A$1:$B$2');
    assert.throws(() => fmt({ row0: 0, col0: 0, row1: 1048576, col1: 0 }), RangeError);
    assert.throws(() => fmt({ row0: 2, col0: 0, row1: 1, col1: 0 }), RangeError);
    assert.throws(() => fmt({ row0: 0, col0: 0, row1: 0 }), TypeError);
    for (const sheetName of ['Sheet1', 'My Sheet', "O'Brien", '売上', 'Sheet.1', '1a', 'A1', 'TRUE']) {
      const range = { sheetName, row0: 1, col0: 2, row1: 4, col1: 6 };
      assert.deepEqual(wb.parseReference(fmt(range)), range, sheetName);
    }
  } finally {
    dispose(wb);
  }
});
