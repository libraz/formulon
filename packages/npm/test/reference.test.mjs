import test from 'node:test';
import { assert, getModule } from './smoke_support.mjs';

const kInvalidArgument = 2;

const dispose = (wb) => wb.delete();

test('spillInfo reports a spill blocked by a value', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault().value;
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
  const wb = mod.Workbook.createDefault().value;
  try {
    const ref = (text, options) => {
      const r = wb.parseReference(text, options);
      assert.ok(r.status.ok, JSON.stringify(r.status));
      return r.value;
    };
    assert.deepEqual(ref('A1'), { sheetName: null, rect: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 } });
    assert.deepEqual(ref('$B$2:C3'), { sheetName: null, rect: { firstRow: 1, firstCol: 1, lastRow: 2, lastCol: 2 } });
    assert.deepEqual(ref("'My Sheet'!A1:B2"), {
      sheetName: 'My Sheet',
      rect: { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 },
    });
    assert.equal(ref('売上!A1').sheetName, '売上');
    assert.equal(ref("'O''Brien'!A1").sheetName, "O'Brien");
    assert.deepEqual(ref('B:D'), { sheetName: null, rect: { firstRow: 0, firstCol: 1, lastRow: 1048575, lastCol: 3 } });
    assert.deepEqual(ref('2:11'), { sheetName: null, rect: { firstRow: 1, firstCol: 0, lastRow: 10, lastCol: 16383 } });
    assert.equal(ref('XFE1'), null);
    assert.equal(ref(''), null);
    assert.equal(ref(' A1'), null);
    assert.equal(ref('A1,B2'), null);
    assert.deepEqual(ref('R2C3', { r1c1: true }), {
      sheetName: null,
      rect: { firstRow: 1, firstCol: 2, lastRow: 1, lastCol: 2 },
    });
    assert.deepEqual(ref('R[-1]C', { r1c1: true, baseRow: 5, baseCol: 5 }), {
      sheetName: null,
      rect: { firstRow: 4, firstCol: 5, lastRow: 4, lastCol: 5 },
    });
    assert.equal(ref('R[-1]C', { r1c1: true }), null);
    assert.throws(() => ref('R[-1]C', { r1c1: true, baseRow: 5 }), TypeError);
    assert.throws(() => ref('R[-1]C', { r1c1: true, baseCol: 5 }), TypeError);
    const offGrid = wb.parseReference('A1', { baseRow: 2000000, baseCol: 0 });
    assert.equal(offGrid.status.code, kInvalidArgument);
    assert.equal(offGrid.value, null);
    assert.throws(() => ref('A1', 5), TypeError);
    assert.throws(() => ref(5), TypeError);
  } finally {
    dispose(wb);
  }
});

test('formatReference writes A1 reference text and round-trips', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault().value;
  try {
    const fmt = (range, options) => {
      const r = wb.formatReference(range, options);
      assert.ok(r.status.ok, JSON.stringify(r.status));
      return r.value;
    };
    const refused = (range) => wb.formatReference(range).status.code;
    assert.equal(fmt({ rect: { firstRow: 0, firstCol: 1, lastRow: 1048575, lastCol: 3 } }), 'B:D');
    assert.equal(fmt({ rect: { firstRow: 1, firstCol: 0, lastRow: 10, lastCol: 16383 } }), '2:11');
    assert.equal(fmt({ rect: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 } }), 'A1');
    assert.equal(
      fmt({ sheetName: 'My Sheet', rect: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 } }),
      "'My Sheet'!A1",
    );
    assert.equal(fmt({ sheetName: 'TRUE', rect: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 } }), "'TRUE'!A1");
    assert.equal(fmt({ sheetName: '売上', rect: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 } }), '売上!A1');
    assert.equal(fmt({ sheetName: null, rect: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 } }), 'A1');
    assert.equal(fmt({ sheetName: '', rect: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 } }), 'A1');
    assert.equal(fmt({ rect: { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 } }, { absolute: true }), '$A$1:$B$2');
    assert.equal(refused({ rect: { firstRow: 0, firstCol: 0, lastRow: 1048576, lastCol: 0 } }), kInvalidArgument);
    assert.equal(refused({ rect: { firstRow: 2, firstCol: 0, lastRow: 1, lastCol: 0 } }), kInvalidArgument);
    assert.throws(() => fmt({ rect: { firstRow: 0, firstCol: 0, lastRow: 0 } }), TypeError);
    assert.throws(() => fmt({ sheetName: 'Sheet1' }), TypeError);
    for (const sheetName of ['Sheet1', 'My Sheet', "O'Brien", '売上', 'Sheet.1', '1a', 'A1', 'TRUE']) {
      const range = { sheetName, rect: { firstRow: 1, firstCol: 2, lastRow: 4, lastCol: 6 } };
      assert.deepEqual(wb.parseReference(fmt(range)).value, range, sheetName);
    }
  } finally {
    dispose(wb);
  }
});
