import test from 'node:test';
import { assert, expectFailure, getModule } from './smoke_support.mjs';

const kInvalidArgument = 2;
const kInvalidReference = 1004;

const dispose = (wb) => wb.dispose();

test('spillInfo reports a spill blocked by a value', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    wb.sheet(0).setFormula(1, 3, '=SEQUENCE(5)');
    wb.sheet(0).setText(3, 3, 'x');
    wb.recalc();
    const blocked = wb.sheet(0).getSpill(1, 3);
    assert.equal('engaged' in blocked, false);
    assert.equal(blocked.blocked, true);
    assert.deepEqual(blocked.range, { firstRow: 1, firstCol: 3, lastRow: 5, lastCol: 3 });
    // Only the anchor reports the attempt.
    assert.equal(wb.sheet(0).getSpill(2, 3), null);
    wb.sheet(0).setBlank(3, 3);
    wb.recalc();
    const engaged = wb.sheet(0).getSpill(1, 3);
    assert.equal(engaged.blocked, false);
    assert.equal(wb.sheet(0).getSpill(0, 0), null);
  } finally {
    dispose(wb);
  }
});

test('parseReference reads the INDIRECT grammar and fails with 1004 on anything else', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const ref = (text, options) => {
      const r = wb.parseReference(text, options);
      return r;
    };
    assert.deepEqual(ref('A1'), { sheetName: null, range: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 } });
    assert.deepEqual(ref('$B$2:C3'), { sheetName: null, range: { firstRow: 1, firstCol: 1, lastRow: 2, lastCol: 2 } });
    assert.deepEqual(ref("'My Sheet'!A1:B2"), {
      sheetName: 'My Sheet',
      range: { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 },
    });
    assert.equal(ref('売上!A1').sheetName, '売上');
    assert.equal(ref("'O''Brien'!A1").sheetName, "O'Brien");
    assert.deepEqual(ref('B:D'), {
      sheetName: null,
      range: { firstRow: 0, firstCol: 1, lastRow: 1048575, lastCol: 3 },
    });
    assert.deepEqual(ref('2:11'), {
      sheetName: null,
      range: { firstRow: 1, firstCol: 0, lastRow: 10, lastCol: 16383 },
    });
    expectFailure(() => ref('XFE1'), kInvalidReference);
    expectFailure(() => ref(''), kInvalidReference);
    expectFailure(() => ref(' A1'), kInvalidReference);
    expectFailure(() => ref('A1,B2'), kInvalidReference);
    assert.deepEqual(ref('R2C3', { r1c1: true }), {
      sheetName: null,
      range: { firstRow: 1, firstCol: 2, lastRow: 1, lastCol: 2 },
    });
    assert.deepEqual(ref('R[-1]C', { r1c1: true, baseRow: 5, baseCol: 5 }), {
      sheetName: null,
      range: { firstRow: 4, firstCol: 5, lastRow: 4, lastCol: 5 },
    });
    expectFailure(() => ref('R[-1]C', { r1c1: true }), kInvalidReference);
    assert.throws(() => ref('R[-1]C', { r1c1: true, baseRow: 5 }), TypeError);
    assert.throws(() => ref('R[-1]C', { r1c1: true, baseCol: 5 }), TypeError);
    expectFailure(() => wb.parseReference('A1', { baseRow: 2000000, baseCol: 0 }), kInvalidArgument);
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
    const fmt = (range, options) => {
      const r = wb.formatReference(range, options);
      return r;
    };
    const refused = (range) => {
      try {
        wb.formatReference(range);
      } catch (e) {
        assert.equal(e.name, 'FormulonError');
        return e.code;
      }
      return assert.fail('expected formatReference to throw');
    };
    assert.equal(fmt({ range: { firstRow: 0, firstCol: 1, lastRow: 1048575, lastCol: 3 } }), 'B:D');
    assert.equal(fmt({ range: { firstRow: 1, firstCol: 0, lastRow: 10, lastCol: 16383 } }), '2:11');
    assert.equal(fmt({ range: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 } }), 'A1');
    assert.equal(
      fmt({ sheetName: 'My Sheet', range: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 } }),
      "'My Sheet'!A1",
    );
    assert.equal(fmt({ sheetName: 'TRUE', range: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 } }), "'TRUE'!A1");
    assert.equal(fmt({ sheetName: '売上', range: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 } }), '売上!A1');
    assert.equal(fmt({ sheetName: null, range: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 } }), 'A1');
    assert.equal(fmt({ sheetName: '', range: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 } }), 'A1');
    assert.equal(fmt({ range: { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 } }, { absolute: true }), '$A$1:$B$2');
    assert.equal(refused({ range: { firstRow: 0, firstCol: 0, lastRow: 1048576, lastCol: 0 } }), kInvalidArgument);
    assert.equal(refused({ range: { firstRow: 2, firstCol: 0, lastRow: 1, lastCol: 0 } }), kInvalidArgument);
    assert.throws(() => fmt({ range: { firstRow: 0, firstCol: 0, lastRow: 0 } }), TypeError);
    assert.throws(() => fmt({ sheetName: 'Sheet1' }), TypeError);
    for (const sheetName of ['Sheet1', 'My Sheet', "O'Brien", '売上', 'Sheet.1', '1a', 'A1', 'TRUE']) {
      const range = { sheetName, range: { firstRow: 1, firstCol: 2, lastRow: 4, lastCol: 6 } };
      assert.deepEqual(wb.parseReference(fmt(range)), range, sheetName);
    }
  } finally {
    dispose(wb);
  }
});
