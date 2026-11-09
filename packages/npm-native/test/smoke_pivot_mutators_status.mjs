import test from 'node:test';
import { inflateRawSync } from 'node:zlib';
import { assert, getModule, isFormulonError, path, pkgRoot, readFile } from './smoke_support.mjs';

// ---- Pivot mutator marshalling -----------------------------------------
//
// Every pivot mutator is driven once against a small pivot and checked
// through a read of the result -- the projected grid wherever the mutation
// reaches it -- so an argument dropped, reordered or mis-converted on the
// way to the C ABI shows up as a wrong grid rather than a green `.ok`. The
// completeness test below keeps the table in step with the declaration
// file.

// Region / Product carry shared items and their records store item
// indices; Amount and Date are plain numbers. Only Region is laid out, so
// the baseline grid is two regions, their sums and the grand total.
function buildPivot(mod) {
  const wb = mod.Workbook.createDefault();
  wb.setProfileId('win-365-ja_JP');
  const cache = wb.addPivotCache();
  const { cacheId } = cache;
  wb.pivotCache(cacheId).setWorksheetSource({ range: 'A1:D5', sheet: 'Sheet1' });
  for (const name of ['Region', 'Product', 'Amount', 'Date']) wb.pivotCache(cacheId).addField(name);
  for (const item of ['East', 'West']) wb.pivotCache(cacheId).addFieldSharedItem(0, item);
  for (const item of ['Apple', 'Pear']) wb.pivotCache(cacheId).addFieldSharedItem(1, item);
  for (const values of [
    [0, 0, 10, 45292],
    [1, 0, 30, 45323],
    [0, 1, 5, 45658],
    [1, 1, 7, 45689],
  ]) {
    const rec = wb.pivotCache(cacheId).addRecord();
    for (const [field, v] of values.entries()) wb.pivotCache(cacheId).setRecordValue(rec, field, v);
  }
  const pivot = wb.sheet(0).addPivot({ name: 'Pivot1', cacheId, anchor: { row: 0, col: 6 } });
  const addField = (spec) => pivot.addField(spec);
  const region = addField({ sourceName: 'Region', axis: 0 });
  pivot.addFieldItem(region, 'East');
  pivot.addFieldItem(region, 'West');
  const product = addField({ sourceName: 'Product', axis: 2 });
  pivot.addFieldItem(product, 'Apple');
  pivot.addFieldItem(product, 'Pear');
  const amount = addField({ sourceName: 'Amount', axis: 2 });
  const date = addField({ sourceName: 'Date', axis: 2 });
  pivot.setRowFieldOrder([region]);
  pivot.addDataField({ name: 'Sum of Amount', fieldIndex: amount, aggregation: 0 });
  return { wb, mod, cacheId, pivot, region, product, amount, date };
}

// The projected grid as one `|`-joined string per sheet row.
function pivotGrid(wb, pivot = 0) {
  const layout = wb.sheet(0).pivot(pivot).getLayout();
  const rows = new Map();
  for (const cell of layout.cells) {
    const v = cell.value;
    const text =
      typeof v === 'number'
        ? String(v)
        : typeof v === 'string'
          ? v
          : v !== null && typeof v === 'object'
            ? `#${v.code}`
            : '';
    if (!rows.has(cell.row)) rows.set(cell.row, new Array(layout.range.lastCol - layout.range.firstCol + 1).fill(''));
    rows.get(cell.row)[cell.col - layout.range.firstCol] = text;
  }
  return [...rows.keys()].sort((a, b) => a - b).map((row) => rows.get(row).join('|'));
}

const BASE_GRID = ['行ラベル|Sum of Amount', 'East|15', 'West|37', '総計|52'];
const TWO_LEVEL_GRID = [
  '行ラベル|Sum of Amount',
  'East|15',
  'Apple|10',
  'Pear|5',
  'West|37',
  'Apple|30',
  'Pear|7',
  '総計|52',
];
const DATE_SERIAL_GRID = ['行ラベル|Sum of Amount', '45292|10', '45323|30', '45658|5', '45689|7', '総計|52'];
const twoLevels = (b) => b.pivot.setRowFieldOrder([b.region, b.product]);
const dateRows = (b) => b.pivot.setRowFieldOrder([b.date]);
const hideEast = (b) => b.pivot.updateFieldItem(b.region, 0, { visible: false });
const WEST_ONLY = ['行ラベル|Sum of Amount', 'West|37', '総計|37'];
const GREATER_THAN_20 = Object.freeze({ axis: 0, fieldName: 'Region', type: 1, value: 20 });
const BETWEEN_10_20 = Object.freeze({
  axis: 0,
  fieldName: 'Region',
  type: 2,
  value: 10,
  valueHigh: 20,
});
// Points record 0's Region at the shared item a row has just appended.
const pointAtNewItem = (b) => b.wb.pivotCache(b.cacheId).setRecordValue(0, 0, 2);

// The raw text of one stored part of a saved package.
function packagePart(bytes, name) {
  const buf = Buffer.from(bytes);
  let eocd = buf.length - 22;
  while (eocd >= 0 && buf.readUInt32LE(eocd) !== 0x06054b50) eocd -= 1;
  assert.ok(eocd >= 0, 'missing ZIP end record');
  let at = buf.readUInt32LE(eocd + 16);
  for (let i = buf.readUInt16LE(eocd + 10); i > 0; i -= 1) {
    const method = buf.readUInt16LE(at + 10);
    const size = buf.readUInt32LE(at + 20);
    const nameLen = buf.readUInt16LE(at + 28);
    const next = at + 46 + nameLen + buf.readUInt16LE(at + 30) + buf.readUInt16LE(at + 32);
    if (buf.toString('utf8', at + 46, at + 46 + nameLen) === name) {
      const local = buf.readUInt32LE(at + 42);
      const start = local + 30 + buf.readUInt16LE(local + 26) + buf.readUInt16LE(local + 28);
      const raw = buf.subarray(start, start + size);
      return (method === 0 ? raw : inflateRawSync(raw)).toString('utf8');
    }
    at = next;
  }
  assert.fail(`${name} not in the package`);
}

// [method, { setup, act, after, expect }]. `act` performs the one call
// under test; `expect` reads the result back.
const PIVOT_MUTATORS = [
  [
    'Workbook.addPivotCache',
    {
      act: (b) => b.wb.addPivotCache(77),
      expect: (b, r) => {
        assert.equal(r.cacheId, 77);
        assert.equal(r, b.wb.pivotCache(77));
        assert.equal(b.wb.listPivotCaches().length, 2);
      },
    },
  ],
  ['Pivot.getCache', { act: (b) => b.pivot.getCache(), expect: (b, r) => assert.equal(r, b.wb.pivotCache(b.cacheId)) }],
  [
    'Workbook.removePivotCache',
    {
      setup: (b) => {
        b.spare = b.wb.addPivotCache();
      },
      act: (b) => b.wb.removePivotCache(b.spare),
      expect: (b) => {
        assert.equal(b.wb.listPivotCaches().length, 1);
        assert.deepEqual(
          b.wb.listPivotCaches().map((c) => c.cacheId),
          [b.cacheId],
        );
        assert.throws(
          () => b.spare.getOptions(),
          (e) => isFormulonError(e, b.mod.StatusCode.NotFound),
        );
      },
    },
  ],
  [
    'PivotCache.setWorksheetSource',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).setWorksheetSource({ range: 'B2:E6', sheet: 'Sheet1' }),
      expect: (b) => {
        const source = b.wb.pivotCache(b.cacheId).getWorksheetSource();
        assert.equal(source.range.sheet, 0);
        assert.deepEqual({ ...source.range.rect }, { firstRow: 1, firstCol: 1, lastRow: 5, lastCol: 4 });
        assert.equal(source.name, null);
        assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
      },
    },
  ],
  [
    'PivotCache.addField',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).addField('Extra'),
      expect: (b, r) => {
        assert.equal(r, 4);
        assert.equal(b.wb.pivotCache(b.cacheId).listFields()[4].name, 'Extra');
      },
    },
  ],
  [
    'PivotCache.clearFields',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).clearFields(),
      expect: (b) => {
        assert.equal(b.wb.pivotCache(b.cacheId).listFields().length, 0);
        // The data field now names a cache field that no longer exists.
        assert.throws(() => b.pivot.getLayout(), isFormulonError);
      },
    },
  ],
  [
    'PivotCache.addFieldSharedItem (number)',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).addFieldSharedItem(0, 42.5),
      after: pointAtNewItem,
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '42.5|10', 'East|5', 'West|37', '総計|52']),
    },
  ],
  [
    'PivotCache.addFieldSharedItem (date)',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).addFieldSharedItem(3, new Date(2024, 0, 1)),
      expect: (b) => assert.equal(b.wb.pivotCache(b.cacheId).listFieldSharedItems(3)[0].isDate, true),
    },
  ],
  [
    'PivotCache.addFieldSharedItem',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).addFieldSharedItem(0, 'North'),
      after: pointAtNewItem,
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'North|10', 'West|37', '総計|52']),
    },
  ],
  [
    'PivotCache.addFieldSharedItem (bool)',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).addFieldSharedItem(0, true),
      after: pointAtNewItem,
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', 'TRUE|10', '総計|52']),
    },
  ],
  [
    'PivotCache.addFieldSharedItem (blank)',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).addFieldSharedItem(0, null),
      after: pointAtNewItem,
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '(空白)|10', '総計|52']),
    },
  ],
  [
    'PivotCache.addFieldSharedItem (error)',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).addFieldSharedItem(0, { kind: b.mod.ValueKind.Error, errorCode: 6 }),
      after: pointAtNewItem,
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '#N/A|10', '総計|52']),
    },
  ],
  [
    'PivotCache.clearFieldSharedItems',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).clearFieldSharedItems(0),
      // With no shared items the records' indices render as themselves.
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '0|15', '1|37', '総計|52']),
    },
  ],
  [
    'PivotCache.addRecord',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).addRecord(),
      expect: (b, r) => {
        assert.equal(r, 4);
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|15', 'West|37', '(空白)|', '総計|52']);
      },
    },
  ],
  [
    'PivotCache.clearRecords',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).clearRecords(),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '総計|']),
    },
  ],
  [
    'PivotCache.setRecordValue',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).setRecordValue(0, 2, 1000),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|1005', 'West|37', '総計|1042']),
    },
  ],
  [
    'PivotCache.setRecordValue (date)',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).setRecordValue(0, 3, new Date(2024, 0, 9)),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), BASE_GRID),
    },
  ],
  [
    'PivotCache.setRecordValue (text)',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).setRecordValue(0, 0, 'West'),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|47', '総計|52']),
    },
  ],
  [
    'PivotCache.setRecordValue (bool)',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).setRecordValue(0, 2, true),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '総計|42']),
    },
  ],
  [
    'PivotCache.setRecordValue (blank)',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).setRecordValue(0, 2, null),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '総計|42']),
    },
  ],
  [
    'PivotCache.setRecordValue (error)',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).setRecordValue(0, 2, { kind: b.mod.ValueKind.Error, errorCode: 1 }),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|#1', 'West|37', '総計|#1']),
    },
  ],
  [
    'Sheet.addPivot',
    {
      act: (b) => b.wb.sheet(0).addPivot({ name: 'Pivot2', cacheId: b.cacheId, anchor: { row: 20, col: 1 } }),
      expect: (b, r) => {
        assert.equal(r.index, 1);
        assert.equal(r.name, 'Pivot2');
        const layout = r.getLayout();
        assert.equal(layout.range.firstRow, 20);
        assert.equal(layout.range.firstCol, 1);
        assert.deepEqual(pivotGrid(b.wb, 0), BASE_GRID);
      },
    },
  ],
  [
    'Sheet.removePivot',
    {
      act: (b) => b.wb.sheet(0).removePivot(b.pivot),
      expect: (b) => {
        assert.equal(b.wb.sheet(0).listPivots().length, 0);
        assert.throws(() => b.wb.sheet(0).pivot(0), isFormulonError);
        assert.throws(
          () => b.pivot.getLayout(),
          (e) => isFormulonError(e, b.mod.StatusCode.NotFound),
        );
      },
    },
  ],
  [
    'Pivot.rename',
    {
      act: (b) => b.pivot.rename('Renamed'),
      // A pivot's name is only observable in the part the writer emits.
      expect: (b) => {
        const saved = b.wb.save();
        assert.match(packagePart(saved, 'xl/pivotTables/pivotTable1.xml'), /<pivotTableDefinition [^>]*name="Renamed"/);
      },
    },
  ],
  [
    'Pivot.setAnchor',
    {
      act: (b) => b.pivot.setAnchor({ row: 10, col: 2 }),
      expect: (b) => {
        const layout = b.pivot.getLayout();
        assert.equal(layout.range.firstRow, 10);
        assert.equal(layout.range.firstCol, 2);
        assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
      },
    },
  ],
  [
    'Pivot.updateGrandTotals',
    {
      setup: (b) => b.pivot.setColumnFieldOrder([b.product]),
      // Rows off drops the per-row total column; columns on keeps the
      // bottom total row, so a swapped pair would show the opposite.
      act: (b) => b.pivot.updateGrandTotals({ rows: false, columns: true }),
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), [
          'Sum of Amount|列ラベル|',
          '行ラベル|Apple|Pear',
          'East|10|5',
          'West|30|7',
          '総計|40|12',
        ]),
    },
  ],
  [
    'Pivot.updateStyleOptions',
    {
      act: (b) =>
        b.pivot.updateStyleOptions({
          styleName: 'PivotStyleMedium9',
          showRowHeaders: false,
          showColumnHeaders: true,
          showRowStripes: true,
          showColumnStripes: false,
          showLastColumn: true,
        }),
      expect: (b) => {
        const style = b.pivot.getStyleOptions();
        assert.equal(style.styleName, 'PivotStyleMedium9');
        assert.deepEqual(
          [
            style.showRowHeaders,
            style.showColumnHeaders,
            style.showRowStripes,
            style.showColumnStripes,
            style.showLastColumn,
          ],
          [false, true, true, false, true],
        );
        assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
      },
    },
  ],
  [
    'Pivot.setReportLayout',
    {
      act: (b) => b.pivot.setReportLayout(1),
      expect: (b) => {
        assert.equal(b.pivot.getReportLayout(), 1);
        assert.deepEqual(pivotGrid(b.wb), ['Region|Sum of Amount', 'East|15', 'West|37', '総計|52']);
      },
    },
  ],
  [
    'Pivot.addField',
    {
      act: (b) => b.pivot.addField({ sourceName: 'Product', customName: 'Fruit', axis: 3 }),
      expect: (b, r) => {
        // Product is a cache column the fixture already added, so it is updated in place.
        assert.equal(r, b.product);
        assert.deepEqual(pivotGrid(b.wb), ['Fruit|(すべて)', '|', ...BASE_GRID]);
      },
    },
  ],
  [
    'Pivot.clearFields',
    {
      setup: hideEast,
      act: (b) => b.pivot.clearFields(),
      expect: (b) => {
        assert.equal(b.pivot.listFields().length, 0);
        // The row order empties with the fields; East's hidden flag goes too.
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '|52']);
      },
    },
  ],
  [
    'Pivot.setFieldAxis',
    {
      act: (b) => b.pivot.setFieldAxis(b.product, 3),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['Product|(すべて)', '|', ...BASE_GRID]),
    },
  ],
  [
    'Pivot.updateField',
    {
      // East now outsums West, so sorting ascending by the data field
      // differs from both label orders.
      setup: (b) => b.wb.pivotCache(b.cacheId).setRecordValue(0, 2, 40),
      act: (b) => b.pivot.updateField(b.region, { sortOrder: 1, sortByField: 'Sum of Amount' }),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'West|37', 'East|45', '総計|82']),
    },
  ],
  [
    'Pivot.updateField (subtotalTop)',
    {
      setup: twoLevels,
      act: (b) => b.pivot.updateField(b.region, { subtotalTop: false }),
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), [
          '行ラベル|Sum of Amount',
          'East|',
          'Apple|10',
          'Pear|5',
          'East 集計|15',
          'West|',
          'Apple|30',
          'Pear|7',
          'West 集計|37',
          '総計|52',
        ]),
    },
  ],
  [
    'Pivot.addFieldItem',
    {
      act: (b) => b.pivot.addFieldItem(b.region, 'East', { visible: false }),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), WEST_ONLY),
    },
  ],
  [
    'Pivot.addFieldItem (cacheIndex)',
    {
      // Only the index-addressed form can name the blank item.
      setup: (b) => {
        b.wb.pivotCache(b.cacheId).addFieldSharedItem(0, null);
        b.wb.pivotCache(b.cacheId).setRecordValue(0, 0, null);
      },
      act: (b) => b.pivot.addFieldItem(b.region, { cacheIndex: 2 }, { visible: false }),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '総計|42']),
    },
  ],
  [
    'Pivot.clearFieldItems',
    {
      setup: hideEast,
      act: (b) => b.pivot.clearFieldItems(b.region),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), BASE_GRID),
    },
  ],
  [
    'Pivot.updateFieldItem (visible)',
    {
      act: (b) => b.pivot.updateFieldItem(b.region, 1, { visible: false }),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|15', '総計|15']),
    },
  ],
  [
    'Pivot.addFieldSubtotalFunction',
    {
      setup: twoLevels,
      act: (b) => b.pivot.addFieldSubtotalFunction(b.region, 3),
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), [
          '行ラベル|Sum of Amount',
          'East|10',
          'Apple|10',
          'Pear|5',
          'West|30',
          'Apple|30',
          'Pear|7',
          '総計|52',
        ]),
    },
  ],
  [
    'Pivot.clearFieldSubtotalFunctions',
    {
      setup: (b) => {
        twoLevels(b);
        b.pivot.addFieldSubtotalFunction(b.region, 3);
      },
      act: (b) => b.pivot.clearFieldSubtotalFunctions(b.region),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), TWO_LEVEL_GRID),
    },
  ],
  [
    'Pivot.setFieldDateGroup',
    {
      setup: dateRows,
      act: (b) => b.pivot.setFieldDateGroup(b.date, { periods: 64, calendar: 1 }),
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '令和6年|40', '令和7年|12', '総計|52']),
    },
  ],
  [
    'Pivot.setFieldDateGroup (null ungroups)',
    {
      setup: (b) => {
        dateRows(b);
        b.pivot.setFieldDateGroup(b.date, { periods: 64, calendar: 0 });
      },
      act: (b) => b.pivot.setFieldDateGroup(b.date, null),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), DATE_SERIAL_GRID),
    },
  ],
  [
    'Pivot.updateField (numberFormat)',
    {
      act: (b) => b.pivot.updateField(b.amount, { numberFormat: '#,##0.00' }),
      // A field's numFmtId is only observable in the part the writer emits.
      expect: (b) => {
        assert.throws(
          () => b.pivot.updateField(99, { numberFormat: '#,##0.00' }),
          (e) => isFormulonError(e, 13),
        );
        assert.equal(b.pivot.listFields()[b.amount].numberFormat, '#,##0.00');
        const saved = b.wb.save();
        assert.match(packagePart(saved, 'xl/pivotTables/pivotTable1.xml'), /<pivotField [^>]*numFmtId="4"/);
        assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
      },
    },
  ],
  [
    'Pivot.setRowFieldOrder',
    {
      act: (b) => b.pivot.setRowFieldOrder([b.product]),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'Apple|40', 'Pear|12', '総計|52']),
    },
  ],
  [
    'Pivot.setColumnFieldOrder',
    {
      act: (b) => b.pivot.setColumnFieldOrder([b.product]),
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), [
          'Sum of Amount|列ラベル||',
          '行ラベル|Apple|Pear|総計',
          'East|10|5|15',
          'West|30|7|37',
          '総計|40|12|52',
        ]),
    },
  ],
  [
    'Pivot.addDataField',
    {
      act: (b) => b.pivot.addDataField({ name: 'Max of Amount', fieldIndex: b.amount, aggregation: 3 }),
      expect: (b, r) => {
        assert.equal(r, 1);
        assert.deepEqual(pivotGrid(b.wb), [
          '行ラベル|Sum of Amount|Max of Amount',
          'East|15|10',
          'West|37|30',
          '総計|52|30',
        ]);
      },
    },
  ],
  [
    'Pivot.clearDataFields',
    {
      act: (b) => b.pivot.clearDataFields(),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル', 'East', 'West', '総計']),
    },
  ],
  [
    'Pivot.setDataField',
    {
      act: (b) =>
        b.pivot.setDataField(0, {
          name: 'Count of Amount',
          fieldIndex: b.amount,
          aggregation: 1,
          numberFormat: '2',
        }),
      expect: (b) => {
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Count of Amount', 'East|2', 'West|2', '総計|4']);
        const counts = b.wb
          .sheet(0)
          .pivot(b.pivot)
          .getLayout()
          .cells.filter((c) => typeof c.value === 'number');
        assert.deepEqual(
          counts.map((c) => c.numberFormat),
          ['2', '2', '2'],
        );
      },
    },
  ],
  [
    'Pivot.addFilter',
    {
      act: (b) => b.pivot.addFilter(GREATER_THAN_20),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), WEST_ONLY),
    },
  ],
  [
    'Pivot.clearFilters',
    {
      setup: (b) => b.pivot.addFilter(GREATER_THAN_20),
      act: (b) => b.pivot.clearFilters(),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), BASE_GRID),
    },
  ],
  [
    'Pivot.removeFilter',
    {
      // Each filter alone keeps a different region, so removing the wrong
      // index leaves the wrong one.
      setup: (b) => {
        b.pivot.addFilter(GREATER_THAN_20);
        b.pivot.addFilter(BETWEEN_10_20);
      },
      act: (b) => b.pivot.removeFilter(1),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), WEST_ONLY),
    },
  ],
  [
    'Pivot.updateField (defaultSubtotal)',
    {
      act: (b) => b.pivot.updateField(b.region, { defaultSubtotal: false }),
      expect: (b) => assert.equal(b.pivot.listFields()[b.region].defaultSubtotal, false),
    },
  ],
  [
    'Pivot.updateField (insertBlankRow)',
    {
      act: (b) => b.pivot.updateField(b.region, { insertBlankRow: true }),
      expect: (b) => assert.equal(b.pivot.listFields()[b.region].insertBlankRow, true),
    },
  ],
  [
    'Pivot.updateField (fillDownLabels)',
    {
      act: (b) => b.pivot.updateField(b.region, { fillDownLabels: true }),
      expect: (b) => assert.equal(b.pivot.listFields()[b.region].fillDownLabels, true),
    },
  ],
  [
    'Pivot.updateField',
    {
      act: (b) => b.pivot.updateField(b.region, { showDetail: false }),
      expect: (b) => assert.equal(b.pivot.getFieldItem(b.region, 0).showDetail, false),
    },
  ],
  [
    'Pivot.updateFieldItem (showDetail)',
    {
      act: (b) => b.pivot.updateFieldItem(b.region, 1, { showDetail: false }),
      expect: (b) => {
        assert.equal(b.pivot.getFieldItem(b.region, 0).showDetail, true);
        assert.equal(b.pivot.getFieldItem(b.region, 1).showDetail, false);
      },
    },
  ],
  [
    'Pivot.updateFieldItem (caption)',
    {
      act: (b) => b.pivot.updateFieldItem(b.region, 0, { caption: 'Eastern' }),
      expect: (b) => {
        assert.equal(b.pivot.getFieldItem(b.region, 0).caption, 'Eastern');
        assert.equal(b.pivot.getFieldItem(b.region, 1).caption, null);
      },
    },
  ],
  [
    'Pivot.showDetails',
    {
      act: (b) => b.pivot.showDetails(1, 7),
      expect: (b, r) => {
        assert.equal(r, b.wb.sheet(0));
        assert.equal(r.index, 0);
        assert.equal(b.wb.listSheets().length, 2);
      },
    },
  ],
  [
    'Pivot.updateOptions',
    {
      setup: (b) => {
        b.before = b.pivot.getOptions();
      },
      act: (b) => b.pivot.updateOptions({ errorCaption: 'oops', indent: 3, pageWrap: 2, showHeaders: false }),
      expect: (b) => {
        const o = b.pivot.getOptions();
        assert.equal(o.errorCaption, 'oops');
        assert.equal(o.indent, 3);
        assert.equal(o.pageWrap, 2);
        assert.equal(o.showHeaders, false);
        // Omitted keys keep their value.
        assert.equal(o.showDrill, b.before.showDrill);
      },
    },
  ],
  [
    'PivotCache.updateOptions',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).updateOptions({ saveData: false, missingItemsLimit: 5 }),
      expect: (b) => {
        const o = b.wb.pivotCache(b.cacheId).getOptions();
        assert.equal(o.saveData, false);
        assert.equal(o.missingItemsLimit, 5);
        assert.equal(o.refreshOnLoad, false);
      },
    },
  ],
  [
    'PivotCache.addCalculatedField',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).addCalculatedField('Double', 'Amount*2'),
      expect: (b, r) => {
        assert.equal(r, 4);
        const fields = b.wb.pivotCache(b.cacheId).listFields();
        assert.equal(fields[4].name, 'Double');
        assert.equal(fields[4].formula, '=Amount*2');
        assert.equal(fields[2].formula, null);
      },
    },
  ],
  [
    'PivotCache.setCalculatedFieldFormula',
    {
      setup: (b) => b.wb.pivotCache(b.cacheId).addCalculatedField('Double', 'Amount*2'),
      act: (b) => b.wb.pivotCache(b.cacheId).setCalculatedFieldFormula(4, 'Amount*3'),
      expect: (b) => assert.equal(b.wb.pivotCache(b.cacheId).listFields()[4].formula, '=Amount*3'),
    },
  ],
  [
    'PivotCache.removeCalculatedField',
    {
      setup: (b) => b.wb.pivotCache(b.cacheId).addCalculatedField('Double', 'Amount*2'),
      act: (b) => b.wb.pivotCache(b.cacheId).removeCalculatedField(4),
      expect: (b) => assert.equal(b.wb.pivotCache(b.cacheId).listFields().length, 4),
    },
  ],
  [
    'PivotCache.addCalculatedItem',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).addCalculatedItem(b.region, 'Both', 'East+West'),
      expect: (b) => {
        assert.equal(b.wb.pivotCache(b.cacheId).listCalculatedItems().length, 1);
        const item = b.wb.pivotCache(b.cacheId).listCalculatedItems()[0];
        assert.equal(item.fieldIndex, b.region);
        assert.equal(item.name, 'Both');
        assert.equal(item.formula, '=Region[East]+Region[West]');
      },
    },
  ],
  [
    'PivotCache.setCalculatedItemFormula',
    {
      setup: (b) => b.wb.pivotCache(b.cacheId).addCalculatedItem(b.region, 'Both', 'East+West'),
      act: (b) => b.wb.pivotCache(b.cacheId).setCalculatedItemFormula(0, 'East'),
      expect: (b) => assert.equal(b.wb.pivotCache(b.cacheId).listCalculatedItems()[0].formula, '=Region[East]'),
    },
  ],
  [
    'PivotCache.moveCalculatedItem',
    {
      setup: (b) => {
        b.wb.pivotCache(b.cacheId).addCalculatedItem(b.region, 'First', 'East');
        b.wb.pivotCache(b.cacheId).addCalculatedItem(b.region, 'Second', 'West');
      },
      act: (b) => b.wb.pivotCache(b.cacheId).moveCalculatedItem(0, 1),
      expect: (b) => {
        assert.equal(b.wb.pivotCache(b.cacheId).listCalculatedItems()[0].name, 'Second');
        assert.equal(b.wb.pivotCache(b.cacheId).listCalculatedItems()[1].name, 'First');
      },
    },
  ],
  [
    'PivotCache.removeCalculatedItem',
    {
      setup: (b) => b.wb.pivotCache(b.cacheId).addCalculatedItem(b.region, 'Both', 'East+West'),
      act: (b) => b.wb.pivotCache(b.cacheId).removeCalculatedItem(0),
      expect: (b) => assert.equal(b.wb.pivotCache(b.cacheId).listCalculatedItems().length, 0),
    },
  ],
  [
    'Pivot.updateField (customName)',
    {
      act: (b) => b.pivot.updateField(b.region, { customName: 'Area' }),
      expect: (b) => {
        assert.equal(b.pivot.listFields()[b.region].customName, 'Area');
        b.pivot.updateField(b.region, { customName: null });
        assert.equal(b.pivot.listFields()[b.region].customName, null);
      },
    },
  ],
  [
    'Pivot.updateField (showAll)',
    {
      act: (b) => b.pivot.updateField(b.product, { showAll: true }),
      expect: (b) => {
        assert.equal(b.pivot.listFields()[b.product].showAll, true);
        b.pivot.updateField(b.product, { showAll: false });
        assert.equal(b.pivot.listFields()[b.product].showAll, false);
      },
    },
  ],
  [
    'Pivot.project',
    {
      act: (b) => b.pivot.project(),
      expect: (b) => {
        assert.equal(b.pivot.isProjectionPending(), false);
        assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
      },
    },
  ],
];

test('the pivot mutator table covers every declared pivot mutator', async () => {
  const dts = await readFile(path.join(pkgRoot, 'dist', 'index.d.ts'), 'utf8');
  const members = (cls) => {
    const body = dts.match(new RegExp(`^export interface ${cls} \\{\\n([\\s\\S]*?)\\n\\}`, 'm'));
    assert.ok(body, `${cls} is declared`);
    return [...body[1].matchAll(/^ {2}(\w+)\(/gm)].map((m) => `${cls}.${m[1]}`);
  };
  const owners = ['Workbook.addPivotCache', 'Workbook.removePivotCache', 'Sheet.addPivot', 'Sheet.removePivot'];
  for (const name of owners) {
    const [cls, member] = name.split('.');
    assert.ok(members(cls).includes(name), `${member} is declared on ${cls}`);
  }
  const declaredAll = [...new Set([...members('Pivot'), ...members('PivotCache'), ...owners])];
  const declared = declaredAll.filter((name) => !/^\w+\.(?:get|is|list)/.test(name));
  assert.ok(declared.length > 40, `expected the pivot mutator family, found ${declared.length}`);
  const covered = PIVOT_MUTATORS.map(([name]) => name.replace(/ \(.*\)$/, ''));
  assert.deepEqual(
    declared.filter((name) => !covered.includes(name)),
    [],
    'pivot mutators without a row',
  );
  assert.deepEqual(
    covered.filter((name) => !declaredAll.includes(name)),
    [],
    'rows for methods that are not declared pivot mutators',
  );
});

test('pivot mutators forward their arguments to the C ABI', async () => {
  const mod = await getModule();
  const base = buildPivot(mod);
  try {
    assert.deepEqual(pivotGrid(base.wb), BASE_GRID);
  } finally {
    base.wb.dispose();
  }
  for (const [name, row] of PIVOT_MUTATORS) {
    const b = buildPivot(mod);
    try {
      row.setup?.(b);
      const r = row.act(b);
      row.after?.(b);
      row.expect(b, r);
    } catch (err) {
      err.message = `${name}: ${err.message}`;
      throw err;
    } finally {
      b.wb.dispose();
    }
  }
});

test('manual pivot sort order keeps item order and rejects bad arguments', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    b.pivot.clearFieldItems(b.region);
    b.pivot.addFieldItem(b.region, 'West');
    b.pivot.addFieldItem(b.region, 'East');
    b.pivot.updateField(b.region, { sortOrder: mod.PivotSortOrder.Manual });
    assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'West|37', 'East|15', '総計|52']);
    assert.throws(() => b.pivot.updateField(b.region, { sortOrder: 3 }), isFormulonError);
    assert.throws(
      () => b.pivot.updateField(b.region, { sortOrder: mod.PivotSortOrder.Manual, sortByField: 'Sum of Amount' }),
      isFormulonError,
    );
    b.pivot.updateField(b.region, { sortOrder: mod.PivotSortOrder.Ascending });
    assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
  } finally {
    b.wb.dispose();
  }
});

test('pivot filters and subtotal functions survive save and reload', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    b.pivot.addFilter(GREATER_THAN_20);
    b.pivot.addFieldSubtotalFunction(b.region, 2);
    b.pivot.addFieldSubtotalFunction(b.region, 0);
    const region = b.pivot.listFields()[b.region];
    assert.equal(region.defaultSubtotal, false);
    assert.deepEqual(region.subtotalFunctions, [0, 2]);
    assert.equal(b.pivot.listFields()[b.amount].defaultSubtotal, true);
    assert.deepEqual(b.pivot.listFields()[b.amount].subtotalFunctions, []);

    const saved = b.wb.save();
    assert.match(
      packagePart(saved, 'xl/pivotTables/pivotTable1.xml'),
      /<filters count="1"><filter fld="0" type="valueGreaterThan" evalOrder="-1" id="1" iMeasureFld="0">/,
    );
    const back = mod.Workbook.load(saved);
    try {
      back.setProfileId('win-365-ja_JP');
      assert.equal(back.sheet(0).pivot(0).listFilters().length, 1);
      const filter = back.sheet(0).pivot(0).listFilters()[0];
      assert.equal(filter.type, 1);
      assert.equal(filter.fieldName, 'Region');
      assert.equal(filter.value, 20);
      assert.deepEqual(pivotGrid(back), WEST_ONLY);
      const field = back.sheet(0).pivot(0).listFields()[0];
      assert.equal(field.defaultSubtotal, false);
      assert.deepEqual(field.subtotalFunctions, [0, 2]);
      back.sheet(0).pivot(0).clearFilters();
      assert.doesNotMatch(packagePart(back.save(), 'xl/pivotTables/pivotTable1.xml'), /<filters/);
    } finally {
      back.dispose();
    }
  } finally {
    b.wb.dispose();
  }
});

// ---- Failures on the fallible accessors --------------------------------
//
// Every accessor backed by a fallible C ABI call throws, so a rejected
// argument or a released handle cannot read as a legitimate zero count,
// empty name, or unspilled cell. Each row is
// [method, args that succeed, args the engine rejects (or null when only
// a released handle can fail it)].
const STATUS_ACCESSORS = [
  ['listSheets', [], null],
  ['listCells', [0], [99]],
  ['getCellXfIndex', [0, 0, 0], [99, 0, 0]],
  [
    'recalcViewport',
    [0, { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 }],
    [99, { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 }],
  ],
  ['listCfRules', [0], [99]],
  ['listValidations', [0], [99]],
  ['listProtectedRanges', [0], [99]],
  ['listSparklineGroups', [0], [99]],
  ['getCellPhoneticRuns', [0, 0, 0], [99, 0, 0]],
  ['listExternalLinks', [], null],
  ['listDefinedNames', [], null],
  ['listTables', [0], [99]],
  ['listPassthroughs', [], null],
  ['listPivots', [0], [99]],
  ['listPivotCaches', [], null],
  ['PivotCache.listFields', null, [9999]],
  ['PivotCache.listFieldSharedItems', null, [9999, 0]],
  ['PivotCache.getRecordCount', null, [9999]],
  ['Pivot.listFields', null, [0, 99]],
  ['Pivot.listFieldItems', null, [0, 99, 0]],
  ['Pivot.getProperties', null, [0, 99]],
  ['Pivot.getReportLayout', null, [0, 99]],
  ['Pivot.listDataFields', null, [0, 99]],
  ['Pivot.listFilters', null, [0, 99]],
  ['PivotCache.listCalculatedItems', null, [9999]],
  ['listFonts', [], null],
  ['listFills', [], null],
  ['listBorders', [], null],
  ['listCellXfs', [], null],
  ['listDxfs', [], null],
  ['listCellStyles', [], null],
  ['listTableStyles', [], null],
  ['listCellStyleXfs', [], null],
  ['getCalcMode', [], null],
  ['getFullPrecision', [], null],
  ['getCalcOnSave', [], null],
  ['getCommentVisible', null, [99, 0, 0]],
  ['getRefMode', [], null],
  ['getDate1904', [], null],
  ['getProfileId', [], null],
  ['getPinnedNow', [], null],
  ['getSpill', [0, 0, 0], [99, 0, 0]],
  ['getArrayFormula', [0, 0, 0], [99, 0, 0]],
  ['getColumnWidthPt', [0, 0, 0], [99, 0, 0]],
  ['getRowHeightPt', [0, 0], [99, 0]],
  ['convertColumnCharsToPt', [0, 8, 0], [99, 8, 0]],
  ['convertColumnPtToChars', [0, 10, 0], [99, 10, 0]],
  ['getPrecedents', [0, 0, 0, 1], [99, 0, 0, 1]],
  ['getDependents', [0, 0, 0, 1], [99, 0, 0, 1]],
];

// Calls a STATUS_ACCESSORS row on the workbook, or on the sheet its first argument names.
// A `Pivot.x` row takes [sheet, pivot, ...rest], a `PivotCache.x` row [cacheId, ...rest].
function callAccessor(wb, name, args) {
  if (name.startsWith('Pivot.'))
    return wb
      .sheet(args[0])
      .pivot(args[1])
      [name.slice(6)](...args.slice(2));
  if (name.startsWith('PivotCache.')) return wb.pivotCache(args[0])[name.slice(11)](...args.slice(1));
  if (name in wb) return wb[name](...args);
  if (name in wb.ooxml) return wb.ooxml[name](...args);
  const sheet = wb.sheet(args[0]);
  return name in sheet ? sheet[name](...args.slice(1)) : sheet.ooxml[name](...args.slice(1));
}

test('every row of the accessor table names a declared method', async () => {
  const dts = await readFile(path.join(pkgRoot, 'dist', 'index.d.ts'), 'utf8');
  const declared = new Set();
  for (const [, cls, body] of dts.matchAll(/^export interface (\w+) \{\n([\s\S]*?)\n\}/gm)) {
    for (const [, member] of body.matchAll(/^ {2}([A-Za-z]\w*)\(/gm)) {
      declared.add(member);
      declared.add(`${cls}.${member}`);
    }
  }
  assert.deepEqual(
    STATUS_ACCESSORS.map(([name]) => name).filter((name) => !declared.has(name)),
    [],
    'STATUS_ACCESSORS rows for methods that are not declared',
  );
});

test('fallible accessors succeed, throw on a rejection and on a released handle', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  for (const [name, okArgs, badArgs] of STATUS_ACCESSORS) {
    if (okArgs !== null) {
      assert.doesNotThrow(() => callAccessor(wb, name, okArgs), name);
    }
    if (badArgs !== null) {
      assert.throws(
        () => callAccessor(wb, name, badArgs),
        isFormulonError,
        `${name} must report the rejected arguments`,
      );
    }
  }
  assert.equal(wb.listSheets().length, 1);
  assert.equal(wb.getProfileId(), 'win-365-en_US');
  assert.equal(mod.localizeFunctionName('SUM', 'win-365-ja_JP'), 'SUM');
  assert.throws(() => mod.localizeFunctionName('SUM', 'nope'), isFormulonError);
  assert.equal(mod.listFunctionNames().includes('SUM'), true);
  assert.throws(() => mod.listLocaleErrorNames('nope'), isFormulonError);
  assert.throws(() => mod.compareText('a', 'b', 0, 'nope'), isFormulonError);

  wb.dispose();
  const dead = wb;
  // Every accessor, including the ones no argument can fail, reports the
  // released handle instead of a default reading.
  for (const [name, okArgs, badArgs] of STATUS_ACCESSORS) {
    assert.throws(
      () => callAccessor(dead, name, okArgs ?? badArgs),
      (e) => isFormulonError(e, 7000),
      `${name} after release`,
    );
  }
});

test('list getters return arrays and throw on an unreadable sheet', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    // A readable sheet: every list is empty but the enumeration succeeded,
    // which is what distinguishes it from a sheet that could not be read.
    for (const [name, call] of [
      ['listMerges', () => wb.sheet(0).listMerges()],
      ['listComments', () => wb.sheet(0).listComments()],
      ['listHyperlinks', () => wb.sheet(0).listHyperlinks()],
      ['listDrawings', () => wb.sheet(0).listDrawings()],
    ]) {
      const list = call();
      assert.ok(Array.isArray(list), `${name} must stay an array`);
      assert.equal(list.length, 0, name);
    }

    // An unreadable sheet is reported, not silently rendered as "empty".
    for (const [name, call] of [
      ['listMerges', () => wb.sheet(99).listMerges()],
      ['listComments', () => wb.sheet(99).listComments()],
      ['listHyperlinks', () => wb.sheet(99).listHyperlinks()],
      ['listDrawings', () => wb.sheet(99).listDrawings()],
    ]) {
      assert.throws(call, isFormulonError, `${name} must signal an unreadable sheet`);
    }
  } finally {
    wb.dispose();
  }
});

test('setPivotFieldDateGroup defaults its optional group fields', async () => {
  const mod = await getModule();
  const explicit = buildPivot(mod);
  const defaulted = buildPivot(mod);
  try {
    dateRows(explicit);
    dateRows(defaulted);
    const days = 8; // PivotDatePeriod.Days
    explicit.wb.sheet(0).pivot(explicit.pivot).setFieldDateGroup(explicit.date, {
      periods: days,
      calendar: 0,
      intervalDays: 1,
    });
    defaulted.wb.sheet(0).pivot(defaulted.pivot).setFieldDateGroup(defaulted.date, { periods: days });
    const field = defaulted.wb.sheet(0).pivot(defaulted.pivot).listFields()[defaulted.date];
    assert.equal(field.datePeriods, days);
    assert.deepEqual(field, explicit.wb.sheet(0).pivot(explicit.pivot).listFields()[explicit.date]);
    assert.deepEqual(pivotGrid(defaulted.wb), pivotGrid(explicit.wb));
  } finally {
    explicit.wb.dispose();
    defaulted.wb.dispose();
  }
});

test('pivot cache date shared items report as dates and reject dates the engine cannot store', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    b.wb.pivotCache(b.cacheId).addFieldSharedItem(3, new Date(2024, 0, 1));
    b.wb.pivotCache(b.cacheId).addFieldSharedItem(3, new Date(2024, 1, 1, 12));
    b.wb.pivotCache(b.cacheId).addFieldSharedItem(3, 7);
    assert.deepEqual(
      b.wb
        .pivotCache(b.cacheId)
        .listFieldSharedItems(3)
        .map((item) => item.isDate),
      [true, true, false],
    );
    const before = new Date(1899, 11, 30);
    assert.throws(() => b.wb.pivotCache(b.cacheId).addFieldSharedItem(3, before), isFormulonError);
    assert.throws(() => b.wb.pivotCache(b.cacheId).setRecordValue(0, 3, before), isFormulonError);
    const invalid = new Date(Number.NaN);
    assert.throws(() => b.wb.pivotCache(b.cacheId).addFieldSharedItem(3, invalid), RangeError);
    assert.throws(() => b.wb.pivotCache(b.cacheId).setRecordValue(0, 3, invalid), RangeError);
    // The date option is gone: a third argument is refused.
    assert.throws(() => b.wb.pivotCache(b.cacheId).addFieldSharedItem(3, 'x', { date: true }), TypeError);
    assert.throws(() => b.wb.pivotCache(b.cacheId).setRecordValue(0, 3, true, { date: true }), TypeError);
    assert.equal(b.wb.pivotCache(b.cacheId).listFieldSharedItems(3).length, 3);
  } finally {
    b.wb.dispose();
  }
});

test('a pivot cache shared item list reads back in order', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const items = b.wb.pivotCache(b.cacheId).listFieldSharedItems(0);
    assert.deepEqual(
      items.map((item) => [item.value, item.isDate]),
      [
        ['East', false],
        ['West', false],
      ],
    );
    assert.throws(() => b.wb.pivotCache(b.cacheId).listFieldSharedItems(99), isFormulonError);
  } finally {
    b.wb.dispose();
  }
});

test('a pivot cache error value without errorCode throws and changes nothing', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const items = b.wb.pivotCache(b.cacheId).listFieldSharedItems(0).length;
    const partial = { kind: mod.ValueKind.Error };
    assert.throws(() => b.wb.pivotCache(b.cacheId).addFieldSharedItem(0, partial), TypeError);
    assert.equal(b.wb.pivotCache(b.cacheId).listFieldSharedItems(0).length, items);
    assert.throws(() => b.wb.pivotCache(b.cacheId).setRecordValue(0, 2, partial), TypeError);
    assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
  } finally {
    b.wb.dispose();
  }
});

test('pivot report options and the Values field order report bad arguments', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    assert.equal(mod.PIVOT_VALUES_FIELD, -2);
    // A single data field has no Values level to place.
    assert.throws(() => b.pivot.setRowFieldOrder([b.region, mod.PIVOT_VALUES_FIELD]), isFormulonError);
    assert.throws(() => b.pivot.updateField(99, { defaultSubtotal: true }), isFormulonError);
    assert.throws(() => b.pivot.updateField(99, { insertBlankRow: true }), isFormulonError);
    assert.throws(() => b.pivot.updateField(99, { fillDownLabels: true }), isFormulonError);
    assert.throws(() => b.pivot.updateField(99, { showDetail: true }), isFormulonError);
    assert.throws(() => b.pivot.updateFieldItem(b.region, 99, { showDetail: true }), isFormulonError);
    assert.throws(() => b.pivot.updateFieldItem(b.region, 99, { caption: 'x' }), isFormulonError);
    // Captions collide with another item's label.
    assert.throws(() => b.pivot.updateFieldItem(b.region, 0, { caption: 'West' }), isFormulonError);
    assert.throws(() => b.wb.sheet(99).pivot(0).showDetails(0, 0), isFormulonError);
    assert.throws(() => b.wb.sheet(0).chart(9999).rename('x'), isFormulonError);
    assert.throws(() => b.wb.sheet(0).chart(9999).rename(''), isFormulonError);
  } finally {
    b.wb.dispose();
  }
});

test('pivot options read back every member and keep what a partial write omits', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const before = b.pivot.getOptions();
    assert.equal(before.errorCaption, null);
    assert.equal(typeof before.indent, 'number');
    const patch = {
      showError: true,
      errorCaption: '#n/a',
      showMissing: false,
      missingCaption: 'none',
      gridDropZones: true,
      mergeItem: true,
      pageOverThenDown: true,
      altText: 'Sales',
      altTextSummary: 'line one\nline two',
      grandTotalCaption: 'Everything',
      autoFormatting: false,
      enableDrill: false,
    };
    b.pivot.updateOptions(patch);
    const after = b.pivot.getOptions();
    for (const [key, value] of Object.entries(patch)) assert.equal(after[key], value, key);
    assert.equal(after.showHeaders, before.showHeaders);
    assert.equal(after.indent, before.indent);
    // A write of the whole read-back is a no-op, and null or '' clears a string.
    b.pivot.updateOptions({ ...after, errorCaption: null });
    assert.equal(b.pivot.getOptions().errorCaption, null);
    assert.equal(b.pivot.getOptions().altText, 'Sales');
    b.pivot.updateOptions({ grandTotalCaption: '' });
    assert.equal(b.pivot.getOptions().grandTotalCaption, null);
  } finally {
    b.wb.dispose();
  }
});

test('pivot options reject out-of-range values and wrong types without changing anything', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const base = b.pivot.getOptions();
    const unchanged = () => assert.deepEqual(b.pivot.getOptions(), base);
    for (const bad of [{ indent: 127 }, { pageWrap: 256 }, { errorCaption: 'x'.repeat(256) }]) {
      assert.throws(() => b.pivot.updateOptions(bad), isFormulonError, JSON.stringify(bad).slice(0, 40));
      unchanged();
    }
    for (const bad of [{ indent: 'a' }, { errorCaption: 5 }]) {
      assert.throws(() => b.pivot.updateOptions(bad), TypeError, JSON.stringify(bad));
      unchanged();
    }
    for (const bad of [{ indent: 1.5 }, { indent: -1 }, { pageWrap: Number.NaN }]) {
      assert.throws(() => b.pivot.updateOptions(bad), RangeError, JSON.stringify(bad));
      unchanged();
    }
    assert.throws(() => b.pivot.updateOptions('x'), TypeError);
    assert.throws(() => b.pivot.updateOptions(), TypeError);
    // A string is a pivot name, even one that reads as a number.
    assert.throws(
      () => b.wb.sheet(0).pivot(String(b.pivot)).updateOptions({}),
      (e) => isFormulonError(e, mod.StatusCode.NotFound),
    );
    assert.throws(() => b.wb.sheet(0).pivot(true).updateOptions({}), TypeError);
    assert.throws(() =>
      b.wb
        .sheet(2 ** 32)
        .pivot(b.pivot)
        .updateOptions({}),
    );
    unchanged();
    assert.throws(() => b.wb.sheet(99).pivot(0).updateOptions({}), isFormulonError);
    assert.throws(() => b.wb.sheet(99).pivot(0).getOptions(), isFormulonError);

    const cacheBase = b.wb.pivotCache(b.cacheId).getOptions();
    assert.throws(() => b.wb.pivotCache(b.cacheId).updateOptions({ missingItemsLimit: 1048577 }), isFormulonError);
    assert.throws(() => b.wb.pivotCache(b.cacheId).updateOptions({ missingItemsLimit: 'a' }), TypeError);
    for (const bad of [{ missingItemsLimit: 1.5 }, { missingItemsLimit: Number.NaN }, { missingItemsLimit: -2 }]) {
      assert.throws(() => b.wb.pivotCache(b.cacheId).updateOptions(bad), RangeError, JSON.stringify(bad));
    }
    assert.throws(() => b.wb.pivotCache(b.cacheId).updateOptions(null), TypeError);
    assert.throws(() => b.wb.pivotCache('1').updateOptions({}), TypeError);
    assert.deepEqual(b.wb.pivotCache(b.cacheId).getOptions(), cacheBase);
    assert.throws(() => b.wb.pivotCache(9999).updateOptions({}), isFormulonError);
    assert.throws(() => b.wb.pivotCache(9999).getOptions(), isFormulonError);
  } finally {
    b.wb.dispose();
  }
});

test('calculated fields and items reject bad arguments and change nothing', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const fields = b.wb.pivotCache(b.cacheId).listFields().length;
    // Wrong types throw before anything runs.
    assert.throws(() => b.wb.pivotCache(b.cacheId).addCalculatedField('D', 5), TypeError);
    assert.throws(() => b.wb.pivotCache(b.cacheId).addCalculatedField(null, 'Amount'), TypeError);
    assert.throws(() => b.wb.pivotCache(b.cacheId).addCalculatedField('D'), TypeError);
    assert.throws(() => b.wb.pivotCache('1').addCalculatedField('D', 'Amount'), TypeError);
    assert.throws(() => b.wb.pivotCache(b.cacheId).addCalculatedItem(0.5, 'N', 'East'), RangeError);
    assert.throws(() => b.wb.pivotCache(b.cacheId).addCalculatedItem(b.region, 'N', undefined), TypeError);
    assert.throws(() => b.wb.pivotCache(b.cacheId).moveCalculatedItem(0, -1), RangeError);
    assert.throws(() => b.wb.pivotCache('a').listCalculatedItems(), TypeError);
    // Rejected by the engine: empty or reserved names, bad formulas.
    for (const [name, formula] of [
      ['', 'Amount'],
      ['All', 'Amount'],
      ['A[1]', 'Amount'],
      ['Amount', 'Amount'],
      ['D', ''],
      ['D', 'Amount+'],
      ['D', 'Nope*2'],
      ['D', 'A1+Amount'],
      ['D', 'x'.repeat(8193)],
    ]) {
      assert.throws(
        () => b.wb.pivotCache(b.cacheId).addCalculatedField(name, formula),
        isFormulonError,
        `${name}=${formula.slice(0, 8)}`,
      );
    }
    assert.throws(() => b.wb.pivotCache(9999).addCalculatedField('D', 'Amount'), isFormulonError);
    assert.equal(b.wb.pivotCache(b.cacheId).listFields().length, fields);
    // Not a calculated field.
    assert.throws(() => b.wb.pivotCache(b.cacheId).setCalculatedFieldFormula(0, 'Amount'), isFormulonError);
    assert.throws(() => b.wb.pivotCache(b.cacheId).removeCalculatedField(0), isFormulonError);
    assert.throws(() => b.wb.pivotCache(b.cacheId).removeCalculatedField(2 ** 31), isFormulonError);
    assert.throws(() => b.wb.pivotCache(9999).listFields(), isFormulonError);
    // A calculated field a second one names cannot be removed.
    b.wb.pivotCache(b.cacheId).addCalculatedField('Base', 'Amount');
    b.wb.pivotCache(b.cacheId).addCalculatedField('Top', 'Base*2');
    assert.throws(() => b.wb.pivotCache(b.cacheId).removeCalculatedField(fields), isFormulonError);
    assert.equal(b.wb.pivotCache(b.cacheId).listFields().length, fields + 2);

    // Items: unplaced fields, duplicates, and out-of-range indices.
    assert.throws(() => b.wb.pivotCache(b.cacheId).addCalculatedItem(b.date, 'N', 'Amount'), isFormulonError);
    assert.throws(() => b.wb.pivotCache(b.cacheId).addCalculatedItem(99, 'N', 'East'), isFormulonError);
    assert.throws(() => b.wb.pivotCache(b.cacheId).addCalculatedItem(b.region, 'East', 'West'), isFormulonError);
    assert.throws(() => b.wb.pivotCache(b.cacheId).addCalculatedItem(b.region, 'N', 'Nope'), isFormulonError);
    assert.throws(() => b.wb.pivotCache(9999).addCalculatedItem(b.region, 'N', 'East'), isFormulonError);
    assert.equal(b.wb.pivotCache(b.cacheId).listCalculatedItems().length, 0);
    assert.throws(() => b.wb.pivotCache(9999).listCalculatedItems(), isFormulonError);
    assert.throws(() => b.wb.pivotCache(b.cacheId).setCalculatedItemFormula(0, 'East'), isFormulonError);
    assert.throws(() => b.wb.pivotCache(b.cacheId).moveCalculatedItem(0, 0), isFormulonError);
    assert.throws(() => b.wb.pivotCache(b.cacheId).removeCalculatedItem(0), isFormulonError);
    assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
  } finally {
    b.wb.dispose();
  }
});

test('pivot field custom names, show-all and default data field names report bad arguments', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    assert.throws(() => b.pivot.updateField(b.region, { customName: 5 }), TypeError);
    assert.throws(() => b.pivot.updateField(b.region), TypeError);
    assert.throws(() => b.pivot.updateField(-1, { customName: 'x' }), RangeError);
    assert.throws(() => b.pivot.updateField(b.region, { showAll: null }), TypeError);
    assert.throws(() => b.pivot.updateField(b.region, { bogus: 0 }), TypeError);
    assert.throws(() => b.pivot.updateField(99, { customName: 'x' }), isFormulonError);
    assert.throws(() => b.pivot.updateField(b.region, { customName: 'x'.repeat(256) }), isFormulonError);
    // Another field's source name or display name is taken.
    assert.throws(() => b.pivot.updateField(b.region, { customName: 'Product' }), isFormulonError);
    assert.throws(() => b.pivot.updateField(b.region, { customName: 'Sum of Amount' }), isFormulonError);
    // The field's own source name is accepted.
    b.pivot.updateField(b.region, { customName: 'Region' });
    assert.throws(() => b.pivot.updateField(99, { showAll: true }), isFormulonError);

    // The first derived name is free; once taken, a new one is numbered
    // unless the entry that holds it is the one being replaced.
    assert.equal(b.pivot.getDataFieldDefaultName(b.amount, 0), '合計 / Amount');
    b.pivot.addDataField({ fieldIndex: b.amount, aggregation: 0 });
    assert.equal(b.pivot.listDataFields()[1].name, '合計 / Amount');
    assert.equal(b.pivot.getDataFieldDefaultName(b.amount, 0), '合計 / Amount2');
    assert.equal(b.pivot.getDataFieldDefaultName(b.amount, 0, undefined), '合計 / Amount2');
    assert.equal(b.pivot.getDataFieldDefaultName(b.amount, 0, 1), '合計 / Amount');
    assert.throws(() => b.pivot.getDataFieldDefaultName(b.amount, 0, 99), isFormulonError);
    assert.equal(b.pivot.listDataFields().length, 2);
    assert.throws(() => b.pivot.getDataFieldDefaultName(99, 0), isFormulonError);
    assert.throws(() => b.pivot.getDataFieldDefaultName(b.amount, 99), isFormulonError);
    assert.throws(() => b.wb.sheet(99).pivot(0).getDataFieldDefaultName(0, 0), isFormulonError);
    assert.throws(() => b.pivot.getDataFieldDefaultName(b.amount, 'sum'), TypeError);
    assert.throws(() => b.pivot.getDataFieldDefaultName(b.amount, 0, null), TypeError);
    assert.throws(() => b.pivot.getDataFieldDefaultName(b.amount), TypeError);
    assert.equal(b.pivot.listDataFields()[0].name, 'Sum of Amount');
    assert.throws(() => b.wb.sheet(99).pivot(0).project(), isFormulonError);
    assert.throws(() => b.wb.sheet(99).pivot(0).isProjectionPending(), isFormulonError);
    assert.throws(() => b.wb.sheet(0).projectPivot(), TypeError);
  } finally {
    b.wb.dispose();
  }
});

test('label-between filters carry valueHigh', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const spec = {
      axis: 0,
      fieldName: 'Region',
      type: mod.PivotFilterKind.LabelBetween,
      value: 'A',
    };
    // The upper bound is required for the between types.
    assert.throws(() => b.pivot.addFilter(spec), isFormulonError);
    b.pivot.addFilter({ ...spec, valueHigh: 'F' });
    const at = b.pivot.listFilters()[0];
    assert.equal(at.type, 22);
    assert.equal(at.value, 'A');
    assert.equal(at.valueHigh, 'F');
    // And rejected for every other type.
    assert.throws(
      () => b.pivot.addFilter({ ...spec, type: mod.PivotFilterKind.LabelEqual, valueHigh: 'F' }),
      isFormulonError,
    );
    assert.throws(() => b.pivot.addFilter({ ...spec, valueHigh: 5 }), isFormulonError);
    assert.equal(b.pivot.listFilters().length, 1);
  } finally {
    b.wb.dispose();
  }
});
