import assert from 'node:assert/strict';
import { StatusCode } from '../../packages/npm/common.mjs';
import { appendEmptyZipEntry, expectFailure, readAll, readZipEntryText, VAL } from './smoke_support.mjs';

const nestedErrorClass = (value) => (typeof value === 'symbol' ? TypeError : RangeError);

/** A literal `#AARRGGBB` colour spec as the raw layer reads it back. */
const rgbColor = (argb) => ({ kind: 1, argb: argb, theme: 0, tint: 0, indexed: 0 });

export function registerCellsIo(Module, test) {
  test('setText / setBool / setBlank round-trip', () => {
    const wb = Module.Workbook.createDefault();
    try {
      wb.setText(0, 0, 0, 'hello');
      wb.setBool(0, 1, 0, true);
      wb.setBlank(0, 2, 0);
      wb.recalc();

      const a1 = wb.getValue(0, 0, 0);
      assert.equal(a1.kind, VAL.TEXT);
      assert.equal(a1.text, 'hello');

      const a2 = wb.getValue(0, 1, 0);
      assert.equal(a2.kind, VAL.BOOL);
      assert.equal(a2.boolean, true);

      const a3 = wb.getValue(0, 2, 0);
      assert.equal(a3.kind, VAL.BLANK);
    } finally {
      wb.delete();
    }
  });

  test('setCellPhonetic / getCellPhonetic preserve and clear a Japanese guide', () => {
    const wb = Module.Workbook.createDefault();
    try {
      wb.setText(0, 0, 0, '漢字');
      wb.setCellPhonetic(0, 0, 0, 'かんじ');
      const phonetic = wb.getCellPhonetic(0, 0, 0);
      assert.equal(phonetic, 'かんじ');

      wb.setCellPhonetic(0, 0, 0, '');
      assert.equal(wb.getCellPhonetic(0, 0, 0), null);

      wb.setCellPhonetic(0, 0, 0, 'かんじ');
      wb.setText(0, 0, 0, '文字列');
      assert.equal(wb.getCellPhonetic(0, 0, 0), null);
    } finally {
      wb.delete();
    }
  });

  test('setCellPhoneticRuns / getCellPhoneticRun keep the spans apart', () => {
    const wb = Module.Workbook.createDefault();
    try {
      wb.setText(0, 0, 0, '東京都');
      assert.deepEqual(
        readAll(
          () => wb.getCellPhoneticRunCount(0, 0, 0),
          (i) => wb.getCellPhoneticRun(0, 0, 0, i),
        ),
        [],
      );

      const runs = [
        { start: 0, end: 2, text: 'トウキョウ' },
        { start: 2, end: 3, text: 'ト' },
      ];
      wb.setCellPhoneticRuns(0, 0, 0, runs);
      assert.deepEqual(
        readAll(
          () => wb.getCellPhoneticRunCount(0, 0, 0),
          (i) => wb.getCellPhoneticRun(0, 0, 0, i),
        ),
        runs,
      );
      // The flattening getter still reports the concatenation.
      assert.equal(wb.getCellPhonetic(0, 0, 0), 'トウキョウト');

      // Writing that back through the whole-cell setter is the collapse the
      // run API exists to avoid.
      wb.setCellPhonetic(0, 0, 0, 'トウキョウト');
      assert.deepEqual(
        readAll(
          () => wb.getCellPhoneticRunCount(0, 0, 0),
          (i) => wb.getCellPhoneticRun(0, 0, 0, i),
        ),
        [{ start: 0, end: 3, text: 'トウキョウト' }],
      );

      wb.setCellPhoneticRuns(0, 0, 0, []);
      assert.deepEqual(
        readAll(
          () => wb.getCellPhoneticRunCount(0, 0, 0),
          (i) => wb.getCellPhoneticRun(0, 0, 0, i),
        ),
        [],
      );

      expectFailure(() =>
        wb.setCellPhoneticRuns(0, 0, 0, [
          { start: 2, end: 3, text: 'ト' },
          { start: 0, end: 2, text: 'トウ' },
        ]),
      );
    } finally {
      wb.delete();
    }
  });

  test('setCellPhoneticProperties survives a run edit and a save/load cycle', () => {
    // The result carries a `status` alongside the triple; compare the
    // triple alone so a status-shape change does not read as a value change.
    const phoneticProps = (book, sheet, row, col) => {
      const { fontIndex, type, alignment } = book.getCellPhoneticProperties(sheet, row, col);
      return { fontIndex, type, alignment };
    };
    const wb = Module.Workbook.createDefault();
    try {
      wb.setText(0, 0, 0, '大阪');
      // An unannotated cell reports what Excel infers for a guide written
      // with no <phoneticPr> at all.
      assert.deepEqual(phoneticProps(wb, 0, 0, 0), { fontIndex: 0, type: 0, alignment: 0 });

      wb.setCellPhoneticProperties(0, 0, 0, { fontIndex: 3, type: 2, alignment: 2 });
      // Setting the readings must not reset the rendering.
      wb.setCellPhoneticRuns(0, 0, 0, [{ start: 0, end: 2, text: 'おおさか' }]);
      assert.deepEqual(phoneticProps(wb, 0, 0, 0), { fontIndex: 3, type: 2, alignment: 2 });

      const saved = wb.save();
      const loaded = Module.Workbook.load(saved);
      try {
        assert.deepEqual(phoneticProps(loaded, 0, 0, 0), { fontIndex: 3, type: 2, alignment: 2 });
      } finally {
        loaded.delete();
      }

      // Two bits each on the binary side, so a wider ordinal is refused
      // rather than truncated into a different, valid-looking one.
      expectFailure(() => wb.setCellPhoneticProperties(0, 0, 0, { fontIndex: 0, type: 4, alignment: 0 }));
      assert.deepEqual(phoneticProps(wb, 0, 0, 0), { fontIndex: 3, type: 2, alignment: 2 });

      // A value write discards the guide and its rendering together.
      wb.setText(0, 0, 0, '京都');
      assert.deepEqual(phoneticProps(wb, 0, 0, 0), { fontIndex: 0, type: 0, alignment: 0 });
    } finally {
      wb.delete();
    }
  });

  test('cell nested integer fields reject narrowing before C mutations', () => {
    const invalid = [2 ** 32, 1.5, Number.NaN, Number.POSITIVE_INFINITY, Symbol('cell integer')];
    for (const value of invalid) {
      const wb = Module.Workbook.createDefault();
      try {
        wb.setText(0, 0, 0, '東京都');
        const runs = [
          { start: 0, end: 2, text: 'トウキョウ' },
          { start: 2, end: 3, text: 'ト' },
        ];
        wb.setCellPhoneticRuns(0, 0, 0, runs);
        const beforeRuns = readAll(
          () => wb.getCellPhoneticRunCount(0, 0, 0),
          (i) => wb.getCellPhoneticRun(0, 0, 0, i),
        );
        assert.throws(
          () => wb.setCellPhoneticRuns(0, 0, 0, [{ start: value, end: 2, text: '壊' }]),
          nestedErrorClass(value),
          `setCellPhoneticRuns start=${String(value)} accepted`,
        );
        assert.deepEqual(
          readAll(
            () => wb.getCellPhoneticRunCount(0, 0, 0),
            (i) => wb.getCellPhoneticRun(0, 0, 0, i),
          ),
          beforeRuns,
        );

        wb.setCellPhoneticProperties(0, 0, 0, { fontIndex: 3, type: 2, alignment: 2 });
        const beforeProps = wb.getCellPhoneticProperties(0, 0, 0);
        assert.throws(
          () => wb.setCellPhoneticProperties(0, 0, 0, { fontIndex: value, type: 2, alignment: 2 }),
          nestedErrorClass(value),
          `setCellPhoneticProperties fontIndex=${String(value)} accepted`,
        );
        assert.deepEqual(
          {
            fontIndex: wb.getCellPhoneticProperties(0, 0, 0).fontIndex,
            type: wb.getCellPhoneticProperties(0, 0, 0).type,
            alignment: wb.getCellPhoneticProperties(0, 0, 0).alignment,
          },
          { fontIndex: beforeProps.fontIndex, type: beforeProps.type, alignment: beforeProps.alignment },
        );

        assert.throws(
          () => wb.formatValue({ kind: value, number: 1 }, 'General'),
          TypeError,
          `formatValue kind=${String(value)} accepted`,
        );
      } finally {
        wb.delete();
      }
    }
  });

  test('cell Value boolean and errorCode use checked signed integers', () => {
    const invalid = [2 ** 32, 1.5, Number.NaN, Number.POSITIVE_INFINITY, Symbol('cell value integer')];
    for (const value of invalid) {
      const wb = Module.Workbook.createDefault();
      try {
        assert.throws(
          () => wb.formatValue({ kind: 2, boolean: value }, 'General'),
          TypeError,
          `formatValue boolean=${String(value)} accepted`,
        );
        assert.throws(
          () => wb.formatValue({ kind: 4, errorCode: value }, 'General'),
          nestedErrorClass(value),
          `formatValue errorCode=${String(value)} accepted`,
        );
      } finally {
        wb.delete();
      }
    }
  });

  test('setDefaultFont replaces font 0, which addFont can only append beside', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.equal(wb.getFont(0).name, 'Calibri');
      const appended = wb.addFont({ name: '游ゴシック', sizePt: 11 });
      assert.ok(appended > 0);
      assert.equal(wb.getFont(0).name, 'Calibri');

      wb.setDefaultFont({ name: '游ゴシック', sizePt: 11, charset: 128 });
      assert.equal(wb.getFont(0).name, '游ゴシック');
      assert.equal(wb.getFont(0).charset, 128);

      const added = wb.addFont({ name: 'Meiryo', sizePt: 12 });
      wb.setFont(added, { name: 'MS Gothic', sizePt: 9 });
      assert.equal(wb.getFont(added).name, 'MS Gothic');
      expectFailure(() => wb.setFont(wb.getFontCount(), { name: 'MS Gothic', sizePt: 9 }));
    } finally {
      wb.delete();
    }
  });

  test('table update omits fields without resetting existing metadata', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const created = wb.addTable(0, {
        range: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1 },
        name: 'Sales',
        columns: ['Product', 'Amount'],
        styleName: 'TableStyleMedium2',
        headerRow: false,
        totalsRow: true,
      });
      const before = wb.getTable(0, created);
      assert.deepEqual(before.range.rect, { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1 });
      assert.equal(before.headerRow, false);
      assert.equal(before.totalsRow, true);
      assert.equal('headerRowCount' in before, false);
      assert.equal('styleName' in before, false);
      assert.equal('displayName' in before, false);

      const beforeSaved = wb.save();
      const beforeXml = readZipEntryText(beforeSaved, 'xl/tables/table1.xml');
      assert.match(beforeXml, /name="TableStyleMedium2"/);
      assert.match(beforeXml, /headerRowCount="0"/);
      assert.match(beforeXml, /totalsRowCount="1"/);

      // An empty update object is a partial no-op. This specifically covers
      // the WASM-to-C mapping of omitted rect/style/row flags to preservation
      // sentinels rather than empty/false defaults.
      wb.updateTable(0, created, {});
      const after = wb.getTable(0, created);
      assert.deepEqual(after.range.rect, before.range.rect);

      const saved = wb.save();
      const afterXml = readZipEntryText(saved, 'xl/tables/table1.xml');
      assert.match(afterXml, /name="TableStyleMedium2"/);
      assert.match(afterXml, /headerRowCount="0"/);
      assert.match(afterXml, /totalsRowCount="1"/);
      const loaded = Module.Workbook.load(saved);
      try {
        assert.equal(loaded.disposed, false);
        const reloaded = loaded.getTable(0, created);
        assert.deepEqual(reloaded.range.rect, { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1 });
        const roundTripped = loaded.save();
        const roundTrippedXml = readZipEntryText(roundTripped, 'xl/tables/table1.xml');
        assert.match(roundTrippedXml, /name="TableStyleMedium2"/);
        assert.match(roundTrippedXml, /headerRowCount="0"/);
        assert.match(roundTrippedXml, /totalsRowCount="1"/);
      } finally {
        loaded.delete();
      }
    } finally {
      wb.delete();
    }
  });

  test('table update rethrows nested getter failures before changing the table', () => {
    const getterFields = ['rect', 'headerRow', 'totalsRow'];
    for (const field of getterFields) {
      const wb = Module.Workbook.createDefault();
      try {
        const created = wb.addTable(0, {
          range: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1 },
          name: 'Sales',
          columns: ['Product', 'Amount'],
          styleName: 'TableStyleMedium2',
          headerRow: false,
          totalsRow: true,
        });
        const before = readZipEntryText(wb.save(), 'xl/tables/table1.xml');
        const spec = {};
        Object.defineProperty(spec, field, {
          enumerable: true,
          get() {
            throw new Error(`table update ${field} getter`);
          },
        });
        assert.throws(() => wb.updateTable(0, created, spec), new RegExp(`table update ${field} getter`));
        assert.equal(readZipEntryText(wb.save(), 'xl/tables/table1.xml'), before);
      } finally {
        wb.delete();
      }
    }
  });

  test('table update reads each optional field once before the C call', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const created = wb.addTable(0, {
        range: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1 },
        name: 'Sales',
        columns: ['Product', 'Amount'],
        styleName: 'TableStyleMedium2',
        headerRow: false,
        totalsRow: true,
      });
      const values = {
        rect: { firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 1 },
        headerRow: true,
        totalsRow: false,
      };
      const reads = Object.fromEntries(Object.keys(values).map((key) => [key, 0]));
      const spec = {};
      for (const [key, value] of Object.entries(values)) {
        Object.defineProperty(spec, key, {
          enumerable: true,
          get() {
            reads[key] += 1;
            if (reads[key] > 1) throw new Error(`table update ${key} getter read twice`);
            return value;
          },
        });
      }
      wb.updateTable(0, created, spec);
      assert.deepEqual(reads, { rect: 1, headerRow: 1, totalsRow: 1 });
      assert.deepEqual(wb.getTable(0, created).range.rect, { firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 1 });
    } finally {
      wb.delete();
    }
  });

  test('addTable rejects a narrowing range coordinate before the table is created', () => {
    const invalid = [2 ** 32, 1.5, Number.NaN, Number.POSITIVE_INFINITY, Symbol('table sheet')];
    for (const value of invalid) {
      const wb = Module.Workbook.createDefault();
      try {
        assert.throws(
          () =>
            wb.addTable(0, {
              range: { firstRow: value, firstCol: 0, lastRow: 2, lastCol: 1 },
              name: 'BadTable',
              columns: ['Product', 'Amount'],
            }),
          // A symbol is a type error; a fraction, NaN, infinity and 2^32 are range errors.
          nestedErrorClass(value),
          `addTable range.firstRow=${String(value)} accepted`,
        );
        assert.equal(wb.getTableCount(0), 0);
      } finally {
        wb.delete();
      }
    }
  });

  test('save() + Workbook.load() round-trip preserves a literal', () => {
    const wb = Module.Workbook.createDefault();
    let saved = null;
    try {
      wb.setNumber(0, 0, 0, 7);
      wb.setFormula(0, 1, 0, '=A1+1');
      wb.recalc();

      const r = wb.save();
      assert.ok(r instanceof Uint8Array);
      assert.ok(r.length > 0);
      saved = r;
    } finally {
      wb.delete();
    }

    const loaded = Module.Workbook.load(saved);
    try {
      assert.equal(loaded.disposed, false);
      assert.ok(loaded.getSheetCount() >= 1);
      loaded.recalc();
      const a1 = loaded.getValue(0, 0, 0);
      assert.equal(a1.kind, VAL.NUMBER);
      assert.equal(a1.number, 7);
    } finally {
      loaded.delete();
    }
  });

  test('a CF match carries the fields of its kind and a CF predicate evaluates to a boolean', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const dxf = wb.addDxf({ font: { bold: true } });
      wb.setNumber(0, 0, 0, 20);
      wb.addCfRule(0, {
        ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
        type: 1,
        operator: 6,
        formula1: '10',
        dxfIndex: dxf,
        stopIfTrue: false,
      });
      const cells = wb.evaluateCfRange(0, { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }, {});
      const match = cells[0].matches[0];
      assert.deepEqual({ ...match }, { kind: 0, priority: 1, dxfIndex: dxf });
      assert.strictEqual(wb.evaluateCfFormula(0, 0, 0, 0, 0, '=A1>10'), true);
      assert.strictEqual(wb.evaluateCfFormula(0, 0, 0, 0, 0, '=A1>100'), false);
    } finally {
      wb.delete();
    }
  });

  test('CF rules are keyed by the id addCfRule returns', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const cellA1 = [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }];
      const first = wb.addCfRule(0, { ranges: cellA1, type: 1, operator: 6, formula1: '10' });
      const second = wb.addCfRule(0, { ranges: cellA1, type: 1, operator: 6, formula1: '20' });
      assert.equal(typeof first, 'string');
      assert.notEqual(first, second);
      assert.equal(wb.getCfRule(0, 1).id, second);
      expectFailure(() => wb.removeCfRule(0, '{00000000-0000-0000-0000-000000000000}'), StatusCode.NotFound);
      wb.removeCfRule(0, first);
      assert.equal(wb.getCfRuleCount(0), 1);
      assert.equal(wb.getCfRule(0, 0).id, second);
    } finally {
      wb.delete();
    }
  });

  test('updateCfRule patches the named keys and refuses a key the rule type does not use', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const dxf = wb.addDxf({ font: { bold: true } });
      const other = wb.addDxf({ font: { italic: true } });
      const cellA1 = [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }];
      const ruleId = wb.addCfRule(0, { ranges: cellA1, type: 1, operator: 6, formula1: '10', dxfIndex: dxf });
      const barId = wb.addCfRule(0, {
        ranges: cellA1,
        type: 3,
        dataBar: { min: { type: 3 }, max: { type: 4 }, fill: rgbColor(0xff0000ff), gradient: false },
      });

      wb.updateCfRule(0, ruleId, { dxfIndex: other, formula1: '20', stopIfTrue: true });
      const changed = wb.getCfRule(0, 0);
      assert.equal(changed.dxfIndex, other);
      assert.match(changed.formula1, /^=?20$/);
      assert.equal(changed.stopIfTrue, true);
      assert.equal(changed.operator, 6);

      // `null` clears a nullable key; the others keep their values.
      wb.updateCfRule(0, ruleId, { dxfIndex: null });
      assert.equal(wb.getCfRule(0, 0).dxfIndex, null);
      assert.match(wb.getCfRule(0, 0).formula1, /^=?20$/);
      wb.updateCfRule(0, ruleId, {});
      assert.match(wb.getCfRule(0, 0).formula1, /^=?20$/);

      // A nested payload patches its own keys only.
      wb.updateCfRule(0, barId, { dataBar: { gradient: true, showValue: false } });
      const bar = wb.getCfRule(0, 1).dataBar;
      assert.equal(bar.gradient, true);
      assert.equal(bar.showValue, false);
      assert.equal(bar.fill.argb, 0xff0000ff);

      assert.throws(() => wb.updateCfRule(0, ruleId, { stopIfTrue: null }), TypeError);
      assert.throws(() => wb.updateCfRule(0, ruleId, { dxfIndex: 1.5 }), RangeError);
      expectFailure(() => wb.updateCfRule(0, ruleId, { rank: 3 }), StatusCode.InvalidArgument);
      expectFailure(() => wb.updateCfRule(0, ruleId, { dxfIndex: 99 }), StatusCode.IndexOutOfRange);
      expectFailure(() => wb.updateCfRule(0, '{00000000-0000-0000-0000-000000000000}', {}), StatusCode.NotFound);
      expectFailure(() => wb.updateCfRule(9, ruleId, {}), StatusCode.SheetIndexOutOfRange);
    } finally {
      wb.delete();
    }
  });

  test('data bar x14 fields survive save and load', () => {
    // The six live in the `x14` extension, not the legacy `<dataBar>`
    // element. An in-session round-trip would still pass with a writer that
    // never emits the extension -- which is how they were lost before -- so
    // this goes through a real save and load.
    const wb = Module.Workbook.createDefault();
    try {
      wb.addCfRule(0, {
        ranges: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 0 }],
        type: 3,
        dataBar: {
          min: { type: 3 },
          max: { type: 4 },
          fill: rgbColor(0xff0000ff),
          gradient: false,
          axisPosition: 1,
          negativeFill: rgbColor(0xffff0000),
          border: rgbColor(0xff090909),
          negativeBorder: rgbColor(0xff080808),
          axisColor: rgbColor(0xff010203),
        },
      });
      const saved = wb.save();
      const loaded = Module.Workbook.load(saved);
      try {
        assert.equal(loaded.disposed, false);
        const bar = readAll(
          () => loaded.getCfRuleCount(0),
          (i) => loaded.getCfRule(0, i),
        )[0].dataBar;
        assert.equal(bar.gradient, false);
        assert.equal(bar.axisPosition, 1);
        assert.deepEqual(bar.fill, rgbColor(0xff0000ff));
        assert.deepEqual(bar.negativeFill, rgbColor(0xffff0000));
        assert.deepEqual(bar.border, rgbColor(0xff090909));
        assert.deepEqual(bar.negativeBorder, rgbColor(0xff080808));
        assert.deepEqual(bar.axisColor, rgbColor(0xff010203));

        // The decoded bar fed straight back reproduces the same rule.
        loaded.addCfRule(0, {
          ranges: [{ firstRow: 4, firstCol: 0, lastRow: 6, lastCol: 0 }],
          type: 3,
          dataBar: bar,
        });
        const reread = readAll(
          () => loaded.getCfRuleCount(0),
          (i) => loaded.getCfRule(0, i),
        )[1].dataBar;
        assert.equal(reread.gradient, false);
        assert.equal(reread.axisPosition, 1);
        assert.deepEqual(reread.axisColor, rgbColor(0xff010203));
      } finally {
        loaded.delete();
      }
    } finally {
      wb.delete();
    }
  });

  test('omitted data bar x14 fields keep the model defaults', () => {
    const wb = Module.Workbook.createDefault();
    try {
      wb.addCfRule(0, {
        ranges: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 0 }],
        type: 3,
        dataBar: { min: { type: 3 }, max: { type: 4 }, fill: rgbColor(0xff0000ff) },
      });
      // The getter engages all six, so they read back as the defaults
      // rather than as absent keys.
      const bar = readAll(
        () => wb.getCfRuleCount(0),
        (i) => wb.getCfRule(0, i),
      )[0].dataBar;
      assert.equal(bar.gradient, true);
      assert.equal(bar.axisPosition, 0);
      assert.deepEqual(bar.negativeFill, rgbColor(0xff0000ff));
    } finally {
      wb.delete();
    }
  });

  test('saveWithDiagnostics() and readDiagnostics() expose stable counters', () => {
    const wb = Module.Workbook.createDefault();
    try {
      wb.setFormula(0, 0, 0, '=SUM(T[C])');
      wb.addValidation(0, {
        ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
        type: 3,
        formula1: '"Yes,No"',
      });
      // `dxfIndex` is a real index into `<dxfs>`, not a sentinel: every binding
      // signals "no differential format" by omitting the field, which leaves
      // the C ABI's `has_dxf_index` at 0. Supplying `0` therefore claims the
      // first dxf, so the table has to have one.
      const dxf = wb.addDxf({ font: { bold: true } });
      wb.addCfRule(0, {
        ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
        type: 1,
        operator: 6,
        formula1: '10',
        dxfIndex: dxf,
        stopIfTrue: false,
      });
      const xlsx = wb.saveWithDiagnostics(0);
      assert.equal(xlsx.downgradedFormulaCount, 0);
      assert.equal(xlsx.deferredFeatureCount, 0);
      assert.equal(xlsx.droppedPartCount, 0);
      assert.equal(xlsx.droppedRelationshipCount, 0);
      assert.equal(xlsx.renumberedPartCount, 0);

      const xlsb = wb.saveWithDiagnostics(1);
      assert.equal(xlsb.downgradedFormulaCount, 1);
      assert.equal(xlsb.deferredFeatureCount, 0);
      // The binary writer never reassigns a part id.
      assert.equal(xlsb.renumberedPartCount, 0);

      const loaded = Module.Workbook.load(xlsb.bytes);
      try {
        const read = loaded.getDiagnostics();
        assert.equal(read.undecodedFormulaCount, 0);
        assert.equal(read.undecodedDefinedNameCount, 0);
        assert.equal(read.undecodedPartCount, 0);
        assert.equal(read.skippedFeatureCount, 0);
        assert.equal(read.unknownContentTypeCount, 0);
      } finally {
        loaded.delete();
      }
      const preservedLoaded = Module.Workbook.load(appendEmptyZipEntry(xlsb.bytes, 'xl/preserved.bin'));
      try {
        const preserved = preservedLoaded.getDiagnostics();
        assert.equal(preserved.undecodedPartCount, 0);
      } finally {
        preservedLoaded.delete();
      }
      expectFailure(() => wb.saveWithDiagnostics(2), StatusCode.InvalidArgument);
      wb.save(1);
    } finally {
      wb.delete();
    }
  });

  test('Workbook.load reports empty input as a FormulonError', () => {
    expectFailure(() => Module.Workbook.load(new Uint8Array()), StatusCode.InvalidArgument);
  });

  test('Workbook.load rejects non-Uint8Array input like the native binding', () => {
    assert.throws(() => Module.Workbook.load(new Uint16Array([1])), TypeError);
  });

  test('read diagnostics of a disposed workbook fail with BindingInvalidHandle', () => {
    const wb = Module.Workbook.createDefault();
    wb.dispose();
    try {
      expectFailure(() => wb.getDiagnostics(), StatusCode.BindingInvalidHandle);
    } finally {
      wb.delete();
    }
  });

  test('formula error surfaces as ERROR-kind value', () => {
    const r = Module.evaluateFormula('=1/0').cells[0][0];
    assert.equal(r.kind, VAL.ERROR);
  });

  test('parser failure carries a readable message', () => {
    // An obviously unparseable formula. The C ABI either surfaces a
    // non-zero status (parser error) or returns ok with an Error
    // value — accept either, but the diagnostic should be readable
    // when the call throws.
    let r;
    try {
      r = Module.evaluateFormula('=#GARBAGE').cells[0][0];
    } catch (e) {
      assert.equal(e.name, 'FormulonError');
      assert.equal(typeof e.message, 'string');
      return;
    }
    // Acceptable alternative: the formula parses to something that
    // evaluates to an error value.
    assert.equal(r.kind, VAL.ERROR);
  });

  test('compareText folds by strength and rejects bad arguments', () => {
    const profile = 'mac-365-zh_CN';
    const strict = Module.compareText('A', '\uFF21', 0, profile);
    assert.notEqual(strict, 0);
    const dynamic = Module.compareText('A', '\uFF21', 1, profile);
    assert.equal(dynamic, 0);
    const ab = Module.compareText('a', 'b', 0, 'mac-365-en_US');
    const ba = Module.compareText('b', 'a', 0, 'mac-365-en_US');
    assert.equal(ab, -1);
    assert.equal(ba, 1);
    expectFailure(() => Module.compareText('a', 'b', 0, 'nope'), StatusCode.InvalidArgument);
    expectFailure(() => Module.compareText('a', 'b', -1, profile), StatusCode.InvalidArgument);
    expectFailure(() => Module.compareText('a', 'b', 3, profile));
  });

  test('convertCivilToSerial and convertSerialToCivil map the engine date rules', () => {
    const civil = (year, month, day, hour = 0, minute = 0, second = 0) => ({ year, month, day, hour, minute, second });
    assert.equal(Module.convertCivilToSerial(civil(1900, 1, 1), false), 1);
    assert.equal(Module.convertCivilToSerial(civil(1900, 2, 29), false), 60);
    assert.equal(Module.convertCivilToSerial(civil(1900, 3, 1), false), 61);
    assert.equal(Module.convertCivilToSerial(civil(1904, 1, 1), true), 0);
    assert.equal(Module.convertCivilToSerial(civil(2024, 1, 1, 12), false), 45292.5);
    assert.deepEqual(Module.convertSerialToCivil(60, false), civil(1900, 2, 29));
    assert.deepEqual(Module.convertSerialToCivil(45292.5, false), civil(2024, 1, 1, 12));
    assert.deepEqual(Module.convertSerialToCivil(0, true), civil(1904, 1, 1));
    expectFailure(() => Module.convertCivilToSerial(civil(1900, 2, 29), true), StatusCode.InvalidArgument);
    expectFailure(() => Module.convertCivilToSerial(civil(2024, 13, 1), false), StatusCode.InvalidArgument);
    expectFailure(() => Module.convertCivilToSerial(civil(1903, 12, 31), true), StatusCode.InvalidArgument);
    expectFailure(() => Module.convertSerialToCivil(0.5, false), StatusCode.InvalidArgument);
    expectFailure(() => Module.convertSerialToCivil(1e9, false), StatusCode.InvalidArgument);
    assert.throws(() => Module.convertCivilToSerial(civil(2024, 1.5, 1), false), RangeError);
    assert.throws(() => Module.convertCivilToSerial(civil(2 ** 31, 1, 1), false), RangeError);
    assert.throws(() => Module.convertCivilToSerial('2024-01-01', false), TypeError);
    assert.throws(() => Module.convertSerialToCivil('1', false), TypeError);
    expectFailure(() => Module.convertSerialToCivil(Number.NaN, false), StatusCode.InvalidArgument);
  });

  test('getCollationClasses orders by kind, folds text by strength, and copies its output', () => {
    const values = [
      { kind: 1, number: 1 },
      { kind: 3, text: '1' },
      { kind: 2, boolean: true },
      { kind: 3, text: 'TRUE' },
      { kind: 0 },
      { kind: 3, text: '' },
      { kind: 1, number: 0 },
      { kind: 2, boolean: false },
    ];
    const r = Module.getCollationClasses(values, 1, 'mac-365-en_US');
    assert.ok(r.classIds instanceof Uint32Array);
    assert.equal(r.classIds.length, 8);
    assert.equal(r.classCount, 8);
    assert.equal(new Set(r.classIds).size, 8);
    const [num1, str1, boolTrue, strTrue, blank, strEmpty, num0, boolFalse] = r.classIds;
    assert.ok(Math.max(num0, num1) < Math.min(str1, strTrue, strEmpty));
    assert.ok(Math.max(str1, strTrue, strEmpty) < Math.min(boolFalse, boolTrue));
    assert.ok(Math.max(boolFalse, boolTrue) < blank);
    assert.ok(num0 < num1 && boolFalse < boolTrue);

    const snapshot = Array.from(r.classIds);
    Module.getCollationClasses(
      [
        { kind: 3, text: 'z' },
        { kind: 3, text: 'y' },
      ],
      1,
      'mac-365-en_US',
    );
    assert.deepEqual(Array.from(r.classIds), snapshot);

    const folded = Module.getCollationClasses(
      [
        { kind: 3, text: 'a' },
        { kind: 3, text: 'A' },
      ],
      1,
      'mac-365-en_US',
    );
    assert.equal(folded.classIds[0], folded.classIds[1]);
    assert.equal(folded.classCount, 1);

    const empty = Module.getCollationClasses([], 1, 'mac-365-en_US');
    assert.equal(empty.classIds.length, 0);
    assert.equal(empty.classCount, 0);

    assert.throws(() => Module.getCollationClasses([{ number: 1 }], 1, 'mac-365-en_US'), TypeError);
    assert.throws(() => Module.getCollationClasses({}, 1, 'mac-365-en_US'), TypeError);
    expectFailure(() => Module.getCollationClasses([{ kind: 1, number: 1 }], 1, 'nope'));
    expectFailure(() => Module.getCollationClasses(values, -1, 'mac-365-en_US'), StatusCode.InvalidArgument);
  });

  test('getCollationClasses timing on 100k mixed CJK/ASCII texts', () => {
    const cjk = [
      '\u6771\u4EAC',
      '\u5927\u962A',
      '\u3055\u304F\u3089',
      '\u30B5\u30AF\u30E9',
      '\u4E2D\u6587',
      '\uD55C\uAE00',
    ];
    const values = [];
    for (let i = 0; i < 100000; i += 1) {
      const text = i % 2 === 0 ? `${cjk[i % cjk.length]}${i % 5000}` : `item${i % 7000}`;
      values.push({ kind: 3, text });
    }
    const t0 = process.hrtime.bigint();
    const r = Module.getCollationClasses(values, 1, 'mac-365-ja_JP');
    const ms = Number(process.hrtime.bigint() - t0) / 1e6;
    assert.equal(r.classIds.length, 100000);
    console.log(`getCollationClasses 100k mixed CJK/ASCII texts: ${ms.toFixed(0)} ms (${r.classCount} classes)`);
  });

  test('getLocaleFacts reports collationMeasured per host', () => {
    assert.equal(Module.getLocaleFacts('mac-365-de_DE').collationMeasured, true);
    assert.equal(Module.getLocaleFacts('win-365-ja_JP').collationMeasured, false);
    assert.equal(Module.getLocaleFacts('win-365-en_US').collationMeasured, false);
  });
}
