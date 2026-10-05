import assert from 'node:assert/strict';
import { appendEmptyZipEntry, readZipEntryText, VAL } from './smoke_support.mjs';

export function registerCellsIo(Module, test) {
  test('setText / setBool / setBlank round-trip', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setText(0, 0, 0, 'hello').ok);
      assert.ok(wb.setBool(0, 1, 0, true).ok);
      assert.ok(wb.setBlank(0, 2, 0).ok);
      assert.ok(wb.recalc().ok);

      const a1 = wb.getValue(0, 0, 0);
      assert.equal(a1.value.kind, VAL.TEXT);
      assert.equal(a1.value.text, 'hello');

      const a2 = wb.getValue(0, 1, 0);
      assert.equal(a2.value.kind, VAL.BOOL);
      assert.equal(a2.value.boolean, 1);

      const a3 = wb.getValue(0, 2, 0);
      assert.equal(a3.value.kind, VAL.BLANK);
    } finally {
      wb.delete();
    }
  });

  test('setCellPhonetic / getCellPhonetic preserve and clear a Japanese guide', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setText(0, 0, 0, '漢字').ok);
      assert.ok(wb.setCellPhonetic(0, 0, 0, 'かんじ').ok);
      const phonetic = wb.getCellPhonetic(0, 0, 0);
      assert.ok(phonetic.status.ok);
      assert.equal(phonetic.value, 'かんじ');

      assert.ok(wb.setCellPhonetic(0, 0, 0, '').ok);
      assert.equal(wb.getCellPhonetic(0, 0, 0).value, '');

      assert.ok(wb.setCellPhonetic(0, 0, 0, 'かんじ').ok);
      assert.ok(wb.setText(0, 0, 0, '文字列').ok);
      assert.equal(wb.getCellPhonetic(0, 0, 0).value, '');
    } finally {
      wb.delete();
    }
  });

  test('setCellPhoneticRuns / getCellPhoneticRuns keep the spans apart', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setText(0, 0, 0, '東京都').ok);
      assert.deepEqual(wb.getCellPhoneticRuns(0, 0, 0).runs, []);

      const runs = [
        { sb: 0, eb: 2, text: 'トウキョウ' },
        { sb: 2, eb: 3, text: 'ト' },
      ];
      assert.ok(wb.setCellPhoneticRuns(0, 0, 0, runs).ok);
      assert.deepEqual(wb.getCellPhoneticRuns(0, 0, 0).runs, runs);
      // The flattening getter still reports the concatenation.
      assert.equal(wb.getCellPhonetic(0, 0, 0).value, 'トウキョウト');

      // Writing that back through the whole-cell setter is the collapse the
      // run API exists to avoid.
      assert.ok(wb.setCellPhonetic(0, 0, 0, 'トウキョウト').ok);
      assert.deepEqual(wb.getCellPhoneticRuns(0, 0, 0).runs, [{ sb: 0, eb: 3, text: 'トウキョウト' }]);

      assert.ok(wb.setCellPhoneticRuns(0, 0, 0, []).ok);
      assert.deepEqual(wb.getCellPhoneticRuns(0, 0, 0).runs, []);

      assert.equal(
        wb.setCellPhoneticRuns(0, 0, 0, [
          { sb: 2, eb: 3, text: 'ト' },
          { sb: 0, eb: 2, text: 'トウ' },
        ]).ok,
        false,
      );
    } finally {
      wb.delete();
    }
  });

  test('setCellPhoneticProperties survives a run edit and a save/load cycle', () => {
    // The result carries a `status` alongside the triple; compare the
    // triple alone so a status-shape change does not read as a value change.
    const phoneticProps = (book, sheet, row, col) => {
      const { fontId, type, alignment } = book.getCellPhoneticProperties(sheet, row, col);
      return { fontId, type, alignment };
    };
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setText(0, 0, 0, '大阪').ok);
      // An unannotated cell reports what Excel infers for a guide written
      // with no <phoneticPr> at all.
      assert.deepEqual(phoneticProps(wb, 0, 0, 0), { fontId: 0, type: 0, alignment: 0 });

      assert.ok(wb.setCellPhoneticProperties(0, 0, 0, { fontId: 3, type: 2, alignment: 2 }).ok);
      // Setting the readings must not reset the rendering.
      assert.ok(wb.setCellPhoneticRuns(0, 0, 0, [{ sb: 0, eb: 2, text: 'おおさか' }]).ok);
      assert.deepEqual(phoneticProps(wb, 0, 0, 0), { fontId: 3, type: 2, alignment: 2 });

      const saved = wb.save();
      assert.ok(saved.status.ok);
      const loaded = Module.Workbook.loadBytes(saved.bytes);
      try {
        assert.deepEqual(phoneticProps(loaded, 0, 0, 0), { fontId: 3, type: 2, alignment: 2 });
      } finally {
        loaded.delete();
      }

      // Two bits each on the binary side, so a wider ordinal is refused
      // rather than truncated into a different, valid-looking one.
      assert.equal(wb.setCellPhoneticProperties(0, 0, 0, { fontId: 0, type: 4, alignment: 0 }).ok, false);
      assert.deepEqual(phoneticProps(wb, 0, 0, 0), { fontId: 3, type: 2, alignment: 2 });

      // A value write discards the guide and its rendering together.
      assert.ok(wb.setText(0, 0, 0, '京都').ok);
      assert.deepEqual(phoneticProps(wb, 0, 0, 0), { fontId: 0, type: 0, alignment: 0 });
    } finally {
      wb.delete();
    }
  });

  test('setDefaultFont replaces font 0, which addFont can only append beside', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.equal(wb.getFont(0).name, 'Calibri');
      const appended = wb.addFont({ name: '游ゴシック', size: 11 });
      assert.ok(appended.status.ok);
      assert.ok(appended.index > 0);
      assert.equal(wb.getFont(0).name, 'Calibri');

      assert.ok(wb.setDefaultFont({ name: '游ゴシック', size: 11, hasCharset: true, charset: 128 }).ok);
      assert.equal(wb.getFont(0).name, '游ゴシック');
      assert.equal(wb.getFont(0).charset, 128);

      const added = wb.addFont({ name: 'Meiryo', size: 12 });
      assert.ok(wb.setFont(added.index, { name: 'MS Gothic', size: 9 }).ok);
      assert.equal(wb.getFont(added.index).name, 'MS Gothic');
      assert.equal(wb.setFont(wb.fontCount().value, { name: 'MS Gothic', size: 9 }).ok, false);
    } finally {
      wb.delete();
    }
  });

  test('table update omits fields without resetting existing metadata', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const created = wb.createTable({
        sheetIndex: 0,
        ref: 'A1:B3',
        name: 'Sales',
        columns: ['Product', 'Amount'],
        styleName: 'TableStyleMedium2',
        headerRow: false,
        totalsRow: true,
      });
      assert.ok(created.status.ok, `createTable failed: ${JSON.stringify(created.status)}`);
      const before = wb.tableAt(created.index);
      assert.ok(before.status.ok);
      assert.equal(before.ref, 'A1:B3');

      const beforeSaved = wb.save();
      assert.ok(beforeSaved.status.ok);
      const beforeXml = readZipEntryText(beforeSaved.bytes, 'xl/tables/table1.xml');
      assert.match(beforeXml, /name="TableStyleMedium2"/);
      assert.match(beforeXml, /headerRowCount="0"/);
      assert.match(beforeXml, /totalsRowCount="1"/);

      // An empty update object is a partial no-op. This specifically covers
      // the WASM-to-C mapping of omitted ref/style/row flags to preservation
      // sentinels rather than empty/false defaults.
      assert.ok(wb.updateTable(created.index, {}).ok);
      const after = wb.tableAt(created.index);
      assert.ok(after.status.ok);
      assert.equal(after.ref, before.ref);

      const saved = wb.save();
      assert.ok(saved.status.ok, `save failed: ${JSON.stringify(saved.status)}`);
      const afterXml = readZipEntryText(saved.bytes, 'xl/tables/table1.xml');
      assert.match(afterXml, /name="TableStyleMedium2"/);
      assert.match(afterXml, /headerRowCount="0"/);
      assert.match(afterXml, /totalsRowCount="1"/);
      const loaded = Module.Workbook.loadBytes(saved.bytes);
      try {
        assert.ok(loaded.isValid(), `load failed: ${Module.lastErrorMessage()}`);
        const reloaded = loaded.tableAt(created.index);
        assert.ok(reloaded.status.ok);
        assert.equal(reloaded.ref, 'A1:B3');
        const roundTripped = loaded.save();
        assert.ok(roundTripped.status.ok);
        const roundTrippedXml = readZipEntryText(roundTripped.bytes, 'xl/tables/table1.xml');
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

  test('save() + Workbook.loadBytes() round-trip preserves a literal', () => {
    const wb = Module.Workbook.createDefault();
    let saved = null;
    try {
      assert.ok(wb.setNumber(0, 0, 0, 7).ok);
      assert.ok(wb.setFormula(0, 1, 0, '=A1+1').ok);
      assert.ok(wb.recalc().ok);

      const r = wb.save();
      assert.ok(r.status.ok, `save failed: ${JSON.stringify(r.status)}`);
      assert.ok(r.bytes instanceof Uint8Array);
      assert.ok(r.bytes.length > 0);
      saved = r.bytes;
    } finally {
      wb.delete();
    }

    const loaded = Module.Workbook.loadBytes(saved);
    try {
      assert.ok(loaded.isValid(), `load failed: ${Module.lastErrorMessage()}`);
      assert.ok(loaded.sheetCount().value >= 1);
      assert.ok(loaded.recalc().ok);
      const a1 = loaded.getValue(0, 0, 0);
      assert.ok(a1.status.ok);
      assert.equal(a1.value.kind, VAL.NUMBER);
      assert.equal(a1.value.number, 7);
    } finally {
      loaded.delete();
    }
  });

  test('data bar x14 fields survive save and load', () => {
    // The six live in the `x14` extension, not the legacy `<dataBar>`
    // element. An in-session round-trip would still pass with a writer that
    // never emits the extension -- which is how they were lost before -- so
    // this goes through a real save and load.
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(
        wb.addConditionalFormat(0, {
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 0 }],
          type: 3,
          dataBar: {
            min: { type: 3 },
            max: { type: 4 },
            fill: { r: 0, g: 0, b: 255 },
            gradient: false,
            axisPosition: 1,
            negativeFill: { r: 255, g: 0, b: 0 },
            border: { r: 9, g: 9, b: 9 },
            negativeBorder: { r: 8, g: 8, b: 8 },
            axisColor: { r: 1, g: 2, b: 3 },
          },
        }).status.ok,
      );
      const saved = wb.save();
      assert.ok(saved.status.ok, `save failed: ${JSON.stringify(saved.status)}`);
      const loaded = Module.Workbook.loadBytes(saved.bytes);
      try {
        assert.ok(loaded.isValid(), `load failed: ${Module.lastErrorMessage()}`);
        const bar = loaded.getConditionalFormats(0)[0].dataBar;
        assert.equal(bar.gradient, false);
        assert.equal(bar.axisPosition, 1);
        assert.equal(bar.negativeFill.r, 255);
        assert.equal(bar.border.r, 9);
        assert.equal(bar.negativeBorder.r, 8);
        assert.equal(bar.axisColor.b, 3);

        // The decoded bar fed straight back reproduces the same rule.
        assert.ok(
          loaded.addConditionalFormat(0, {
            sqref: [{ firstRow: 4, firstCol: 0, lastRow: 6, lastCol: 0 }],
            type: 3,
            dataBar: bar,
          }).status.ok,
        );
        const reread = loaded.getConditionalFormats(0)[1].dataBar;
        assert.equal(reread.gradient, false);
        assert.equal(reread.axisPosition, 1);
        assert.equal(reread.axisColor.b, 3);
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
      assert.ok(
        wb.addConditionalFormat(0, {
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 0 }],
          type: 3,
          dataBar: { min: { type: 3 }, max: { type: 4 }, fill: { r: 0, g: 0, b: 255 } },
        }).status.ok,
      );
      // The getter engages all six, so they read back as the defaults
      // rather than as absent keys.
      const bar = wb.getConditionalFormats(0)[0].dataBar;
      assert.equal(bar.gradient, true);
      assert.equal(bar.axisPosition, 0);
      assert.equal(bar.negativeFill.b, 255);
    } finally {
      wb.delete();
    }
  });

  test('saveWithDiagnostics() and readDiagnostics() expose stable counters', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setFormula(0, 0, 0, '=SUM(T[C])').ok);
      assert.ok(
        wb.addValidation(0, {
          ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
          type: 3,
          formula1: '"Yes,No"',
        }).ok,
      );
      // `dxfId` is a real index into `<dxfs>`, not a sentinel: every binding
      // signals "no differential format" by omitting the field, which leaves
      // the C ABI's `dxf_id_engaged` at 0. Supplying `0` therefore claims the
      // first dxf, so the table has to have one.
      const dxf = wb.addDxf({ font: { bold: true } });
      assert.ok(dxf.status.ok, `addDxf failed: ${JSON.stringify(dxf.status)}`);
      assert.ok(
        wb.addConditionalFormat(0, {
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
          type: 1,
          op: 5,
          formula1: '10',
          dxfId: dxf.index,
          stopIfTrue: false,
        }).status.ok,
      );
      const xlsx = wb.saveWithDiagnostics(1);
      assert.ok(xlsx.status.ok, `xlsx save failed: ${JSON.stringify(xlsx.status)}`);
      assert.equal(xlsx.downgradedFormulaCount, 0);
      assert.equal(xlsx.deferredFeatureCount, 0);
      assert.equal(xlsx.droppedPartCount, 0);
      assert.equal(xlsx.droppedRelationshipCount, 0);
      assert.equal(xlsx.renumberedPartCount, 0);

      const xlsb = wb.saveWithDiagnostics(2);
      assert.ok(xlsb.status.ok, `xlsb save failed: ${JSON.stringify(xlsb.status)}`);
      assert.equal(xlsb.downgradedFormulaCount, 1);
      assert.equal(xlsb.deferredFeatureCount, 0);
      // The binary writer never reassigns a part id.
      assert.equal(xlsb.renumberedPartCount, 0);

      const loaded = Module.Workbook.loadBytes(xlsb.bytes);
      try {
        const read = loaded.readDiagnostics();
        assert.ok(read.status.ok, `read diagnostics failed: ${JSON.stringify(read.status)}`);
        assert.equal(read.undecodedFormulaCount, 0);
        assert.equal(read.undecodedDefinedNameCount, 0);
        assert.equal(read.undecodedPartCount, 0);
        assert.equal(read.skippedFeatureCount, 0);
        assert.equal(read.unknownContentTypeCount, 0);
      } finally {
        loaded.delete();
      }
      const preservedLoaded = Module.Workbook.loadBytes(appendEmptyZipEntry(xlsb.bytes, 'xl/preserved.bin'));
      try {
        const preserved = preservedLoaded.readDiagnostics();
        assert.ok(preserved.status.ok, `passthrough diagnostics failed: ${JSON.stringify(preserved.status)}`);
        assert.equal(preserved.undecodedPartCount, 0);
      } finally {
        preservedLoaded.delete();
      }
      const invalidSave = wb.saveWithDiagnostics(0);
      assert.ok(!invalidSave.status.ok, `unknown format unexpectedly succeeded: ${JSON.stringify(invalidSave)}`);
      assert.equal(invalidSave.bytes, null);
      assert.equal(invalidSave.downgradedFormulaCount, 0);
      assert.equal(invalidSave.deferredFeatureCount, 0);
      assert.equal(invalidSave.droppedPartCount, 0);
      assert.equal(invalidSave.droppedRelationshipCount, 0);
      assert.equal(invalidSave.renumberedPartCount, 0);
      assert.ok(wb.saveAs(2).status.ok);
    } finally {
      wb.delete();
    }
  });

  test('Workbook.loadBytes reports empty input instead of leaving stale diagnostics', () => {
    const wb = Module.Workbook.loadBytes(new Uint8Array());
    try {
      assert.equal(wb.isValid(), false);
      assert.match(Module.lastErrorMessage(), /NULL or empty input/);
    } finally {
      wb.delete();
    }
  });

  test('Workbook.loadBytes rejects non-Uint8Array input like the native binding', () => {
    const wb = Module.Workbook.loadBytes(new Uint16Array([1]));
    try {
      assert.equal(wb.isValid(), false);
      assert.match(Module.lastErrorMessage(), /NULL or empty input/);
    } finally {
      wb.delete();
    }
  });

  test('invalid workbook read diagnostics return a zeroed failure envelope', () => {
    const wb = Module.Workbook.loadBytes(new Uint8Array());
    try {
      const read = wb.readDiagnostics();
      assert.ok(!read.status.ok, `invalid handle unexpectedly succeeded: ${JSON.stringify(read)}`);
      assert.equal(read.undecodedFormulaCount, 0);
      assert.equal(read.undecodedDefinedNameCount, 0);
      assert.equal(read.undecodedPartCount, 0);
      assert.equal(read.skippedFeatureCount, 0);
      assert.equal(read.unknownContentTypeCount, 0);
    } finally {
      wb.delete();
    }
  });

  test('formula error surfaces as ERROR-kind value', () => {
    const r = Module.evalFormula('=1/0');
    assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
    assert.equal(r.value.kind, VAL.ERROR);
  });

  test('parser failure populates lastErrorMessage', () => {
    // An obviously unparseable formula. The C ABI either surfaces a
    // non-zero status (parser error) or returns ok with an Error
    // value — accept either, but the diagnostic should be readable
    // when the status is non-ok.
    const r = Module.evalFormula('=#GARBAGE');
    if (!r.status.ok) {
      assert.equal(typeof r.status.message, 'string');
      const msg = Module.lastErrorMessage();
      assert.equal(typeof msg, 'string');
    } else {
      // Acceptable alternative: the formula parses to something that
      // evaluates to an error value.
      assert.equal(r.value.kind, VAL.ERROR);
    }
  });
}
