import test from 'node:test';
import { appendEmptyZipEntry, assert, getModule } from './smoke_support.mjs';

test('comments round-trip: setComment + getComment', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Pre-existing cell so the comment attaches to a stored slot.
  assert.ok(wb.setNumber(0, 1, 1, 7).ok);
  const sc = wb.setComment(0, 1, 1, 'libraz', 'hello world');
  assert.ok(sc.ok, `setComment: ${JSON.stringify(sc)}`);
  const c = wb.getComment(0, 1, 1);
  assert.ok(c !== null, 'expected non-null CommentEntry');
  assert.equal(c.author, 'libraz');
  assert.equal(c.text, 'hello world');
  // Removing surfaces null on the next read.
  assert.ok(wb.setComment(0, 1, 1, '', '').ok);
  assert.equal(wb.getComment(0, 1, 1), null);
});

test('getCommentResult distinguishes absence from an invalid sheet', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const missing = wb.getCommentResult(0, 1, 1);
  assert.equal(missing.status.ok, false);
  assert.equal(missing.comment, null);
  const invalid = wb.getCommentResult(99, 1, 1);
  assert.equal(invalid.status.ok, false);
  assert.equal(invalid.comment, null);
  assert.notEqual(missing.status.status, invalid.status.status);
  assert.ok(wb.setComment(0, 1, 1, 'libraz', 'hello').ok);
  const found = wb.getCommentResult(0, 1, 1);
  assert.equal(found.status.ok, true);
  assert.equal(found.comment.author, 'libraz');
  assert.equal(found.comment.text, 'hello');
});

test('cellCount + cellAt enumerate stored cells', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setNumber(0, 0, 0, 1).ok);
  assert.ok(wb.setNumber(0, 0, 1, 2).ok);
  assert.ok(wb.setFormula(0, 1, 0, '=A1+B1').ok);
  assert.ok(wb.recalc().ok);
  const n = wb.cellCount(0).value;
  assert.equal(typeof n, 'number');
  assert.ok(n >= 3, `expected >= 3 stored cells, got ${n}`);
  // Iterate; one of the entries should carry our formula.
  let foundFormula = false;
  for (let i = 0; i < n; ++i) {
    const e = wb.cellAt(0, i);
    assert.ok(e.status.ok, `cellAt(${i}): ${JSON.stringify(e.status)}`);
    assert.equal(typeof e.row, 'number');
    assert.equal(typeof e.col, 'number');
    if (e.formula === '=A1+B1') {
      foundFormula = true;
      assert.equal(e.value.kind, mod.ValueKind.Number);
      assert.equal(e.value.number, 3);
    }
  }
  assert.ok(foundFormula, 'expected to find the =A1+B1 cell during iteration');
});

test('default export exposes CfMatchKind ordinals', async () => {
  const mod = await getModule();
  assert.equal(typeof mod.CfMatchKind, 'object');
  assert.equal(mod.CfMatchKind.DifferentialFormat, 0);
  assert.equal(mod.CfMatchKind.ColorScale, 1);
  assert.equal(mod.CfMatchKind.DataBar, 2);
  assert.equal(mod.CfMatchKind.IconSet, 3);
});

test('default export exposes ErrorCode ordinals and versionString alias', async () => {
  const mod = await getModule();
  assert.equal(typeof mod.ErrorCode, 'object');
  assert.equal(mod.ErrorCode.Div0, 1);
  assert.equal(mod.ErrorCode.Ref, 3);
  assert.equal(mod.ErrorCode.NA, 6);
  assert.equal(typeof mod.versionString, 'function');
  assert.equal(mod.versionString(), mod.version());
});

test('setError() rejects a missing errorCode argument', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.throws(() => wb.setError(0, 0, 0), TypeError);
});

test('setError() rejects a non-number errorCode before mutating the cell', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setNumber(0, 0, 0, 7).ok);
  assert.throws(() => wb.setError(0, 0, 0, 'not a number'), TypeError);
  // The rejected call must not have overwritten the cell.
  const got = wb.getValue(0, 0, 0);
  assert.ok(got.status.ok, JSON.stringify(got.status));
  assert.equal(got.value.number, 7);
});

test('saveAs() writes xlsx and xlsb bytes; rejects a missing format argument', async () => {
  const mod = await getModule();
  assert.equal(typeof mod.WorkbookFormat, 'object');
  const wb = mod.Workbook.createDefault();
  const xlsx = wb.saveAs(mod.WorkbookFormat.Xlsx);
  assert.ok(xlsx.status.ok, JSON.stringify(xlsx.status));
  assert.ok(xlsx.bytes.length > 0);
  const xlsb = wb.saveAs(mod.WorkbookFormat.Xlsb);
  assert.ok(xlsb.status.ok, JSON.stringify(xlsb.status));
  assert.ok(xlsb.bytes.length > 0);
  assert.throws(() => wb.saveAs(), TypeError);
});

test('saveWithDiagnostics() reports counters and readDiagnostics keeps a stable shape', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setFormula(0, 0, 0, '=SUM(T[C])').ok);
  assert.ok(
    wb.addValidation(0, {
      ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
      type: 3,
      formula1: '"Yes,No"',
    }).ok,
  );
  // A CF rule's dxfId must resolve against a registered dxf.
  const dxf = wb.addDxf({ font: { bold: true } });
  assert.ok(dxf.status.ok, JSON.stringify(dxf.status));
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

  const xlsx = wb.saveWithDiagnostics(mod.WorkbookFormat.Xlsx);
  assert.ok(xlsx.status.ok, JSON.stringify(xlsx.status));
  assert.ok(xlsx.bytes instanceof Uint8Array);
  assert.equal(xlsx.downgradedFormulaCount, 0);
  assert.equal(xlsx.deferredFeatureCount, 0);
  assert.equal(xlsx.droppedPartCount, 0);
  assert.equal(xlsx.droppedRelationshipCount, 0);
  assert.equal(xlsx.renumberedPartCount, 0);

  const xlsb = wb.saveWithDiagnostics(mod.WorkbookFormat.Xlsb);
  assert.ok(xlsb.status.ok, JSON.stringify(xlsb.status));
  assert.ok(xlsb.bytes instanceof Uint8Array);
  assert.equal(xlsb.downgradedFormulaCount, 1);
  // The validation and the CF rule, with its dxf, are both written.
  assert.equal(xlsb.deferredFeatureCount, 0);
  // The binary writer never reassigns a part id.
  assert.equal(xlsb.renumberedPartCount, 0);

  const loaded = mod.Workbook.loadBytes(xlsb.bytes);
  const read = loaded.readDiagnostics();
  assert.ok(read.status.ok, JSON.stringify(read.status));
  assert.equal(read.undecodedFormulaCount, 0);
  assert.equal(read.undecodedDefinedNameCount, 0);
  assert.equal(read.undecodedPartCount, 0);
  assert.equal(read.skippedFeatureCount, 0);
  assert.equal(read.unknownContentTypeCount, 0);
  const preservedLoaded = mod.Workbook.loadBytes(appendEmptyZipEntry(xlsb.bytes, 'xl/preserved.bin'));
  const preserved = preservedLoaded.readDiagnostics();
  assert.ok(preserved.status.ok, JSON.stringify(preserved.status));
  assert.equal(preserved.undecodedPartCount, 0);
  preservedLoaded.dispose();
  const invalidSave = wb.saveWithDiagnostics(0);
  assert.ok(!invalidSave.status.ok, JSON.stringify(invalidSave));
  assert.equal(invalidSave.bytes, null);
  assert.equal(invalidSave.downgradedFormulaCount, 0);
  assert.equal(invalidSave.deferredFeatureCount, 0);
  assert.equal(invalidSave.droppedPartCount, 0);
  assert.equal(invalidSave.droppedRelationshipCount, 0);
  assert.equal(invalidSave.renumberedPartCount, 0);
  assert.ok(wb.saveAs(mod.WorkbookFormat.Xlsb).bytes instanceof Uint8Array);
});

test('invalid workbook read diagnostics return a zeroed failure envelope', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.loadBytes(new Uint8Array());
  const read = wb.readDiagnostics();
  assert.ok(!read.status.ok, JSON.stringify(read));
  assert.equal(read.undecodedFormulaCount, 0);
  assert.equal(read.undecodedDefinedNameCount, 0);
  assert.equal(read.undecodedPartCount, 0);
  assert.equal(read.skippedFeatureCount, 0);
  assert.equal(read.unknownContentTypeCount, 0);
  wb.dispose();
});

test('addValidation / getValidations / removeValidationAt / clearValidations round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Empty by default.
  assert.equal(wb.getValidations(0).length, 0);

  // List-type validation at A1:B3.
  const listRule = {
    ranges: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1 }],
    type: 3,
    op: 0,
    errorStyle: 1,
    allowBlank: true,
    showInputMessage: true,
    showErrorMessage: true,
    formula1: '"Yes,No,Maybe"',
    promptTitle: 'Choose',
    promptMessage: 'Pick one',
    errorTitle: 'Bad value',
    errorMessage: 'Pick from the list',
  };
  assert.ok(wb.addValidation(0, listRule).ok);

  // Decimal between [0, 100] across two rectangles.
  const decimalRule = {
    ranges: [
      { firstRow: 4, firstCol: 0, lastRow: 4, lastCol: 0 },
      { firstRow: 6, firstCol: 0, lastRow: 9, lastCol: 0 },
    ],
    type: 2,
    op: 0,
    formula1: '0',
    formula2: '100',
    allowBlank: false,
  };
  assert.ok(wb.addValidation(0, decimalRule).ok);

  const list = wb.getValidations(0);
  assert.equal(list.length, 2);
  // Booleans must arrive as JS booleans, not 0/1.
  assert.equal(typeof list[0].allowBlank, 'boolean');
  assert.equal(list[0].allowBlank, true);
  assert.equal(typeof list[0].showInputMessage, 'boolean');
  assert.equal(list[0].showInputMessage, true);
  assert.equal(typeof list[1].allowBlank, 'boolean');
  assert.equal(list[1].allowBlank, false);
  assert.equal(list[0].formula1, '"Yes,No,Maybe"');
  assert.equal(list[1].formula1, '0');
  assert.equal(list[1].formula2, '100');
  assert.equal(list[1].ranges.length, 2);
  assert.equal(list[1].ranges[1].firstRow, 6);

  // removeValidationAt removes the first rule.
  assert.ok(wb.removeValidationAt(0, 0).ok);
  const after = wb.getValidations(0);
  assert.equal(after.length, 1);
  assert.equal(after[0].type, 2);

  // Out-of-range index is rejected.
  assert.ok(!wb.removeValidationAt(0, 99).ok);
  // Sheet-out-of-range is rejected on every entry.
  assert.ok(!wb.addValidation(999, listRule).ok);
  assert.ok(!wb.removeValidationAt(999, 0).ok);
  assert.ok(!wb.clearValidations(999).ok);

  // clearValidations drops everything; safe to call again.
  assert.ok(wb.clearValidations(0).ok);
  assert.equal(wb.getValidations(0).length, 0);
  assert.ok(wb.clearValidations(0).ok);
});

// -- Newly wired surface (audit parity with the WASM binding) ----------
// These tests exercise the methods that were previously only reachable
// from the WASM package. They assert sensible return shapes rather than
// deep semantics; the point is "the addon wires them and does not throw".

test('addConditionalFormat + getConditionalFormats + clearConditionalFormats round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.getConditionalFormats(0).length, 0);
  // A CF rule's dxfId must resolve against a registered dxf.
  const dxf = wb.addDxf({ font: { bold: true } });
  assert.ok(dxf.status.ok, JSON.stringify(dxf.status));
  // cellIs rule (type 1) comparing the cell to a literal.
  const add = wb.addConditionalFormat(0, {
    sqref: [{ firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 0 }],
    type: 1,
    op: 5,
    formula1: '10',
    dxfId: dxf.index,
    stopIfTrue: false,
  });
  assert.ok(add.status.ok, `addConditionalFormat: ${JSON.stringify(add)}`);
  assert.equal(typeof add.index, 'number');
  const list = wb.getConditionalFormats(0);
  assert.equal(list.length, 1);
  assert.equal(list[0].type, 1);
  assert.equal(typeof list[0].id, 'string');
  assert.equal(list[0].sqref.length, 1);
  assert.equal(list[0].sqref[0].lastRow, 4);
  assert.equal(list[0].formula1, '10');
  // removeConditionalFormatAt drops the only rule.
  assert.ok(wb.removeConditionalFormatAt(0, 0).ok);
  assert.equal(wb.getConditionalFormats(0).length, 0);
  // clearConditionalFormats is a safe no-op on an empty sheet.
  assert.ok(wb.clearConditionalFormats(0).ok);
  // Sheet-out-of-range is rejected.
  assert.ok(!wb.clearConditionalFormats(999).ok);
});

test('data bar x14 fields survive save and load', async () => {
  // The six live in the `x14` extension rather than the legacy `<dataBar>`
  // element, so an in-session round-trip would still pass with a writer
  // that never emits the extension -- which is exactly how they used to be
  // lost. Only a save/load cycle pins them.
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const add = wb.addConditionalFormat(0, {
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
      direction: 2,
    },
  });
  assert.ok(add.status.ok, `addConditionalFormat: ${JSON.stringify(add)}`);

  const saved = wb.save();
  assert.ok(saved.status.ok);
  const loaded = mod.Workbook.loadBytes(saved.bytes);
  const bar = loaded.getConditionalFormats(0)[0].dataBar;
  assert.equal(bar.gradient, false);
  assert.equal(bar.axisPosition, 1);
  assert.equal(bar.negativeFill.r, 255);
  assert.equal(bar.border.r, 9);
  assert.equal(bar.negativeBorder.r, 8);
  assert.equal(bar.axisColor.b, 3);
  assert.equal(bar.direction, 2);

  // The decoded bar fed straight back must reproduce the same rule.
  const again = loaded.addConditionalFormat(0, {
    sqref: [{ firstRow: 4, firstCol: 0, lastRow: 6, lastCol: 0 }],
    type: 3,
    dataBar: bar,
  });
  assert.ok(again.status.ok, `re-add: ${JSON.stringify(again)}`);
  const reread = loaded.getConditionalFormats(0)[1].dataBar;
  assert.equal(reread.gradient, false);
  assert.equal(reread.axisPosition, 1);
  assert.equal(reread.axisColor.b, 3);
  assert.equal(reread.direction, 2);

  loaded.dispose();
  wb.dispose();
});

test('icon set floor survives save and load', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const add = wb.addConditionalFormat(0, {
    sqref: [{ firstRow: 0, firstCol: 0, lastRow: 9, lastCol: 0 }],
    type: 4,
    iconSet: {
      name: 2,
      thresholds: [
        { type: 0, value: '7' },
        { type: 0, value: '9' },
      ],
      floor: { type: 0, value: '5', gte: false },
    },
  });
  assert.ok(add.status.ok, `addConditionalFormat: ${JSON.stringify(add)}`);
  const saved = wb.save();
  assert.ok(saved.status.ok);
  const loaded = mod.Workbook.loadBytes(saved.bytes);
  const floor = loaded.getConditionalFormats(0)[0].iconSet.floor;
  assert.equal(floor.type, 0);
  assert.equal(floor.value, '5');
  assert.equal(floor.gte, false);
  loaded.dispose();
  wb.dispose();
});

test('omitted data bar x14 fields keep the model defaults', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const add = wb.addConditionalFormat(0, {
    sqref: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 0 }],
    type: 3,
    dataBar: { min: { type: 3 }, max: { type: 4 }, fill: { r: 0, g: 0, b: 255 } },
  });
  assert.ok(add.status.ok, `addConditionalFormat: ${JSON.stringify(add)}`);
  const bar = wb.getConditionalFormats(0)[0].dataBar;
  // The getter engages all six, so they read back as the defaults rather
  // than as absent keys.
  assert.equal(bar.gradient, true);
  assert.equal(bar.axisPosition, 0);
  assert.equal(bar.negativeFill.b, 255);
  wb.dispose();
});

test('named cell styles: cellStyleCount / getCellStyle / cellStyleXfCount', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // A fresh workbook may carry zero named styles; the accessors must
  // still return well-formed values rather than throwing.
  const count = wb.cellStyleCount().value;
  assert.equal(typeof count, 'number');
  assert.ok(count >= 0);
  assert.equal(typeof wb.cellStyleXfCount().value, 'number');
  if (count > 0) {
    const cs = wb.getCellStyle(0);
    assert.ok(cs.status.ok, `getCellStyle: ${JSON.stringify(cs.status)}`);
    assert.equal(typeof cs.name, 'string');
    assert.equal(typeof cs.xfId, 'number');
  } else {
    // Out-of-range index surfaces an error status, not a throw.
    const cs = wb.getCellStyle(0);
    assert.equal(cs.status.ok, false);
  }
});

test('setCalcMode / calcMode round-trip the calc policy', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Default is automatic (0).
  assert.equal(wb.calcMode().value, 0);
  assert.ok(wb.setCalcMode(1).ok);
  assert.equal(wb.calcMode().value, 1);
  assert.ok(wb.setCalcMode(2).ok);
  assert.equal(wb.calcMode().value, 2);
});

test('setPinnedNow / pinnedNow / clearPinnedNow drive the clock seam', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Unpinned by default: the workbook follows the host clock.
  assert.equal(wb.pinnedNow().now, null);

  assert.ok(wb.setPinnedNow(2026, 4, 23, 15, 30, 45).ok);
  assert.deepEqual(wb.pinnedNow().now, {
    year: 2026,
    month: 4,
    day: 23,
    hour: 15,
    minute: 30,
    second: 45,
  });

  // 2026-04-23 is serial 46135 under the 1900 date system.
  const today = wb.evaluateFormulaText(0, 0, 0, '=TODAY()');
  assert.ok(today.status.ok, `evaluateFormulaText: ${JSON.stringify(today.status)}`);
  assert.equal(today.value.number, 46135);

  // The pin is a calendar instant, not a normalising constructor.
  assert.equal(wb.setPinnedNow(2026, 13, 1, 0, 0, 0).ok, false);
  assert.equal(wb.setPinnedNow(2025, 2, 29, 0, 0, 0).ok, false);

  assert.ok(wb.clearPinnedNow().ok);
  assert.equal(wb.pinnedNow().now, null);
});

test('excelProfileId / setExcelProfileId round-trip the profile id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const def = wb.excelProfileId().value;
  assert.equal(typeof def, 'string');
  assert.ok(def.length > 0);
  assert.ok(wb.setExcelProfileId('mac-365-ja_JP').ok);
  assert.equal(wb.excelProfileId().value, 'mac-365-ja_JP');
});

// Values are the mac-365-* oracle captures of `DOLLAR(-1234.567)`; the
// estimated win-* ids are checked for the id round trip only.
const PROFILE_DOLLAR = {
  'mac-365-ja_JP': '¥-1,235',
  'mac-365-en_US': '($1,234.57)',
  'mac-365-de_DE': '-1.234,57 €',
  'mac-365-fr_FR': '(1 234,57 €)',
  'mac-365-zh_CN': '(¥1,234.57)',
  'mac-365-ko_KR': '(₩1,235)',
  'mac-365-th_TH': '(฿1,234.57)',
};
const PROFILE_LOCALES = ['ja_JP', 'en_US', 'de_DE', 'fr_FR', 'zh_CN', 'ko_KR', 'th_TH'];

test('setExcelProfileId switches a live workbook through every profile id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.excelProfileId().value, 'win-365-en_US');
  assert.ok(wb.setFormula(0, 0, 0, '=DOLLAR(-1234.567)').ok);
  for (const host of ['mac', 'win']) {
    for (const locale of PROFILE_LOCALES) {
      const id = `${host}-365-${locale}`;
      assert.ok(wb.setExcelProfileId(id).ok, id);
      assert.equal(wb.excelProfileId().value, id);
      assert.ok(wb.recalc().ok, id);
      if (id in PROFILE_DOLLAR) {
        assert.equal(wb.getValue(0, 0, 0).value.text, PROFILE_DOLLAR[id], id);
      }
    }
  }
  assert.equal(wb.setExcelProfileId('mac-365-xx_XX').ok, false);
});

test('functionNames + functionMetadata expose the catalog', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const names = wb.functionNames();
  assert.ok(Array.isArray(names));
  assert.ok(names.length > 0, 'expected a non-empty function catalog');
  assert.ok(names.includes('SUM'), 'expected SUM in the catalog');
  const md = wb.functionMetadata('SUM');
  assert.equal(md.ok, true);
  assert.equal(md.name, 'SUM');
  assert.equal(typeof md.minArity, 'number');
  // Unknown function returns { ok: false }.
  const miss = wb.functionMetadata('NOT_A_REAL_FUNCTION_XYZ');
  assert.equal(miss.ok, false);
});

test('mergeFunctionMetadata overlays provider metadata; is identity without a provider', async () => {
  const mod = await getModule();
  assert.equal(typeof mod.mergeFunctionMetadata, 'function');
  const base = mod.Workbook.createDefault().functionMetadata('XLOOKUP');
  assert.equal(base.ok, true);
  // The engine leaves display metadata empty.
  assert.equal(base.signatureTemplate, undefined);
  assert.equal(base.description, undefined);

  const entry = {
    signature: 'XLOOKUP(lookup_value, lookup_array, return_array)',
    description: 'Searches a range or an array.',
    aliases: { 'fr-FR': 'RECHERCHEX' },
    localized: {
      'fr-FR': { signature: 'RECHERCHEX(...)', description: 'Recherche.' },
    },
  };

  // Localized override wins for the matching locale.
  const fr = mod.mergeFunctionMetadata(base, entry, 'fr-FR');
  assert.equal(fr.signatureTemplate, 'RECHERCHEX(...)');
  assert.equal(fr.description, 'Recherche.');
  assert.equal(fr.localizedName, 'RECHERCHEX');
  // Structural fields survive the merge.
  assert.equal(fr.ok, true);
  assert.equal(fr.name, 'XLOOKUP');

  // A locale with no localized/alias entry falls back to the default
  // signature/description and the canonical display name.
  const de = mod.mergeFunctionMetadata(base, entry, 'de-DE');
  assert.equal(de.signatureTemplate, 'XLOOKUP(lookup_value, lookup_array, return_array)');
  assert.equal(de.description, 'Searches a range or an array.');
  assert.equal(de.localizedName, 'XLOOKUP');

  // No provider entry -> base returned verbatim, display metadata stays NULL.
  const none = mod.mergeFunctionMetadata(base, undefined, 'fr-FR');
  assert.equal(none, base);
  assert.equal(none.signatureTemplate, undefined);
  assert.equal(none.description, undefined);
});

test('localizeFunctionName / canonicalizeFunctionName round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // en-US profile: the localized name is the canonical name itself.
  assert.equal(wb.localizeFunctionName('SUM', 'win-365-en_US').value, 'SUM');
  assert.equal(wb.canonicalizeFunctionName('SUM', 'win-365-en_US').value, 'SUM');
  // An unknown name returns the empty string.
  assert.equal(wb.canonicalizeFunctionName('NOPE_XYZ', 'win-365-en_US').value, '');
});

test('formula and function-name localization is keyed by profile id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.equal(wb.localizeFormula('=SUM(1.5,2)', 'mac-365-de_DE').value, '=SUMME(1,5;2)');
    assert.equal(wb.canonicalizeFormula('=SUMME(1,5;2)', 'mac-365-de_DE').value, '=SUM(1.5,2)');
    assert.equal(wb.localizeFormula('=IF(TRUE,1,2)', 'mac-365-fr_FR').value, '=SI(VRAI;1;2)');
    assert.equal(wb.localizeFormula('=#N/A', 'mac-365-de_DE').value, '=#NV');
    assert.equal(wb.canonicalizeFormula('=#NV', 'mac-365-de_DE').value, '=#N/A');
    assert.equal(wb.localizeFormula('=A1&";"', 'mac-365-de_DE').value, '=A1&";"');
    assert.equal(wb.localizeFunctionName('SUM', 'mac-365-de_DE').value, 'SUMME');
    assert.equal(wb.canonicalizeFunctionName('SUMME', 'win-365-de_DE').value, 'SUM');
    assert.equal(wb.localizeFormula('=SUM(1)', 'nope').status.ok, false);
    assert.equal(wb.canonicalizeFormula('=SUM(1)', 'nope').status.ok, false);
    assert.equal(wb.canonicalizeFunctionName('SUM', 'nope').status.ok, false);
  } finally {
    wb.dispose();
  }
});

test('localeFacts reports separators, names, error spellings and measurement', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const mac = wb.localeFacts('mac-365-de_DE');
    assert.ok(mac.status.ok);
    assert.equal(mac.facts.decimalSeparator, ',');
    assert.equal(mac.facts.listSeparator, ';');
    assert.equal(mac.facts.dateOrder, 'dmy');
    assert.equal(mac.facts.measured, true);
    const na = mac.facts.errorNames.find((e) => e.canonical === '#N/A');
    assert.equal(na.localized, '#NV');
    assert.equal(wb.localeFacts('win-365-de_DE').facts.measured, false);
    assert.equal(wb.localeFacts('win-365-en_US').facts.decimalSeparator, '.');
    const bad = wb.localeFacts('nope');
    assert.equal(bad.status.ok, false);
    assert.equal(bad.facts, null);
  } finally {
    wb.dispose();
  }
});

test('precedents / dependents return arrays for a small formula graph', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setNumber(0, 0, 0, 1).ok); // A1
  assert.ok(wb.setNumber(0, 0, 1, 2).ok); // B1
  assert.ok(wb.setFormula(0, 1, 0, '=A1+B1').ok); // A2 = A1 + B1
  assert.ok(wb.recalc().ok);
  // A2 reads A1 and B1.
  const prec = wb.precedents(0, 1, 0, 1);
  assert.ok(Array.isArray(prec));
  assert.equal(prec.length, 2);
  for (const node of prec) {
    assert.equal(typeof node.sheet, 'number');
    assert.equal(typeof node.row, 'number');
    assert.equal(typeof node.col, 'number');
  }
  // A1 is read by A2.
  const dep = wb.dependents(0, 0, 0, 1);
  assert.ok(Array.isArray(dep));
  assert.equal(dep.length, 1);
  assert.equal(dep[0].row, 1);
  assert.equal(dep[0].col, 0);
});

test('spillInfo reports a dynamic-array region', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Non-spill cell: engaged is false, fields zeroed.
  const none = wb.spillInfo(0, 0, 0);
  assert.equal(typeof none.engaged, 'boolean');
  assert.equal(none.engaged, false);
  // A spilling dynamic array anchored at A1.
  assert.ok(wb.setFormula(0, 0, 0, '=SEQUENCE(3,1)').ok);
  assert.ok(wb.recalc().ok);
  const info = wb.spillInfo(0, 0, 0);
  assert.equal(info.engaged, true);
  assert.equal(info.anchorRow, 0);
  assert.equal(info.anchorCol, 0);
  assert.equal(info.rows, 3);
  assert.equal(info.cols, 1);
});

test('getLambdaText is wired and reports the documented contract', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // A bare top-level LAMBDA renders as #CALC! in Excel 365 and does not
  // retain a renderable closure; the engine mirrors that, so the lambda
  // text surface returns an error status (not a throw). A plain number
  // cell is likewise not a lambda. Both prove the method marshals
  // arguments and propagates the Expected failure as a JS error status.
  assert.ok(wb.setFormula(0, 0, 0, '=LAMBDA(x,x+1)').ok);
  assert.ok(wb.recalc().ok);
  const bare = wb.getLambdaText(0, 0, 0);
  assert.equal(bare.status.ok, false);
  assert.equal(typeof bare.text, 'string');
  assert.equal(bare.text, '');

  assert.ok(wb.setNumber(0, 1, 0, 5).ok);
  const num = wb.getLambdaText(0, 1, 0);
  assert.equal(num.status.ok, false);
});

test('getExternalLinks returns an array (empty for a fresh workbook)', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const links = wb.getExternalLinks();
  assert.ok(Array.isArray(links));
  assert.equal(links.length, 0);
});

test('getSheetProtection / setSheetProtection round-trip the protection flags', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const before = wb.getSheetProtection(0);
  assert.ok(before.status.ok, `getSheetProtection: ${JSON.stringify(before.status)}`);
  assert.equal(before.protection.enabled, 0);
  // Enable protection and lock cell selection.
  const next = {
    ...before.protection,
    enabled: 1,
    selectLockedCells: 1,
    sort: 1,
  };
  assert.ok(wb.setSheetProtection(0, next).ok);
  const after = wb.getSheetProtection(0);
  assert.ok(after.status.ok);
  assert.equal(after.protection.enabled, 1);
  assert.equal(after.protection.selectLockedCells, 1);
  assert.equal(after.protection.sort, 1);
  // Sheet-out-of-range is rejected.
  assert.ok(!wb.getSheetProtection(999).status.ok);
  assert.ok(!wb.setSheetProtection(999, next).ok);
});
