import test from 'node:test';
import { appendEmptyZipEntry, assert, getModule, readAll } from './smoke_support.mjs';

test('comments round-trip: setComment + getComment', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Pre-existing cell so the comment attaches to a stored slot.
  assert.ok(wb.setNumber(0, 1, 1, 7).ok);
  const sc = wb.setComment(0, 1, 1, 'libraz', 'hello world');
  assert.ok(sc.ok, `setComment: ${JSON.stringify(sc)}`);
  const c = wb.getComment(0, 1, 1);
  assert.ok(c !== null, 'expected non-null Comment');
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
  const n = wb.getCellCount(0).value;
  assert.equal(typeof n, 'number');
  assert.ok(n >= 3, `expected >= 3 stored cells, got ${n}`);
  // Iterate; one of the entries should carry our formula.
  let foundFormula = false;
  for (let i = 0; i < n; ++i) {
    const e = wb.getCellAt(0, i);
    assert.ok(e.status.ok, `getCellAt(${i}): ${JSON.stringify(e.status)}`);
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

test('default export exposes ErrorCode ordinals and getVersion', async () => {
  const mod = await getModule();
  assert.equal(typeof mod.ErrorCode, 'object');
  assert.equal(mod.ErrorCode.Div0, 1);
  assert.equal(mod.ErrorCode.Ref, 3);
  assert.equal(mod.ErrorCode.NA, 6);
  assert.equal(typeof mod.getVersion, 'function');
  assert.ok(mod.getVersion().length > 0);
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
    }).status.ok,
  );
  // A CF rule's dxfId must resolve against a registered dxf.
  const dxf = wb.addDxf({ font: { bold: true } });
  assert.ok(dxf.status.ok, JSON.stringify(dxf.status));
  assert.ok(
    wb.addCfRule(0, {
      ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
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

  const loaded = mod.Workbook.load(xlsb.bytes);
  const read = loaded.getDiagnostics();
  assert.ok(read.status.ok, JSON.stringify(read.status));
  assert.equal(read.undecodedFormulaCount, 0);
  assert.equal(read.undecodedDefinedNameCount, 0);
  assert.equal(read.undecodedPartCount, 0);
  assert.equal(read.skippedFeatureCount, 0);
  assert.equal(read.unknownContentTypeCount, 0);
  const preservedLoaded = mod.Workbook.load(appendEmptyZipEntry(xlsb.bytes, 'xl/preserved.bin'));
  const preserved = preservedLoaded.getDiagnostics();
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
  const wb = mod.Workbook.load(new Uint8Array());
  const read = wb.getDiagnostics();
  assert.ok(!read.status.ok, JSON.stringify(read));
  assert.equal(read.undecodedFormulaCount, 0);
  assert.equal(read.undecodedDefinedNameCount, 0);
  assert.equal(read.undecodedPartCount, 0);
  assert.equal(read.skippedFeatureCount, 0);
  assert.equal(read.unknownContentTypeCount, 0);
  wb.dispose();
});

test('addValidation / getValidation / removeValidationAt / clearValidations round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Empty by default.
  assert.equal(
    readAll(
      () => wb.getValidationCount(0),
      (i) => wb.getValidation(0, i),
    ).length,
    0,
  );

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
  assert.ok(wb.addValidation(0, listRule).status.ok);

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
  assert.ok(wb.addValidation(0, decimalRule).status.ok);

  const list = readAll(
    () => wb.getValidationCount(0),
    (i) => wb.getValidation(0, i),
  );
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
  assert.ok(wb.removeValidation(0, 0).ok);
  const after = readAll(
    () => wb.getValidationCount(0),
    (i) => wb.getValidation(0, i),
  );
  assert.equal(after.length, 1);
  assert.equal(after[0].type, 2);

  // Out-of-range index is rejected.
  assert.ok(!wb.removeValidation(0, 99).ok);
  // Sheet-out-of-range is rejected on every entry.
  assert.ok(!wb.addValidation(999, listRule).status.ok);
  assert.ok(!wb.removeValidation(999, 0).ok);
  assert.ok(!wb.clearValidations(999).ok);

  // clearValidations drops everything; safe to call again.
  assert.ok(wb.clearValidations(0).ok);
  assert.equal(
    readAll(
      () => wb.getValidationCount(0),
      (i) => wb.getValidation(0, i),
    ).length,
    0,
  );
  assert.ok(wb.clearValidations(0).ok);
});

// -- Newly wired surface (audit parity with the WASM binding) ----------
// These tests exercise the methods that were previously only reachable
// from the WASM package. They assert sensible return shapes rather than
// deep semantics; the point is "the addon wires them and does not throw".

test('addConditionalFormat + getCfRule + clearConditionalFormats round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(
    readAll(
      () => wb.getCfRuleCount(0),
      (i) => wb.getCfRule(0, i),
    ).length,
    0,
  );
  // A CF rule's dxfId must resolve against a registered dxf.
  const dxf = wb.addDxf({ font: { bold: true } });
  assert.ok(dxf.status.ok, JSON.stringify(dxf.status));
  // cellIs rule (type 1) comparing the cell to a literal.
  const add = wb.addCfRule(0, {
    ranges: [{ firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 0 }],
    type: 1,
    op: 5,
    formula1: '10',
    dxfId: dxf.index,
    stopIfTrue: false,
  });
  assert.ok(add.status.ok, `addConditionalFormat: ${JSON.stringify(add)}`);
  assert.equal(typeof add.index, 'number');
  const list = readAll(
    () => wb.getCfRuleCount(0),
    (i) => wb.getCfRule(0, i),
  );
  assert.equal(list.length, 1);
  assert.equal(list[0].type, 1);
  assert.equal(typeof list[0].id, 'string');
  assert.equal(list[0].ranges.length, 1);
  assert.equal(list[0].ranges[0].lastRow, 4);
  assert.equal(list[0].formula1, '10');
  // removeConditionalFormatAt drops the only rule.
  assert.ok(wb.removeCfRule(0, 0).ok);
  assert.equal(
    readAll(
      () => wb.getCfRuleCount(0),
      (i) => wb.getCfRule(0, i),
    ).length,
    0,
  );
  // clearConditionalFormats is a safe no-op on an empty sheet.
  assert.ok(wb.clearCfRules(0).ok);
  // Sheet-out-of-range is rejected.
  assert.ok(!wb.clearCfRules(999).ok);
});

test('data bar x14 fields survive save and load', async () => {
  // The six live in the `x14` extension rather than the legacy `<dataBar>`
  // element, so an in-session round-trip would still pass with a writer
  // that never emits the extension -- which is exactly how they used to be
  // lost. Only a save/load cycle pins them.
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const add = wb.addCfRule(0, {
    ranges: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 0 }],
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
  const loaded = mod.Workbook.load(saved.bytes);
  const bar = readAll(
    () => loaded.getCfRuleCount(0),
    (i) => loaded.getCfRule(0, i),
  )[0].dataBar;
  assert.equal(bar.gradient, false);
  assert.equal(bar.axisPosition, 1);
  assert.equal(bar.negativeFill.r, 255);
  assert.equal(bar.border.r, 9);
  assert.equal(bar.negativeBorder.r, 8);
  assert.equal(bar.axisColor.b, 3);
  assert.equal(bar.direction, 2);

  // The decoded bar fed straight back must reproduce the same rule.
  const again = loaded.addCfRule(0, {
    ranges: [{ firstRow: 4, firstCol: 0, lastRow: 6, lastCol: 0 }],
    type: 3,
    dataBar: bar,
  });
  assert.ok(again.status.ok, `re-add: ${JSON.stringify(again)}`);
  const reread = readAll(
    () => loaded.getCfRuleCount(0),
    (i) => loaded.getCfRule(0, i),
  )[1].dataBar;
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
  const add = wb.addCfRule(0, {
    ranges: [{ firstRow: 0, firstCol: 0, lastRow: 9, lastCol: 0 }],
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
  const loaded = mod.Workbook.load(saved.bytes);
  const floor = readAll(
    () => loaded.getCfRuleCount(0),
    (i) => loaded.getCfRule(0, i),
  )[0].iconSet.floor;
  assert.equal(floor.type, 0);
  assert.equal(floor.value, '5');
  assert.equal(floor.gte, false);
  loaded.dispose();
  wb.dispose();
});

test('omitted data bar x14 fields keep the model defaults', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const add = wb.addCfRule(0, {
    ranges: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 0 }],
    type: 3,
    dataBar: { min: { type: 3 }, max: { type: 4 }, fill: { r: 0, g: 0, b: 255 } },
  });
  assert.ok(add.status.ok, `addConditionalFormat: ${JSON.stringify(add)}`);
  const bar = readAll(
    () => wb.getCfRuleCount(0),
    (i) => wb.getCfRule(0, i),
  )[0].dataBar;
  // The getter engages all six, so they read back as the defaults rather
  // than as absent keys.
  assert.equal(bar.gradient, true);
  assert.equal(bar.axisPosition, 0);
  assert.equal(bar.negativeFill.b, 255);
  wb.dispose();
});

test('named cell styles: listCellStyles / getCellStyle / getCellStyleXfCount', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // A fresh workbook may carry zero named styles; the accessors must
  // still return well-formed values rather than throwing.
  const count = wb.listCellStyles().length;
  assert.equal(typeof count, 'number');
  assert.ok(count >= 0);
  assert.equal(typeof wb.getCellStyleXfCount().value, 'number');
  if (count > 0) {
    const cs = wb.listCellStyles()[0];
    assert.equal(typeof cs.name, 'string');
    assert.equal(typeof cs.xfId, 'number');
    assert.deepEqual(wb.getCellStyle(cs.name).value, cs);
  }
  // An unknown name is ok with a null value, not a throw.
  const none = wb.getCellStyle('No Such Style');
  assert.ok(none.status.ok);
  assert.equal(none.value, null);
});

test('setCalcMode / calcMode round-trip the calc policy', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Default is automatic (0).
  assert.equal(wb.getCalcMode().value, 0);
  assert.ok(wb.setCalcMode(1).ok);
  assert.equal(wb.getCalcMode().value, 1);
  assert.ok(wb.setCalcMode(2).ok);
  assert.equal(wb.getCalcMode().value, 2);
});

test('setPinnedNow / pinnedNow / clearPinnedNow drive the clock seam', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Unpinned by default: the workbook follows the host clock.
  assert.equal(wb.getPinnedNow().now, null);

  assert.ok(wb.setPinnedNow(2026, 4, 23, 15, 30, 45).ok);
  assert.deepEqual(wb.getPinnedNow().now, {
    year: 2026,
    month: 4,
    day: 23,
    hour: 15,
    minute: 30,
    second: 45,
  });

  // 2026-04-23 is serial 46135 under the 1900 date system.
  const today = wb.evaluateFormula(0, 0, 0, '=TODAY()');
  assert.ok(today.status.ok, `evaluateFormulaText: ${JSON.stringify(today.status)}`);
  assert.equal(today.value.number, 46135);

  // The pin is a calendar instant, not a normalising constructor.
  assert.equal(wb.setPinnedNow(2026, 13, 1, 0, 0, 0).ok, false);
  assert.equal(wb.setPinnedNow(2025, 2, 29, 0, 0, 0).ok, false);

  assert.ok(wb.clearPinnedNow().ok);
  assert.equal(wb.getPinnedNow().now, null);
});

test('excelProfileId / setExcelProfileId round-trip the profile id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const def = wb.getExcelProfileId().value;
  assert.equal(typeof def, 'string');
  assert.ok(def.length > 0);
  assert.ok(wb.setExcelProfileId('mac-365-ja_JP').ok);
  assert.equal(wb.getExcelProfileId().value, 'mac-365-ja_JP');
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
  'mac-365-es_ES': '-1.234,57 €',
  'mac-365-es_MX': '-$1,234.57',
  'mac-365-pt_BR': '-R$ 1.234,57',
  'mac-365-ru_RU': '-1 234,57 ₽',
  'mac-365-zh_TW': '($1,234.57)',
  'mac-365-it_IT': '-1.234,57 €',
  'mac-365-nl_NL': '(€ 1.234,57)',
};
const PROFILE_LOCALES = [
  'ja_JP',
  'en_US',
  'de_DE',
  'fr_FR',
  'zh_CN',
  'ko_KR',
  'th_TH',
  'es_ES',
  'es_MX',
  'pt_BR',
  'ru_RU',
  'zh_TW',
  'it_IT',
  'nl_NL',
];

test('setExcelProfileId switches a live workbook through every profile id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.getExcelProfileId().value, 'win-365-en_US');
  assert.ok(wb.setFormula(0, 0, 0, '=DOLLAR(-1234.567)').ok);
  for (const host of ['mac', 'win']) {
    for (const locale of PROFILE_LOCALES) {
      const id = `${host}-365-${locale}`;
      assert.ok(wb.setExcelProfileId(id).ok, id);
      assert.equal(wb.getExcelProfileId().value, id);
      assert.ok(wb.recalc().ok, id);
      if (id in PROFILE_DOLLAR) {
        assert.equal(wb.getValue(0, 0, 0).value.text, PROFILE_DOLLAR[id], id);
      }
    }
  }
  assert.equal(wb.setExcelProfileId('mac-365-xx_XX').ok, false);
});

test('functionNames + getFunctionMetadata expose the catalog', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const names = wb.listFunctionNames();
  assert.ok(Array.isArray(names));
  assert.ok(names.length > 0, 'expected a non-empty function catalog');
  assert.ok(names.includes('SUM'), 'expected SUM in the catalog');
  const md = wb.getFunctionMetadata('SUM');
  assert.equal(md.ok, true);
  assert.equal(md.name, 'SUM');
  assert.equal(typeof md.minArity, 'number');
  // Unknown function returns { ok: false }.
  const miss = wb.getFunctionMetadata('NOT_A_REAL_FUNCTION_XYZ');
  assert.equal(miss.ok, false);
});

test('mergeFunctionMetadata overlays provider metadata; is identity without a provider', async () => {
  const mod = await getModule();
  assert.equal(typeof mod.mergeFunctionMetadata, 'function');
  const base = mod.Workbook.createDefault().getFunctionMetadata('XLOOKUP');
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

test('getLocaleFacts reports separators, names, error spellings and measurement', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const mac = wb.getLocaleFacts('mac-365-de_DE');
    assert.ok(mac.status.ok);
    assert.equal(mac.facts.decimalSeparator, ',');
    assert.equal(mac.facts.listSeparator, ';');
    assert.equal(mac.facts.dateOrder, 'dmy');
    assert.equal(mac.facts.measured, true);
    assert.equal(mac.facts.collationMeasured, true);
    const na = mac.facts.errorNames.find((e) => e.canonical === '#N/A');
    assert.equal(na.localized, '#NV');
    const listed = wb.listLocaleErrorNames('mac-365-de_DE');
    assert.ok(listed.status.ok);
    assert.deepEqual([...listed], mac.facts.errorNames);
    assert.equal(wb.listLocaleErrorNames('nope').status.ok, false);
    assert.equal(wb.getLocaleFacts('win-365-de_DE').facts.measured, false);
    assert.equal(wb.getLocaleFacts('win-365-ja_JP').facts.collationMeasured, false);
    assert.equal(wb.getLocaleFacts('win-365-en_US').facts.decimalSeparator, '.');
    const bad = wb.getLocaleFacts('nope');
    assert.equal(bad.status.ok, false);
    assert.equal(bad.facts, null);
  } finally {
    wb.dispose();
  }
});

test('compareText folds by strength and rejects bad arguments', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const profile = 'mac-365-zh_CN';
    const strict = wb.compareText('A', '\uFF21', 0, profile);
    assert.ok(strict.status.ok);
    assert.notEqual(strict.order, 0);
    const dynamic = wb.compareText('A', '\uFF21', 1, profile);
    assert.ok(dynamic.status.ok);
    assert.equal(dynamic.order, 0);
    const ab = wb.compareText('a', 'b', 0, 'mac-365-en_US');
    const ba = wb.compareText('b', 'a', 0, 'mac-365-en_US');
    assert.equal(ab.order, -1);
    assert.equal(ba.order, 1);
    const bad = wb.compareText('a', 'b', 0, 'nope');
    assert.equal(bad.status.ok, false);
    assert.equal(bad.order, 0);
    assert.throws(() => wb.compareText('a', 'b', -1, profile), RangeError);
  } finally {
    wb.dispose();
  }
});

test('getCollationClasses orders by kind, folds text by strength, and copies its output', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const values = [
      { kind: 1, number: 1 },
      { kind: 3, text: '1' },
      { kind: 2, boolean: 1 },
      { kind: 3, text: 'TRUE' },
      { kind: 0 },
      { kind: 3, text: '' },
      { kind: 1, number: 0 },
      { kind: 2, boolean: 0 },
    ];
    const r = wb.getCollationClasses(values, 1, 'mac-365-en_US');
    assert.ok(r.status.ok);
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
    const other = wb.getCollationClasses(
      [
        { kind: 3, text: 'z' },
        { kind: 3, text: 'y' },
      ],
      1,
      'mac-365-en_US',
    );
    assert.ok(other.status.ok);
    assert.deepEqual(Array.from(r.classIds), snapshot);

    const folded = wb.getCollationClasses(
      [
        { kind: 3, text: 'a' },
        { kind: 3, text: 'A' },
      ],
      1,
      'mac-365-en_US',
    );
    assert.equal(folded.classIds[0], folded.classIds[1]);
    assert.equal(folded.classCount, 1);

    const empty = wb.getCollationClasses([], 1, 'mac-365-en_US');
    assert.ok(empty.status.ok);
    assert.equal(empty.classIds.length, 0);
    assert.equal(empty.classCount, 0);

    const noKind = wb.getCollationClasses([{ number: 1 }], 1, 'mac-365-en_US');
    assert.equal(noKind.status.ok, false);
    assert.equal(noKind.classIds.length, 0);
    assert.equal(noKind.classCount, 0);
    assert.equal(wb.getCollationClasses({}, 1, 'mac-365-en_US').status.ok, false);
    assert.equal(wb.getCollationClasses([{ kind: 1, number: 1 }], 1, 'nope').status.ok, false);
    assert.throws(() => wb.getCollationClasses(values, -1, 'mac-365-en_US'), RangeError);
  } finally {
    wb.dispose();
  }
});

test('getCollationClasses timing on 100k mixed CJK/ASCII texts', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
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
    const r = wb.getCollationClasses(values, 1, 'mac-365-ja_JP');
    const ms = Number(process.hrtime.bigint() - t0) / 1e6;
    assert.ok(r.status.ok);
    assert.equal(r.classIds.length, 100000);
    console.log(`getCollationClasses 100k mixed CJK/ASCII texts: ${ms.toFixed(0)} ms (${r.classCount} classes)`);
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
  const prec = wb.getPrecedents(0, 1, 0, 1);
  assert.ok(Array.isArray(prec));
  assert.equal(prec.length, 2);
  for (const node of prec) {
    assert.equal(typeof node.sheet, 'number');
    assert.equal(typeof node.row, 'number');
    assert.equal(typeof node.col, 'number');
  }
  // A1 is read by A2.
  const dep = wb.getDependents(0, 0, 0, 1);
  assert.ok(Array.isArray(dep));
  assert.equal(dep.length, 1);
  assert.equal(dep[0].row, 1);
  assert.equal(dep[0].col, 0);
});

test('spillInfo reports a dynamic-array region', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Non-spill cell: engaged is false, fields zeroed.
  const none = wb.getSpillInfo(0, 0, 0);
  assert.equal(typeof none.engaged, 'boolean');
  assert.equal(none.engaged, false);
  // A spilling dynamic array anchored at A1.
  assert.ok(wb.setFormula(0, 0, 0, '=SEQUENCE(3,1)').ok);
  assert.ok(wb.recalc().ok);
  const info = wb.getSpillInfo(0, 0, 0);
  assert.equal(info.engaged, true);
  assert.deepEqual(info.rect, { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 0 });
});

test('array formulas can be authored, read from every cell and cleared', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  for (let r = 0; r < 3; r += 1) assert.ok(wb.setNumber(0, r, 0, r + 1).ok);
  const none = wb.getArrayFormula(0, 0, 2);
  assert.ok(none.status.ok);
  assert.equal(none.arrayFormula, null);
  assert.ok(wb.setArrayFormula(0, { firstRow: 0, firstCol: 2, lastRow: 2, lastCol: 2 }, '=A1:A3*2').ok);
  assert.ok(wb.recalc().ok);
  for (let r = 0; r < 3; r += 1) {
    const found = wb.getArrayFormula(0, r, 2);
    assert.ok(found.status.ok);
    assert.deepEqual(found.arrayFormula.rect, { firstRow: 0, firstCol: 2, lastRow: 2, lastCol: 2 });
    assert.ok(found.arrayFormula.formula.includes('A1:A3'));
  }
  assert.equal(wb.getValue(0, 2, 2).value.number, 6);
  // A partial overlap and a write into a non-anchor cell are refused.
  assert.equal(wb.setArrayFormula(0, { firstRow: 1, firstCol: 2, lastRow: 2, lastCol: 2 }, '=1').ok, false);
  assert.equal(wb.setNumber(0, 1, 2, 1).ok, false);
  assert.ok(wb.setBlank(0, 0, 2).ok);
  assert.ok(wb.recalc().ok);
  assert.equal(wb.getArrayFormula(0, 1, 2).arrayFormula, null);
  assert.equal(wb.getArrayFormula(99, 0, 0).status.ok, false);
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

test('a fresh workbook has no external links', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.deepEqual([wb.getExternalLinkCount().status.ok, wb.getExternalLinkCount().value], [true, 0]);
  assert.equal(wb.getExternalLink(0).status.ok, false);
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
