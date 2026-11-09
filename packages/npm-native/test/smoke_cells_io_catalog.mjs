import test from 'node:test';
import { appendEmptyZipEntry, assert, getModule, isFormulonError } from './smoke_support.mjs';

test('comments round-trip: setComment + getComment', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Pre-existing cell so the comment attaches to a stored slot.
  wb.sheet(0).setNumber(1, 1, 7);
  wb.sheet(0).setComment(1, 1, { author: 'libraz', text: 'hello world' });
  const c = wb.sheet(0).getComment(1, 1);
  assert.ok(c !== null, 'expected non-null Comment');
  assert.equal(c.author, 'libraz');
  assert.equal(c.text, 'hello world');
  // Removing surfaces null on the next read.
  wb.sheet(0).removeComment(1, 1);
  assert.equal(wb.sheet(0).getComment(1, 1), null);
});

test('getComment distinguishes absence from an invalid sheet', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const missing = wb.sheet(0).getComment(1, 1);
  assert.equal(missing, null);
  assert.throws(
    () => wb.sheet(99).getComment(1, 1),
    (e) => isFormulonError(e, mod.StatusCode.SheetIndexOutOfRange),
  );
  wb.sheet(0).setComment(1, 1, { author: 'libraz', text: 'hello' });
  const found = wb.sheet(0).getComment(1, 1);
  assert.equal(found.author, 'libraz');
  assert.equal(found.text, 'hello');
});

test('listCells enumerates stored cells', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setNumber(0, 0, 1);
  wb.sheet(0).setNumber(0, 1, 2);
  wb.sheet(0).setFormula(1, 0, '=A1+B1');
  wb.recalc();
  const cells = wb.sheet(0).listCells();
  assert.ok(cells.length >= 3, `expected >= 3 stored cells, got ${cells.length}`);
  // One of the entries should carry our formula.
  let foundFormula = false;
  for (const e of cells) {
    assert.equal(typeof e.row, 'number');
    assert.equal(typeof e.col, 'number');
    if (e.formula === '=A1+B1') {
      foundFormula = true;
      assert.equal(e.value, 3);
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
  assert.throws(() => wb.sheet(0).setError(0, 0), TypeError);
});

test('setError() rejects a non-number errorCode before mutating the cell', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setNumber(0, 0, 7);
  assert.throws(() => wb.sheet(0).setError(0, 0, 'not a number'), TypeError);
  // The rejected call must not have overwritten the cell.
  const got = wb.sheet(0).getTypedValue(0, 0);
  assert.equal(got.number, 7);
});

test('save() writes xlsx and xlsb bytes; a missing format selects the default', async () => {
  const mod = await getModule();
  assert.equal(typeof mod.WorkbookFormat, 'object');
  const wb = mod.Workbook.createDefault();
  const xlsx = wb.save(mod.WorkbookFormat.Xlsx);
  assert.ok(xlsx.length > 0);
  const xlsb = wb.save(mod.WorkbookFormat.Xlsb);
  assert.ok(xlsb.length > 0);
  assert.ok(wb.save().length > 0);
});

test('saveWithDiagnostics() reports counters and readDiagnostics keeps a stable shape', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setFormula(0, 0, '=SUM(T[C])');
  wb.sheet(0).addValidation({
    ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
    type: 3,
    formula1: '"Yes,No"',
  });
  // A CF rule's dxfIndex must resolve against a registered dxf.
  const dxf = wb.ooxml.addDxf({ font: { bold: true } });
  wb.sheet(0).addCfRule({
    ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
    type: 1,
    operator: 5,
    formula1: '10',
    dxfIndex: dxf,
    stopIfTrue: false,
  });

  const xlsx = wb.saveWithDiagnostics(mod.WorkbookFormat.Xlsx);
  assert.ok(xlsx.bytes instanceof Uint8Array);
  assert.equal(xlsx.downgradedFormulaCount, 0);
  assert.equal(xlsx.deferredFeatureCount, 0);
  assert.equal(xlsx.droppedPartCount, 0);
  assert.equal(xlsx.droppedRelationshipCount, 0);
  assert.equal(xlsx.renumberedPartCount, 0);

  const xlsb = wb.saveWithDiagnostics(mod.WorkbookFormat.Xlsb);
  assert.ok(xlsb.bytes instanceof Uint8Array);
  assert.equal(xlsb.downgradedFormulaCount, 1);
  // The validation and the CF rule, with its dxf, are both written.
  assert.equal(xlsb.deferredFeatureCount, 0);
  // The binary writer never reassigns a part id.
  assert.equal(xlsb.renumberedPartCount, 0);

  const loaded = mod.Workbook.load(xlsb.bytes);
  const read = loaded.getDiagnostics();
  assert.equal(read.undecodedFormulaCount, 0);
  assert.equal(read.undecodedDefinedNameCount, 0);
  assert.equal(read.undecodedPartCount, 0);
  assert.equal(read.skippedFeatureCount, 0);
  assert.equal(read.unknownContentTypeCount, 0);
  const preservedLoaded = mod.Workbook.load(appendEmptyZipEntry(xlsb.bytes, 'xl/preserved.bin'));
  const preserved = preservedLoaded.getDiagnostics();
  assert.equal(preserved.undecodedPartCount, 0);
  preservedLoaded.dispose();
  assert.throws(() => wb.saveWithDiagnostics(0), isFormulonError);
  assert.ok(wb.save(mod.WorkbookFormat.Xlsb) instanceof Uint8Array);
});

test('read diagnostics of a disposed workbook throw', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.dispose();
  assert.throws(
    () => wb.getDiagnostics(),
    (e) => isFormulonError(e, mod.StatusCode.BindingInvalidHandle),
  );
});

test('addValidation / getValidation / removeValidationAt / clearValidations round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Empty by default.
  assert.equal(wb.sheet(0).listValidations().length, 0);

  // List-type validation at A1:B3.
  const listRule = {
    ranges: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1 }],
    type: 3,
    operator: 0,
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
  wb.sheet(0).addValidation(listRule);

  // Decimal between [0, 100] across two rectangles.
  const decimalRule = {
    ranges: [
      { firstRow: 4, firstCol: 0, lastRow: 4, lastCol: 0 },
      { firstRow: 6, firstCol: 0, lastRow: 9, lastCol: 0 },
    ],
    type: 2,
    operator: 0,
    formula1: '0',
    formula2: '100',
    allowBlank: false,
  };
  wb.sheet(0).addValidation(decimalRule);

  const list = wb.sheet(0).listValidations();
  assert.equal(list.length, 2);
  // Booleans must arrive as JS booleans, not 0/1.
  assert.equal(typeof list[0].allowBlank, 'boolean');
  assert.equal(list[0].allowBlank, true);
  assert.equal(typeof list[0].showInputMessage, 'boolean');
  assert.equal(list[0].showInputMessage, true);
  assert.equal(typeof list[1].allowBlank, 'boolean');
  assert.equal(list[1].allowBlank, false);
  assert.equal(list[0].formula1, '="Yes,No,Maybe"');
  assert.equal(list[1].formula1, '=0');
  assert.equal(list[1].formula2, '=100');
  assert.equal(list[1].ranges.length, 2);
  assert.equal(list[1].ranges[1].firstRow, 6);

  // removeValidationAt removes the first rule.
  wb.sheet(0).removeValidation(0);
  const after = wb.sheet(0).listValidations();
  assert.equal(after.length, 1);
  assert.equal(after[0].type, 2);

  // Out-of-range index is rejected.
  assert.throws(() => wb.sheet(0).removeValidation(99), isFormulonError);
  // Sheet-out-of-range is rejected on every entry.
  assert.throws(() => wb.sheet(999).addValidation(listRule), isFormulonError);
  assert.throws(() => wb.sheet(999).removeValidation(0), isFormulonError);
  assert.throws(() => wb.sheet(999).clearValidations(), isFormulonError);

  // clearValidations drops everything; safe to call again.
  wb.sheet(0).clearValidations();
  assert.equal(wb.sheet(0).listValidations().length, 0);
  wb.sheet(0).clearValidations();
});

// -- Newly wired surface (audit parity with the WASM binding) ----------
// These tests exercise the methods that were previously only reachable
// from the WASM package. They assert sensible return shapes rather than
// deep semantics; the point is "the addon wires them and does not throw".

test('addConditionalFormat + getCfRule + clearConditionalFormats round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.sheet(0).listCfRules().length, 0);
  // A CF rule's dxfIndex must resolve against a registered dxf.
  const dxf = wb.ooxml.addDxf({ font: { bold: true } });
  // cellIs rule (type 1) comparing the cell to a literal.
  const add = wb.sheet(0).addCfRule({
    ranges: [{ firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 0 }],
    type: 1,
    operator: 5,
    formula1: '10',
    dxfIndex: dxf,
    stopIfTrue: false,
  });
  assert.equal(typeof add, 'number');
  const list = wb.sheet(0).listCfRules();
  assert.equal(list.length, 1);
  assert.equal(list[0].type, 1);
  assert.equal(typeof list[0].id, 'string');
  assert.equal(list[0].ranges.length, 1);
  assert.equal(list[0].ranges[0].lastRow, 4);
  assert.equal(list[0].formula1, '=10');
  // removeConditionalFormatAt drops the only rule.
  wb.sheet(0).removeCfRule(0);
  assert.equal(wb.sheet(0).listCfRules().length, 0);
  // clearConditionalFormats is a safe no-op on an empty sheet.
  wb.sheet(0).clearCfRules();
  // Sheet-out-of-range is rejected.
  assert.throws(() => wb.sheet(999).clearCfRules(), isFormulonError);
});

test('data bar x14 fields survive save and load', async () => {
  // The six live in the `x14` extension rather than the legacy `<dataBar>`
  // element, so an in-session round-trip would still pass with a writer
  // that never emits the extension -- which is exactly how they used to be
  // lost. Only a save/load cycle pins them.
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).addCfRule({
    ranges: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 0 }],
    type: 3,
    dataBar: {
      min: { type: 3 },
      max: { type: 4 },
      fill: '#0000FF',
      gradient: false,
      axisPosition: 1,
      negativeFill: '#FF0000',
      border: '#090909',
      negativeBorder: '#080808',
      axisColor: '#010203',
      direction: 2,
    },
  });

  const saved = wb.save();
  const loaded = mod.Workbook.load(saved);
  const bar = loaded.sheet(0).listCfRules()[0].dataBar;
  assert.equal(bar.gradient, false);
  assert.equal(bar.axisPosition, 1);
  assert.equal(bar.negativeFill.argb, 0xffff0000);
  assert.equal(bar.border.argb, 0xff090909);
  assert.equal(bar.negativeBorder.argb, 0xff080808);
  assert.equal(bar.axisColor.argb, 0xff010203);
  assert.equal(bar.direction, 2);

  // The decoded bar fed straight back must reproduce the same rule.
  loaded.sheet(0).addCfRule({
    ranges: [{ firstRow: 4, firstCol: 0, lastRow: 6, lastCol: 0 }],
    type: 3,
    dataBar: bar,
  });
  const reread = loaded.sheet(0).listCfRules()[1].dataBar;
  assert.equal(reread.gradient, false);
  assert.equal(reread.axisPosition, 1);
  assert.equal(reread.axisColor.argb, 0xff010203);
  assert.equal(reread.direction, 2);

  loaded.dispose();
  wb.dispose();
});

test('icon set floor survives save and load', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).addCfRule({
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
  const saved = wb.save();
  const loaded = mod.Workbook.load(saved);
  const floor = loaded.sheet(0).listCfRules()[0].iconSet.floor;
  assert.equal(floor.type, 0);
  assert.equal(floor.value, '5');
  assert.equal(floor.gte, false);
  loaded.dispose();
  wb.dispose();
});

test('omitted data bar x14 fields keep the model defaults', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).addCfRule({
    ranges: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 0 }],
    type: 3,
    dataBar: { min: { type: 3 }, max: { type: 4 }, fill: '#0000FF' },
  });
  const bar = wb.sheet(0).listCfRules()[0].dataBar;
  // The getter engages all six, so they read back as the defaults rather
  // than as absent keys.
  assert.equal(bar.gradient, true);
  assert.equal(bar.axisPosition, 0);
  assert.equal(bar.negativeFill.argb, 0xff0000ff);
  wb.dispose();
});

test('named cell styles: listCellStyles / getCellStyle / listCellStyleXfs', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // A fresh workbook may carry zero named styles; the accessors must
  // still return well-formed values rather than throwing.
  const count = wb.listCellStyles().length;
  assert.equal(typeof count, 'number');
  assert.ok(count >= 0);
  assert.equal(typeof wb.ooxml.listCellStyleXfs().length, 'number');
  if (count > 0) {
    const cs = wb.listCellStyles()[0];
    assert.equal(typeof cs.name, 'string');
    assert.equal(typeof cs.styleXfIndex, 'number');
    assert.deepEqual(wb.getCellStyle(cs.name), cs);
  }
  // An unknown name is ok with a null value, not a throw.
  const none = wb.getCellStyle('No Such Style');
  assert.equal(none, null);
});

test('setCalcMode / calcMode round-trip the calc policy', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Default is automatic (0).
  assert.equal(wb.getCalcMode(), 0);
  wb.setCalcMode(1);
  assert.equal(wb.getCalcMode(), 1);
  wb.setCalcMode(2);
  assert.equal(wb.getCalcMode(), 2);
});

test('setPinnedNow / getPinnedNow / setPinnedNow(null) drive the clock seam', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Unpinned by default: the workbook follows the host clock.
  assert.equal(wb.getPinnedNow(), null);

  wb.setPinnedNow({ year: 2026, month: 4, day: 23, hour: 15, minute: 30, second: 45 });
  assert.deepEqual(wb.getPinnedNow(), {
    year: 2026,
    month: 4,
    day: 23,
    hour: 15,
    minute: 30,
    second: 45,
  });

  // 2026-04-23 is serial 46135 under the 1900 date system.
  const today = wb.sheet(0).evaluateTypedFormula(0, 0, '=TODAY()');
  assert.equal(today.number, 46135);

  // The pin is a calendar instant, not a normalising constructor.
  for (const [year, month, day] of [
    [2026, 13, 1],
    [2025, 2, 29],
  ]) {
    assert.throws(() => wb.setPinnedNow({ year, month, day, hour: 0, minute: 0, second: 0 }), isFormulonError);
  }

  wb.setPinnedNow(null);
  assert.equal(wb.getPinnedNow(), null);
});

test('excelProfileId / setProfileId round-trip the profile id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const def = wb.getProfileId();
  assert.equal(typeof def, 'string');
  assert.ok(def.length > 0);
  wb.setProfileId('mac-365-ja_JP');
  assert.equal(wb.getProfileId(), 'mac-365-ja_JP');
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

test('setProfileId switches a live workbook through every profile id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.getProfileId(), 'win-365-en_US');
  wb.sheet(0).setFormula(0, 0, '=DOLLAR(-1234.567)');
  for (const host of ['mac', 'win']) {
    for (const locale of PROFILE_LOCALES) {
      const id = `${host}-365-${locale}`;
      wb.setProfileId(id);
      assert.equal(wb.getProfileId(), id);
      wb.recalc();
      if (id in PROFILE_DOLLAR) {
        assert.equal(wb.sheet(0).getTypedValue(0, 0).text, PROFILE_DOLLAR[id], id);
      }
    }
  }
  assert.throws(() => wb.setProfileId('mac-365-xx_XX'), isFormulonError);
});

test('listFunctionNames + getFunctionMetadata expose the catalog', async () => {
  const mod = await getModule();
  const names = mod.listFunctionNames();
  assert.ok(Array.isArray(names));
  assert.ok(names.length > 0, 'expected a non-empty function catalog');
  assert.ok(names.includes('SUM'), 'expected SUM in the catalog');
  const md = mod.getFunctionMetadata('SUM');
  assert.equal(md.name, 'SUM');
  assert.equal(typeof md.minArity, 'number');
  // An unknown function reads as null.
  const miss = mod.getFunctionMetadata('NOT_A_REAL_FUNCTION_XYZ');
  assert.equal(miss, null);
});

test('mergeFunctionMetadata overlays provider metadata; without a provider it adds only the canonical name', async () => {
  const mod = await getModule();
  assert.equal(typeof mod.mergeFunctionMetadata, 'function');
  const base = mod.getFunctionMetadata('XLOOKUP');
  assert.equal(base.name, 'XLOOKUP');
  // The engine leaves display metadata empty.
  assert.equal(base.signature, null);
  assert.equal(base.description, null);

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
  assert.equal(fr.signature, 'RECHERCHEX(...)');
  assert.equal(fr.description, 'Recherche.');
  assert.equal(fr.localizedName, 'RECHERCHEX');
  // Structural fields survive the merge.
  assert.equal(fr.minArity, base.minArity);
  assert.equal(fr.name, 'XLOOKUP');

  // A locale with no localized/alias entry falls back to the default
  // signature/description and the canonical display name.
  const de = mod.mergeFunctionMetadata(base, entry, 'de-DE');
  assert.equal(de.signature, 'XLOOKUP(lookup_value, lookup_array, return_array)');
  assert.equal(de.description, 'Searches a range or an array.');
  assert.equal(de.localizedName, 'XLOOKUP');

  // No provider entry -> the base fields plus localizedName = name; display metadata stays absent.
  const none = mod.mergeFunctionMetadata(base, undefined, 'fr-FR');
  assert.deepEqual({ ...none }, { ...base, localizedName: 'XLOOKUP' });
  assert.equal(none.signature, null);
  assert.equal(none.description, null);
});

test('localizeFunctionName / canonicalizeFunctionName round-trip', async () => {
  const mod = await getModule();
  // en-US profile: the localized name is the canonical name itself.
  assert.equal(mod.localizeFunctionName('SUM', 'win-365-en_US'), 'SUM');
  assert.equal(mod.canonicalizeFunctionName('SUM', 'win-365-en_US'), 'SUM');
  // An unknown name is refused.
  assert.throws(
    () => mod.canonicalizeFunctionName('NOPE_XYZ', 'win-365-en_US'),
    (e) => isFormulonError(e, mod.StatusCode.NotFound),
  );
});

test('formula and function-name localization is keyed by profile id', async () => {
  const mod = await getModule();
  assert.equal(mod.localizeFormula('=SUM(1.5,2)', 'mac-365-de_DE'), '=SUMME(1,5;2)');
  assert.equal(mod.canonicalizeFormula('=SUMME(1,5;2)', 'mac-365-de_DE'), '=SUM(1.5,2)');
  assert.equal(mod.localizeFormula('=IF(TRUE,1,2)', 'mac-365-fr_FR'), '=SI(VRAI;1;2)');
  assert.equal(mod.localizeFormula('=#N/A', 'mac-365-de_DE'), '=#NV');
  assert.equal(mod.canonicalizeFormula('=#NV', 'mac-365-de_DE'), '=#N/A');
  assert.equal(mod.localizeFormula('=A1&";"', 'mac-365-de_DE'), '=A1&";"');
  assert.equal(mod.localizeFunctionName('SUM', 'mac-365-de_DE'), 'SUMME');
  assert.equal(mod.canonicalizeFunctionName('SUMME', 'win-365-de_DE'), 'SUM');
  assert.throws(() => mod.localizeFormula('=SUM(1)', 'nope'), isFormulonError);
  assert.throws(() => mod.canonicalizeFormula('=SUM(1)', 'nope'), isFormulonError);
  assert.throws(() => mod.canonicalizeFunctionName('SUM', 'nope'), isFormulonError);
});

test('getLocaleFacts reports separators, names, error spellings and measurement', async () => {
  const mod = await getModule();
  const mac = mod.getLocaleFacts('mac-365-de_DE');
  assert.equal(mac.decimalSeparator, ',');
  assert.equal(mac.listSeparator, ';');
  assert.equal(mac.dateOrder, mod.LocaleDateOrder.Dmy);
  assert.equal(mac.measured, true);
  assert.equal(mac.collationMeasured, true);
  const na = mac.errorNames.find((e) => e.canonical === '#N/A');
  assert.equal(na.localized, '#NV');
  const listed = mod.listLocaleErrorNames('mac-365-de_DE');
  assert.deepEqual([...listed], mac.errorNames);
  assert.throws(() => mod.listLocaleErrorNames('nope'), isFormulonError);
  assert.equal(mod.getLocaleFacts('win-365-de_DE').measured, false);
  assert.equal(mod.getLocaleFacts('win-365-ja_JP').collationMeasured, false);
  assert.equal(mod.getLocaleFacts('win-365-en_US').decimalSeparator, '.');
  assert.throws(() => mod.getLocaleFacts('nope'), isFormulonError);
});

test('compareText folds by strength and rejects bad arguments', async () => {
  const mod = await getModule();
  const profile = 'mac-365-zh_CN';
  const strict = mod.compareText('A', '\uFF21', 0, profile);
  assert.notEqual(strict, 0);
  const dynamic = mod.compareText('A', '\uFF21', 1, profile);
  assert.equal(dynamic, 0);
  const ab = mod.compareText('a', 'b', 0, 'mac-365-en_US');
  const ba = mod.compareText('b', 'a', 0, 'mac-365-en_US');
  assert.equal(ab, -1);
  assert.equal(ba, 1);
  assert.throws(() => mod.compareText('a', 'b', 0, 'nope'), isFormulonError);
  assert.throws(() => mod.compareText('a', 'b', -1, profile), RangeError);
});

test('getCollationClasses orders by kind, folds text by strength, and copies its output', async () => {
  const mod = await getModule();
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
  const r = mod.getCollationClasses(values, 1, 'mac-365-en_US');
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
  mod.getCollationClasses(
    [
      { kind: 3, text: 'z' },
      { kind: 3, text: 'y' },
    ],
    1,
    'mac-365-en_US',
  );
  assert.deepEqual(Array.from(r.classIds), snapshot);

  const folded = mod.getCollationClasses(
    [
      { kind: 3, text: 'a' },
      { kind: 3, text: 'A' },
    ],
    1,
    'mac-365-en_US',
  );
  assert.equal(folded.classIds[0], folded.classIds[1]);
  assert.equal(folded.classCount, 1);

  const empty = mod.getCollationClasses([], 1, 'mac-365-en_US');
  assert.equal(empty.classIds.length, 0);
  assert.equal(empty.classCount, 0);

  // A value without `kind` is rejected, not read as blank.
  assert.throws(() => mod.getCollationClasses([{ number: 1 }], 1, 'mac-365-en_US'), TypeError);
  assert.throws(() => mod.getCollationClasses({}, 1, 'mac-365-en_US'), TypeError);
  assert.throws(() => mod.getCollationClasses([{ kind: 1, number: 1 }], 1, 'nope'), isFormulonError);
  assert.throws(() => mod.getCollationClasses(values, -1, 'mac-365-en_US'), RangeError);
});

test('getCollationClasses timing on 100k mixed CJK/ASCII texts', async () => {
  const mod = await getModule();
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
  const r = mod.getCollationClasses(values, 1, 'mac-365-ja_JP');
  const ms = Number(process.hrtime.bigint() - t0) / 1e6;
  assert.equal(r.classIds.length, 100000);
  console.log(`getCollationClasses 100k mixed CJK/ASCII texts: ${ms.toFixed(0)} ms (${r.classCount} classes)`);
});

test('precedents / dependents return arrays for a small formula graph', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setNumber(0, 0, 1); // A1
  wb.sheet(0).setNumber(0, 1, 2); // B1
  wb.sheet(0).setFormula(1, 0, '=A1+B1'); // A2 = A1 + B1
  wb.recalc();
  // A2 reads A1 and B1.
  const prec = wb.sheet(0).getPrecedents(1, 0, 1);
  assert.ok(Array.isArray(prec));
  assert.equal(prec.length, 2);
  for (const node of prec) {
    assert.equal(typeof node.sheet, 'number');
    assert.equal(typeof node.row, 'number');
    assert.equal(typeof node.col, 'number');
  }
  // A1 is read by A2.
  const dep = wb.sheet(0).getDependents(0, 0, 1);
  assert.ok(Array.isArray(dep));
  assert.equal(dep.length, 1);
  assert.equal(dep[0].row, 1);
  assert.equal(dep[0].col, 0);
});

test('spillInfo reports a dynamic-array region', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // A cell outside any spill reads null.
  assert.equal(wb.sheet(0).getSpill(0, 0), null);
  // A spilling dynamic array anchored at A1.
  wb.sheet(0).setFormula(0, 0, '=SEQUENCE(3,1)');
  wb.recalc();
  const info = wb.sheet(0).getSpill(0, 0);
  assert.equal(info.blocked, false);
  assert.equal('engaged' in info, false);
  assert.deepEqual(info.range, { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 0 });
});

test('array formulas can be authored, read from every cell and cleared', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  for (let r = 0; r < 3; r += 1) wb.sheet(0).setNumber(r, 0, r + 1);
  const none = wb.sheet(0).getArrayFormula(0, 2);
  assert.equal(none, null);
  wb.sheet(0).setArrayFormula({ range: { firstRow: 0, firstCol: 2, lastRow: 2, lastCol: 2 }, formula: '=A1:A3*2' });
  wb.recalc();
  for (let r = 0; r < 3; r += 1) {
    const found = wb.sheet(0).getArrayFormula(r, 2);
    assert.deepEqual(found.range, { firstRow: 0, firstCol: 2, lastRow: 2, lastCol: 2 });
    assert.ok(found.formula.includes('A1:A3'));
  }
  assert.equal(wb.sheet(0).getTypedValue(2, 2).number, 6);
  // A partial overlap and a write into a non-anchor cell are refused.
  assert.throws(
    () => wb.sheet(0).setArrayFormula({ range: { firstRow: 1, firstCol: 2, lastRow: 2, lastCol: 2 }, formula: '=1' }),
    isFormulonError,
  );
  assert.throws(() => wb.sheet(0).setNumber(1, 2, 1), isFormulonError);
  wb.sheet(0).setBlank(0, 2);
  wb.recalc();
  assert.equal(wb.sheet(0).getArrayFormula(1, 2), null);
  assert.throws(() => wb.sheet(99).getArrayFormula(0, 0), isFormulonError);
});

test('getLambdaText is wired and reports the documented contract', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // A bare top-level LAMBDA renders as #CALC! in Excel 365 and does not
  // retain a renderable closure; the engine mirrors that, so no lambda is
  // present. A plain number cell is likewise not a lambda, and a bad sheet
  // throws.
  wb.sheet(0).setFormula(0, 0, '=LAMBDA(x,x+1)');
  wb.recalc();
  const bare = wb.sheet(0).getLambdaText(0, 0);
  assert.equal(bare, null);

  wb.sheet(0).setNumber(1, 0, 5);
  const num = wb.sheet(0).getLambdaText(1, 0);
  assert.equal(num, null);
  assert.throws(
    () => wb.sheet(99).getLambdaText(0, 0),
    (e) => isFormulonError(e, mod.StatusCode.SheetIndexOutOfRange),
  );
});

test('a fresh workbook has no external links', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.deepEqual(wb.listExternalLinks(), []);
});

test('sheet protection is a record whose presence is the protection, with dialog defaults', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.sheet(0).getSheetProtection(), null);

  // An empty input protects the sheet with Excel's dialog defaults.
  wb.sheet(0).setSheetProtection({});
  const defaults = wb.sheet(0).getSheetProtection();
  assert.equal(defaults.allowSelectLockedCells, true);
  assert.equal(defaults.allowSelectUnlockedCells, true);
  for (const key of [
    'allowEditObjects',
    'allowEditScenarios',
    'allowFormatCells',
    'allowFormatColumns',
    'allowFormatRows',
    'allowInsertColumns',
    'allowInsertRows',
    'allowInsertHyperlinks',
    'allowDeleteColumns',
    'allowDeleteRows',
    'allowSort',
    'allowAutoFilter',
    'allowPivotTables',
  ]) {
    assert.equal(defaults[key], false, key);
  }
  assert.equal('enabled' in defaults, false);
  assert.equal('sheet' in defaults, false);

  // The record a getter returns is accepted as is; a set replaces the whole protection.
  const next = { ...defaults, allowSelectLockedCells: false, allowSort: true };
  wb.sheet(0).setSheetProtection(next);
  const after = wb.sheet(0).getSheetProtection();
  assert.equal(after.allowSelectLockedCells, false);
  assert.equal(after.allowSort, true);

  // A patch merges into the stored protection; null removes it.
  wb.sheet(0).updateSheetProtection({ allowFormatRows: true });
  const patched = wb.sheet(0).getSheetProtection();
  assert.equal(patched.allowFormatRows, true);
  assert.equal(patched.allowSort, true);
  wb.sheet(0).setSheetProtection(null);
  assert.equal(wb.sheet(0).getSheetProtection(), null);
  wb.sheet(0).setSheetProtection(null);
  assert.equal(wb.sheet(0).getSheetProtection(), null);

  // A patch on an unprotected sheet applies to the defaults and protects it; an empty patch does not.
  wb.sheet(0).updateSheetProtection({});
  assert.equal(wb.sheet(0).getSheetProtection(), null);
  wb.sheet(0).updateSheetProtection({ allowAutoFilter: true });
  assert.equal(wb.sheet(0).getSheetProtection().allowAutoFilter, true);
  assert.equal(wb.sheet(0).getSheetProtection().allowSelectLockedCells, true);

  // Sheet-out-of-range is rejected.
  assert.throws(() => wb.sheet(999).getSheetProtection(), isFormulonError);
  assert.throws(() => wb.sheet(999).setSheetProtection(next), isFormulonError);
});
