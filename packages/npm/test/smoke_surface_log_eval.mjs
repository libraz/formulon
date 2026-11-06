import test from 'node:test';
import {
  assert,
  emitXlsbWarning,
  expectFailure,
  getModule,
  loadStagedFactory,
  loadStagedModule,
  path,
  pkgJsonPath,
  pkgRoot,
  readFile,
  VAL,
} from './smoke_support.mjs';

// The pthread workers must boot from the core, whose top level starts the
// runtime; a bundler can empty the entry shim when it is used as a worker.
test('threads build spawns its workers from the core, not the entry shim', async () => {
  const glue = await readFile(path.join(pkgRoot, 'dist', 'formulon_threads_core.js'), 'utf8');
  assert.match(glue, /new Worker\(new URL\("formulon_threads_core\.js",import\.meta\.url\)/);
  assert.doesNotMatch(glue, /"formulon_threads\.js"/);
  const pkg = JSON.parse(await readFile(pkgJsonPath, 'utf8'));
  assert.ok(pkg.sideEffects.includes('./dist/formulon_threads_core.js'));
});
test('default export is callable factory returning a Module', async () => {
  const factory = await loadStagedFactory();
  assert.equal(typeof factory, 'function');
  const Module = await factory();
  assert.ok(Module && typeof Module === 'object');
  assert.equal(typeof Module.getVersion, 'function');
});

test('all declared enum and constant exports resolve at runtime', async () => {
  const mod = await loadStagedModule();
  const dts = await readFile(path.join(pkgRoot, 'dist', 'formulon.d.ts'), 'utf8');
  const names = [...dts.matchAll(/^export (?:const )?(?:enum|const) (\w+)/gm)].map((match) => match[1]);
  for (const name of names) {
    assert.ok(name in mod, `missing runtime export ${name}`);
  }
  assert.equal(mod.ValueKind.Error, 4);
  assert.equal(mod.PivotAxis.Row, 0);
  assert.equal(mod.ExternalLinkKind.Dde, 3);
  // Named constants a consumer needs in order to interpret shared record
  // fields (`Value.errorCode`, `ExternalLink.kind`) or to drive the
  // calc policy. The native package exports the same names and ordinals;
  // `check_binding_drift dts-enums` holds the two in step.
  assert.equal(mod.ErrorCode.Div0, 1);
  assert.equal(mod.ErrorCode.Unknown, 16);
  assert.equal(mod.CalcMode.Manual, 1);
  for (const name of ['ErrorCode', 'CalcMode', 'ExternalLinkKind']) {
    assert.ok(Object.isFrozen(mod[name]), `${name} must be frozen`);
  }
});

test('setLogMinLevel / setLogSink are module-level controls on the Module', async () => {
  const Module = await getModule();
  assert.equal(typeof Module.setLogMinLevel, 'function');
  assert.equal(typeof Module.setLogSink, 'function');
  const mod = await loadStagedModule();
  for (const level of Object.values(mod.LogLevel)) {
    Module.setLogMinLevel(level);
  }
  for (const bad of [-1, 5, 99]) {
    expectFailure(() => Module.setLogMinLevel(bad), Module.StatusCode.InvalidArgument);
  }
  Module.setLogMinLevel(mod.LogLevel.Off);
});

test('a registered sink receives records only once the threshold admits them', async () => {
  const Module = await getModule();
  const mod = await loadStagedModule();
  const records = [];
  Module.setLogSink((record) => records.push(new Uint8Array(record)));
  try {
    // Default threshold is Off: the same save must produce nothing.
    emitXlsbWarning(Module, mod.WorkbookFormat.Xlsb);
    assert.deepEqual(records, []);

    Module.setLogMinLevel(mod.LogLevel.Warn);
    emitXlsbWarning(Module, mod.WorkbookFormat.Xlsb);
    assert.ok(records.length > 0, 'expected at least one record');
    for (const r of records) {
      // The sink must receive a copy, never a live view into the WASM
      // instance's own memory: the /threads build backs that memory with
      // a `SharedArrayBuffer`, over which most browsers' `TextDecoder`
      // refuses to operate -- a failure this suite cannot reproduce under
      // Node, so the buffer identity is what it checks instead.
      assert.ok(!(r.buffer instanceof SharedArrayBuffer), 'log sink record must not be SharedArrayBuffer-backed');
    }
    const text = records.map((r) => new TextDecoder().decode(r)).join('');
    assert.match(text, /xlsb\.writer\.formula_downgraded/);
    assert.match(text, /"level":"warn"/);
  } finally {
    Module.setLogMinLevel(mod.LogLevel.Off);
    Module.setLogSink(null);
  }
});

test('a throwing log sink surfaces its own exception and leaves the engine usable', async () => {
  const Module = await getModule();
  const mod = await loadStagedModule();
  let calls = 0;
  const thrown = new Error('sink failure');
  Module.setLogSink(() => {
    calls += 1;
    throw thrown;
  });
  try {
    Module.setLogMinLevel(mod.LogLevel.Warn);
    // The sink runs from deep inside the writer. If its exception reached
    // the C++ frames it would tear them down without running a single
    // destructor, because the binary is linked without exception support;
    // it is re-thrown from the outer call instead.
    assert.throws(
      () => emitXlsbWarning(Module, mod.WorkbookFormat.Xlsb),
      (e) => e === thrown,
    );
    assert.ok(calls > 0, 'expected the sink to be invoked');
  } finally {
    Module.setLogMinLevel(mod.LogLevel.Off);
    Module.setLogSink(null);
  }

  const wb = Module.Workbook.createDefault();
  try {
    wb.sheet(0).setFormula(0, 0, '=1+1');
    wb.recalc();
    assert.equal(wb.sheet(0).getValue(0, 0).number, 2);
  } finally {
    wb.dispose();
  }
});

// One-shot evaluation is anchored at Sheet1!A1 and read-only on every
// shipped surface, so these must agree value-for-value with the WASM
// package and the Python wheel. The anchor-referencing cases are the ones
// that diverge if a surface writes the formula into A1 and recalcs.
const ONE_SHOT_CASES = [
  { formula: '=1+2', kind: 1, number: 3 },
  { formula: '=A1', kind: 1, number: 0 },
  { formula: '=COUNTA(A1)', kind: 1, number: 0 },
  { formula: '=ISBLANK(A1)', kind: 2, boolean: true },
  { formula: '=SUM(A1:A3)', kind: 1, number: 0 },
  { formula: '=SEQUENCE(3)', kind: 1, number: 1 },
  { formula: '=ROWS(SEQUENCE(3))', kind: 1, number: 3 },
];
test('evalFormula evaluates read-only against a blank anchor', async () => {
  const Module = await getModule();
  for (const spec of ONE_SHOT_CASES) {
    const r = Module.evalFormula(spec.formula);
    assert.equal(r.kind, spec.kind, spec.formula);
    if (spec.number !== undefined) {
      assert.equal(r.number, spec.number, spec.formula);
    }
    if (spec.boolean !== undefined) {
      assert.equal(r.boolean, spec.boolean, spec.formula);
    }
    assert.equal('errorCode' in r, false, spec.formula);
  }
});

// Parses the staged `export const X: Readonly<{ A: 0; ... }>` declarations into
// {name: {member: ordinal}}. Name presence alone is not enough: a staged
// bundle carrying a stale ordinal ships a constant that silently means
// something else, and nothing upstream of dist/ can see that.
function declaredOrdinalTables(dts) {
  const tables = {};
  const blocks = dts.matchAll(/^export const (\w+): Readonly<\{([^}]*)\}>;/gm);
  for (const [, name, body] of blocks) {
    const members = {};
    for (const line of body.split(';')) {
      const m = line.trim().match(/^(\w+):\s*(-?\d+)$/);
      if (m) members[m[1]] = Number(m[2]);
    }
    tables[name] = members;
  }
  return tables;
}

test('staged constant tables carry the ordinals the staged .d.ts declares', async () => {
  const mod = await loadStagedModule();
  const dts = await readFile(path.join(pkgRoot, 'dist', 'formulon.d.ts'), 'utf8');
  const tables = declaredOrdinalTables(dts);
  assert.ok(Object.keys(tables).length > 0, 'expected declared ordinal tables');
  for (const [name, members] of Object.entries(tables)) {
    const runtime = mod[name];
    assert.ok(runtime, `missing runtime table ${name}`);
    assert.deepEqual({ ...runtime }, members, `${name} ordinals differ from dist/formulon.d.ts`);
    assert.ok(Object.isFrozen(runtime), `${name} must be frozen`);
  }
});

// Builds a workbook with one pivot, ready for the filter-contract cases.
// `delete` is embind's auto-added finaliser; the caller must call it.
function makePivotWorkbook(Workbook) {
  const wb = Workbook.createDefault();
  const cacheId = wb.addPivotCache(0);
  // A cache without a worksheet source cannot be saved, so the fixture
  // declares the range its two fields and two records occupy.
  wb.setPivotCacheWorksheetSource(cacheId, { present: true, ref: 'A1:B3', sheet: 'Sheet1' });
  wb.addPivotCacheField(cacheId, 'Region');
  wb.addPivotCacheField(cacheId, 'Amount');
  for (const [region, amount] of [
    ['East', 10],
    ['West', 30],
  ]) {
    const rec = wb.addPivotCacheRecord(cacheId);
    wb.setPivotCacheRecordValue(cacheId, rec, 0, region);
    wb.setPivotCacheRecordValue(cacheId, rec, 1, amount);
  }
  const pivot = wb.sheet(0).addPivot({ name: 'Pivot1', cacheId, anchorRow: 0, anchorCol: 4 });
  wb.sheet(0).addPivotField(pivot, { sourceName: 'Region', axis: 0 });
  const amountField = wb.sheet(0).addPivotField(pivot, { sourceName: 'Amount', axis: 2 });
  wb.sheet(0).addPivotDataField(pivot, {
    name: 'Sum of Amount',
    fieldIndex: amountField,
    aggregation: 0,
  });
  return { wb, pivot: pivot };
}

// greaterThan on a double payload: the shape pivotFilterAt must read back.
const PIVOT_FILTER = Object.freeze({
  axis: 0,
  fieldName: 'Region',
  type: 1,
  valueKind: 1,
  valueDouble: 15,
});

test('pivotFilterAt reads back an added filter field for field', async () => {
  const Module = await getModule();
  const { wb, pivot } = makePivotWorkbook(Module.Workbook);
  try {
    wb.sheet(0).addPivotFilter(pivot, PIVOT_FILTER);
    assert.equal(wb.sheet(0).listPivotFilters(pivot).length, 1);
    const got = wb.sheet(0).listPivotFilters(pivot)[0];
    assert.equal(got.axis, PIVOT_FILTER.axis);
    assert.equal(got.fieldName, PIVOT_FILTER.fieldName);
    assert.equal(got.type, PIVOT_FILTER.type);
    assert.equal(got.valueKind, PIVOT_FILTER.valueKind);
    assert.equal(got.valueDouble, PIVOT_FILTER.valueDouble);
  } finally {
    wb.dispose();
  }
});

test('pivotFilterCount counts the filters on the pivot', async () => {
  const Module = await getModule();
  const { wb, pivot } = makePivotWorkbook(Module.Workbook);
  try {
    assert.equal(wb.sheet(0).listPivotFilters(pivot).length, 0);
    wb.sheet(0).addPivotFilter(pivot, PIVOT_FILTER);
    assert.equal(wb.sheet(0).listPivotFilters(pivot).length, 1);
  } finally {
    wb.dispose();
  }
});

test('pivot filters survive save/load', async () => {
  const Module = await getModule();
  const { wb, pivot } = makePivotWorkbook(Module.Workbook);
  let bytes;
  try {
    wb.sheet(0).addPivotFilter(pivot, PIVOT_FILTER);
    assert.equal(wb.sheet(0).listPivotFilters(pivot).length, 1);
    const saved = wb.save();
    bytes = saved;
  } finally {
    wb.dispose();
  }
  const reloaded = Module.Workbook.load(bytes);
  try {
    assert.equal(reloaded.sheet(0).listPivots().length, 1);
    // The filter is written as `<filters>` and decodes back into the list.
    assert.equal(reloaded.sheet(0).listPivotFilters(0).length, 1);
    const got = reloaded.sheet(0).listPivotFilters(0)[0];
    assert.equal(got.fieldName, PIVOT_FILTER.fieldName);
    assert.equal(got.type, PIVOT_FILTER.type);
    assert.equal(got.valueKind, PIVOT_FILTER.valueKind);
    assert.equal(got.valueDouble, PIVOT_FILTER.valueDouble);
  } finally {
    reloaded.dispose();
  }
});

test('pivotFilterAt rejects an out-of-range index', async () => {
  const Module = await getModule();
  const { wb } = makePivotWorkbook(Module.Workbook);
  try {
    expectFailure(() => wb.sheet(0).listPivotFilters(99));
  } finally {
    wb.dispose();
  }
});

test('a required pivot spec string field left out of the object is rejected, not silently empty', async () => {
  const Module = await getModule();
  const { wb, pivot } = makePivotWorkbook(Module.Workbook);
  try {
    const fieldCountBefore = wb.sheet(0).listPivotFields(pivot).length;
    // no `sourceName`
    expectFailure(() => wb.sheet(0).addPivotField(pivot, { axis: 0 }), Module.StatusCode.BindingNullPointer);
    assert.equal(wb.sheet(0).listPivotFields(pivot).length, fieldCountBefore, 'no field must have been added');

    // `name` is optional: an omitted one is derived from the workbook profile.
    const dataFieldCountBefore = wb.sheet(0).listPivotDataFields(pivot).length;
    const dataField = wb.sheet(0).addPivotDataField(pivot, { fieldIndex: 0, aggregation: 0 }); // no `name`
    assert.equal(wb.sheet(0).listPivotDataFields(pivot).length, dataFieldCountBefore + 1);
    assert.equal(wb.sheet(0).listPivotDataFields(pivot)[dataField].name, 'Sum of Region');

    const filterCountBefore = wb.sheet(0).listPivotFilters(pivot).length;
    // no `fieldName`
    expectFailure(
      () => wb.sheet(0).addPivotFilter(pivot, { axis: 0, type: 1, valueKind: 1, valueDouble: 15 }),
      Module.StatusCode.BindingNullPointer,
    );
    assert.equal(wb.sheet(0).listPivotFilters(pivot).length, filterCountBefore, 'no filter must have been added');
  } finally {
    wb.dispose();
  }
});

test('getVersion returns a non-empty string', async () => {
  const Module = await getModule();
  const v = Module.getVersion();
  assert.equal(typeof v, 'string');
  assert.ok(v.length > 0, `expected non-empty version, got ${JSON.stringify(v)}`);
});

test('evalFormula(=SUM(1,2,3)) returns NUMBER 6', async () => {
  const Module = await getModule();
  const r = Module.evalFormula('=SUM(1,2,3)');
  assert.equal(r.kind, VAL.NUMBER);
  assert.equal(r.number, 6);
});

test('evalFormula(=1/0) surfaces Excel error as a value (not failed status)', async () => {
  const Module = await getModule();
  const r = Module.evalFormula('=1/0');
  assert.equal(r.kind, VAL.ERROR);
});

test('evalFormula uses read-only evaluation for the intersection operator', async () => {
  const Module = await getModule();
  const r = Module.evalFormula('=A1 B1');
  assert.equal(r.kind, VAL.ERROR);
  assert.equal(r.errorCode, 0); // ErrorCode::Null / #NULL!
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
  try {
    assert.equal(wb.getExcelProfileId(), 'win-365-en_US');
    wb.sheet(0).setFormula(0, 0, '=DOLLAR(-1234.567)');
    for (const host of ['mac', 'win']) {
      for (const locale of PROFILE_LOCALES) {
        const id = `${host}-365-${locale}`;
        wb.setExcelProfileId(id);
        assert.equal(wb.getExcelProfileId(), id);
        wb.recalc();
        if (id in PROFILE_DOLLAR) {
          assert.equal(wb.sheet(0).getValue(0, 0).text, PROFILE_DOLLAR[id], id);
        }
      }
    }
    expectFailure(() => wb.setExcelProfileId('mac-365-xx_XX'));
  } finally {
    wb.dispose();
  }
});

test('formula and function-name localization is keyed by profile id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.equal(mod.localizeFormula('=SUM(1.5,2)', 'mac-365-de_DE'), '=SUMME(1,5;2)');
    assert.equal(mod.canonicalizeFormula('=SUMME(1,5;2)', 'mac-365-de_DE'), '=SUM(1.5,2)');
    assert.equal(mod.localizeFormula('=IF(TRUE,1,2)', 'mac-365-fr_FR'), '=SI(VRAI;1;2)');
    assert.equal(mod.localizeFormula('=#N/A', 'mac-365-de_DE'), '=#NV');
    assert.equal(mod.canonicalizeFormula('=#NV', 'mac-365-de_DE'), '=#N/A');
    assert.equal(mod.localizeFormula('=A1&";"', 'mac-365-de_DE'), '=A1&";"');
    assert.equal(mod.localizeFunctionName('SUM', 'mac-365-de_DE'), 'SUMME');
    assert.equal(mod.canonicalizeFunctionName('SUMME', 'win-365-de_DE'), 'SUM');
    expectFailure(() => mod.localizeFormula('=SUM(1)', 'nope'));
    expectFailure(() => mod.canonicalizeFormula('=SUM(1)', 'nope'));
    expectFailure(() => mod.canonicalizeFunctionName('SUM', 'nope'));
  } finally {
    wb.dispose();
  }
});

test('getLocaleFacts reports separators, names, error spellings and measurement', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const mac = mod.getLocaleFacts('mac-365-de_DE');
    assert.equal(mac.decimalSeparator, ',');
    assert.equal(mac.listSeparator, ';');
    assert.equal(mac.dateOrder, mod.LocaleDateOrder.Dmy);
    assert.equal(mac.measured, true);
    const na = mac.errorNames.find((e) => e.canonical === '#N/A');
    assert.equal(na.localized, '#NV');
    const listed = mod.listLocaleErrorNames('mac-365-de_DE');
    assert.deepEqual([...listed], mac.errorNames);
    expectFailure(() => mod.listLocaleErrorNames('nope'));
    assert.equal(mod.getLocaleFacts('win-365-de_DE').measured, false);
    assert.equal(mod.getLocaleFacts('win-365-en_US').decimalSeparator, '.');
    expectFailure(() => mod.getLocaleFacts('nope'));
  } finally {
    wb.dispose();
  }
});

test('pivot definition read-backs report what the setters wrote', async () => {
  const Module = await getModule();
  const { wb, pivot } = makePivotWorkbook(Module.Workbook);
  try {
    assert.equal(wb.sheet(0).getPivotCacheId(pivot), wb.listPivotCaches()[0]);
    assert.equal(wb.sheet(0).getPivotName(pivot), 'Pivot1');
    assert.deepEqual({ ...wb.sheet(0).getPivotGrandTotals(pivot) }, { rows: true, columns: true });
    wb.sheet(0).setPivotGrandTotals(pivot, { rows: false, columns: true });
    const totals = wb.sheet(0).getPivotGrandTotals(pivot);
    assert.equal(totals.rows, false);
    assert.equal(totals.columns, true);

    assert.deepEqual(wb.sheet(0).getPivotRowFieldOrder(pivot), [0]);
    assert.deepEqual(wb.sheet(0).getPivotColumnFieldOrder(pivot), []);

    wb.sheet(0).setPivotFieldSort(pivot, 0, 0, { byField: 'Sum of Amount' });
    const region = wb.sheet(0).listPivotFields(pivot)[0];
    assert.equal(region.sourceName, 'Region');
    assert.equal(region.axis, 0);
    assert.equal(region.cacheFieldIndex, 0);
    assert.equal(region.sortOrder, 0);
    assert.equal(region.sortByField, 'Sum of Amount');
    assert.equal(region.datePeriods, 0);
    assert.equal(region.groupBaseIndex, null);
    assert.equal(wb.sheet(0).listPivotFields(pivot)[1].axis, 2);

    wb.sheet(0).addPivotFieldItem(pivot, 0, 'East', { visible: true });
    wb.sheet(0).addPivotFieldItem(pivot, 0, 'West', { visible: true });
    wb.sheet(0).updatePivotItem(pivot, 0, 1, { visible: false });
    assert.equal(wb.sheet(0).listPivotFieldItems(pivot, 0).length, 2);
    const east = wb.sheet(0).getPivotFieldItem(pivot, 0, 0);
    assert.equal(east.name, 'East');
    // The scratch cache holds no shared items, so a label binds to nothing.
    assert.equal(east.cacheIndex, null);
    assert.equal(east.visible, true);
    assert.equal(wb.sheet(0).getPivotFieldItem(pivot, 0, 1).visible, false);
    wb.sheet(0).addPivotFieldSharedItem(pivot, 0, 5, true);
    const bound = wb.sheet(0).getPivotFieldItem(pivot, 0, 2);
    assert.equal(bound.cacheIndex, 5);
    assert.equal(bound.name, '');

    const dataField = wb.sheet(0).listPivotDataFields(pivot)[0];
    assert.equal(dataField.name, 'Sum of Amount');
    assert.equal(dataField.fieldIndex, 1);
    assert.equal(dataField.aggregation, 0);
    assert.equal(dataField.showAsBaseField, null);
    assert.equal(dataField.showAsBaseItem, null);
  } finally {
    wb.dispose();
  }
});

test('pivot style info defaults on create, reads back what the setter wrote and rejects bad arguments', async () => {
  const Module = await getModule();
  const { wb, pivot } = makePivotWorkbook(Module.Workbook);
  try {
    const initial = wb.sheet(0).getPivotStyleInfo(pivot);
    assert.notEqual(initial, null);
    assert.equal(initial.name, 'PivotStyleLight16');
    assert.deepEqual(
      [initial.showRowHeaders, initial.showColHeaders, initial.showRowStripes, initial.showColStripes],
      [true, true, false, false],
    );
    assert.equal(initial.showLastColumn, true);

    wb.sheet(0).setPivotStyleInfo(pivot, { name: 'PivotStyleDark3', showRowStripes: true });
    const set = wb.sheet(0).getPivotStyleInfo(pivot);
    assert.equal(set.name, 'PivotStyleDark3');
    assert.deepEqual(
      [set.showRowHeaders, set.showColHeaders, set.showRowStripes, set.showColStripes, set.showLastColumn],
      [true, true, true, false, true],
    );

    // Survives a save/load.
    const saved = wb.save();
    const reloaded = Module.Workbook.load(saved);
    try {
      assert.equal(reloaded.sheet(0).getPivotStyleInfo(0).name, 'PivotStyleDark3');
      assert.equal(reloaded.sheet(0).getPivotStyleInfo(0).showRowStripes, true);
    } finally {
      reloaded.dispose();
    }

    // An empty patch changes nothing; an empty name leaves the element without a style name.
    wb.sheet(0).setPivotStyleInfo(pivot, {});
    assert.equal(wb.sheet(0).getPivotStyleInfo(pivot).name, 'PivotStyleDark3');
    wb.sheet(0).setPivotStyleInfo(pivot, { name: '' });
    const bare = wb.sheet(0).getPivotStyleInfo(pivot);
    assert.notEqual(bare, null);
    assert.equal(bare.name, '');

    expectFailure(() => wb.sheet(0).getPivotStyleInfo(9));
    expectFailure(() => wb.sheet(0).setPivotStyleInfo(9, { name: 'X' }));
  } finally {
    wb.dispose();
  }
});

test('pivot definition read-backs reject out-of-range indices', async () => {
  const Module = await getModule();
  const { wb, pivot } = makePivotWorkbook(Module.Workbook);
  try {
    for (const call of [
      () => wb.sheet(0).getPivotCacheId(9),
      () => wb.sheet(0).getPivotName(9),
      () => wb.sheet(9).getPivotGrandTotals(pivot),
      () => wb.sheet(0).getPivotRowFieldOrder(9),
      () => wb.sheet(0).getPivotColumnFieldOrder(9),
      () => wb.sheet(0).listPivotFields(99),
      () => wb.sheet(0).listPivotFieldItems(pivot, 99),
      () => wb.sheet(0).getPivotFieldItem(pivot, 0, 99),
      () => wb.sheet(0).listPivotDataFields(99),
    ]) {
      expectFailure(call);
    }
  } finally {
    wb.dispose();
  }
});

test('an empty or omitted data field name is derived from the workbook profile', async () => {
  const Module = await getModule();
  const { wb, pivot } = makePivotWorkbook(Module.Workbook);
  try {
    for (const [profile, expected] of [
      ['win-365-ja_JP', '合計 / Amount'],
      ['win-365-en_US', 'Sum of Amount'],
    ]) {
      wb.setExcelProfileId(profile);
      for (const spec of [
        { fieldIndex: 1, aggregation: 0 },
        { name: '', fieldIndex: 1, aggregation: 0 },
      ]) {
        // A repeat of the same source field is numbered, so each name starts from none.
        wb.sheet(0).clearPivotDataFields(pivot);
        const added = wb.sheet(0).addPivotDataField(pivot, spec);
        assert.equal(wb.sheet(0).listPivotDataFields(pivot)[added].name, expected);
      }
      wb.sheet(0).setPivotDataField(pivot, 0, { fieldIndex: 1, aggregation: 2 });
      assert.ok(wb.sheet(0).listPivotDataFields(pivot)[0].name.endsWith(' Amount'));
    }
    expectFailure(() => wb.sheet(0).addPivotDataField(pivot, { fieldIndex: 99, aggregation: 0 }));
  } finally {
    wb.dispose();
  }
});

test('active sheet is recorded in bookViews and survives a round trip', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  wb.addSheet('B');
  wb.addSheet('C');
  assert.equal(wb.getActiveSheet(), wb.sheet(0));
  wb.setActiveSheet(2);
  const active = wb.getActiveSheet();
  assert.equal(active, wb.sheet('C'));
  expectFailure(() => wb.setActiveSheet(9));
  assert.equal(wb.getActiveSheet().index, 2);
  const saved = wb.save();
  const loaded = Module.Workbook.load(saved);
  try {
    assert.equal(loaded.getActiveSheet().name, 'C');
  } finally {
    loaded.dispose();
    wb.dispose();
  }
});

test('the show-phonetic flag survives value writes and a round trip', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  assert.equal(wb.sheet(0).getCellShowPhonetic(0, 0), false);
  wb.sheet(0).setText(0, 0, '漢字');
  wb.sheet(0).setCellShowPhonetic(0, 0, true);
  wb.sheet(0).setText(0, 0, '漢字');
  assert.equal(wb.sheet(0).getCellShowPhonetic(0, 0), true);
  const saved = wb.save();
  const loaded = Module.Workbook.load(saved);
  try {
    assert.equal(loaded.sheet(0).getCellShowPhonetic(0, 0), true);
    loaded.sheet(0).setCellShowPhonetic(0, 0, false);
    assert.equal(loaded.sheet(0).getCellShowPhonetic(0, 0), false);
    expectFailure(() => loaded.sheet(99).getCellShowPhonetic(0, 0));
    expectFailure(() => loaded.sheet(99).setCellShowPhonetic(0, 0, true));
  } finally {
    loaded.dispose();
    wb.dispose();
  }
});

test('an explicit all-zero phoneticPr is reported as present and survives a round trip', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  wb.sheet(0).setText(0, 0, '漢字');
  wb.sheet(0).setCellPhoneticRuns(0, 0, [{ sb: 0, eb: 2, text: 'かんじ' }]);
  assert.equal(wb.sheet(0).getCellPhoneticProperties(0, 0).present, false);
  wb.sheet(0).setCellPhoneticProperties(0, 0, { fontId: 0, type: 0, alignment: 0 });
  assert.equal(wb.sheet(0).getCellPhoneticProperties(0, 0).present, true);
  const saved = wb.save();
  const loaded = Module.Workbook.load(saved);
  try {
    const got = loaded.sheet(0).getCellPhoneticProperties(0, 0);
    assert.equal(got.present, true);
    assert.deepEqual([got.fontId, got.type, got.alignment], [0, 0, 0]);
  } finally {
    loaded.dispose();
    wb.dispose();
  }
});

test('scroll origin, pane origin and selection round-trip through save', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  wb.sheet(0).setView({ freezeRows: 2, freezeCols: 1 });
  const view = wb.sheet(0).getView();
  assert.equal(view.paneTopLeftRow, 2);
  assert.equal(view.paneTopLeftCol, 1);
  wb.sheet(0).setView({ topLeftRow: 5, topLeftCol: 0, paneTopLeftRow: 40, paneTopLeftCol: 6 });
  expectFailure(() => wb.sheet(0).setView({ paneTopLeftRow: 1, paneTopLeftCol: 6 }));
  const ranges = [
    { firstRow: 41, firstCol: 7, lastRow: 43, lastCol: 8 },
    { firstRow: 50, firstCol: 9, lastRow: 50, lastCol: 9 },
  ];
  wb.sheet(0).setSelection({ activeRow: 41, activeCol: 7, ranges });
  expectFailure(() =>
    wb.sheet(0).setSelection({
      activeRow: 0,
      activeCol: 0,
      ranges: [{ firstRow: 3, firstCol: 0, lastRow: 1, lastCol: 0 }],
    }),
  );
  const saved = wb.save();
  const loaded = Module.Workbook.load(saved);
  try {
    const got = loaded.sheet(0).getView();
    assert.deepEqual([got.topLeftRow, got.topLeftCol, got.paneTopLeftRow, got.paneTopLeftCol], [5, 0, 40, 6]);
    const selection = loaded.sheet(0).getSelection();
    assert.equal(selection.activeRow, 41);
    assert.equal(selection.activeCol, 7);
    assert.deepEqual(
      selection.ranges.map((r) => [r.firstRow, r.firstCol, r.lastRow, r.lastCol]),
      [
        [41, 7, 43, 8],
        [50, 9, 50, 9],
      ],
    );
    expectFailure(() => loaded.sheet(99).getSelection());
  } finally {
    loaded.dispose();
    wb.dispose();
  }
});

test('sheet split reads back, round-trips and exchanges with the freeze', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    assert.deepEqual([wb.sheet(0).getView().splitX, wb.sheet(0).getView().splitY], [0, 0]);
    wb.sheet(0).setView({ splitX: 2000, splitY: 1500 });
    const view = wb.sheet(0).getView();
    assert.deepEqual([view.splitX, view.splitY, view.freezeRows, view.freezeCols], [2000, 1500, 0, 0]);
    const saved = wb.save();
    const loaded = Module.Workbook.load(saved);
    try {
      const got = loaded.sheet(0).getView();
      assert.deepEqual([got.splitX, got.splitY, got.freezeRows, got.freezeCols], [2000, 1500, 0, 0]);
    } finally {
      loaded.dispose();
    }
    wb.sheet(0).setView({ freezeRows: 2, freezeCols: 1 });
    assert.deepEqual([wb.sheet(0).getView().splitX, wb.sheet(0).getView().splitY], [0, 0]);
    wb.sheet(0).setView({ splitX: 0, splitY: 900 });
    const split = wb.sheet(0).getView();
    assert.deepEqual([split.splitX, split.splitY, split.freezeRows, split.freezeCols], [0, 900, 0, 0]);
    assert.throws(() => wb.sheet(0).setView({ splitX: -1, splitY: 0 }));
    assert.throws(() => wb.sheet(0).setView({ splitX: 0, splitY: Number.NaN }));
    assert.throws(() => wb.sheet(0).setView({ splitX: Number.POSITIVE_INFINITY, splitY: 0 }));
    assert.equal(wb.sheet(0).getView().splitY, 900);
    wb.sheet(0).setView({ splitX: 0, splitY: 0 });
    assert.equal(wb.sheet(0).getView().splitY, 0);
  } finally {
    wb.dispose();
  }
});

test('pivotFieldAdd binds by cache column name and a null axis takes a field off the report', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    const cacheId = wb.addPivotCache();
    wb.setPivotCacheWorksheetSource(cacheId, { present: true, ref: 'A1:D3', sheet: 'Sheet1' });
    for (const name of ['Region', 'Product', 'Sales', 'Qty']) wb.addPivotCacheField(cacheId, name);
    for (const sales of [100, 200]) {
      const rec = wb.addPivotCacheRecord(cacheId);
      for (const [col, value] of [0, 0, sales, 1].entries()) wb.setPivotCacheRecordValue(cacheId, rec, col, value);
    }
    const pivot = wb.sheet(0).addPivot({ name: 'PT', cacheId, anchorRow: 0, anchorCol: 6 });
    const region = wb.sheet(0).addPivotField(pivot, { sourceName: 'Region', axis: 0 });
    const sales = wb.sheet(0).addPivotField(pivot, { sourceName: 'Sales', axis: 2 });
    assert.equal(region, 0);
    assert.equal(sales, 2);
    assert.equal(wb.sheet(0).listPivotFields(pivot).length, 4);
    assert.equal(wb.sheet(0).listPivotFields(pivot)[1].axis, null);
    wb.sheet(0).addPivotDataField(pivot, { fieldIndex: 2, aggregation: 0 });
    assert.equal(wb.sheet(0).listPivotDataFields(pivot).length, 1);

    wb.sheet(0).setPivotFieldAxis(pivot, 2, null);
    assert.equal(wb.sheet(0).listPivotFields(pivot)[2].axis, null);
    assert.equal(wb.sheet(0).listPivotDataFields(pivot).length, 0);
    wb.sheet(0).setPivotFieldAxis(pivot, 0, null);
    assert.deepEqual(wb.sheet(0).getPivotRowFieldOrder(pivot), []);
    expectFailure(() => wb.sheet(0).setPivotFieldAxis(pivot, 0, 4));
    assert.throws(() => wb.sheet(0).setPivotFieldAxis(pivot, 0, -1), RangeError);
    assert.throws(() => wb.sheet(0).setPivotFieldAxis(pivot, 0, -2), RangeError);
  } finally {
    wb.dispose();
  }
});

test('copySheet names the copy automatically or explicitly and reports bad arguments', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    wb.sheet(0).setNumber(0, 0, 5);
    wb.sheet(0).setFormula(0, 1, '=A1*2');
    const auto = wb.copySheet(0, 1);
    assert.equal(auto.index, 1);
    assert.equal(auto.name, 'Sheet1 (2)');
    const named = wb.copySheet(0, 0, 'Archive');
    assert.equal(named, wb.sheet(0));
    assert.equal(wb.sheet(0).name, 'Archive');
    assert.equal(auto.index, 2);
    wb.recalc();
    assert.equal(wb.sheet(0).getValue(0, 1).number, 10);
    expectFailure(() => wb.copySheet(9, 0), Module.StatusCode.SheetIndexOutOfRange);
    expectFailure(() => wb.copySheet(0, 9), Module.StatusCode.SheetIndexOutOfRange);
    expectFailure(() => wb.copySheet(0, 0, ''), Module.StatusCode.InvalidSheetName);
    assert.equal(wb.listSheets().length, 3);
  } finally {
    wb.dispose();
  }
});
