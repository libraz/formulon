import test from 'node:test';
import {
  assert,
  emitXlsbWarning,
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
  assert.equal(typeof Module.versionString, 'function');
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
  // fields (`Value.errorCode`, `ExternalLinkRecord.kind`) or to drive the
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
    assert.ok(Module.setLogMinLevel(level).ok, `level=${level}`);
  }
  for (const bad of [-1, 5, 99]) {
    const r = Module.setLogMinLevel(bad);
    assert.equal(r.ok, false, `level=${bad} should be rejected`);
    assert.equal(r.status, 2);
  }
  assert.ok(Module.setLogMinLevel(mod.LogLevel.Off).ok);
});

test('a registered sink receives records only once the threshold admits them', async () => {
  const Module = await getModule();
  const mod = await loadStagedModule();
  const records = [];
  assert.ok(Module.setLogSink((record) => records.push(new Uint8Array(record))).ok);
  try {
    // Default threshold is Off: the same save must produce nothing.
    emitXlsbWarning(Module, mod.WorkbookFormat.Xlsb);
    assert.deepEqual(records, []);

    assert.ok(Module.setLogMinLevel(mod.LogLevel.Warn).ok);
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
    assert.ok(Module.setLogMinLevel(mod.LogLevel.Off).ok);
    assert.ok(Module.setLogSink(null).ok);
  }
});

test('a throwing log sink is dropped and leaves the engine usable', async () => {
  const Module = await getModule();
  const mod = await loadStagedModule();
  let calls = 0;
  assert.ok(
    Module.setLogSink(() => {
      calls += 1;
      throw new Error('sink failure');
    }).ok,
  );
  try {
    assert.ok(Module.setLogMinLevel(mod.LogLevel.Warn).ok);
    // The sink runs from deep inside the writer. If its exception reached
    // the C++ frames it would tear them down without running a single
    // destructor, because the binary is linked without exception support.
    emitXlsbWarning(Module, mod.WorkbookFormat.Xlsb);
    assert.ok(calls > 0, 'expected the sink to be invoked');
  } finally {
    assert.ok(Module.setLogMinLevel(mod.LogLevel.Off).ok);
    assert.ok(Module.setLogSink(null).ok);
  }

  const wb = Module.Workbook.createDefault();
  try {
    assert.ok(wb.setFormula(0, 0, 0, '=1+1').ok);
    const after = wb.recalc();
    assert.ok(after.ok, `recalc after a throwing sink: ${JSON.stringify(after)}`);
    assert.equal(wb.getValue(0, 0, 0).value.number, 2);
  } finally {
    wb.delete();
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
  { formula: '=ISBLANK(A1)', kind: 2, boolean: 1 },
  { formula: '=SUM(A1:A3)', kind: 1, number: 0 },
  { formula: '=SEQUENCE(3)', kind: 1, number: 1 },
  { formula: '=ROWS(SEQUENCE(3))', kind: 1, number: 3 },
];
test('evalFormula evaluates read-only against a blank anchor', async () => {
  const Module = await getModule();
  for (const spec of ONE_SHOT_CASES) {
    const r = Module.evalFormula(spec.formula);
    assert.ok(r.status.ok, `${spec.formula}: ${JSON.stringify(r.status)}`);
    assert.equal(r.value.kind, spec.kind, spec.formula);
    if (spec.number !== undefined) {
      assert.equal(r.value.number, spec.number, spec.formula);
    }
    if (spec.boolean !== undefined) {
      assert.equal(r.value.boolean, spec.boolean, spec.formula);
    }
    assert.equal(r.value.errorCode, 0, spec.formula);
  }
});

// Parses the staged `export enum X { A = 0, ... }` declarations into
// {name: {member: ordinal}}. Name presence alone is not enough: a staged
// bundle carrying a stale ordinal ships a constant that silently means
// something else, and nothing upstream of dist/ can see that.
function declaredOrdinalTables(dts) {
  const tables = {};
  const blocks = dts.matchAll(/^export enum (\w+) \{([^}]*)\}/gm);
  for (const [, name, body] of blocks) {
    const members = {};
    let next = 0;
    for (const line of body.split(',')) {
      const m = line.trim().match(/^(\w+)\s*(?:=\s*(-?\d+))?$/);
      if (!m) continue;
      if (m[2] !== undefined) next = Number(m[2]);
      members[m[1]] = next;
      next += 1;
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
  const cacheId = wb.pivotCacheCreate(0).index;
  // A cache without a worksheet source cannot be saved, so the fixture
  // declares the range its two fields and two records occupy.
  assert.ok(wb.pivotCacheSetWorksheetSource(cacheId, { present: true, ref: 'A1:B3', sheet: 'Sheet1' }).ok);
  assert.ok(wb.pivotCacheFieldAdd(cacheId, 'Region').status.ok);
  assert.ok(wb.pivotCacheFieldAdd(cacheId, 'Amount').status.ok);
  for (const [region, amount] of [
    ['East', 10],
    ['West', 30],
  ]) {
    const rec = wb.pivotCacheRecordAdd(cacheId).index;
    assert.ok(wb.pivotCacheRecordSetText(cacheId, rec, 0, region).ok);
    assert.ok(wb.pivotCacheRecordSetNumber(cacheId, rec, 1, amount).ok);
  }
  const pivot = wb.pivotCreate(0, 'Pivot1', cacheId, 0, 4);
  assert.ok(pivot.status.ok, JSON.stringify(pivot.status));
  assert.ok(wb.pivotFieldAdd(0, pivot.index, { sourceName: 'Region', axis: 0 }).status.ok);
  const amountField = wb.pivotFieldAdd(0, pivot.index, { sourceName: 'Amount', axis: 2 });
  assert.ok(amountField.status.ok);
  assert.ok(
    wb.pivotDataFieldAdd(0, pivot.index, {
      name: 'Sum of Amount',
      fieldIndex: amountField.index,
      aggregation: 0,
    }).status.ok,
  );
  return { wb, pivot: pivot.index };
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
    assert.ok(wb.pivotFilterAdd(0, pivot, PIVOT_FILTER).ok);
    assert.equal(wb.pivotFilterCount(0, pivot).value, 1);
    const got = wb.pivotFilterAt(0, pivot, 0);
    assert.ok(got.status.ok, JSON.stringify(got.status));
    assert.equal(got.axis, PIVOT_FILTER.axis);
    assert.equal(got.fieldName, PIVOT_FILTER.fieldName);
    assert.equal(got.type, PIVOT_FILTER.type);
    assert.equal(got.valueKind, PIVOT_FILTER.valueKind);
    assert.equal(got.valueDouble, PIVOT_FILTER.valueDouble);
  } finally {
    wb.delete();
  }
});

test('pivotFilterCount counts the filters on the pivot', async () => {
  const Module = await getModule();
  const { wb, pivot } = makePivotWorkbook(Module.Workbook);
  try {
    assert.equal(wb.pivotFilterCount(0, pivot).value, 0);
    assert.ok(wb.pivotFilterAdd(0, pivot, PIVOT_FILTER).ok);
    assert.equal(wb.pivotFilterCount(0, pivot).value, 1);
  } finally {
    wb.delete();
  }
});

test('pivot filters survive save/load', async () => {
  const Module = await getModule();
  const { wb, pivot } = makePivotWorkbook(Module.Workbook);
  let bytes;
  try {
    assert.ok(wb.pivotFilterAdd(0, pivot, PIVOT_FILTER).ok);
    assert.equal(wb.pivotFilterCount(0, pivot).value, 1);
    const saved = wb.save();
    assert.ok(saved.status.ok, JSON.stringify(saved.status));
    bytes = saved.bytes;
  } finally {
    wb.delete();
  }
  const reloaded = Module.Workbook.loadBytes(bytes);
  try {
    assert.equal(reloaded.pivotCount(0).value, 1);
    // The filter is written as `<filters>` and decodes back into the list.
    assert.equal(reloaded.pivotFilterCount(0, 0).value, 1);
    const got = reloaded.pivotFilterAt(0, 0, 0);
    assert.ok(got.status.ok, JSON.stringify(got.status));
    assert.equal(got.fieldName, PIVOT_FILTER.fieldName);
    assert.equal(got.type, PIVOT_FILTER.type);
    assert.equal(got.valueKind, PIVOT_FILTER.valueKind);
    assert.equal(got.valueDouble, PIVOT_FILTER.valueDouble);
  } finally {
    reloaded.delete();
  }
});

test('pivotFilterAt rejects an out-of-range index', async () => {
  const Module = await getModule();
  const { wb, pivot } = makePivotWorkbook(Module.Workbook);
  try {
    const got = wb.pivotFilterAt(0, pivot, 99);
    assert.equal(got.status.ok, false);
  } finally {
    wb.delete();
  }
});

test('a required pivot spec string field left out of the object is rejected, not silently empty', async () => {
  const Module = await getModule();
  const { wb, pivot } = makePivotWorkbook(Module.Workbook);
  try {
    const fieldCountBefore = wb.pivotFieldCount(0, pivot).value;
    const field = wb.pivotFieldAdd(0, pivot, { axis: 0 }); // no `sourceName`
    assert.equal(field.status.ok, false);
    assert.equal(field.status.status, 7001);
    assert.equal(wb.pivotFieldCount(0, pivot).value, fieldCountBefore, 'no field must have been added');

    // `name` is optional: an omitted one is derived from the workbook profile.
    const dataFieldCountBefore = wb.pivotDataFieldCount(0, pivot).value;
    const dataField = wb.pivotDataFieldAdd(0, pivot, { fieldIndex: 0, aggregation: 0 }); // no `name`
    assert.ok(dataField.status.ok, JSON.stringify(dataField.status));
    assert.equal(wb.pivotDataFieldCount(0, pivot).value, dataFieldCountBefore + 1);
    assert.equal(wb.pivotDataFieldAt(0, pivot, dataField.index).name, 'Sum of Region');

    const filterCountBefore = wb.pivotFilterCount(0, pivot).value;
    const filter = wb.pivotFilterAdd(0, pivot, { axis: 0, type: 1, valueKind: 1, valueDouble: 15 }); // no `fieldName`
    assert.equal(filter.ok, false);
    assert.equal(filter.status, 7001);
    assert.equal(wb.pivotFilterCount(0, pivot).value, filterCountBefore, 'no filter must have been added');
  } finally {
    wb.delete();
  }
});

test('versionString returns a non-empty string', async () => {
  const Module = await getModule();
  const v = Module.versionString();
  assert.equal(typeof v, 'string');
  assert.ok(v.length > 0, `expected non-empty version, got ${JSON.stringify(v)}`);
});

test('evalFormula(=SUM(1,2,3)) returns NUMBER 6', async () => {
  const Module = await getModule();
  const r = Module.evalFormula('=SUM(1,2,3)');
  assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
  assert.equal(r.value.kind, VAL.NUMBER);
  assert.equal(r.value.number, 6);
});

test('evalFormula(=1/0) surfaces Excel error as a value (not failed status)', async () => {
  const Module = await getModule();
  const r = Module.evalFormula('=1/0');
  assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
  assert.equal(r.value.kind, VAL.ERROR);
});

test('evalFormula uses read-only evaluation for the intersection operator', async () => {
  const Module = await getModule();
  const r = Module.evalFormula('=A1 B1');
  assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
  assert.equal(r.value.kind, VAL.ERROR);
  assert.equal(r.value.errorCode, 0); // ErrorCode::Null / #NULL!
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
  } finally {
    wb.delete();
  }
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
    wb.delete();
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
    wb.delete();
  }
});

test('pivot definition read-backs report what the setters wrote', async () => {
  const Module = await getModule();
  const { wb, pivot } = makePivotWorkbook(Module.Workbook);
  try {
    assert.equal(wb.pivotCacheId(0, pivot).index, wb.pivotCacheIdAt(0).index);
    assert.equal(wb.pivotName(0, pivot).value, 'Pivot1');
    assert.deepEqual((({ rows, cols }) => ({ rows, cols }))(wb.pivotGetGrandTotals(0, pivot)), {
      rows: true,
      cols: true,
    });
    assert.ok(wb.pivotSetGrandTotals(0, pivot, false, true).ok);
    const totals = wb.pivotGetGrandTotals(0, pivot);
    assert.ok(totals.status.ok);
    assert.equal(totals.rows, false);
    assert.equal(totals.cols, true);

    assert.deepEqual(wb.pivotRowFieldOrder(0, pivot).indices, [0]);
    assert.deepEqual(wb.pivotColFieldOrder(0, pivot).indices, []);

    assert.ok(wb.pivotFieldSetSort(0, pivot, 0, 0, 'Sum of Amount').ok);
    const region = wb.pivotFieldAt(0, pivot, 0);
    assert.ok(region.status.ok, JSON.stringify(region.status));
    assert.equal(region.sourceName, 'Region');
    assert.equal(region.axis, 0);
    assert.equal(region.cacheFieldIndex, 0);
    assert.equal(region.sortOrder, 0);
    assert.equal(region.sortByField, 'Sum of Amount');
    assert.equal(region.hasDateGroup, false);
    assert.equal(region.dateStartYear, -1);
    assert.equal(wb.pivotFieldAt(0, pivot, 1).axis, 2);

    assert.ok(wb.pivotFieldAddItem(0, pivot, 0, 'East', true).ok);
    assert.ok(wb.pivotFieldAddItem(0, pivot, 0, 'West', true).ok);
    assert.ok(wb.pivotFieldSetItemVisible(0, pivot, 0, 1, false).ok);
    assert.equal(wb.pivotFieldItemCount(0, pivot, 0).value, 2);
    const east = wb.pivotFieldItemAt(0, pivot, 0, 0);
    assert.equal(east.name, 'East');
    // The scratch cache holds no shared items, so a label binds to nothing.
    assert.equal(east.cacheIndex, -1);
    assert.equal(east.visible, true);
    assert.equal(wb.pivotFieldItemAt(0, pivot, 0, 1).visible, false);
    assert.ok(wb.pivotFieldAddItemAt(0, pivot, 0, 5, true).ok);
    const bound = wb.pivotFieldItemAt(0, pivot, 0, 2);
    assert.equal(bound.cacheIndex, 5);
    assert.equal(bound.name, '');

    const dataField = wb.pivotDataFieldAt(0, pivot, 0);
    assert.ok(dataField.status.ok, JSON.stringify(dataField.status));
    assert.equal(dataField.name, 'Sum of Amount');
    assert.equal(dataField.fieldIndex, 1);
    assert.equal(dataField.aggregation, 0);
    assert.equal(dataField.showAsBaseField, -1);
    assert.equal(dataField.showAsBaseItem, -1);
  } finally {
    wb.delete();
  }
});

test('pivot style info defaults on create, reads back what the setter wrote and rejects bad arguments', async () => {
  const Module = await getModule();
  const { wb, pivot } = makePivotWorkbook(Module.Workbook);
  try {
    const initial = wb.pivotGetStyleInfo(0, pivot);
    assert.ok(initial.status.ok, JSON.stringify(initial.status));
    assert.equal(initial.present, true);
    assert.equal(initial.name, 'PivotStyleLight16');
    assert.deepEqual(
      [initial.showRowHeaders, initial.showColHeaders, initial.showRowStripes, initial.showColStripes],
      [true, true, false, false],
    );
    assert.equal(initial.showLastColumn, true);

    assert.ok(wb.pivotSetStyleInfo(0, pivot, { name: 'PivotStyleDark3', showRowStripes: true }).ok);
    const set = wb.pivotGetStyleInfo(0, pivot);
    assert.equal(set.name, 'PivotStyleDark3');
    assert.deepEqual(
      [set.showRowHeaders, set.showColHeaders, set.showRowStripes, set.showColStripes, set.showLastColumn],
      [false, false, true, false, false],
    );

    // Survives a save/load.
    const saved = wb.save();
    assert.ok(saved.status.ok, JSON.stringify(saved.status));
    const reloaded = Module.Workbook.loadBytes(saved.bytes);
    try {
      assert.equal(reloaded.pivotGetStyleInfo(0, 0).name, 'PivotStyleDark3');
      assert.equal(reloaded.pivotGetStyleInfo(0, 0).showRowStripes, true);
    } finally {
      reloaded.delete();
    }

    // An omitted name keeps the element without a style name.
    assert.ok(wb.pivotSetStyleInfo(0, pivot, {}).ok);
    const bare = wb.pivotGetStyleInfo(0, pivot);
    assert.equal(bare.present, true);
    assert.equal(bare.name, '');

    assert.equal(wb.pivotGetStyleInfo(0, 9).status.ok, false);
    assert.equal(wb.pivotSetStyleInfo(0, 9, { name: 'X' }).ok, false);
  } finally {
    wb.delete();
  }
});

test('pivot definition read-backs reject out-of-range indices', async () => {
  const Module = await getModule();
  const { wb, pivot } = makePivotWorkbook(Module.Workbook);
  try {
    for (const r of [
      wb.pivotCacheId(0, 9),
      wb.pivotName(0, 9),
      wb.pivotGetGrandTotals(9, pivot),
      wb.pivotRowFieldOrder(0, 9),
      wb.pivotColFieldOrder(0, 9),
      wb.pivotFieldAt(0, pivot, 99),
      wb.pivotFieldItemCount(0, pivot, 99),
      wb.pivotFieldItemAt(0, pivot, 0, 99),
      wb.pivotDataFieldAt(0, pivot, 99),
    ]) {
      const status = r.status;
      assert.equal(status.ok, false, JSON.stringify(status));
    }
    assert.equal(wb.pivotFieldAt(0, pivot, 99).axis, -1);
  } finally {
    wb.delete();
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
      assert.ok(wb.setExcelProfileId(profile).ok);
      for (const spec of [
        { fieldIndex: 1, aggregation: 0 },
        { name: '', fieldIndex: 1, aggregation: 0 },
      ]) {
        const added = wb.pivotDataFieldAdd(0, pivot, spec);
        assert.ok(added.status.ok, JSON.stringify(added.status));
        assert.equal(wb.pivotDataFieldAt(0, pivot, added.index).name, expected);
      }
      assert.ok(wb.pivotDataFieldSet(0, pivot, 0, { fieldIndex: 1, aggregation: 2 }).ok);
      assert.ok(wb.pivotDataFieldAt(0, pivot, 0).name.endsWith(' Amount'));
    }
    assert.equal(wb.pivotDataFieldAdd(0, pivot, { fieldIndex: 99, aggregation: 0 }).status.ok, false);
  } finally {
    wb.delete();
  }
});

test('active sheet is recorded in bookViews and survives a round trip', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  assert.ok(wb.addSheet('B').ok);
  assert.ok(wb.addSheet('C').ok);
  assert.equal(wb.activeSheet().index, 0);
  assert.ok(wb.setActiveSheet(2).ok);
  const active = wb.activeSheet();
  assert.ok(active.status.ok, JSON.stringify(active.status));
  assert.equal(active.index, 2);
  assert.equal(wb.setActiveSheet(9).ok, false);
  assert.equal(wb.activeSheet().index, 2);
  const saved = wb.save();
  assert.ok(saved.status.ok, JSON.stringify(saved.status));
  const loaded = Module.Workbook.loadBytes(saved.bytes);
  try {
    assert.equal(loaded.activeSheet().index, 2);
  } finally {
    loaded.delete();
    wb.delete();
  }
});

test('the show-phonetic flag survives value writes and a round trip', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  assert.equal(wb.getCellShowPhonetic(0, 0, 0).show, false);
  assert.ok(wb.setText(0, 0, 0, '漢字').ok);
  assert.ok(wb.setCellShowPhonetic(0, 0, 0, true).ok);
  assert.ok(wb.setText(0, 0, 0, '漢字').ok);
  assert.equal(wb.getCellShowPhonetic(0, 0, 0).show, true);
  const saved = wb.save();
  assert.ok(saved.status.ok, JSON.stringify(saved.status));
  const loaded = Module.Workbook.loadBytes(saved.bytes);
  try {
    assert.equal(loaded.getCellShowPhonetic(0, 0, 0).show, true);
    assert.ok(loaded.setCellShowPhonetic(0, 0, 0, false).ok);
    assert.equal(loaded.getCellShowPhonetic(0, 0, 0).show, false);
    const rejected = loaded.getCellShowPhonetic(99, 0, 0);
    assert.equal(rejected.status.ok, false);
    assert.equal(rejected.show, false);
    assert.equal(loaded.setCellShowPhonetic(99, 0, 0, true).ok, false);
  } finally {
    loaded.delete();
    wb.delete();
  }
});

test('an explicit all-zero phoneticPr is reported as present and survives a round trip', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  assert.ok(wb.setText(0, 0, 0, '漢字').ok);
  assert.ok(wb.setCellPhoneticRuns(0, 0, 0, [{ sb: 0, eb: 2, text: 'かんじ' }]).ok);
  assert.equal(wb.getCellPhoneticProperties(0, 0, 0).present, false);
  assert.ok(wb.setCellPhoneticProperties(0, 0, 0, { fontId: 0, type: 0, alignment: 0 }).ok);
  assert.equal(wb.getCellPhoneticProperties(0, 0, 0).present, true);
  const saved = wb.save();
  assert.ok(saved.status.ok, JSON.stringify(saved.status));
  const loaded = Module.Workbook.loadBytes(saved.bytes);
  try {
    const got = loaded.getCellPhoneticProperties(0, 0, 0);
    assert.equal(got.present, true);
    assert.deepEqual([got.fontId, got.type, got.alignment], [0, 0, 0]);
  } finally {
    loaded.delete();
    wb.delete();
  }
});

test('scroll origin, pane origin and selection round-trip through save', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  assert.ok(wb.setSheetFreeze(0, 2, 1).ok);
  const view = wb.getSheetView(0).view;
  assert.equal(view.paneTopLeftRow, 2);
  assert.equal(view.paneTopLeftCol, 1);
  assert.ok(wb.setSheetTopLeft(0, 5, 0).ok);
  assert.ok(wb.setSheetPaneTopLeft(0, 40, 6).ok);
  assert.equal(wb.setSheetPaneTopLeft(0, 1, 6).ok, false);
  const ranges = [
    { firstRow: 41, firstCol: 7, lastRow: 43, lastCol: 8 },
    { firstRow: 50, firstCol: 9, lastRow: 50, lastCol: 9 },
  ];
  assert.ok(wb.setSheetSelection(0, { activeRow: 41, activeCol: 7, ranges }).ok);
  assert.equal(
    wb.setSheetSelection(0, {
      activeRow: 0,
      activeCol: 0,
      ranges: [{ firstRow: 3, firstCol: 0, lastRow: 1, lastCol: 0 }],
    }).ok,
    false,
  );
  const saved = wb.save();
  assert.ok(saved.status.ok, JSON.stringify(saved.status));
  const loaded = Module.Workbook.loadBytes(saved.bytes);
  try {
    const got = loaded.getSheetView(0).view;
    assert.deepEqual([got.topLeftRow, got.topLeftCol, got.paneTopLeftRow, got.paneTopLeftCol], [5, 0, 40, 6]);
    const selection = loaded.getSheetSelection(0);
    assert.ok(selection.status.ok, JSON.stringify(selection.status));
    assert.equal(selection.activeRow, 41);
    assert.equal(selection.activeCol, 7);
    assert.deepEqual(
      selection.ranges.map((r) => [r.firstRow, r.firstCol, r.lastRow, r.lastCol]),
      [
        [41, 7, 43, 8],
        [50, 9, 50, 9],
      ],
    );
    const missing = loaded.getSheetSelection(99);
    assert.equal(missing.status.ok, false);
    assert.deepEqual(missing.ranges, []);
  } finally {
    loaded.delete();
    wb.delete();
  }
});

test('sheet split reads back, round-trips and exchanges with the freeze', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    assert.deepEqual([wb.getSheetView(0).view.splitX, wb.getSheetView(0).view.splitY], [0, 0]);
    assert.ok(wb.setSheetSplit(0, 2000, 1500).ok);
    const view = wb.getSheetView(0).view;
    assert.deepEqual([view.splitX, view.splitY, view.freezeRows, view.freezeCols], [2000, 1500, 0, 0]);
    const saved = wb.save();
    assert.ok(saved.status.ok, JSON.stringify(saved.status));
    const loaded = Module.Workbook.loadBytes(saved.bytes);
    try {
      const got = loaded.getSheetView(0).view;
      assert.deepEqual([got.splitX, got.splitY, got.freezeRows, got.freezeCols], [2000, 1500, 0, 0]);
    } finally {
      loaded.delete();
    }
    assert.ok(wb.setSheetFreeze(0, 2, 1).ok);
    assert.deepEqual([wb.getSheetView(0).view.splitX, wb.getSheetView(0).view.splitY], [0, 0]);
    assert.ok(wb.setSheetSplit(0, 0, 900).ok);
    const split = wb.getSheetView(0).view;
    assert.deepEqual([split.splitX, split.splitY, split.freezeRows, split.freezeCols], [0, 900, 0, 0]);
    assert.equal(wb.setSheetSplit(0, -1, 0).ok, false);
    assert.equal(wb.setSheetSplit(0, 0, Number.NaN).ok, false);
    assert.equal(wb.setSheetSplit(0, Number.POSITIVE_INFINITY, 0).ok, false);
    assert.equal(wb.getSheetView(0).view.splitY, 900);
    assert.ok(wb.setSheetSplit(0, 0, 0).ok);
    assert.equal(wb.getSheetView(0).view.splitY, 0);
  } finally {
    wb.delete();
  }
});

test('pivotFieldAdd binds by cache column name and axis -1 takes a field off the report', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    const cacheId = wb.pivotCacheCreate(0).index;
    assert.ok(wb.pivotCacheSetWorksheetSource(cacheId, { present: true, ref: 'A1:D3', sheet: 'Sheet1' }).ok);
    for (const name of ['Region', 'Product', 'Sales', 'Qty']) assert.ok(wb.pivotCacheFieldAdd(cacheId, name).status.ok);
    for (const sales of [100, 200]) {
      const rec = wb.pivotCacheRecordAdd(cacheId).index;
      for (const [col, value] of [0, 0, sales, 1].entries())
        assert.ok(wb.pivotCacheRecordSetNumber(cacheId, rec, col, value).ok);
    }
    const pivot = wb.pivotCreate(0, 'PT', cacheId, 0, 6).index;
    const region = wb.pivotFieldAdd(0, pivot, { sourceName: 'Region', axis: 0 });
    const sales = wb.pivotFieldAdd(0, pivot, { sourceName: 'Sales', axis: 2 });
    assert.equal(region.index, 0);
    assert.equal(sales.index, 2);
    assert.equal(wb.pivotFieldCount(0, pivot).value, 4);
    assert.equal(wb.pivotFieldAt(0, pivot, 1).axis, -1);
    assert.ok(wb.pivotDataFieldAdd(0, pivot, { fieldIndex: 2, aggregation: 0 }).status.ok);
    assert.equal(wb.pivotDataFieldCount(0, pivot).value, 1);

    assert.ok(wb.pivotFieldSetAxis(0, pivot, 2, -1).ok);
    assert.equal(wb.pivotFieldAt(0, pivot, 2).axis, -1);
    assert.equal(wb.pivotDataFieldCount(0, pivot).value, 0);
    assert.ok(wb.pivotFieldSetAxis(0, pivot, 0, -1).ok);
    assert.deepEqual(wb.pivotRowFieldOrder(0, pivot).indices, []);
    assert.equal(wb.pivotFieldSetAxis(0, pivot, 0, 4).ok, false);
    assert.equal(wb.pivotFieldSetAxis(0, pivot, 0, -2).ok, false);
  } finally {
    wb.delete();
  }
});
