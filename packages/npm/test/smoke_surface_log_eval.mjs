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

test('pivotFilterCount reports only what this session added', async () => {
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

test('active filters are session state and do not survive save/load', async () => {
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
    // The pivot round-trips; its active-filter list deliberately does not.
    assert.equal(reloaded.pivotFilterCount(0, 0).value, 0);
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

    const dataFieldCountBefore = wb.pivotDataFieldCount(0, pivot).value;
    const dataField = wb.pivotDataFieldAdd(0, pivot, { fieldIndex: 0, aggregation: 0 }); // no `name`
    assert.equal(dataField.status.ok, false);
    assert.equal(dataField.status.status, 7001);
    assert.equal(wb.pivotDataFieldCount(0, pivot).value, dataFieldCountBefore, 'no data field must have been added');

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
};
const PROFILE_LOCALES = ['ja_JP', 'en_US', 'de_DE', 'fr_FR', 'zh_CN', 'ko_KR', 'th_TH'];

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
