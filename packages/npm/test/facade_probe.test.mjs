// The Workbook / Sheet facade of common.mjs over a stub raw Workbook, without
// an engine build: handle tracking across sheet-structure changes, removed and
// disposed handles, SheetRef resolution, patch expansion, and the public key
// sets of both classes against the `Workbook` / `Sheet` / `WorkbookCtor`
// interfaces of the two published declaration files.

import assert from 'node:assert/strict';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

import * as common from '../common.mjs';
import { memberNames, prototypeKeys } from './dts_members.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(here, '..', '..', '..');
const DTS = {
  wasm: path.join(repoRoot, 'src', 'wasm', 'formulon.d.ts'),
  native: path.join(repoRoot, 'packages', 'npm-native', 'index.d.ts'),
};
const { FormulonError, StatusCode } = common;

// --- stub raw Workbook --------------------------------------------------------

function fail(code, message) {
  return new FormulonError(code, message, '');
}

/**
 * A raw Workbook class: flat methods taking the sheet index first, a sheet list
 * that the structure ops edit, and a call log. Every declared method exists, so
 * the facade's forwarding loop finds it.
 */
function makeRaw(surface) {
  const declared = [...memberNames(DTS[surface], 'Workbook'), ...memberNames(DTS[surface], 'Sheet')];
  const instances = [];

  class RawWorkbook {
    constructor(names) {
      this.names = names;
      this.active = 0;
      this.log = [];
      this.isDisposed = false;
      this.grandTotals = { rows: true, columns: false };
      this.threaded = new Map([['c1', { id: 'c1', text: 'old', mentions: [] }]]);
      instances.push(this);
    }
    static createDefault() {
      return new RawWorkbook(['Sheet1']);
    }
    static createEmpty() {
      return new RawWorkbook([]);
    }
    static createForProfile(_id) {
      return new RawWorkbook(['Sheet1']);
    }
    static load(_bytes) {
      return new RawWorkbook(['Data', 'Calc', 'Notes']);
    }
    get disposed() {
      return this.isDisposed;
    }
    dispose() {
      this.isDisposed = true;
    }
    alive() {
      if (this.isDisposed) throw fail(StatusCode.BindingInvalidHandle, 'the workbook is disposed');
    }
    checkSheet(i) {
      if (!Number.isInteger(i) || i < 0 || i >= this.names.length) {
        throw fail(StatusCode.SheetIndexOutOfRange, `sheet ${i}`);
      }
    }
    getSheetCount() {
      this.alive();
      return this.names.length;
    }
    getSheetName(i) {
      this.alive();
      this.checkSheet(i);
      return this.names[i];
    }
    getSheetIndex(name) {
      this.alive();
      const i = this.names.findIndex((n) => n.toLowerCase() === name.toLowerCase());
      return i < 0 ? null : i;
    }
    addSheet(name) {
      this.alive();
      this.names.push(name);
      return this.names.length - 1;
    }
    copySheet(source, insertBefore, newName) {
      this.alive();
      this.checkSheet(source);
      if (insertBefore > this.names.length) throw fail(StatusCode.SheetIndexOutOfRange, 'insertBefore');
      this.names.splice(insertBefore, 0, newName ?? `${this.names[source]} (2)`);
      return insertBefore;
    }
    moveSheet(from, to) {
      this.alive();
      this.checkSheet(from);
      this.checkSheet(to);
      const [name] = this.names.splice(from, 1);
      this.names.splice(to, 0, name);
    }
    removeSheet(i) {
      this.alive();
      this.checkSheet(i);
      if (this.names.length === 1) throw fail(StatusCode.CannotRemoveLastSheet, 'last sheet');
      this.names.splice(i, 1);
    }
    renameSheet(i, name) {
      this.alive();
      this.checkSheet(i);
      this.names[i] = name;
    }
    activeSheet() {
      this.alive();
      return this.active;
    }
    setActiveSheet(i) {
      this.alive();
      this.checkSheet(i);
      this.active = i;
    }
    showPivotDetails(sheet, _pivotIndex, _row, _col) {
      this.alive();
      this.checkSheet(sheet);
      this.names.splice(sheet, 0, 'Details');
      return sheet;
    }
    parseReference(text) {
      this.alive();
      const m = text.match(/^\$?([A-Za-z]+)\$?(\d+)$/);
      if (!m) return null;
      let col = 0;
      for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
      const row = Number(m[2]) - 1;
      return { sheetName: null, rect: { firstRow: row, firstCol: col - 1, lastRow: row, lastCol: col - 1 } };
    }
    getPivotGrandTotals(_sheet, _pivotIndex) {
      this.alive();
      return { ...this.grandTotals };
    }
    updatePivotGrandTotals(sheet, pivotIndex, rows, columns) {
      this.alive();
      this.log.push(['updatePivotGrandTotals', [sheet, pivotIndex, rows, columns]]);
      this.grandTotals = { rows, columns };
    }
    getThreadedComment(_sheet, id) {
      this.alive();
      return this.threaded.get(id) ?? null;
    }
    updateThreadedComment(sheet, id, patch) {
      this.alive();
      this.log.push(['updateThreadedComment', [sheet, id, patch]]);
      this.threaded.set(id, { ...this.threaded.get(id), ...patch });
    }
    getPivotCount(_sheet) {
      this.alive();
      return 1;
    }
    getPivotRuntimeId(_sheet, pivotIndex) {
      this.alive();
      return pivotIndex + 1;
    }
    resolvePivotRuntimeId(_sheet, id) {
      this.alive();
      if (id !== 1) throw fail(StatusCode.NotFound, 'pivot');
      return 0;
    }
    getCfRuleCount(_sheet) {
      this.alive();
      return 2;
    }
    getCfRule(sheet, i) {
      this.alive();
      return { sheet, i };
    }
  }

  const defined = new Set(Object.getOwnPropertyNames(RawWorkbook.prototype));
  const record = (name) =>
    function (...args) {
      this.alive();
      this.log.push([name, args]);
      return null;
    };
  for (const name of declared) {
    if (name.startsWith('[') || defined.has(name)) continue;
    Object.defineProperty(RawWorkbook.prototype, name, { value: record(name), writable: true, configurable: true });
  }
  return { RawWorkbook, instances };
}

function buildFacade(surface = 'wasm') {
  const { RawWorkbook, instances } = makeRaw(surface);
  const runtime = {
    Workbook: RawWorkbook,
    evaluateFormula: () => null,
    getVersion: () => 'stub',
    getStatusName: () => 'OK',
    getFunctionMetadata: () => null,
    listFunctionNames: () => [],
    localizeFunctionName: (n) => n,
    canonicalizeFunctionName: (n) => n,
    localizeFormula: (f) => f,
    canonicalizeFormula: (f) => f,
    getLocaleFacts: () => null,
    listLocaleErrorNames: () => [],
    compareText: () => 0,
    getCollationClasses: () => null,
    getImageInfo: () => null,
    setLogMinLevel: () => {},
    setLogSink: () => {},
    convertCivilToSerial: () => 0,
    convertSerialToCivil: () => null,
  };
  const formulon = common.buildFormulon(runtime, common);
  return { formulon, Workbook: formulon.Workbook, instances };
}

function expectCode(fn, code) {
  assert.throws(fn, (e) => {
    assert.equal(e.name, 'FormulonError', `expected a FormulonError, got ${e}`);
    assert.equal(e.code, code, `code ${e.code} (${e.message})`);
    return true;
  });
}

function loaded() {
  const { Workbook, instances } = buildFacade();
  const wb = Workbook.load(new Uint8Array());
  return { wb, raw: instances.at(-1), Workbook };
}

/** The sheet index the raw class receives for a call through `ws`. */
function rawSheetOf(raw, ws) {
  raw.log.length = 0;
  ws.setNumber(0, 0, 1);
  const call = raw.log.find(([n]) => n === 'setNumber');
  assert.ok(call, 'no raw setNumber call');
  return call[1][0];
}

// H11: the facade's sheet list matches the raw one after every structure change.
function assertInSync(wb, raw) {
  const handles = wb.listSheets();
  assert.equal(raw.getSheetCount(), handles.length);
  assert.deepEqual(
    handles.map((s) => s.name),
    raw.names,
  );
  for (const [i, s] of handles.entries()) assert.equal(s.index, i);
}

// --- tests --------------------------------------------------------------------

test('the facade replaces the raw class: no public constructor, one class per createFormulon', () => {
  const a = buildFacade();
  const b = buildFacade();
  assert.notEqual(a.Workbook, b.Workbook);
  assert.throws(() => new a.Workbook(), TypeError);
  const wb = a.Workbook.createDefault();
  assert.equal(wb.constructor.name, 'Workbook');
  assert.equal(Object.getPrototypeOf(wb.sheet(0)).constructor.name, 'Sheet');
  assert.notEqual(wb.sheet(0), b.Workbook.createDefault().sheet(0));
});

test('H1: one handle per sheet, by index or case-insensitive name', () => {
  const { wb } = loaded();
  const ws = wb.sheet(0);
  assert.equal(wb.sheet(0), ws);
  assert.equal(wb.sheet('Data'), ws);
  assert.equal(wb.sheet('data'), ws);
  assert.equal(wb.sheet(ws), ws);
  assert.equal(wb.listSheets()[0], ws);
  assert.equal(ws.workbook, wb);
  expectCode(() => wb.sheet('nope'), StatusCode.NotFound);
  expectCode(() => wb.sheet(99), StatusCode.SheetIndexOutOfRange);
  expectCode(() => wb.sheet(-1), StatusCode.SheetIndexOutOfRange);
  assert.throws(() => wb.sheet(1.5), RangeError);
  assert.throws(() => wb.sheet(true), TypeError);
  assert.throws(() => wb.sheet(null), TypeError);
});

test('H2: copySheet before a handle shifts it and returns the copy', () => {
  const { wb, raw } = loaded();
  const calc = wb.sheet('Calc');
  const copy = wb.copySheet(calc, 0, 'Copy');
  assert.equal(copy.name, 'Copy');
  assert.equal(copy.index, 0);
  assert.equal(wb.sheet(0), copy);
  assert.equal(calc.index, 2);
  assert.equal(rawSheetOf(raw, calc), 2);
  assertInSync(wb, raw);
});

test('H3 / H4: Sheet.move and Sheet.rename keep the handle on its sheet', () => {
  const { wb, raw } = loaded();
  const data = wb.sheet('Data');
  data.move(2);
  assert.equal(data.index, 2);
  assert.equal(data.name, 'Data');
  assert.equal(rawSheetOf(raw, data), 2);
  assertInSync(wb, raw);
  wb.sheet('data').rename('Renamed');
  assert.equal(data.name, 'Renamed');
  assert.equal(wb.sheet('Renamed'), data);
  assertInSync(wb, raw);
});

test('H5 / H6 / H15: a removed handle throws NotFound; later handles shift down', () => {
  const { wb, raw } = loaded();
  const [data, calc, notes] = wb.listSheets();
  wb.removeSheet(calc);
  expectCode(() => calc.index, StatusCode.NotFound);
  expectCode(() => calc.name, StatusCode.NotFound);
  expectCode(() => calc.setNumber(0, 0, 1), StatusCode.NotFound);
  assert.equal(data.index, 0);
  assert.equal(notes.index, 1);
  assert.equal(wb.sheet(1), notes);
  assert.equal(rawSheetOf(raw, notes), 1);
  assertInSync(wb, raw);
});

test('H9: Pivot.showDetails returns the new sheet and shifts the handles after it', () => {
  const { wb, raw } = loaded();
  const [, calc, notes] = wb.listSheets();
  const details = calc.pivot(0).showDetails(3, 4);
  assert.equal(details.name, 'Details');
  assert.equal(details.index, 1);
  assert.equal(calc.index, 2);
  assert.equal(notes.index, 3);
  assertInSync(wb, raw);
});

test('SheetRef accepts a Sheet, an index or a name; a foreign Sheet is InvalidArgument', () => {
  const { wb, raw, Workbook } = loaded();
  const added = wb.addSheet('Extra');
  assert.equal(added.index, 3);
  assert.equal(wb.sheet('Extra'), added);
  wb.removeSheet('Extra');
  wb.removeSheet(2);
  assertInSync(wb, raw);
  wb.setActiveSheet('Calc');
  assert.equal(wb.activeSheet(), wb.sheet('Calc'));
  assert.equal(raw.active, 1);
  const other = Workbook.createDefault();
  expectCode(() => wb.removeSheet(other.sheet(0)), StatusCode.InvalidArgument);
  expectCode(() => wb.sheet(other.sheet(0)), StatusCode.InvalidArgument);
});

test('H13 / H7: after dispose, Sheet members throw BindingInvalidHandle', () => {
  const { wb } = loaded();
  const ws = wb.sheet(0);
  wb.dispose();
  assert.equal(wb.disposed, true);
  expectCode(() => ws.index, StatusCode.BindingInvalidHandle);
  expectCode(() => ws.name, StatusCode.BindingInvalidHandle);
  expectCode(() => ws.setNumber(0, 0, 1), StatusCode.BindingInvalidHandle);
  wb.dispose();
});

test('H14: copySheet at the end appends; past the end fails; Sheet.move in place is a no-op', () => {
  const { wb, raw } = loaded();
  const data = wb.sheet(0);
  const copy = wb.copySheet(data, 3);
  assert.equal(copy.index, 3);
  expectCode(() => wb.copySheet(data, 5), StatusCode.SheetIndexOutOfRange);
  assertInSync(wb, raw);
  data.move(data.index);
  assertInSync(wb, raw);
  const single = buildFacade().Workbook.createDefault();
  expectCode(() => single.removeSheet(0), StatusCode.CannotRemoveLastSheet);
  assert.equal(single.listSheets().length, 1);
});

test('A6: a patch leaves omitted and undefined keys unchanged, and refuses null on a non-nullable key', () => {
  const { wb, raw } = loaded();
  const ws = wb.sheet('Calc');
  const pivot = ws.pivot(0);
  pivot.updateGrandTotals({ columns: true });
  assert.deepEqual(raw.log.at(-1), ['updatePivotGrandTotals', [1, 0, true, true]]);
  pivot.updateGrandTotals({ rows: false, columns: undefined });
  assert.deepEqual(raw.log.at(-1), ['updatePivotGrandTotals', [1, 0, false, true]]);
  assert.throws(() => pivot.updateGrandTotals({ rows: null }), TypeError);
  ws.updateThreadedComment('c1', { text: 'new' });
  assert.deepEqual(raw.log.at(-1), ['updateThreadedComment', [1, 'c1', { text: 'new' }]]);
  ws.updateThreadedComment('c1', { done: true });
  assert.deepEqual(raw.log.at(-1), ['updateThreadedComment', [1, 'c1', { done: true }]]);
  assert.throws(() => ws.updateThreadedComment('c1', { done: null }), TypeError);
});

test('the hand-written guards count arguments as the facade method sees them', () => {
  const { wb } = loaded();
  const ws = wb.sheet(0);
  assert.throws(() => ws.setValue(0, 0, []), /^TypeError: Sheet\.setValue argument 3 /);
  assert.throws(() => wb.sheet({}), /^TypeError: Workbook\.sheet argument 1 /);
});

test('list families assemble the raw count and element getter', () => {
  const { wb } = loaded();
  assert.deepEqual(wb.sheet(1).listCfRules(), [
    { sheet: 1, i: 0 },
    { sheet: 1, i: 1 },
  ]);
});

for (const surface of ['wasm', 'native']) {
  test(`S2 (${surface}): the facade's keys are the declared Workbook / Sheet / WorkbookCtor members`, () => {
    const { Workbook } = buildFacade(surface);
    const wb = Workbook.createDefault();
    const declared = (name) => [...memberNames(DTS[surface], name)].sort();
    assert.deepEqual(prototypeKeys(Object.getPrototypeOf(wb)), declared('Workbook'));
    assert.deepEqual(prototypeKeys(Object.getPrototypeOf(wb.sheet(0))), declared('Sheet'));
    const statics = Object.getOwnPropertyNames(Workbook).filter((k) => !['length', 'name', 'prototype'].includes(k));
    assert.deepEqual(statics.sort(), declared('WorkbookCtor'));
  });
}

for (const surface of ['wasm', 'native']) {
  test(`S2 (${surface}): the Formulon object's keys are the declared Formulon members`, () => {
    const { formulon } = buildFacade(surface);
    const declared = [...memberNames(DTS[surface], 'Formulon')].sort();
    assert.deepEqual(Object.keys(formulon).sort(), declared);
  });
}
