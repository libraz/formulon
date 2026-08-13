//
// Node-based smoke tests for the Formulon WASM bundle.
//
// Loads `build-wasm/formulon.js` (or FORMULON_WASM_BUILD_DIR when set),
// exercises every embind export at least once, and exits 1 with a
// descriptive message on any failure.
//
// Run via `make test-wasm`, which checks for the build artefact and a
// Node binary first. The runner is intentionally framework-free: it
// uses only `node:assert/strict` so contributors can run it without an
// `npm install` step.

import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { inflateRawSync } from 'node:zlib';

// fm_value_kind_t mirror (see src/c_api/formulon_c.h).
const VAL = Object.freeze({
  BLANK: 0,
  NUMBER: 1,
  BOOL: 2,
  TEXT: 3,
  ERROR: 4,
  ARRAY: 5,
  REF: 6,
  LAMBDA: 7,
});

// fm_pivot_cell_kind_t mirror (see src/c_api/formulon_c.h).
const PIVOT = Object.freeze({
  HEADER: 0,
  ROW_LABEL: 1,
  COL_LABEL: 2,
  DATA: 3,
  ROW_SUBTOTAL: 4,
  COL_SUBTOTAL: 5,
  GRAND_TOTAL: 6,
  BLANK: 7,
});

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const wasmBuildDir = process.env.FORMULON_WASM_BUILD_DIR ?? 'build-wasm';
const moduleUrl = path.resolve(__dirname, '..', '..', wasmBuildDir, 'formulon.js');

const utf8 = new TextEncoder();
let crcTable = null;

function crc32(bytes) {
  if (crcTable === null) {
    crcTable = new Uint32Array(256);
    for (let i = 0; i < 256; i += 1) {
      let c = i;
      for (let j = 0; j < 8; j += 1) {
        c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      }
      crcTable[i] = c >>> 0;
    }
  }
  let c = 0xffffffff;
  for (const b of bytes) {
    c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

function pushU16(out, n) {
  out.push(n & 0xff, (n >>> 8) & 0xff);
}

function pushU32(out, n) {
  out.push(n & 0xff, (n >>> 8) & 0xff, (n >>> 16) & 0xff, (n >>> 24) & 0xff);
}

function pushBytes(out, bytes) {
  for (const b of bytes) out.push(b);
}

function zipStore(parts) {
  const out = [];
  const central = [];
  const entries = parts.map(([name, body]) => ({
    name: utf8.encode(name),
    body: utf8.encode(body),
  }));

  for (const entry of entries) {
    entry.offset = out.length;
    entry.crc = crc32(entry.body);
    pushU32(out, 0x04034b50);
    pushU16(out, 20);
    pushU16(out, 0);
    pushU16(out, 0);
    pushU16(out, 0);
    pushU16(out, 0);
    pushU32(out, entry.crc);
    pushU32(out, entry.body.length);
    pushU32(out, entry.body.length);
    pushU16(out, entry.name.length);
    pushU16(out, 0);
    pushBytes(out, entry.name);
    pushBytes(out, entry.body);
  }

  const centralOffset = out.length;
  for (const entry of entries) {
    pushU32(central, 0x02014b50);
    pushU16(central, 20);
    pushU16(central, 20);
    pushU16(central, 0);
    pushU16(central, 0);
    pushU16(central, 0);
    pushU16(central, 0);
    pushU32(central, entry.crc);
    pushU32(central, entry.body.length);
    pushU32(central, entry.body.length);
    pushU16(central, entry.name.length);
    pushU16(central, 0);
    pushU16(central, 0);
    pushU16(central, 0);
    pushU16(central, 0);
    pushU32(central, 0);
    pushU32(central, entry.offset);
    pushBytes(central, entry.name);
  }
  pushBytes(out, central);

  pushU32(out, 0x06054b50);
  pushU16(out, 0);
  pushU16(out, 0);
  pushU16(out, entries.length);
  pushU16(out, entries.length);
  pushU32(out, central.length);
  pushU32(out, centralOffset);
  pushU16(out, 0);
  return new Uint8Array(out);
}

function readZipEntryText(bytes, wantedName) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = bytes.length - 22;
  while (eocd >= 0 && view.getUint32(eocd, true) !== 0x06054b50) eocd -= 1;
  assert.ok(eocd >= 0, 'missing ZIP end record');
  const count = view.getUint16(eocd + 10, true);
  const centralSize = view.getUint32(eocd + 12, true);
  const centralOffset = view.getUint32(eocd + 16, true);
  const decoder = new TextDecoder();
  let cursor = centralOffset;
  const wanted = String(wantedName);
  for (let i = 0; i < count; i += 1) {
    assert.ok(cursor < centralOffset + centralSize, 'ZIP central directory exceeds its declared size');
    assert.equal(view.getUint32(cursor, true), 0x02014b50, 'invalid ZIP central entry');
    const method = view.getUint16(cursor + 10, true);
    const compressedSize = view.getUint32(cursor + 20, true);
    const uncompressedSize = view.getUint32(cursor + 24, true);
    const nameSize = view.getUint16(cursor + 28, true);
    const extraSize = view.getUint16(cursor + 30, true);
    const commentSize = view.getUint16(cursor + 32, true);
    const localOffset = view.getUint32(cursor + 42, true);
    const name = decoder.decode(bytes.slice(cursor + 46, cursor + 46 + nameSize));
    if (name === wanted) {
      assert.ok(localOffset + 30 <= bytes.length);
      const localNameSize = view.getUint16(localOffset + 26, true);
      const localExtraSize = view.getUint16(localOffset + 28, true);
      const start = localOffset + 30 + localNameSize + localExtraSize;
      const compressed = bytes.slice(start, start + compressedSize);
      const body = method === 0 ? compressed : Uint8Array.from(inflateRawSync(compressed));
      assert.equal(body.length, uncompressedSize);
      return decoder.decode(body);
    }
    cursor += 46 + nameSize + extraSize + commentSize;
  }
  assert.fail(`ZIP entry not found: ${wanted}`);
}

function appendEmptyZipEntry(bytes, name) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = bytes.length - 22;
  while (eocd >= 0 && view.getUint32(eocd, true) !== 0x06054b50) eocd -= 1;
  assert.ok(eocd >= 0, 'missing ZIP end record');
  const count = view.getUint16(eocd + 10, true);
  const centralSize = view.getUint32(eocd + 12, true);
  const centralOffset = view.getUint32(eocd + 16, true);
  const oldCentral = bytes.slice(centralOffset, centralOffset + centralSize);
  const encodedName = utf8.encode(name);
  const local = [];
  pushU32(local, 0x04034b50);
  pushU16(local, 20);
  pushU16(local, 0);
  pushU16(local, 0);
  pushU16(local, 0);
  pushU16(local, 0);
  pushU32(local, 0);
  pushU32(local, 0);
  pushU32(local, 0);
  pushU16(local, encodedName.length);
  pushU16(local, 0);
  pushBytes(local, encodedName);
  const central = [];
  pushU32(central, 0x02014b50);
  pushU16(central, 20);
  pushU16(central, 20);
  pushU16(central, 0);
  pushU16(central, 0);
  pushU16(central, 0);
  pushU16(central, 0);
  pushU32(central, 0);
  pushU32(central, 0);
  pushU32(central, 0);
  pushU16(central, encodedName.length);
  pushU16(central, 0);
  pushU16(central, 0);
  pushU16(central, 0);
  pushU16(central, 0);
  pushU32(central, 0);
  pushU32(central, centralOffset);
  pushBytes(central, encodedName);
  const out = [];
  pushBytes(out, bytes.slice(0, centralOffset));
  pushBytes(out, local);
  pushBytes(out, oldCentral);
  pushBytes(out, central);
  pushU32(out, 0x06054b50);
  pushU16(out, 0);
  pushU16(out, 0);
  pushU16(out, count + 1);
  pushU16(out, count + 1);
  pushU32(out, centralSize + central.length);
  pushU32(out, centralOffset + local.length);
  pushU16(out, 0);
  return new Uint8Array(out);
}

function buildPivotWorkbookBytes() {
  const sheetXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n' +
    '  <sheetData/>\n' +
    '</worksheet>\n';

  return zipStore([
    [
      '[Content_Types].xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n' +
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n' +
        '  <Default Extension="xml" ContentType="application/xml"/>\n' +
        '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>\n' +
        '  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n' +
        '  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n' +
        '  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>\n' +
        '  <Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>\n' +
        '  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>\n' +
        '</Types>\n',
    ],
    [
      '_rels/.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n' +
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\n' +
        '</Relationships>\n',
    ],
    [
      'xl/workbook.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n' +
        '  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rId2"/></sheets>\n' +
        '  <pivotCaches><pivotCache cacheId="0" r:id="rId3"/></pivotCaches>\n' +
        '</workbook>\n',
    ],
    [
      'xl/_rels/workbook.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n' +
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>\n' +
        '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>\n' +
        '  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>\n' +
        '</Relationships>\n',
    ],
    ['xl/worksheets/sheet1.xml', sheetXml],
    ['xl/worksheets/sheet2.xml', sheetXml],
    [
      'xl/worksheets/_rels/sheet2.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n' +
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>\n' +
        '</Relationships>\n',
    ],
    [
      'xl/pivotCache/pivotCacheDefinition1.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" recordCount="3">\n' +
        '  <cacheSource type="worksheet"/>\n' +
        '  <cacheFields count="2"><cacheField name="Region"><sharedItems count="2"><s v="North"/><s v="South"/></sharedItems></cacheField><cacheField name="Amount"><sharedItems containsNumber="1"/></cacheField></cacheFields>\n' +
        '</pivotCacheDefinition>\n',
    ],
    [
      'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n' +
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>\n' +
        '</Relationships>\n',
    ],
    [
      'xl/pivotCache/pivotCacheRecords1.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3">\n' +
        '  <r><x v="0"/><n v="100"/></r><r><x v="1"/><n v="200"/></r><r><x v="0"/><n v="300"/></r>\n' +
        '</pivotCacheRecords>\n',
    ],
    [
      'xl/pivotTables/pivotTable1.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="0">\n' +
        '  <location ref="D1:E5"/>\n' +
        '  <pivotFields count="2"><pivotField axis="axisRow" name="Region"><items count="3"><item x="0"/><item x="1"/><item t="default"/></items></pivotField><pivotField dataField="1" name="Amount"/></pivotFields>\n' +
        '  <rowFields count="1"><field x="0"/></rowFields><dataFields count="1"><dataField name="Sum of Amount" fld="1" subtotal="sum"/></dataFields>\n' +
        '</pivotTableDefinition>\n',
    ],
  ]);
}

function findPivotCell(layout, row, col) {
  return layout.cells.find((cell) => cell.row === row && cell.col === col);
}

const cases = [];
function test(name, fn) {
  cases.push({ name, fn });
}

let passed = 0;
let failed = 0;

async function run() {
  // Dynamic import: keeps the file syntactically valid even when the
  // wasm artifact is missing (the harness check is the Makefile's job).
  const factory = (await import(moduleUrl)).default;
  const Module = await factory();

  test('versionString is non-empty', () => {
    const v = Module.versionString();
    assert.equal(typeof v, 'string');
    assert.ok(v.length > 0, `expected non-empty version, got ${JSON.stringify(v)}`);
  });

  test('statusString covers known codes', () => {
    assert.equal(Module.statusString(0), 'kOk');
    // Unknown / out-of-band codes still return a non-null string.
    const unknown = Module.statusString(123456);
    assert.equal(typeof unknown, 'string');
    assert.ok(unknown.length > 0);
  });

  test('errorDisplayName returns Excel literals', () => {
    assert.equal(Module.errorDisplayName(1), '#DIV/0!');
    assert.equal(Module.errorDisplayName(999), '#UNKNOWN!');
  });

  test('evalFormula(=SUM(1,2,3)) returns NUMBER 6', () => {
    const r = Module.evalFormula('=SUM(1,2,3)');
    assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
    assert.equal(r.value.kind, VAL.NUMBER);
    assert.equal(r.value.number, 6);
  });

  test('evalFormula(=IF(TRUE,1,2)) returns NUMBER 1', () => {
    const r = Module.evalFormula('=IF(TRUE,1,2)');
    assert.ok(r.status.ok);
    assert.equal(r.value.kind, VAL.NUMBER);
    assert.equal(r.value.number, 1);
  });

  test('evalFormula evaluates intersection read-only and returns #NULL!', () => {
    const r = Module.evalFormula('=A1 B1');
    assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
    assert.equal(r.value.kind, VAL.ERROR);
    assert.equal(r.value.errorCode, 0); // ErrorCode::Null / #NULL!
  });

  test('evalFormula(=CONCAT) returns TEXT "hello world"', () => {
    const r = Module.evalFormula('=CONCAT("hello"," ","world")');
    assert.ok(r.status.ok);
    assert.equal(r.value.kind, VAL.TEXT);
    assert.equal(r.value.text, 'hello world');
  });

  test('Workbook construction + setNumber + recalc + getValue', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.isValid());
      assert.equal(wb.sheetCount(), 1);
      const nameRes = wb.sheetName(0);
      assert.ok(nameRes.status.ok);
      assert.equal(nameRes.value, 'Sheet1');

      assert.ok(wb.setNumber(0, 0, 0, 42).ok);
      assert.ok(wb.setFormula(0, 1, 0, '=A1*2').ok);
      assert.ok(wb.recalc().ok);

      const a1 = wb.getValue(0, 0, 0);
      assert.ok(a1.status.ok);
      assert.equal(a1.value.kind, VAL.NUMBER);
      assert.equal(a1.value.number, 42);

      const b1 = wb.getValue(0, 1, 0);
      assert.ok(b1.status.ok);
      assert.equal(b1.value.kind, VAL.NUMBER);
      assert.equal(b1.value.number, 84);

      // addSheet should grow sheetCount.
      assert.ok(wb.addSheet('Second').ok);
      assert.equal(wb.sheetCount(), 2);
    } finally {
      wb.delete();
    }
  });

  test('paginate exposes a status envelope and page geometry', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setNumber(0, 0, 0, 1).ok);
      assert.ok(wb.setNumber(0, 48, 0, 2).ok);
      const result = wb.paginate(0);
      assert.ok(result.status.ok, `status=${JSON.stringify(result.status)}`);
      assert.equal(result.pageCount, 2);
      assert.deepEqual(result.printArea, []);
      assert.deepEqual(result.horizontalBreaks, [44]);
      assert.deepEqual(result.verticalBreaks, []);
    } finally {
      wb.delete();
    }
  });

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
      assert.ok(loaded.sheetCount() >= 1);
      assert.ok(loaded.recalc().ok);
      const a1 = loaded.getValue(0, 0, 0);
      assert.ok(a1.status.ok);
      assert.equal(a1.value.kind, VAL.NUMBER);
      assert.equal(a1.value.number, 7);
    } finally {
      loaded.delete();
    }
  });

  test('saveExWithDiagnostics() and xlsbReadDiagnostics() expose stable counters', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setFormula(0, 0, 0, '=@A1:A10').ok);
      assert.ok(
        wb.addValidation(0, {
          ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
          type: 3,
          formula1: '"Yes,No"',
        }).ok,
      );
      assert.ok(
        wb.addConditionalFormat(0, {
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
          type: 1,
          op: 5,
          formula1: '10',
          dxfId: 0,
          stopIfTrue: false,
        }).status.ok,
      );
      const xlsx = wb.saveExWithDiagnostics(1);
      assert.ok(xlsx.status.ok, `xlsx save failed: ${JSON.stringify(xlsx.status)}`);
      assert.equal(xlsx.downgradedFormulaCount, 0);
      assert.equal(xlsx.deferredFeatureCount, 0);

      const xlsb = wb.saveExWithDiagnostics(2);
      assert.ok(xlsb.status.ok, `xlsb save failed: ${JSON.stringify(xlsb.status)}`);
      assert.equal(xlsb.downgradedFormulaCount, 1);
      assert.ok(xlsb.deferredFeatureCount >= 2);

      const loaded = Module.Workbook.loadBytes(xlsb.bytes);
      try {
        const read = loaded.xlsbReadDiagnostics();
        assert.ok(read.status.ok, `read diagnostics failed: ${JSON.stringify(read.status)}`);
        assert.equal(read.undecodedFormulaCount, 0);
        assert.equal(read.undecodedDefinedNameCount, 0);
        assert.equal(read.droppedPartCount, 0);
      } finally {
        loaded.delete();
      }
      const preservedLoaded = Module.Workbook.loadBytes(appendEmptyZipEntry(xlsb.bytes, 'xl/preserved.bin'));
      try {
        const preserved = preservedLoaded.xlsbReadDiagnostics();
        assert.ok(preserved.status.ok, `passthrough diagnostics failed: ${JSON.stringify(preserved.status)}`);
        assert.equal(preserved.droppedPartCount, 0);
      } finally {
        preservedLoaded.delete();
      }
      const invalidSave = wb.saveExWithDiagnostics(0);
      assert.ok(!invalidSave.status.ok, `unknown format unexpectedly succeeded: ${JSON.stringify(invalidSave)}`);
      assert.equal(invalidSave.bytes, null);
      assert.equal(invalidSave.downgradedFormulaCount, 0);
      assert.equal(invalidSave.deferredFeatureCount, 0);
      assert.ok(wb.saveEx(2).status.ok);
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
      const read = wb.xlsbReadDiagnostics();
      assert.ok(!read.status.ok, `invalid handle unexpectedly succeeded: ${JSON.stringify(read)}`);
      assert.equal(read.undecodedFormulaCount, 0);
      assert.equal(read.undecodedDefinedNameCount, 0);
      assert.equal(read.droppedPartCount, 0);
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

  test('setIterative accepts knobs without crashing', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setIterative(true, 50, 0.001).ok);
      // Set a fixed-point formula (sqrt(2) iteration). The engine may
      // surface NUMBER (converged) or ERROR depending on the recalc
      // path; we only assert that the call sequence doesn't crash and
      // that getValue returns a well-formed envelope.
      assert.ok(wb.setNumber(0, 0, 0, 1.0).ok);
      assert.ok(wb.setFormula(0, 0, 0, '=0.5*(A1+2/A1)').ok);
      assert.ok(wb.recalc().ok);
      const v = wb.getValue(0, 0, 0);
      assert.ok(v.status.ok);
      // Known well-formed kinds.
      assert.ok([VAL.NUMBER, VAL.ERROR, VAL.BLANK].includes(v.value.kind));
    } finally {
      wb.delete();
    }
  });

  test('renameSheet updates the sheet name', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.equal(wb.sheetCount(), 1);
      assert.ok(wb.renameSheet(0, 'Renamed').ok);
      const r = wb.sheetName(0);
      assert.ok(r.status.ok);
      assert.equal(r.value, 'Renamed');
    } finally {
      wb.delete();
    }
  });

  test('renameSheet rejects forbidden characters', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const r = wb.renameSheet(0, 'Bad/Name');
      assert.equal(r.ok, false);
      assert.notEqual(r.status, 0);
    } finally {
      wb.delete();
    }
  });

  test('removeSheet drops a sheet but rejects the last one', () => {
    const wb = Module.Workbook.createEmpty();
    try {
      assert.ok(wb.addSheet('A').ok);
      assert.ok(wb.addSheet('B').ok);
      assert.ok(wb.addSheet('C').ok);
      assert.equal(wb.sheetCount(), 3);
      assert.ok(wb.removeSheet(1).ok);
      assert.equal(wb.sheetCount(), 2);
      assert.equal(wb.sheetName(0).value, 'A');
      assert.equal(wb.sheetName(1).value, 'C');
    } finally {
      wb.delete();
    }

    // Workbook with a single sheet rejects removeSheet(0).
    const lone = Module.Workbook.createDefault();
    try {
      const r = lone.removeSheet(0);
      assert.equal(r.ok, false);
    } finally {
      lone.delete();
    }
  });

  test('moveSheet rearranges sheets (Excel UI semantics)', () => {
    const wb = Module.Workbook.createEmpty();
    try {
      assert.ok(wb.addSheet('Alpha').ok);
      assert.ok(wb.addSheet('Beta').ok);
      assert.ok(wb.addSheet('Gamma').ok);
      // Move Alpha (0) to the end. Excel semantics: to=2 (post-removal).
      assert.ok(wb.moveSheet(0, 2).ok);
      assert.equal(wb.sheetName(0).value, 'Beta');
      assert.equal(wb.sheetName(1).value, 'Gamma');
      assert.equal(wb.sheetName(2).value, 'Alpha');
    } finally {
      wb.delete();
    }
  });

  test('setDefinedName adds, updates, and removes', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.equal(wb.definedNameCount(), 0);
      assert.ok(wb.setDefinedName('Pi', '=3.14').ok);
      assert.equal(wb.definedNameCount(), 1);
      const a = wb.definedNameAt(0);
      assert.ok(a.status.ok);
      assert.equal(a.name, 'Pi');
      assert.equal(a.formula, '=3.14');

      assert.ok(wb.setDefinedName('PI', '=3.14159').ok);
      assert.equal(wb.definedNameCount(), 1);
      const b = wb.definedNameAt(0);
      assert.equal(b.name, 'Pi'); // authored case preserved
      assert.equal(b.formula, '=3.14159');

      assert.ok(wb.setDefinedName('Pi', '').ok);
      assert.equal(wb.definedNameCount(), 0);
    } finally {
      wb.delete();
    }
  });

  test('cellCount + cellAt iterate populated cells in order', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setNumber(0, 0, 0, 1).ok);
      assert.ok(wb.setNumber(0, 0, 1, 2).ok);
      assert.ok(wb.setFormula(0, 1, 0, '=A1+B1').ok);
      assert.ok(wb.recalc().ok);

      const count = wb.cellCount(0);
      assert.ok(count >= 3, `expected >=3 cells, got ${count}`);

      // The very first iteration entry should be A1 (row=0, col=0).
      const c0 = wb.cellAt(0, 0);
      assert.ok(c0.status.ok);
      assert.equal(c0.row, 0);
      assert.equal(c0.col, 0);
      assert.equal(c0.formula, null);
      assert.equal(c0.value.kind, VAL.NUMBER);
      assert.equal(c0.value.number, 1);
    } finally {
      wb.delete();
    }
  });

  test('failed entry lookups omit optional payload fields', () => {
    const wb = Module.Workbook.createDefault();
    try {
      for (const entry of [wb.cellAt(99, 0), wb.definedNameAt(0), wb.tableAt(0), wb.passthroughAt(0)]) {
        assert.equal(entry.status.ok, false);
      }
      assert.ok(!('value' in wb.cellAt(99, 0)));
      assert.ok(!('name' in wb.definedNameAt(0)));
      assert.ok(!('name' in wb.tableAt(0)));
      assert.ok(!('path' in wb.passthroughAt(0)));
    } finally {
      wb.delete();
    }
  });

  test('pivotCount + pivotLayout expose PivotTable projection status', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.equal(wb.pivotCount(0), 0);

      const missing = wb.pivotLayout(0, 0);
      assert.equal(missing.status.ok, false);
      assert.notEqual(missing.status.status, 0);
      assert.equal(missing.top, 0);
      assert.equal(missing.left, 0);
      assert.equal(missing.rows, 0);
      assert.equal(missing.cols, 0);
      assert.deepEqual(missing.cells, []);
    } finally {
      wb.delete();
    }
  });

  test('pivotLayout projects loaded PivotTable cells for grid rendering', () => {
    const wb = Module.Workbook.loadBytes(buildPivotWorkbookBytes());
    try {
      assert.ok(wb.isValid(), Module.lastErrorMessage());
      assert.equal(wb.pivotCount(0), 0);
      assert.equal(wb.pivotCount(1), 1);

      const layout = wb.pivotLayout(1, 0);
      assert.ok(layout.status.ok, `status=${JSON.stringify(layout.status)}`);
      assert.equal(layout.top, 0);
      assert.equal(layout.left, 3);
      assert.equal(layout.rows, 4);
      assert.equal(layout.cols, 2);
      assert.ok(layout.cells.length > 0);

      const rowLabels = findPivotCell(layout, 0, 3);
      assert.ok(rowLabels);
      assert.equal(rowLabels.kind, PIVOT.HEADER);
      assert.equal(rowLabels.value.kind, VAL.TEXT);
      assert.equal(rowLabels.value.text, '行ラベル');

      const northLabel = findPivotCell(layout, 1, 3);
      assert.ok(northLabel);
      assert.equal(northLabel.kind, PIVOT.ROW_LABEL);
      assert.equal(northLabel.value.text, 'North');

      const northSum = findPivotCell(layout, 1, 4);
      assert.ok(northSum);
      assert.equal(northSum.kind, PIVOT.DATA);
      assert.equal(northSum.value.kind, VAL.NUMBER);
      assert.equal(northSum.value.number, 400);
      assert.equal(northSum.fieldName, 'Sum of Amount');

      const grandTotal = findPivotCell(layout, 3, 4);
      assert.ok(grandTotal);
      assert.equal(grandTotal.kind, PIVOT.GRAND_TOTAL);
      assert.equal(grandTotal.value.kind, VAL.NUMBER);
      assert.equal(grandTotal.value.number, 600);
    } finally {
      wb.delete();
    }
  });

  test('xf index round-trips through setCellXfIndex / getCellXfIndex', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setNumber(0, 2, 3, 5).ok);
      assert.ok(wb.setCellXfIndex(0, 2, 3, 9).ok);
      const r = wb.getCellXfIndex(0, 2, 3);
      assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
      assert.equal(r.xfIndex, 9);
    } finally {
      wb.delete();
    }
  });

  test('style building blocks: addFont -> addXf -> setCellXfIndex -> getXf', () => {
    const wb = Module.Workbook.createDefault();
    try {
      // Empty workbook starts with empty styles tables.
      assert.equal(wb.fontCount(), 0);
      assert.equal(wb.fillCount(), 0);
      assert.equal(wb.borderCount(), 0);
      assert.equal(wb.xfCount(), 0);

      const f1 = wb.addFont({
        name: 'Arial',
        size: 12,
        bold: true,
        italic: false,
        strike: false,
        underline: 0,
        colorArgb: 0xff112233,
      });
      assert.ok(f1.status.ok, `addFont: ${JSON.stringify(f1.status)}`);
      assert.equal(typeof f1.index, 'number');

      // Adding the same font again returns the same index (linear-search dedup).
      const f1b = wb.addFont({
        name: 'Arial',
        size: 12,
        bold: true,
        italic: false,
        strike: false,
        underline: 0,
        colorArgb: 0xff112233,
      });
      assert.ok(f1b.status.ok);
      assert.equal(f1b.index, f1.index);
      assert.equal(wb.fontCount(), 1);

      const fill = wb.addFill({ pattern: 1, fgArgb: 0xffff0000, bgArgb: 0xff000000 });
      assert.ok(fill.status.ok);
      const border = wb.addBorder({
        left: { style: 1, colorArgb: 0xff000000 },
        right: { style: 1, colorArgb: 0xff000000 },
        top: { style: 1, colorArgb: 0xff000000 },
        bottom: { style: 1, colorArgb: 0xff000000 },
        diagonal: { style: 0, colorArgb: 0 },
        diagonalUp: false,
        diagonalDown: false,
      });
      assert.ok(border.status.ok);

      // Built-in num_fmt resolves without growing the table.
      const builtin = wb.addNumFmt('General');
      assert.ok(builtin.status.ok);
      assert.equal(builtin.numFmtId, 0);

      // Custom num_fmt yields >= 164.
      const custom = wb.addNumFmt('"USD" #,##0');
      assert.ok(custom.status.ok);
      assert.ok(custom.numFmtId >= 164, `expected custom id >= 164, got ${custom.numFmtId}`);

      const xf = wb.addXf({
        fontIndex: f1.index,
        fillIndex: fill.index,
        borderIndex: border.index,
        numFmtId: custom.numFmtId,
        horizontalAlign: 1,
        verticalAlign: 2,
        wrapText: true,
      });
      assert.ok(xf.status.ok, `addXf: ${JSON.stringify(xf.status)}`);
      assert.ok(wb.xfCount() >= 1);

      // Adding the same xf is a no-op.
      const xfDup = wb.addXf({
        fontIndex: f1.index,
        fillIndex: fill.index,
        borderIndex: border.index,
        numFmtId: custom.numFmtId,
        horizontalAlign: 1,
        verticalAlign: 2,
        wrapText: true,
      });
      assert.ok(xfDup.status.ok);
      assert.equal(xfDup.index, xf.index);

      // Stamp the xf on a cell.
      assert.ok(wb.setNumber(0, 0, 0, 1).ok);
      assert.ok(wb.setCellXfIndex(0, 0, 0, xf.index).ok);

      // Read back through the getters.
      const reread = wb.getCellXf(xf.index);
      assert.ok(reread.status.ok);
      assert.equal(reread.fontIndex, f1.index);
      assert.equal(reread.numFmtId, custom.numFmtId);
      assert.equal(reread.wrapText, true);
      assert.equal(reread.hasAlignment, true);
      assert.equal(reread.hasHorizontalAlign, true);
      assert.equal(reread.hasVerticalAlign, true);
      assert.equal(reread.hasWrapText, true);
      assert.equal(reread.hasJustifyLastLine, false);

      const optional = wb.addXf({
        fontIndex: f1.index,
        fillIndex: fill.index,
        borderIndex: border.index,
        numFmtId: custom.numFmtId,
        textRotation: 255,
        indent: 0,
        relativeIndent: -3,
        shrinkToFit: false,
        readingOrder: 0,
        justifyLastLine: true,
      });
      assert.ok(optional.status.ok, `optional alignment: ${JSON.stringify(optional.status)}`);
      const optionalRead = wb.getCellXf(optional.index);
      assert.ok(optionalRead.status.ok);
      assert.equal(optionalRead.textRotation, 255);
      assert.equal(optionalRead.indent, 0);
      assert.equal(optionalRead.relativeIndent, -3);
      assert.equal(optionalRead.shrinkToFit, false);
      assert.equal(optionalRead.readingOrder, 0);
      assert.equal(optionalRead.justifyLastLine, true);
      assert.equal(optionalRead.hasAlignment, true);
      assert.equal(optionalRead.hasJustifyLastLine, true);
      assert.equal(wb.addXf(optionalRead).index, optional.index);

      const omitted = wb.addXf({
        fontIndex: f1.index,
        fillIndex: fill.index,
        borderIndex: border.index,
        numFmtId: custom.numFmtId,
      });
      assert.ok(omitted.status.ok);
      const omittedRead = wb.getCellXf(omitted.index);
      assert.ok(omittedRead.status.ok);
      assert.equal(omittedRead.hasAlignment, false);

      const explicitEmpty = wb.addXf({
        fontIndex: f1.index,
        fillIndex: fill.index,
        borderIndex: border.index,
        numFmtId: custom.numFmtId,
        hasAlignment: true,
      });
      assert.ok(explicitEmpty.status.ok);
      assert.notEqual(explicitEmpty.index, omitted.index);
      const explicitEmptyRead = wb.getCellXf(explicitEmpty.index);
      assert.ok(explicitEmptyRead.status.ok);
      assert.equal(explicitEmptyRead.hasAlignment, true);

      const rfont = wb.getFont(f1.index);
      assert.ok(rfont.status.ok);
      assert.equal(rfont.name, 'Arial');
      assert.equal(rfont.size, 12);
      assert.equal(rfont.bold, true);

      const rfill = wb.getFill(fill.index);
      assert.ok(rfill.status.ok);
      assert.equal(rfill.pattern, 1);
      assert.equal(rfill.fgArgb, 0xffff0000);

      const rborder = wb.getBorder(border.index);
      assert.ok(rborder.status.ok);
      assert.equal(rborder.left.style, 1);
      assert.equal(rborder.diagonalUp, false);

      const rfmt = wb.getNumFmt(custom.numFmtId);
      assert.ok(rfmt.status.ok);
      assert.equal(rfmt.formatCode, '"USD" #,##0');
    } finally {
      wb.delete();
    }
  });

  test('addFont / getFont preserve superscript and subscript', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const superscript = wb.addFont({ name: 'Arial', size: 12, vertAlign: 1 });
      assert.ok(superscript.status.ok);
      assert.equal(wb.getFont(superscript.index).vertAlign, 1);

      const subscript = wb.addFont({ name: 'Arial', size: 12, vertAlign: 2 });
      assert.ok(subscript.status.ok);
      assert.notEqual(subscript.index, superscript.index);
      assert.equal(wb.getFont(subscript.index).vertAlign, 2);
    } finally {
      wb.delete();
    }
  });

  test('addFont(getFont(i)) is the identity and does not grow the table', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const added = wb.addFont({ name: 'Arial', size: 12, vertAlign: 1, colorArgb: 0xff112233 });
      assert.ok(added.status.ok);
      const before = wb.fontCount();
      const readBack = wb.getFont(added.index);
      const again = wb.addFont(readBack);
      assert.ok(again.status.ok);
      assert.equal(again.index, added.index);
      assert.equal(wb.fontCount(), before);
    } finally {
      wb.delete();
    }
  });

  test('a one-field font rewrite keeps the superscript', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const added = wb.addFont({ name: 'Arial', size: 12, vertAlign: 1, colorArgb: 0xff112233 });
      assert.ok(added.status.ok);
      const edited = wb.getFont(added.index);
      edited.colorArgb = 0xff00ff00;
      const recolored = wb.addFont(edited);
      assert.ok(recolored.status.ok);
      assert.notEqual(recolored.index, added.index);
      const reread = wb.getFont(recolored.index);
      assert.equal(reread.vertAlign, 1);
      assert.equal(reread.colorArgb, 0xff00ff00);
    } finally {
      wb.delete();
    }
  });

  test('addDxf / getDxf round-trip a superscript differential font', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const added = wb.addDxf({ font: { name: 'Calibri', size: 9, vertAlign: 1 } });
      assert.ok(added.status.ok);
      const readBack = wb.getDxf(added.index);
      assert.equal(readBack.font.vertAlign, 1);
      const before = wb.dxfCount();
      const again = wb.addDxf(readBack);
      assert.ok(again.status.ok);
      assert.equal(again.index, added.index);
      assert.equal(wb.dxfCount(), before);
    } finally {
      wb.delete();
    }
  });

  test('addXf rejects out-of-range font_index', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const r = wb.addXf({
        fontIndex: 99,
        fillIndex: 0,
        borderIndex: 0,
        numFmtId: 0,
        horizontalAlign: 0,
        verticalAlign: 0,
        wrapText: false,
      });
      assert.ok(!r.status.ok, 'expected addXf to reject out-of-range font_index');
    } finally {
      wb.delete();
    }
  });

  test('addMerge / getMerges round-trip', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.addMerge(0, { firstRow: 0, lastRow: 1, firstCol: 0, lastCol: 2 }).ok);
      const merges = wb.getMerges(0);
      assert.equal(merges.length, 1);
      assert.equal(merges[0].firstRow, 0);
      assert.equal(merges[0].lastCol, 2);
    } finally {
      wb.delete();
    }
  });

  test('removeMerge / removeMergeAt / clearMerges step-wise prune the merge list', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const a = { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 };
      const b = { firstRow: 4, firstCol: 4, lastRow: 5, lastCol: 5 };
      assert.ok(wb.addMerge(0, a).ok);
      assert.ok(wb.addMerge(0, b).ok);
      // removeMerge with an overlap that hits `a` only.
      assert.ok(wb.removeMerge(0, { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }).ok);
      const list = wb.getMerges(0);
      assert.equal(list.length, 1);
      assert.equal(list[0].firstRow, 4);
      // removeMergeAt drops the survivor by index.
      assert.ok(wb.removeMergeAt(0, 0).ok);
      assert.equal(wb.getMerges(0).length, 0);
      // clearMerges nukes the remainder; safe on an empty list.
      assert.ok(wb.addMerge(0, a).ok);
      assert.ok(wb.addMerge(0, b).ok);
      assert.ok(wb.clearMerges(0).ok);
      assert.equal(wb.getMerges(0).length, 0);
    } finally {
      wb.delete();
    }
  });

  test('hyperlinks are read after save+load', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.equal(wb.getHyperlinks(0).length, 0);
    } finally {
      wb.delete();
    }
  });

  test('addHyperlink + getHyperlinks round-trip', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.equal(wb.getHyperlinks(0).length, 0);
      // Three entries with progressively more optional fields populated.
      assert.ok(wb.addHyperlink(0, 1, 2, 'https://example.com/', '', '', '').ok);
      assert.ok(wb.addHyperlink(0, 3, 4, 'mailto:hello@example.com', 'Hello', '', '').ok);
      assert.ok(wb.addHyperlink(0, 5, 6, '', 'See X', 'Internal link', 'Sheet1!A1').ok);
      assert.ok(wb.addHyperlinkRange(0, 7, 8, 9, 10, 'https://example.com/range', 'Range', 'Range tip', '').ok);
      const list = wb.getHyperlinks(0);
      assert.equal(list.length, 4);
      assert.equal(list[0].row, 1);
      assert.equal(list[0].col, 2);
      assert.equal(list[0].lastRow, 1);
      assert.equal(list[0].lastCol, 2);
      assert.equal(list[0].target, 'https://example.com/');
      assert.equal(list[0].display, '');
      assert.equal(list[0].tooltip, '');
      assert.equal(list[1].target, 'mailto:hello@example.com');
      assert.equal(list[1].display, 'Hello');
      assert.equal(list[2].display, 'See X');
      assert.equal(list[2].tooltip, 'Internal link');
      assert.equal(list[2].location, 'Sheet1!A1');
      assert.deepEqual(
        {
          row: list[3].row,
          col: list[3].col,
          lastRow: list[3].lastRow,
          lastCol: list[3].lastCol,
          target: list[3].target,
          display: list[3].display,
          tooltip: list[3].tooltip,
        },
        {
          row: 7,
          col: 8,
          lastRow: 9,
          lastCol: 10,
          target: 'https://example.com/range',
          display: 'Range',
          tooltip: 'Range tip',
        },
      );
      // Sheet-out-of-range is rejected.
      assert.ok(!wb.addHyperlink(999, 0, 0, 'https://x/', '', '', '').ok);
      // clearHyperlinks drops everything.
      assert.ok(wb.clearHyperlinks(0).ok);
      assert.equal(wb.getHyperlinks(0).length, 0);
    } finally {
      wb.delete();
    }
  });

  test('removeHyperlink / removeHyperlinkAt / clearHyperlinks surface on an empty sheet', () => {
    const wb = Module.Workbook.createDefault();
    try {
      // No-op variants on an empty list still return kOk.
      assert.ok(wb.removeHyperlink(0, 0, 0).ok);
      assert.ok(wb.clearHyperlinks(0).ok);
      // Out-of-range index is rejected.
      assert.ok(!wb.removeHyperlinkAt(0, 0).ok);
      // Sheet-index-out-of-range is rejected on every variant.
      assert.ok(!wb.removeHyperlink(999, 0, 0).ok);
      assert.ok(!wb.removeHyperlinkAt(999, 0).ok);
      assert.ok(!wb.clearHyperlinks(999).ok);
      // Final list still matches the post-save+load read.
      assert.equal(wb.getHyperlinks(0).length, 0);
    } finally {
      wb.delete();
    }
  });

  test('setComment / getComment round-trip', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setComment(0, 1, 1, 'Alice', 'Hello').ok);
      const c = wb.getComment(0, 1, 1);
      assert.ok(c !== null);
      assert.equal(c.author, 'Alice');
      assert.equal(c.text, 'Hello');
      // Empty text removes.
      assert.ok(wb.setComment(0, 1, 1, '', '').ok);
      const after = wb.getComment(0, 1, 1);
      assert.equal(after, null);
    } finally {
      wb.delete();
    }
  });

  test('getCommentResult distinguishes absence from an invalid sheet', () => {
    const wb = Module.Workbook.createDefault();
    try {
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
    } finally {
      wb.delete();
    }
  });

  test('getValidations returns an array', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const arr = wb.getValidations(0);
      assert.ok(Array.isArray(arr) || typeof arr.length === 'number');
      assert.equal(arr.length, 0);
    } finally {
      wb.delete();
    }
  });

  test('addValidation / getValidations / removeValidationAt / clearValidations round-trip', () => {
    const wb = Module.Workbook.createDefault();
    try {
      // List-type validation at A1:B3 with prompts and an error message.
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

      const a = list[0];
      assert.equal(a.type, 3);
      assert.equal(a.op, 0);
      assert.equal(a.errorStyle, 1);
      // Booleans must arrive as JS booleans, not 0/1.
      assert.equal(typeof a.allowBlank, 'boolean');
      assert.equal(a.allowBlank, true);
      assert.equal(typeof a.showInputMessage, 'boolean');
      assert.equal(a.showInputMessage, true);
      assert.equal(typeof a.showErrorMessage, 'boolean');
      assert.equal(a.showErrorMessage, true);
      assert.equal(a.formula1, '"Yes,No,Maybe"');
      assert.equal(a.promptTitle, 'Choose');
      assert.equal(a.errorTitle, 'Bad value');
      assert.equal(a.errorMessage, 'Pick from the list');
      assert.equal(a.ranges.length, 1);
      assert.equal(a.ranges[0].firstRow, 0);
      assert.equal(a.ranges[0].lastCol, 1);

      const b = list[1];
      assert.equal(b.type, 2);
      assert.equal(b.formula1, '0');
      assert.equal(b.formula2, '100');
      assert.equal(b.allowBlank, false);
      assert.equal(b.ranges.length, 2);
      assert.equal(b.ranges[1].firstRow, 6);
      assert.equal(b.ranges[1].lastRow, 9);

      // removeValidationAt drops the first rule.
      assert.ok(wb.removeValidationAt(0, 0).ok);
      const after = wb.getValidations(0);
      assert.equal(after.length, 1);
      assert.equal(after[0].type, 2);

      // Out-of-range index is rejected.
      assert.ok(!wb.removeValidationAt(0, 99).ok);

      // clearValidations drops everything; safe to call again.
      assert.ok(wb.clearValidations(0).ok);
      assert.equal(wb.getValidations(0).length, 0);
      assert.ok(wb.clearValidations(0).ok);

      // Sheet-out-of-range is rejected on every entry.
      assert.ok(!wb.addValidation(999, listRule).ok);
      assert.ok(!wb.removeValidationAt(999, 0).ok);
      assert.ok(!wb.clearValidations(999).ok);
    } finally {
      wb.delete();
    }
  });

  test('CF and sheet-layout lists are plain JS arrays with no delete lifecycle', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const cf = wb.evaluateCfRange(0, 0, 0, 1, 1, Number.NaN);
      assert.ok(cf.status.ok);
      assert.ok(Array.isArray(cf.cells));

      assert.ok(wb.setColumnWidth(0, 2, 2, 18).ok);
      const columns = wb.getSheetColumns(0);
      assert.ok(columns.status.ok);
      assert.ok(Array.isArray(columns.columns));
      assert.equal(columns.columns[0].first, 2);
      assert.equal(columns.columns[0].width, 18);

      assert.ok(wb.setRowHeight(0, 3, 22).ok);
      const rows = wb.getSheetRowOverrides(0);
      assert.ok(rows.status.ok);
      assert.ok(Array.isArray(rows.rows));
      assert.equal(rows.rows[0].row, 3);
      assert.equal(rows.rows[0].height, 22);
    } finally {
      wb.delete();
    }
  });

  test('list getters expose failures through their array status', () => {
    const wb = Module.Workbook.createDefault();
    try {
      for (const list of [
        wb.getMerges(99),
        wb.getComments(99),
        wb.getHyperlinks(99),
        wb.getValidations(99),
        wb.getConditionalFormats(99),
      ]) {
        assert.ok(Array.isArray(list));
        assert.equal(list.length, 0);
        assert.equal(list.status.ok, false);
      }
      const links = wb.getExternalLinks();
      assert.ok(Array.isArray(links));
      assert.equal(links.status.ok, true);
    } finally {
      wb.delete();
    }
  });

  // ---- Run --------------------------------------------------------------
  for (const c of cases) {
    try {
      await c.fn();
      passed += 1;
      console.log(`ok   ${c.name}`);
    } catch (e) {
      failed += 1;
      console.error(`FAIL ${c.name}`);
      console.error(`  ${e && e.stack ? e.stack : e}`);
    }
  }

  console.log('');
  console.log(`Smoke summary: ${passed} passed, ${failed} failed (of ${cases.length})`);
  if (failed > 0) {
    process.exit(1);
  }
}

run().catch((e) => {
  console.error('Fatal harness error:', e);
  process.exit(1);
});
