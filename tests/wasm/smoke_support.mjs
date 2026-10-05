import assert from 'node:assert/strict';
import { inflateRawSync } from 'node:zlib';

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

export { appendEmptyZipEntry, buildPivotWorkbookBytes, findPivotCell, PIVOT, readZipEntryText, VAL, zipStore };
