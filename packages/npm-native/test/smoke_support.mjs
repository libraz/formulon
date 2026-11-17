// Shared staged-module loader and ZIP helpers for the native smoke-test topics.

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const pkgRoot = path.resolve(__dirname, '..');
const pkgJsonPath = path.join(pkgRoot, 'package.json');

// Resolve the staged entry point through the package's own package.json
// "main" field. Fails (rather than skips) if the package isn't staged --
// that's the whole point of running these tests against dist/.
async function loadStagedModule() {
  const raw = await readFile(pkgJsonPath, 'utf8');
  const pkg = JSON.parse(raw);
  const main = pkg.main;
  if (!main) {
    throw new Error(`package.json is missing "main": ${pkgJsonPath}`);
  }
  const mainPath = path.resolve(pkgRoot, main);
  // Node ESM requires file:// URLs for absolute paths on Windows;
  // pathToFileURL is the portable form.
  return import(pathToFileURL(mainPath).href);
}

// The resolved `Formulon` object (value tables, Workbook, free functions).
let modPromise;
function getModule() {
  if (!modPromise) {
    modPromise = loadStagedModule().then((mod) => mod.createFormulon());
  }
  return modPromise;
}

function appendEmptyZipEntry(bytes, name) {
  const input = new Uint8Array(bytes);
  const view = new DataView(input.buffer, input.byteOffset, input.byteLength);
  let eocd = input.length - 22;
  while (eocd >= 0 && view.getUint32(eocd, true) !== 0x06054b50) eocd -= 1;
  assert.ok(eocd >= 0, 'missing ZIP end record');
  const count = view.getUint16(eocd + 10, true);
  const centralSize = view.getUint32(eocd + 12, true);
  const centralOffset = view.getUint32(eocd + 16, true);
  const encodedName = new TextEncoder().encode(name);
  const u16 = (out, n) => out.push(n & 0xff, (n >>> 8) & 0xff);
  const u32 = (out, n) => out.push(n & 0xff, (n >>> 8) & 0xff, (n >>> 16) & 0xff, (n >>> 24) & 0xff);
  const bytesTo = (out, source) => {
    source.forEach((b) => {
      out.push(b);
    });
  };
  const local = [];
  u32(local, 0x04034b50);
  u16(local, 20);
  u16(local, 0);
  u16(local, 0);
  u16(local, 0);
  u16(local, 0);
  u32(local, 0);
  u32(local, 0);
  u32(local, 0);
  u16(local, encodedName.length);
  u16(local, 0);
  bytesTo(local, encodedName);
  const central = [];
  u32(central, 0x02014b50);
  u16(central, 20);
  u16(central, 20);
  u16(central, 0);
  u16(central, 0);
  u16(central, 0);
  u16(central, 0);
  u32(central, 0);
  u32(central, 0);
  u32(central, 0);
  u16(central, encodedName.length);
  u16(central, 0);
  u16(central, 0);
  u16(central, 0);
  u16(central, 0);
  u32(central, 0);
  u32(central, centralOffset);
  bytesTo(central, encodedName);
  const out = [];
  bytesTo(out, input.slice(0, centralOffset));
  bytesTo(out, local);
  bytesTo(out, input.slice(centralOffset, centralOffset + centralSize));
  bytesTo(out, central);
  u32(out, 0x06054b50);
  u16(out, 0);
  u16(out, 0);
  u16(out, count + 1);
  u16(out, count + 1);
  u32(out, centralSize + central.length);
  u32(out, centralOffset + local.length);
  u16(out, 0);
  return Uint8Array.from(out);
}

const CRC_TABLE = Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});

function crc32(bytes) {
  let c = 0xffffffff;
  for (const b of bytes) c = CRC_TABLE[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

// A stored (uncompressed) ZIP archive of `parts`, each `[name, text]`.
function zipStore(parts) {
  const encoder = new TextEncoder();
  const out = [];
  const central = [];
  const u16 = (to, v) => to.push(v & 0xff, (v >>> 8) & 0xff);
  const u32 = (to, v) => to.push(v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff);
  for (const [name, text] of parts) {
    const encodedName = encoder.encode(name);
    const body = encoder.encode(text);
    const crc = crc32(body);
    const offset = out.length;
    u32(out, 0x04034b50);
    for (const v of [20, 0, 0, 0, 0]) u16(out, v);
    u32(out, crc);
    u32(out, body.length);
    u32(out, body.length);
    u16(out, encodedName.length);
    u16(out, 0);
    out.push(...encodedName, ...body);
    u32(central, 0x02014b50);
    for (const v of [20, 20, 0, 0, 0, 0]) u16(central, v);
    u32(central, crc);
    u32(central, body.length);
    u32(central, body.length);
    for (const v of [encodedName.length, 0, 0, 0, 0]) u16(central, v);
    u32(central, 0);
    u32(central, offset);
    central.push(...encodedName);
  }
  const centralOffset = out.length;
  out.push(...central);
  u32(out, 0x06054b50);
  for (const v of [0, 0, parts.length, parts.length]) u16(out, v);
  u32(out, central.length);
  u32(out, centralOffset);
  u16(out, 0);
  return Uint8Array.from(out);
}

// Predicate for `assert.throws`: a host failure, optionally with a given code.
function isFormulonError(error, code) {
  return error?.name === 'FormulonError' && (code === undefined || error.code === code);
}

export { appendEmptyZipEntry, assert, getModule, isFormulonError, loadStagedModule, path, pkgRoot, readFile, zipStore };
