// Shared staged-module loader and ZIP helper for the native smoke-test topics.

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

let modPromise;
function getModule() {
  if (!modPromise) {
    modPromise = loadStagedModule();
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

export { appendEmptyZipEntry, assert, getModule, loadStagedModule, path, pkgRoot, readFile };
