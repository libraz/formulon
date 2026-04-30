#!/usr/bin/env node
// Copyright 2026 libraz. Licensed under the MIT License.
//
// Stages WASM build artefacts into packages/npm/dist/ for publication.
//
// Inputs (defaults; override via --build-dir / --out-dir):
//   --build-dir build-wasm
//   --out-dir   packages/npm/dist
//
// Copies:
//   <build-dir>/formulon.js   -> <out-dir>/formulon.js
//   <build-dir>/formulon.wasm -> <out-dir>/formulon.wasm
//   src/wasm/formulon.d.ts    -> <out-dir>/formulon.d.ts
//
// Run via `make npm-package`. No npm dependencies; Node 18 stdlib only.

import { access, copyFile, mkdir } from 'node:fs/promises';
import { constants as FS } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
// Repo root is two levels above this script (packages/npm/scripts/).
const repoRoot = path.resolve(__dirname, '..', '..', '..');

function parseArgs(argv) {
  const out = { buildDir: 'build-wasm', outDir: 'packages/npm/dist' };
  for (let i = 0; i < argv.length; ++i) {
    const a = argv[i];
    if (a === '--build-dir') {
      out.buildDir = argv[++i];
    } else if (a === '--out-dir') {
      out.outDir = argv[++i];
    } else if (a === '-h' || a === '--help') {
      console.log('Usage: stage.mjs [--build-dir <path>] [--out-dir <path>]');
      process.exit(0);
    } else {
      console.error(`stage.mjs: unknown argument: ${a}`);
      process.exit(2);
    }
  }
  return out;
}

async function fileExists(p) {
  try {
    await access(p, FS.R_OK);
    return true;
  } catch {
    return false;
  }
}

async function main() {
  const { buildDir, outDir } = parseArgs(process.argv.slice(2));
  const absBuildDir = path.resolve(repoRoot, buildDir);
  const absOutDir = path.resolve(repoRoot, outDir);

  const jsSrc = path.join(absBuildDir, 'formulon.js');
  const wasmSrc = path.join(absBuildDir, 'formulon.wasm');
  const dtsSrc = path.join(repoRoot, 'src', 'wasm', 'formulon.d.ts');

  for (const p of [jsSrc, wasmSrc, dtsSrc]) {
    if (!(await fileExists(p))) {
      console.error(`stage.mjs: missing input ${p}`);
      console.error('  Run `make wasm` first to produce the WASM artefacts.');
      process.exit(1);
    }
  }

  await mkdir(absOutDir, { recursive: true });

  await copyFile(jsSrc, path.join(absOutDir, 'formulon.js'));
  await copyFile(wasmSrc, path.join(absOutDir, 'formulon.wasm'));
  await copyFile(dtsSrc, path.join(absOutDir, 'formulon.d.ts'));

  console.log(`staged 3 file(s) -> ${absOutDir}`);
}

main().catch((e) => {
  console.error('stage.mjs: fatal:', e && e.stack ? e.stack : e);
  process.exit(1);
});
