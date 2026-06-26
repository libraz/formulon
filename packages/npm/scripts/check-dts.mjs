#!/usr/bin/env node
// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Drift guard for the npm TypeScript surface.
//
// `packages/npm/dist/formulon.d.ts` is a staged COPY of the canonical
// `src/wasm/formulon.d.ts` (see stage.mjs). The two can silently diverge
// when the source is edited but the package is not re-staged, which ships
// a `.d.ts` that omits fields the runtime actually returns (TS consumers
// then get a compile error for a field that exists at runtime).
//
// This script byte-compares the two files and exits non-zero on any
// difference. It needs no WASM build, so it is cheap enough to run in CI
// and from `make npm-check-dts`.
//
// Run via `make npm-check-dts`. No npm dependencies; Node 18 stdlib only.

import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
// Repo root is two levels above this script (packages/npm/scripts/).
const repoRoot = path.resolve(__dirname, '..', '..', '..');

const srcPath = path.join(repoRoot, 'src', 'wasm', 'formulon.d.ts');
const distPath = path.join(repoRoot, 'packages', 'npm', 'dist', 'formulon.d.ts');

async function main() {
  let src;
  let dist;
  try {
    src = await readFile(srcPath, 'utf8');
  } catch (e) {
    console.error(`check-dts: cannot read canonical source ${srcPath}: ${e.message}`);
    process.exit(1);
  }
  try {
    dist = await readFile(distPath, 'utf8');
  } catch (e) {
    console.error(`check-dts: cannot read staged copy ${distPath}: ${e.message}`);
    console.error('  Run `make npm-package` (or copy src/wasm/formulon.d.ts) to stage it.');
    process.exit(1);
  }

  if (src === dist) {
    console.log('check-dts: dist/formulon.d.ts matches src/wasm/formulon.d.ts');
    return;
  }

  console.error('check-dts: dist/formulon.d.ts is OUT OF SYNC with src/wasm/formulon.d.ts');
  console.error('  The dist copy is stale. Re-stage it with:');
  console.error('    cp src/wasm/formulon.d.ts packages/npm/dist/formulon.d.ts');
  console.error('  (or run `make npm-package`, which re-copies it as part of staging).');
  process.exit(1);
}

main().catch((e) => {
  console.error('check-dts: fatal:', e && e.stack ? e.stack : e);
  process.exit(1);
});
