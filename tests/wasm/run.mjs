//
// Node-based smoke tests for the Formulon WASM bundle.
//
// Loads `build-wasm/formulon.js` (or FORMULON_WASM_BUILD_DIR when set;
// FORMULON_WASM_THREADS=1 selects the pthread build's formulon_threads.js),
// exercises every embind export at least once, and exits 1 with a
// descriptive message on any failure.
//
// Run via `make test-wasm`, which checks for the build artefact and a
// Node binary first. The runner is intentionally framework-free: it
// uses only `node:assert/strict` so contributors can run it without an
// `npm install` step.

import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { registerCellsIo } from './run_cells_io.mjs';
import { registerMetadataLayout } from './run_metadata_layout.mjs';
import { registerPivotRegressions } from './run_pivot_regressions.mjs';
import { registerSheetPivot } from './run_sheet_pivot.mjs';
import { registerStyles } from './run_styles.mjs';
import { registerSurfaceRecalc } from './run_surface_recalc.mjs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const wasmBuildDir = process.env.FORMULON_WASM_BUILD_DIR ?? 'build-wasm';
const threadsBuild = process.env.FORMULON_WASM_THREADS === '1';
const moduleUrl = path.resolve(
  __dirname,
  '..',
  '..',
  wasmBuildDir,
  threadsBuild ? 'formulon_threads.js' : 'formulon.js',
);

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

  registerSurfaceRecalc(Module, test, threadsBuild);
  registerCellsIo(Module, test);
  registerSheetPivot(Module, test);
  registerStyles(Module, test);
  registerMetadataLayout(Module, test);
  registerPivotRegressions(Module, test);

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
