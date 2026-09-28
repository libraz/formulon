#!/usr/bin/env node
//
// Stages WASM build artefacts into packages/npm/dist/ for publication.
//
// Inputs (defaults; override via --build-dir / --threads-build-dir / --out-dir):
//   --build-dir         build-wasm          (single-threaded, package root)
//   --threads-build-dir build-wasm-threads  (pthread, `./threads`)
//   --out-dir           packages/npm/dist
//
// Copies:
//   <build-dir>/formulon.js                   -> <out-dir>/formulon_core.js
//   <build-dir>/formulon.wasm                 -> <out-dir>/formulon.wasm
//   <threads-build-dir>/formulon_threads.js   -> <out-dir>/formulon_threads_core.js
//   <threads-build-dir>/formulon_threads.wasm -> <out-dir>/formulon_threads.wasm
//   packages/npm/index.mjs                    -> <out-dir>/formulon.js
//   packages/npm/threads.mjs                  -> <out-dir>/formulon_threads.js
//   packages/npm/common.mjs                   -> <out-dir>/common.js
//   src/wasm/formulon.d.ts                    -> <out-dir>/formulon.d.ts
//
// The pthread glue spawns its workers from the URL it was linked under,
// formulon_threads.js, which staging gives to the entry shim. A bundler
// honouring `"sideEffects"` can reduce that shim, as a worker entry, to an
// empty chunk, so the spawn URL is rewritten to the core itself, whose top
// level starts the pthread runtime.
//
// Run via `make npm-package`. No npm dependencies; Node 18 stdlib only.

import { constants as FS } from 'node:fs';
import { access, copyFile, mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
// Repo root is two levels above this script (packages/npm/scripts/).
const repoRoot = path.resolve(__dirname, '..', '..', '..');

function parseArgs(argv) {
  const out = { buildDir: 'build-wasm', threadsBuildDir: 'build-wasm-threads', outDir: 'packages/npm/dist' };
  for (let i = 0; i < argv.length; ++i) {
    const a = argv[i];
    if (a === '--build-dir') {
      out.buildDir = argv[++i];
    } else if (a === '--threads-build-dir') {
      out.threadsBuildDir = argv[++i];
    } else if (a === '--out-dir') {
      out.outDir = argv[++i];
    } else if (a === '-h' || a === '--help') {
      console.log('Usage: stage.mjs [--build-dir <path>] [--threads-build-dir <path>] [--out-dir <path>]');
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
  const { buildDir, threadsBuildDir, outDir } = parseArgs(process.argv.slice(2));
  const absBuildDir = path.resolve(repoRoot, buildDir);
  const absThreadsBuildDir = path.resolve(repoRoot, threadsBuildDir);
  const absOutDir = path.resolve(repoRoot, outDir);
  const pkgDir = path.join(repoRoot, 'packages', 'npm');

  // [source, staged name]
  const copies = [
    [path.join(absBuildDir, 'formulon.js'), 'formulon_core.js'],
    [path.join(absBuildDir, 'formulon.wasm'), 'formulon.wasm'],
    [path.join(absThreadsBuildDir, 'formulon_threads.js'), 'formulon_threads_core.js'],
    [path.join(absThreadsBuildDir, 'formulon_threads.wasm'), 'formulon_threads.wasm'],
    [path.join(pkgDir, 'index.mjs'), 'formulon.js'],
    [path.join(pkgDir, 'threads.mjs'), 'formulon_threads.js'],
    [path.join(pkgDir, 'common.mjs'), 'common.js'],
    [path.join(repoRoot, 'src', 'wasm', 'formulon.d.ts'), 'formulon.d.ts'],
  ];

  for (const [src] of copies) {
    if (!(await fileExists(src))) {
      console.error(`stage.mjs: missing input ${src}`);
      console.error('  Run `make wasm wasm-threads` first to produce the WASM artefacts.');
      process.exit(1);
    }
  }

  await mkdir(absOutDir, { recursive: true });
  for (const [src, name] of copies) {
    await copyFile(src, path.join(absOutDir, name));
  }

  const threadsCore = path.join(absOutDir, 'formulon_threads_core.js');
  const spawnUrl = 'new URL("formulon_threads.js",import.meta.url)';
  const glue = await readFile(threadsCore, 'utf8');
  const sites = glue.split(spawnUrl).length - 1;
  if (sites !== 1) {
    console.error(`stage.mjs: expected exactly 1 pthread spawn URL in ${threadsCore}, found ${sites}`);
    console.error(`  looked for: ${spawnUrl}`);
    process.exit(1);
  }
  await writeFile(threadsCore, glue.replace(spawnUrl, 'new URL("formulon_threads_core.js",import.meta.url)'));

  console.log(`staged ${copies.length} file(s) -> ${absOutDir}`);
}

main().catch((e) => {
  console.error('stage.mjs: fatal:', e && e.stack ? e.stack : e);
  process.exit(1);
});
