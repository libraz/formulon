#!/usr/bin/env node
// Copyright 2026 libraz. Licensed under the MIT License.
//
// Stages the Formulon N-API addon + JS shim into
// packages/npm-native/dist/ for publication.
//
// Inputs (defaults; override via --build-dir / --out-dir / --platform-arch):
//   --build-dir       build
//   --out-dir         packages/npm-native/dist
//   --platform-arch   <process.platform>-<process.arch>  (e.g. darwin-arm64)
//
// Copies:
//   <build-dir>/bin/formulon.node   -> <out-dir>/prebuilds/<platform-arch>/formulon.node
//   packages/npm-native/index.mjs   -> <out-dir>/index.mjs
//   packages/npm-native/index.d.ts  -> <out-dir>/index.d.ts
//
// Run via `make node-package`. No npm dependencies; Node 18 stdlib only.

import { access, copyFile, mkdir } from 'node:fs/promises';
import { constants as FS } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
// Repo root is two levels above this script (packages/npm-native/scripts/).
const repoRoot = path.resolve(__dirname, '..', '..', '..');
const pkgRoot = path.resolve(__dirname, '..');

function parseArgs(argv) {
  const out = {
    buildDir: 'build',
    outDir: 'packages/npm-native/dist',
    platformArch: `${process.platform}-${process.arch}`,
  };
  for (let i = 0; i < argv.length; ++i) {
    const a = argv[i];
    if (a === '--build-dir') {
      out.buildDir = argv[++i];
    } else if (a === '--out-dir') {
      out.outDir = argv[++i];
    } else if (a === '--platform-arch') {
      out.platformArch = argv[++i];
    } else if (a === '-h' || a === '--help') {
      console.log('Usage: stage.mjs [--build-dir <path>] [--out-dir <path>] [--platform-arch <tag>]');
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
  const { buildDir, outDir, platformArch } = parseArgs(process.argv.slice(2));
  const absBuildDir = path.resolve(repoRoot, buildDir);
  const absOutDir = path.resolve(repoRoot, outDir);
  const prebuildDir = path.join(absOutDir, 'prebuilds', platformArch);

  const nodeSrc = path.join(absBuildDir, 'bin', 'formulon.node');
  const mjsSrc = path.join(pkgRoot, 'index.mjs');
  const dtsSrc = path.join(pkgRoot, 'index.d.ts');

  for (const p of [nodeSrc, mjsSrc, dtsSrc]) {
    if (!(await fileExists(p))) {
      console.error(`stage.mjs: missing input ${p}`);
      console.error('  Run `make node-native` first to produce the addon.');
      process.exit(1);
    }
  }

  await mkdir(prebuildDir, { recursive: true });

  await copyFile(nodeSrc, path.join(prebuildDir, 'formulon.node'));
  await copyFile(mjsSrc, path.join(absOutDir, 'index.mjs'));
  await copyFile(dtsSrc, path.join(absOutDir, 'index.d.ts'));

  console.log(`staged 3 file(s) -> ${absOutDir} (prebuild slot: ${platformArch})`);
}

main().catch((e) => {
  console.error('stage.mjs: fatal:', e && e.stack ? e.stack : e);
  process.exit(1);
});
