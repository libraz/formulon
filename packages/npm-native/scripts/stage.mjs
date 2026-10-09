#!/usr/bin/env node
//
// Stages the Formulon N-API addon + JS shim into
// packages/npm-native/dist/ for publication.
//
// Inputs (defaults; override via --build-dir / --out-dir / --platform-arch):
//   --build-dir       build
//   --out-dir         packages/npm-native/dist
//   --platform-arch   <process.platform>-<process.arch>  (e.g. darwin-arm64)
//   --source-only     stage only the JS shim and declarations; skip the addon
//
// Copies:
//   <build-dir>/bin/formulon.node   -> <out-dir>/prebuilds/<platform-arch>/formulon.node
//   packages/npm-native/index.mjs   -> <out-dir>/index.mjs
//   packages/npm/common.mjs         -> <out-dir>/common.mjs
//   packages/npm-native/index.d.ts  -> <out-dir>/index.d.ts
//
// Run via `make node-package` (or `--source-only` from release bundling).
// No npm dependencies; Node 18 stdlib only.

import { constants as FS } from 'node:fs';
import { access, copyFile, mkdir, readFile, writeFile } from 'node:fs/promises';
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
    sourceOnly: false,
  };
  for (let i = 0; i < argv.length; ++i) {
    const a = argv[i];
    if (a === '--build-dir') {
      out.buildDir = argv[++i];
    } else if (a === '--out-dir') {
      out.outDir = argv[++i];
    } else if (a === '--platform-arch') {
      out.platformArch = argv[++i];
    } else if (a === '--source-only') {
      out.sourceOnly = true;
    } else if (a === '-h' || a === '--help') {
      console.log('Usage: stage.mjs [--build-dir <path>] [--out-dir <path>] [--platform-arch <tag>] [--source-only]');
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
  const { buildDir, outDir, platformArch, sourceOnly } = parseArgs(process.argv.slice(2));
  const absBuildDir = path.resolve(repoRoot, buildDir);
  const absOutDir = path.resolve(repoRoot, outDir);
  const prebuildDir = path.join(absOutDir, 'prebuilds', platformArch);

  const nodeSrc = path.join(absBuildDir, 'bin', 'formulon.node');
  const mjsSrc = path.join(pkgRoot, 'index.mjs');
  const commonSrc = path.join(repoRoot, 'packages', 'npm', 'common.mjs');
  const dtsSrc = path.join(pkgRoot, 'index.d.ts');

  const inputs = sourceOnly ? [mjsSrc, commonSrc, dtsSrc] : [nodeSrc, mjsSrc, commonSrc, dtsSrc];
  for (const p of inputs) {
    if (!(await fileExists(p))) {
      console.error(`stage.mjs: missing input ${p}`);
      if (!sourceOnly) console.error('  Run `make node-native` first to produce the addon.');
      process.exit(1);
    }
  }

  await mkdir(sourceOnly ? absOutDir : prebuildDir, { recursive: true });

  if (!sourceOnly) await copyFile(nodeSrc, path.join(prebuildDir, 'formulon.node'));
  const mjs = await readFile(mjsSrc, 'utf8');
  const commonImport = "from '../npm/common.mjs'";
  const commonImportSites = mjs.split(commonImport).length - 1;
  if (commonImportSites !== 1) {
    console.error(`stage.mjs: expected exactly 1 canonical common import in ${mjsSrc}, found ${commonImportSites}`);
    console.error(`  looked for: ${commonImport}`);
    process.exit(1);
  }
  await writeFile(path.join(absOutDir, 'index.mjs'), mjs.replace(commonImport, "from './common.mjs'"));
  await copyFile(commonSrc, path.join(absOutDir, 'common.mjs'));
  await copyFile(dtsSrc, path.join(absOutDir, 'index.d.ts'));

  const stagedFiles = sourceOnly ? 3 : 4;
  const slot = sourceOnly ? '' : ` (prebuild slot: ${platformArch})`;
  console.log(`staged ${stagedFiles} file(s) -> ${absOutDir}${slot}`);
}

main().catch((e) => {
  console.error('stage.mjs: fatal:', e && e.stack ? e.stack : e);
  process.exit(1);
});
