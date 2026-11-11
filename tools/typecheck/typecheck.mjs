#!/usr/bin/env node
// Type-checks the `js` blocks of the README files and the d.ts type tests against the staged declarations.
//
//   node typecheck.mjs                 README blocks + packages/npm/test/types/api.typecheck.ts
//   node typecheck.mjs --types-only    api.typecheck.ts only
//   node typecheck.mjs --docs <dir>    also the `js` blocks of every .md under <dir>
//   node typecheck.mjs --docs <dir> --marker no-run   only the doc blocks fenced as `js no-run`
//
// Each block is written to a temp directory as .mjs and checked with checkJs; a diagnostic is reported
// at the block's file:line in the markdown source.
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readdirSync, realpathSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, '../..');
const readmes = ['README.md', 'README_ja.md', 'packages/npm/README.md', 'packages/npm-native/README.md'];
const typeTest = join(repo, 'packages/npm/test/types/api.typecheck.ts');
// The floor run drops @types/node (it conflicts with TypeScript 5.7's lib.dom) and so skips blocks importing node: modules.
const compilers = [
  { label: 'typescript', tsc: join(here, 'node_modules/typescript/bin/tsc'), compilerOptions: {} },
  { label: 'typescript-floor', tsc: join(here, 'node_modules/typescript-floor/bin/tsc'), compilerOptions: { types: [] }, skipNode: true },
];

function parseArgs(argv) {
  const out = { typesOnly: false, docs: null, marker: null };
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === '--types-only') out.typesOnly = true;
    else if (argv[i] === '--docs' && i + 1 < argv.length) out.docs = resolve(argv[++i]);
    else if (argv[i] === '--marker' && i + 1 < argv.length) out.marker = argv[++i];
    else {
      console.error(`typecheck: unknown or incomplete argument '${argv[i]}'`);
      process.exit(2);
    }
  }
  return out;
}

function* markdownFiles(dir) {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (entry.name === 'node_modules' || entry.name.startsWith('.')) continue;
    const path = join(dir, entry.name);
    if (entry.isDirectory()) yield* markdownFiles(path);
    else if (entry.name.endsWith('.md')) yield path;
  }
}

/** Fenced blocks whose info string starts with `js`; `line` is the 1-based line of the opening fence. */
function jsBlocks(file, marker = null) {
  const lines = readFileSync(file, 'utf8').split('\n');
  const blocks = [];
  for (let i = 0; i < lines.length; i++) {
    const open = /^(\s*)(`{3,}|~{3,})js(?:\s.*)?$/.exec(lines[i]);
    if (!open) continue;
    const wanted = marker === null || (open[0].trim().split(/\s+/)[1] ?? '') === marker;
    const close = new RegExp(`^\\s*${open[2][0]}{${open[2].length},}\\s*$`);
    const body = [];
    let j = i + 1;
    for (; j < lines.length && !close.test(lines[j]); j++) body.push(lines[j].replace(new RegExp(`^\\s{0,${open[1].length}}`), ''));
    if (wanted) blocks.push({ file, line: i + 1, code: body.join('\n') + '\n' });
    i = j;
  }
  return blocks;
}

const args = parseArgs(process.argv.slice(2));
const sources = [];
if (!args.typesOnly) sources.push(...readmes.map((r) => join(repo, r)));
if (args.docs) sources.push(...markdownFiles(args.docs));
const docFiles = new Set(args.docs ? markdownFiles(args.docs) : []);

const tmp = realpathSync(mkdtempSync(join(tmpdir(), 'formulon-typecheck-')));
const origins = new Map();
const files = [typeTest];
const nodeFiles = new Set();
let blockCount = 0;
try {
  for (const source of sources) {
    for (const block of jsBlocks(source, docFiles.has(source) ? args.marker : null)) {
      const name = join(tmp, `block${String(blockCount++).padStart(4, '0')}.mjs`);
      writeFileSync(name, block.code);
      origins.set(name, block);
      files.push(name);
      if (/from\s+['"]node:/.test(block.code)) nodeFiles.add(name);
    }
  }
  if ((!args.typesOnly || args.marker) && blockCount === 0) {
    console.error('typecheck: no js blocks found');
    process.exit(1);
  }
  const diagnostic = /^(.+?)\((\d+),(\d+)\): (.*)$/;
  let anyFailed = false;
  for (const compiler of compilers) {
    const project = join(tmp, `tsconfig.${compiler.label}.json`);
    writeFileSync(
      project,
      JSON.stringify({ extends: join(here, 'tsconfig.json'), compilerOptions: compiler.compilerOptions, files: compiler.skipNode ? files.filter((f) => !nodeFiles.has(f)) : files }, null, 2),
    );
    const run = spawnSync(process.execPath, [compiler.tsc, '-p', project, '--pretty', 'false'], { encoding: 'utf8', cwd: tmp });
    const output = `${run.stdout}${run.stderr}`;
    let failed = 0;
    for (const text of output.split('\n')) {
      if (!text.trim()) continue;
      const m = diagnostic.exec(text);
      const file = m && resolve(tmp, m[1]);
      const block = file && origins.get(file);
      if (block) console.error(`${relative(repo, block.file)}:${block.line + Number(m[2])}:${m[3]}: ${m[4]}`);
      else if (file) console.error(`${relative(repo, file)}:${m[2]}:${m[3]}: ${m[4]}`);
      else console.error(text);
      failed++;
    }
    const ok = run.status === 0 && failed === 0;
    if (!ok) anyFailed = true;
    (ok ? console.log : console.error)(
      `typecheck [${compiler.label}]: ${ok ? 'ok' : 'failed'} (${blockCount} block(s), 1 type test)`,
    );
  }
  if (anyFailed) process.exit(1);
} finally {
  rmSync(tmp, { recursive: true, force: true });
}
