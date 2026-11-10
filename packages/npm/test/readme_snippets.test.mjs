// Runs the ```js run blocks of the READMEs against the staged package.
// With FORMULON_DOCS_DIR set, every .md under that directory is walked as well and its fences are
// interpreted by marker: `js run` as above; `js run:node-server` starts the script with PORT set,
// POSTs fixtures/readme/input.xlsx to `/` once, requires a 2xx response containing every
// expectation line, then stops it; `js no-run` is type-checked by tools/typecheck/typecheck.mjs.
// A block importing @libraz/formulon-native belongs to the native runner and is skipped here.
// A fence may carry a VitePress code-group label (`js run [JavaScript]`); a `js` fence that mentions
// `run` but does not match that grammar fails instead of being skipped.
// A block is a program: its stdout must equal the expectations written in it,
// either as a trailing `// value` comment on a `console.log(...)` line or as
// a run of comment lines closing the block. A block with no expectation only
// has to exit 0. Each block runs in a scratch copy of fixtures/readme/.

import assert from 'node:assert/strict';
import { spawn, spawnSync } from 'node:child_process';
import { cp, mkdtemp, readdir, readFile, rm, writeFile } from 'node:fs/promises';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const pkgRoot = path.resolve(here, '..');
const repoRoot = path.resolve(pkgRoot, '..', '..');
const fixtureDir = path.join(here, 'fixtures', 'readme');
const readmes = ['README.md', 'README_ja.md', 'packages/npm/README.md'];

// The same package.json "exports" resolution as the other staged-package tests.
const entry = process.env.FORMULON_NPM_ENTRY ?? '.';
const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
const stagedEntry = pathToFileURL(path.resolve(pkgRoot, pkg.exports[entry].import)).href;

const docsDir = process.env.FORMULON_DOCS_DIR ? path.resolve(process.env.FORMULON_DOCS_DIR) : null;

/** A `js <marker>` fence, optionally followed by a VitePress code-group label (`js run [JavaScript]`). */
const FENCE = /^```js (\S+)(?:\s+\[[^\]]*\])?\s*$/;
const KNOWN_MARKERS = new Set(['run', 'run:node-server', 'no-run']);

/** Returns the fenced `js <marker>` blocks of a Markdown text with their first line and marker. */
function runBlocks(markdown, markers = ['run']) {
  return scanFences(markdown, markers).blocks;
}

/**
 * The `js <marker>` blocks of a Markdown text, plus the 1-based lines of `js` fences that mention `run` but do not
 * match the fence grammar; those would otherwise be skipped silently.
 */
function scanFences(markdown, markers = ['run']) {
  const blocks = [];
  const nearMisses = [];
  const lines = markdown.split('\n');
  for (let i = 0; i < lines.length; i += 1) {
    if (!/^```js(?!on)/.test(lines[i])) continue;
    const open = FENCE.exec(lines[i]);
    if (!open || !KNOWN_MARKERS.has(open[1])) {
      if (/\brun\b/.test(lines[i])) nearMisses.push({ line: i + 1, text: lines[i] });
      continue;
    }
    if (!markers.includes(open[1])) continue;
    const start = i + 1;
    const body = [];
    for (i += 1; i < lines.length && !lines[i].startsWith('```'); i += 1) body.push(lines[i]);
    blocks.push({ line: start, marker: open[1], code: body.join('\n'), lines: body });
  }
  return { blocks, nearMisses };
}

/** Registers a failing test for each fence of `source` that looks runnable but does not parse. */
function reportNearMisses(label, text) {
  for (const miss of scanFences(text).nearMisses) {
    test(`${label}:${miss.line} fence matches the runner grammar`, () => {
      assert.fail(
        `${label}:${miss.line}: \`${miss.text}\` mentions run but is not \`\`\`js <marker> [label]\` with a known marker`,
      );
    });
  }
}

/** Every .md file under a directory, recursively and in a stable order. */
async function markdownFiles(dir) {
  const out = [];
  for (const entry of (await readdir(dir, { withFileTypes: true })).sort((a, b) => a.name.localeCompare(b.name))) {
    if (entry.name === 'node_modules' || entry.name.startsWith('.')) continue;
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) out.push(...(await markdownFiles(full)));
    else if (entry.name.endsWith('.md')) out.push(full);
  }
  return out;
}

/** An unused TCP port on the loopback interface. */
function freePort() {
  return new Promise((resolve, reject) => {
    const server = net.createServer();
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => {
      const { port } = server.address();
      server.close(() => resolve(port));
    });
  });
}

/** Starts a `run:node-server` script, sends one request, returns the response and stops the server. */
async function exerciseServer(script, work) {
  const port = await freePort();
  const child = spawn(process.execPath, [script], {
    cwd: work,
    env: { ...process.env, PORT: String(port) },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  let stderr = '';
  child.stderr.on('data', (chunk) => {
    stderr += chunk;
  });
  const exited = new Promise((resolve) => child.once('exit', resolve));
  const body = await readFile(path.join(work, 'input.xlsx'));
  try {
    const deadline = Date.now() + 15000;
    for (;;) {
      try {
        const res = await fetch(`http://127.0.0.1:${port}/`, {
          method: 'POST',
          headers: { 'content-type': 'application/octet-stream' },
          body,
          signal: AbortSignal.timeout(10000),
        });
        return { status: res.status, text: await res.text() };
      } catch (err) {
        if (child.exitCode !== null || Date.now() > deadline) {
          throw new Error(`server did not answer: ${err.message}\n${stderr}`, { cause: err });
        }
        await new Promise((resolve) => setTimeout(resolve, 100));
      }
    }
  } finally {
    child.kill('SIGKILL');
    await exited;
  }
}

/** Expected stdout lines: inline `console.log(...) // x` comments, then the closing comment run. */
function expectations(lines) {
  const expected = [];
  for (const line of lines) {
    const m = /console\.log\(.*\)\s*;?\s*\/\/ (.+)$/.exec(line);
    if (m) expected.push(m[1].trimEnd());
  }
  let end = lines.length;
  while (end > 0 && lines[end - 1].trim() === '') end -= 1;
  let begin = end;
  while (begin > 0 && /^\/\/ ?/.test(lines[begin - 1])) begin -= 1;
  for (const line of lines.slice(begin, end)) expected.push(line.replace(/^\/\/ ?/, '').trimEnd());
  return expected;
}

const sources = readmes.map((readme) => ({ label: readme, file: path.join(repoRoot, readme), markers: ['run'] }));
if (docsDir) {
  for (const file of await markdownFiles(docsDir)) {
    sources.push({ label: path.relative(docsDir, file), file, markers: ['run', 'run:node-server'] });
  }
}

let found = 0;
let noRun = 0;
for (const source of sources) {
  const text = await readFile(source.file, 'utf8');
  reportNearMisses(source.label, text);
  for (const block of runBlocks(text, source.markers)) {
    if (/@libraz\/formulon-native/.test(block.code)) continue;
    found += 1;
    test(`${source.label}:${block.line} ${block.marker} runs and prints what its comments say`, async () => {
      const work = await mkdtemp(path.join(os.tmpdir(), 'formulon-readme-'));
      try {
        await cp(fixtureDir, work, { recursive: true });
        const code = block.code
          .replace(/from '@libraz\/formulon\/threads'/g, `from '${stagedEntry}'`)
          .replace(/from '@libraz\/formulon'/g, `from '${stagedEntry}'`);
        const script = path.join(work, 'snippet.mjs');
        await writeFile(script, code);
        const expected = expectations(block.lines);
        if (block.marker === 'run:node-server') {
          const res = await exerciseServer(script, work);
          assert.ok(res.status >= 200 && res.status < 300, `status ${res.status}: ${res.text}`);
          for (const line of expected) assert.ok(res.text.includes(line), `response lacks "${line}": ${res.text}`);
          return;
        }
        const run = spawnSync(process.execPath, [script], { cwd: work, encoding: 'utf8' });
        assert.equal(run.status, 0, `${run.stderr}`);
        if (expected.length > 0) {
          assert.deepEqual(run.stdout.trimEnd().split('\n'), expected);
        }
      } finally {
        await rm(work, { recursive: true, force: true });
      }
    });
  }
  if (source.markers.includes('run:node-server')) noRun += runBlocks(text, ['no-run']).length;
}

if (docsDir && noRun > 0) {
  test(`${noRun} no-run block(s) under FORMULON_DOCS_DIR type-check`, () => {
    const run = spawnSync(
      process.execPath,
      [
        path.join(repoRoot, 'tools', 'typecheck', 'typecheck.mjs'),
        '--types-only',
        '--docs',
        docsDir,
        '--marker',
        'no-run',
      ],
      { encoding: 'utf8' },
    );
    assert.equal(run.status, 0, `${run.stdout}${run.stderr}`);
  });
}

test('a labelled fence is collected and a malformed runnable fence is reported', async () => {
  const { blocks, nearMisses } = scanFences(await readFile(path.join(here, 'fixtures', 'fences.md'), 'utf8'), [
    'run',
    'no-run',
  ]);
  assert.deepEqual(
    blocks.map((b) => [b.line, b.marker]),
    [
      [1, 'run'],
      [5, 'no-run'],
    ],
  );
  assert.deepEqual(
    nearMisses.map((m) => m.line),
    [8, 11, 14],
  );
});

test('the READMEs carry runnable examples', () => {
  assert.ok(found >= 3, `found ${found} \`js run\` blocks`);
});
