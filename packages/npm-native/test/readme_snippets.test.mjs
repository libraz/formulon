// Runs the ```js run blocks of the native package README against the staged package.
// With FORMULON_DOCS_DIR set, the `js run` and `js run:node-server` fences under that directory
// that import @libraz/formulon-native run too (see the npm runner for the marker rules).
// A block is a program: its stdout must equal the expectations written in it,
// either as a trailing `// value` comment on a `console.log(...)` line or as
// a run of comment lines closing the block. A block with no expectation only
// has to exit 0. Each block runs in a scratch copy of the shared README fixtures.

import assert from 'node:assert/strict';
import { spawn, spawnSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { cp, mkdtemp, readdir, readFile, rm, writeFile } from 'node:fs/promises';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const pkgRoot = path.resolve(here, '..');
const fixtureDir = path.resolve(pkgRoot, '..', 'npm', 'test', 'fixtures', 'readme');
const readme = path.join(pkgRoot, 'README.md');

// The staged entry named by the package's own "main" field.
const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
assert.ok(pkg.main, 'package.json is missing "main"');
const stagedEntry = pathToFileURL(path.resolve(pkgRoot, pkg.main)).href;

const docsDir = process.env.FORMULON_DOCS_DIR ? path.resolve(process.env.FORMULON_DOCS_DIR) : null;

/** Returns the fenced `js <marker>` blocks of a Markdown text with their first line and marker. */
function runBlocks(markdown, markers = ['run']) {
  const blocks = [];
  const lines = markdown.split('\n');
  for (let i = 0; i < lines.length; i += 1) {
    const open = /^```js (\S+)\s*$/.exec(lines[i]);
    if (!open || !markers.includes(open[1])) continue;
    const start = i + 1;
    const body = [];
    for (i += 1; i < lines.length && !lines[i].startsWith('```'); i += 1) body.push(lines[i]);
    blocks.push({ line: start, marker: open[1], code: body.join('\n'), lines: body });
  }
  return blocks;
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

const sources = [{ label: 'README.md', file: readme, markers: ['run'] }];
if (docsDir) {
  for (const file of await markdownFiles(docsDir)) {
    sources.push({ label: path.relative(docsDir, file), file, markers: ['run', 'run:node-server'], native: true });
  }
}

let readmeBlocks = 0;
for (const source of sources) {
  for (const block of runBlocks(await readFile(source.file, 'utf8'), source.markers)) {
    if (source.native && !/@libraz\/formulon-native/.test(block.code)) continue;
    if (!source.native) readmeBlocks += 1;
    test(`${source.label}:${block.line} ${block.marker} runs and prints what its comments say`, async () => {
      const work = await mkdtemp(path.join(os.tmpdir(), 'formulon-native-readme-'));
      try {
        if (existsSync(fixtureDir)) await cp(fixtureDir, work, { recursive: true });
        const code = block.code.replace(/from '@libraz\/formulon-native'/g, `from '${stagedEntry}'`);
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
}

test('the native README carries runnable examples', () => {
  assert.ok(readmeBlocks >= 1, `found ${readmeBlocks} \`js run\` blocks`);
});
