// Runs the ```js run blocks of the READMEs against the staged package.
// A block is a program: its stdout must equal the expectations written in it,
// either as a trailing `// value` comment on a `console.log(...)` line or as
// a run of comment lines closing the block. A block with no expectation only
// has to exit 0. Each block runs in a scratch copy of fixtures/readme/.

import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { cp, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
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

/** Returns the `js run` fenced blocks of a Markdown text with their first line. */
function runBlocks(markdown) {
  const blocks = [];
  const lines = markdown.split('\n');
  for (let i = 0; i < lines.length; i += 1) {
    if (!/^```js run\s*$/.test(lines[i])) continue;
    const start = i + 1;
    const body = [];
    for (i += 1; i < lines.length && !lines[i].startsWith('```'); i += 1) body.push(lines[i]);
    blocks.push({ line: start, code: body.join('\n'), lines: body });
  }
  return blocks;
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

let found = 0;
for (const readme of readmes) {
  const text = await readFile(path.join(repoRoot, readme), 'utf8');
  for (const block of runBlocks(text)) {
    found += 1;
    test(`${readme}:${block.line} runs and prints what its comments say`, async () => {
      const work = await mkdtemp(path.join(os.tmpdir(), 'formulon-readme-'));
      try {
        await cp(fixtureDir, work, { recursive: true });
        const code = block.code
          .replace(/from '@libraz\/formulon\/threads'/g, `from '${stagedEntry}'`)
          .replace(/from '@libraz\/formulon'/g, `from '${stagedEntry}'`);
        const script = path.join(work, 'snippet.mjs');
        await writeFile(script, code);
        const run = spawnSync(process.execPath, [script], { cwd: work, encoding: 'utf8' });
        assert.equal(run.status, 0, `${run.stderr}`);
        const expected = expectations(block.lines);
        if (expected.length > 0) {
          assert.deepEqual(run.stdout.trimEnd().split('\n'), expected);
        }
      } finally {
        await rm(work, { recursive: true, force: true });
      }
    });
  }
}

test('the READMEs carry runnable examples', () => {
  assert.ok(found >= 3, `found ${found} \`js run\` blocks`);
});
