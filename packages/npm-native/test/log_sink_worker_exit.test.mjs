// A log sink set inside a Worker is dropped when that Worker exits: later records from the main thread, and clearing
// the sink there, must not reach the Worker's torn-down function. Runs in a child process so a crash is a failure
// rather than the end of the test run.

import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const entry = pathToFileURL(path.resolve(pkgRoot, JSON.parse(readFileSync(path.join(pkgRoot, 'package.json'))).main));

const CHILD = `
import { Worker } from 'node:worker_threads';
const entry = process.argv[1];
const worker = new Worker(
  "const { workerData } = await import('node:worker_threads');" +
    " const { createFormulon } = await import(workerData); (await createFormulon()).setLogSink(() => {});",
  { eval: true, type: 'module', workerData: entry },
);
await new Promise((resolve, reject) => { worker.on('exit', resolve); worker.on('error', reject); });
const { createFormulon } = await import(entry);
const f = await createFormulon();
f.setLogMinLevel(f.LogLevel.Warn);
for (let i = 0; i < 3; i++) {
  const wb = f.Workbook.createDefault();
  wb.sheet(0).setFormula(0, 0, '=SUM(T[C])');
  wb.recalc();
  wb.save(f.WorkbookFormat.Xlsb);
  wb.dispose();
}
f.setLogSink(null);
f.setLogSink(() => {});
f.setLogSink(null);
f.setLogMinLevel(f.LogLevel.Off);
process.stdout.write('done');
`;

test('a sink set in an exited Worker is never called or released afterwards', () => {
  const child = spawnSync(process.execPath, ['--input-type=module', '-e', CHILD, entry.href], {
    encoding: 'utf8',
    timeout: 60_000,
  });
  assert.equal(child.signal, null, `child died with ${child.signal}: ${child.stderr}`);
  assert.equal(child.status, 0, child.stderr);
  assert.equal(child.stdout, 'done');
});
