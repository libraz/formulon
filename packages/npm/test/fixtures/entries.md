The default entry is the single-threaded build: a recalculation starts no worker.

```js run
import createFormulon from '@libraz/formulon';

const formulon = await createFormulon();
const wb = formulon.Workbook.createDefault();
const ws = wb.sheet(0);
for (let row = 0; row < 50; row += 1) {
  ws.setValue(row, 0, row);
  ws.setFormula(row, 1, `=A${row + 1}*2`);
}
console.log(wb.recalcParallel(4).workerThreadsStarted); // 0
```

The `/threads` entry is the pthread build: the same recalculation starts its workers.

```js run
import createFormulon from '@libraz/formulon/threads';

const formulon = await createFormulon();
const wb = formulon.Workbook.createDefault();
const ws = wb.sheet(0);
for (let row = 0; row < 50; row += 1) {
  ws.setValue(row, 0, row);
  ws.setFormula(row, 1, `=A${row + 1}*2`);
}
console.log(wb.recalcParallel(4).workerThreadsStarted > 0); // true
```
