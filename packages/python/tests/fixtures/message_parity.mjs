// The JS half of test_message_parity.py: each case fails one call and prints [status name, message] per case.
// Usage: node message_parity.mjs <file URL of the staged npm entry>
const CASES = {
  'sheet past the count': (wb) => wb.sheet(5),
  'sheet name matching none': (wb) => wb.sheet('Nope'),
  'malformed defined name': (wb) => wb.setDefinedName({ name: 'A1', formula: '1' }),
  'malformed table name': (wb, ws) => ws.addTable({ range: 'A1:B3', name: 'Has Space', columns: ['a', 'b'] }),
  'write into part of an array': (wb, ws) => {
    ws.setArrayFormula({ range: 'E1:E3', formula: '=A1:A3*2' });
    ws.getValue('E2');
    ws.setValue('E2', 5);
  },
  'setValues given a non-array': (wb, ws) => ws.setValues(0, 0, 'ab'),
  'setValues row that is not an array': (wb, ws) => ws.setValues(0, 0, [[1], 'x']),
  'setValues rows of unequal length': (wb, ws) => ws.setValues(0, 0, [[1, 2], [3]]),
  'setValues element that is not a cell value': (wb, ws) => ws.setValues(0, 0, [[1, {}]]),
  'setValues non-finite number': (wb, ws) => ws.setValues(0, 0, [[1], [Number.NaN]]),
  'getValues past the cell cap': (wb, ws) =>
    ws.getValues({ firstRow: 0, firstCol: 0, lastRow: 1048575, lastCol: 16383 }),
  'getCollationClasses element that is not a Value': () =>
    f.getCollationClasses([{ kind: f.ValueKind.Text, text: 'a' }, 5], f.CollationStrength.Dynamic, 'win-365-en_US'),
  'toNative of a non-Value': () => f.toNative(5),
  'table insertRows last without first': (wb, ws) =>
    ws.addTable({ range: 'A1:B3', name: 'T', columns: ['a', 'b'] }).insertRows(undefined, 1),
  'table insertRows first past the end': (wb, ws) =>
    ws.addTable({ range: 'A1:B3', name: 'T', columns: ['a', 'b'] }).insertRows(3),
  'table insertColumns last below first': (wb, ws) =>
    ws.addTable({ range: 'A1:B3', name: 'T', columns: ['a', 'b'] }).insertColumns(2, 1),
};

const { createFormulon } = await import(process.argv[2]);
const f = await createFormulon();
const out = {};
for (const [id, call] of Object.entries(CASES)) {
  const wb = f.Workbook.createDefault();
  try {
    call(wb, wb.sheet(0));
    out[id] = null;
  } catch (e) {
    const code =
      e.name === 'FormulonError' ? Object.keys(f.StatusCode).find((k) => f.StatusCode[k] === e.code) : e.name;
    out[id] = [code, e.message];
  } finally {
    wb.dispose();
  }
}
console.log(JSON.stringify(out));
