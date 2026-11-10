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
