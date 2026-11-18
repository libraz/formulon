// A defined name spelled as the profile locale's built-in name is stored as that built-in.

import assert from 'node:assert/strict';
import test from 'node:test';

import { getModule } from './smoke_support.mjs';

async function withWorkbook(profileId, fn) {
  const f = await getModule();
  const wb = f.Workbook.createForProfile(profileId);
  try {
    return await fn(wb, f);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

test('Print_Area defines the print area of the sheet its formula names', async () => {
  await withWorkbook('mac-365-ja_JP', (wb) => {
    wb.addSheet('Sheet2');
    wb.setDefinedName({ name: 'print_area', formula: '=Sheet2!$A$1:$B$2', comment: 'note', hidden: true });
    assert.deepEqual(wb.listDefinedNames(), [
      { name: '_xlnm.Print_Area', formula: '=Sheet2!$A$1:$B$2', scope: 1, comment: 'note', hidden: true },
    ]);
    assert.deepEqual(wb.sheet(1).getPrintArea(), [{ firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 }]);
    assert.equal(wb.getDefinedName('Print_Area'), null);
    assert.equal(wb.getDefinedName('Print_Area', 1).name, '_xlnm.Print_Area');
    wb.sheet(1).setPrintArea('A1:C3');
    assert.equal(wb.listDefinedNames().length, 1);
    assert.equal(wb.getDefinedName('PRINT_AREA', 1).formula, '=Sheet2!$A$1:$C$3');
    wb.removeDefinedName('Print_Area', 1);
    assert.deepEqual(wb.listDefinedNames(), []);
  });
});

test('only the locale spelling is a built-in, and the stored spelling needs a sheet', async () => {
  await withWorkbook('mac-365-de_DE', (wb, f) => {
    const sheet = wb.sheet(0).name;
    wb.setDefinedName({ name: 'Print_Area', formula: `=${sheet}!$A$1:$B$2` });
    wb.setDefinedName({ name: 'Druckbereich', formula: `=${sheet}!$A$1:$B$2` });
    assert.deepEqual(
      wb.listDefinedNames().map((n) => [n.name, n.scope]),
      [
        ['Print_Area', null],
        ['_xlnm.Print_Area', 0],
      ],
    );
    assert.throws(
      () => wb.setDefinedName({ name: '_xlnm.Print_Titles', formula: `=${sheet}!$1:$1` }),
      (e) => e.code === f.StatusCode.InvalidArgument,
    );
  });
});

const BUILT_IN_NAMES = [
  'Print_Area',
  'Print_Titles',
  '_FilterDatabase',
  'Criteria',
  'Extract',
  'Database',
  'Consolidate_Area',
  'Sheet_Title',
  'Auto_Open',
  'Recorder',
  'Auto_Close',
  'Data_Form',
  'Auto_Activate',
  'Auto_Deactivate',
];
const SHEET_BUILT_INS = new Set(['Print_Area', 'Print_Titles']);

function macProfiles(f) {
  return f.listProfileIds().filter((id) => id.startsWith('mac-365-'));
}

function code(fn) {
  try {
    fn();
  } catch (e) {
    return e.code;
  }
  return null;
}

test('the resolved key is where setDefinedName stores every built-in in every locale', async () => {
  const f = await getModule();
  const profiles = macProfiles(f);
  assert.equal(profiles.length, 14);
  for (const profile of profiles) {
    for (const english of BUILT_IN_NAMES) {
      const typed = f.localizeDefinedName(`_xlnm.${english}`, profile);
      await withWorkbook(profile, (wb) => {
        wb.addSheet('Sheet2');
        const input = { name: typed, formula: '=Sheet2!$A$1:$B$2' };
        const key = wb.resolveDefinedNameKey(input);
        assert.deepEqual(wb.listDefinedNames(), [], `${profile} ${typed}`);
        assert.equal(key.name, f.canonicalizeDefinedName(typed, profile), `${profile} ${typed}`);
        assert.equal(key.scope, SHEET_BUILT_INS.has(english) ? 1 : null, `${profile} ${typed}`);
        wb.setDefinedName(input);
        assert.deepEqual(
          wb.listDefinedNames().map((n) => [n.name, n.scope]),
          [[key.name, key.scope]],
          `${profile} ${typed}`,
        );
        assert.equal(wb.getDefinedName(key.name, key.scope).name, key.name);
        assert.equal(f.localizeDefinedName(key.name, profile), typed);
      });
    }
  }
});

test('resolveDefinedNameKey refuses what setDefinedName refuses, with the same status', async () => {
  const f = await getModule();
  for (const profile of macProfiles(f)) {
    await withWorkbook(profile, (wb) => {
      wb.addSheet('Sheet2');
      const ws = wb.sheet(0);
      ws.setText('A1', 'a');
      ws.setText('B1', 'b');
      ws.addTable({ range: 'A1:B2', name: 'Sales' });
      wb.setDefinedName({ name: 'Taken', formula: '=Sheet1!$A$1' });
      const cases = [
        { name: '_xlnm.Print_Area', formula: '=$A$1:$B$2' },
        { name: '_xlnm.Print_Titles', formula: '=$1:$1' },
        { name: 'Sales', formula: '=Sheet1!$A$1' },
        { name: 'Fresh', formula: '=Sheet1!$A$1', scope: 7 },
        { name: 'Fresh', formula: '' },
        { name: 'Fresh', formula: '=SUM(' },
        { name: '', formula: '=Sheet1!$A$1' },
        { name: '1Fresh', formula: '=Sheet1!$A$1' },
        { name: 'taken', formula: '=Sheet1!$B$1' },
        { name: 'Fresh', formula: '=Sheet1!$A$1', scope: 1 },
      ];
      for (const input of cases) {
        const before = wb.listDefinedNames().length;
        let key = null;
        const resolved = code(() => {
          key = wb.resolveDefinedNameKey(input);
        });
        assert.equal(wb.listDefinedNames().length, before);
        const set = code(() => wb.setDefinedName(input));
        assert.equal(resolved, set, `${profile} ${JSON.stringify(input)}`);
        if (set === null) {
          assert.ok(wb.getDefinedName(key.name, key.scope), `${profile} ${JSON.stringify(input)}`);
        }
      }
      const statusOf = (input) => code(() => wb.resolveDefinedNameKey(input));
      assert.equal(statusOf({ name: '_xlnm.Print_Area', formula: '=$A$1' }), f.StatusCode.InvalidArgument);
      assert.equal(statusOf({ name: 'Sales', formula: '=Sheet1!$A$2', scope: 1 }), f.StatusCode.AlreadyExists);
      assert.equal(statusOf({ name: 'Fresh', formula: '=Sheet1!$A$1', scope: 9 }), f.StatusCode.SheetIndexOutOfRange);
      assert.throws(() => wb.resolveDefinedNameKey({ formula: '=Sheet1!$A$1' }), TypeError);
      assert.throws(() => wb.resolveDefinedNameKey(null), TypeError);
    });
  }
});

test('the resolved key lets a caller keep the comment and hidden flag a set resets', async () => {
  await withWorkbook('mac-365-de_DE', (wb) => {
    wb.addSheet('Sheet2');
    wb.setDefinedName({ name: 'Druckbereich', formula: '=Sheet2!$A$1:$B$2', comment: 'note', hidden: true });
    const input = { name: 'DRUCKBEREICH', formula: '=Sheet2!$A$1:$C$3' };
    const key = wb.resolveDefinedNameKey(input);
    assert.deepEqual({ ...key }, { name: '_xlnm.Print_Area', scope: 1 });
    const prev = wb.getDefinedName(key.name, key.scope);
    wb.setDefinedName({ ...input, comment: prev?.comment ?? null, hidden: prev?.hidden ?? false });
    assert.deepEqual(wb.listDefinedNames(), [
      { name: '_xlnm.Print_Area', formula: '=Sheet2!$A$1:$C$3', scope: 1, comment: 'note', hidden: true },
    ]);
    wb.setDefinedName(input);
    assert.equal(wb.getDefinedName(key.name, key.scope).comment, null);
    // A read record is a valid input.
    const again = wb.resolveDefinedNameKey(wb.listDefinedNames()[0]);
    assert.deepEqual({ ...again }, { name: '_xlnm.Print_Area', scope: 1 });
  });
});

test('localizeDefinedName and canonicalizeDefinedName map between the stored and the displayed spelling', async () => {
  const f = await getModule();
  assert.equal(f.localizeDefinedName('_xlnm.Print_Area', 'mac-365-de_DE'), 'Druckbereich');
  assert.equal(f.localizeDefinedName('_xlnm.PRINT_AREA', 'mac-365-de_DE'), 'Druckbereich');
  assert.equal(f.localizeDefinedName('_xlnm.Print_Area', 'mac-365-fr_FR'), 'Zone_d_impression');
  assert.equal(f.localizeDefinedName('_xlnm.Print_Area', 'mac-365-ja_JP'), 'Print_Area');
  assert.equal(f.localizeDefinedName('_xlnm.print_area', 'mac-365-en_US'), 'print_area');
  assert.equal(f.localizeDefinedName('_xlnm.Foo', 'mac-365-en_US'), '_xlnm.Foo');
  assert.equal(f.localizeDefinedName('Sales', 'mac-365-de_DE'), 'Sales');
  assert.equal(f.canonicalizeDefinedName('druckbereich', 'mac-365-de_DE'), '_xlnm.Print_Area');
  assert.equal(f.canonicalizeDefinedName('Print_Area', 'mac-365-de_DE'), 'Print_Area');
  assert.equal(f.canonicalizeDefinedName('print_area', 'mac-365-en_US'), '_xlnm.print_area');
  assert.equal(f.canonicalizeDefinedName('_xlnm.Foo', 'mac-365-en_US'), '_xlnm.Foo');
  assert.equal(f.canonicalizeDefinedName('Sales', 'mac-365-en_US'), 'Sales');
  for (const profile of macProfiles(f)) {
    for (const english of BUILT_IN_NAMES) {
      const stored = `_xlnm.${english}`;
      const shown = f.localizeDefinedName(stored, profile);
      assert.equal(f.canonicalizeDefinedName(shown, profile), stored, `${profile} ${stored}`);
    }
  }
  for (const call of [
    () => f.localizeDefinedName('', 'mac-365-en_US'),
    () => f.canonicalizeDefinedName('', 'mac-365-en_US'),
    () => f.localizeDefinedName('Print_Area', 'no-such-profile'),
    () => f.canonicalizeDefinedName('Print_Area', 'no-such-profile'),
  ]) {
    assert.equal(code(call), f.StatusCode.InvalidArgument);
  }
  assert.throws(() => f.localizeDefinedName(null, 'mac-365-en_US'), TypeError);
});
