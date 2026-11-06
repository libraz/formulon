import test from 'node:test';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';

const ANN = '{11111111-1111-1111-1111-111111111111}';
const BOB = '{22222222-2222-2222-2222-222222222222}';
const THREAD = '{AAAAAAAA-0000-0000-0000-000000000001}';
const REPLY = '{AAAAAAAA-0000-0000-0000-000000000002}';
const MENTION = '{BBBBBBBB-0000-0000-0000-000000000001}';

function hiddenRows(wb) {
  return wb
    .listRowOverrides(0)
    .filter((r) => r.hidden)
    .map((r) => r.row);
}

function seedColumn(wb) {
  for (const [row, text] of ['h', 'a', 'b', 'a'].entries()) wb.setText(0, row, 0, text);
}

const RANGE = { firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 0 };

test('filter, sort and validation enums keep their C ordinals', async () => {
  const mod = await getModule();
  assert.equal(mod.FilterKind.Icon, 6);
  assert.equal(mod.FilterOperator.GreaterThan, 5);
  assert.equal(mod.DynamicFilterType.YearToDate, 18);
  assert.equal(mod.DynamicFilterType.M12, 34);
  assert.equal(mod.SortBy.Icon, 3);
  assert.equal(mod.SortMethod.Stroke, 2);
  assert.equal(mod.DateTimeGrouping.Second, 5);
  assert.equal(mod.ValidationErrorStyle.Information, 2);
});

test('AutoFilter round-trips and its criteria evaluate, apply and clear', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    seedColumn(wb);
    const absent = wb.getAutoFilter(0);
    assert.equal(absent, null);

    const filter = {
      range: RANGE,
      columns: [
        {
          colId: 0,
          kind: mod.FilterKind.Values,
          values: ['a'],
          dateGroups: [{ year: 2026, month: 3, grouping: mod.DateTimeGrouping.Month }],
        },
      ],
    };
    wb.setAutoFilter(0, filter);
    const got = wb.getAutoFilter(0);
    assert.deepEqual(got.range, RANGE);
    assert.equal(got.columns.length, 1);
    const col = got.columns[0];
    assert.equal(col.kind, 1);
    assert.deepEqual(col.values, ['a']);
    assert.equal(col.showButton, true);
    assert.equal(col.hiddenButton, false);
    assert.equal(col.dateGroups.length, 1);
    assert.equal(col.dateGroups[0].year, 2026);
    assert.equal(col.dateGroups[0].month, 3);
    assert.equal(col.dateGroups[0].grouping, mod.DateTimeGrouping.Month);
    assert.equal(col.dxfId, 0);
    assert.equal(got.sort, null);
    // Setting the unchanged record back is accepted.
    wb.setAutoFilter(0, got);
    assert.deepEqual(wb.getAutoFilter(0), got);

    // Replace the date group with a plain value filter: dateGroups drops out.
    wb.setAutoFilter(0, { range: RANGE, columns: [{ colId: 0, kind: mod.FilterKind.Values, values: ['a'] }] });

    const ev = wb.evaluateAutoFilter(0);
    assert.equal(ev.firstRow, 1);
    assert.deepEqual(ev.match, [true, false, true]);
    assert.deepEqual(hiddenRows(wb), []);

    wb.applyAutoFilter(0);
    assert.deepEqual(hiddenRows(wb), [2]);
    wb.clearAutoFilter(0);
    assert.deepEqual(hiddenRows(wb), []);
    assert.equal(wb.getAutoFilter(0).columns.length, 0);

    // A custom criterion and a sort state survive a round trip.
    wb.setAutoFilter(0, {
      range: RANGE,
      columns: [{ colId: 0, kind: mod.FilterKind.Custom, customCount: 1, op1: mod.FilterOperator.NotEqual, val1: 'b' }],
      sort: {
        ref: { firstRow: 1, firstCol: 0, lastRow: 3, lastCol: 0 },
        conditions: [{ ref: { firstRow: 1, firstCol: 0, lastRow: 3, lastCol: 0 }, descending: true }],
      },
    });
    const custom = wb.getAutoFilter(0);
    assert.equal(custom.columns[0].kind, 2);
    assert.equal(custom.columns[0].op1, mod.FilterOperator.NotEqual);
    assert.equal(custom.columns[0].val1, 'b');
    assert.equal(custom.sort.conditions.length, 1);
    assert.equal(custom.sort.conditions[0].descending, true);
    assert.equal(custom.sort.columnSort, false);
    assert.deepEqual(wb.evaluateAutoFilter(0).match, [true, false, true]);

    wb.removeAutoFilter(0);
    assert.equal(wb.getAutoFilter(0), null);
    assert.throws(() => wb.evaluateAutoFilter(0), isFormulonError);

    // An out-of-range column id is rejected and leaves the sheet unchanged.
    assert.throws(
      () =>
        wb.setAutoFilter(0, {
          range: RANGE,
          columns: [{ colId: 5, kind: mod.FilterKind.Values, values: ['a'] }],
        }),
      isFormulonError,
    );
    assert.equal(wb.getAutoFilter(0), null);
    assert.throws(() => wb.setAutoFilter(0, null), TypeError);
  } finally {
    wb.dispose();
  }
});

test('table AutoFilter entry points reject a table index that does not exist', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.equal(wb.getTableCount(), 0);
    assert.throws(() => wb.getTableAutoFilter(0), isFormulonError);
    assert.throws(() => wb.evaluateTableAutoFilter(0), isFormulonError);
    assert.throws(() => wb.setTableAutoFilter(0, { range: RANGE, columns: [] }), isFormulonError);
    assert.throws(() => wb.removeTableAutoFilter(0), isFormulonError);
    assert.throws(() => wb.applyTableAutoFilter(0), isFormulonError);
    assert.throws(() => wb.clearTableAutoFilter(0), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('validateValue and listInvalidCells follow the rule covering a cell', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const free = wb.validateValue(0, 0, 1, { kind: mod.ValueKind.Number, number: 99 });
    assert.equal(free.hasRule, false);
    assert.equal(free.valid, true);

    wb.addValidation(0, {
      ranges: [{ firstRow: 0, lastRow: 3, firstCol: 1, lastCol: 1 }],
      type: 1,
      op: 0,
      errorStyle: mod.ValidationErrorStyle.Warning,
      formula1: '1',
      formula2: '10',
      showErrorMessage: true,
    });
    const ok = wb.validateValue(0, 0, 1, { kind: mod.ValueKind.Number, number: 5 });
    assert.deepEqual(
      { hasRule: ok.hasRule, valid: ok.valid, ruleIndex: ok.ruleIndex, errorStyle: ok.errorStyle },
      { hasRule: true, valid: true, ruleIndex: 0, errorStyle: mod.ValidationErrorStyle.Warning },
    );
    const bad = wb.validateValue(0, 0, 1, { kind: mod.ValueKind.Number, number: 50 });
    assert.equal(bad.valid, false);
    assert.equal(bad.hasRule, true);
    const text = wb.validateValue(0, 0, 1, { kind: mod.ValueKind.Text, text: '5' });
    assert.equal(text.valid, false, 'text "5" is text, not a number');
    assert.throws(() => wb.validateValue(99, 0, 0, { kind: mod.ValueKind.Number, number: 1 }), isFormulonError);
    assert.throws(() => wb.validateValue(0, 0, 0, null), TypeError);
    assert.throws(
      () => wb.validateValue(0, 0, 0, { number: 50 }),
      TypeError,
      'a value without `kind` is rejected, not read as blank',
    );

    wb.setNumber(0, 0, 1, 5);
    wb.setNumber(0, 1, 1, 50);
    wb.setNumber(0, 2, 1, 70);
    const page = wb.getInvalidCells(0);
    assert.deepEqual(
      page.cells.map((c) => [c.row, c.col, c.value.number]),
      [
        [1, 1, 50],
        [2, 1, 70],
      ],
    );
    assert.equal(page.nextCursor, null);

    const first = wb.getInvalidCells(0, 0, 1);
    assert.equal(first.cells.length, 1);
    assert.equal(typeof first.nextCursor, 'number');
    const second = wb.getInvalidCells(0, first.nextCursor, 1);
    assert.deepEqual(
      second.cells.map((c) => c.row),
      [2],
    );
    assert.throws(() => wb.getInvalidCells(99), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('threaded comments cover a thread, a reply with a mention, resolution and removal', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.equal(wb.listPeople().length, 0);
    wb.listPeople();
    wb.addPerson({ id: ANN, displayName: 'Ann', userId: 'ann@example.com', providerId: 'AD' });
    wb.addPerson({ id: BOB, displayName: 'Bob', userId: '', providerId: 'None' });
    assert.throws(() => wb.addPerson({ id: ANN, displayName: 'Ann again' }), isFormulonError);
    const persons = wb.listPeople();
    assert.deepEqual(
      persons.map((p) => [p.id, p.displayName, p.userId, p.providerId]),
      [
        [ANN, 'Ann', 'ann@example.com', 'AD'],
        [BOB, 'Bob', '', 'None'],
      ],
    );

    wb.addThreadedComment(0, {
      id: THREAD,
      row: 2,
      col: 3,
      personId: ANN,
      created: '2026-01-01T10:00:00.00',
      text: 'Review this',
    });
    wb.addThreadedComment(0, {
      id: REPLY,
      personId: BOB,
      created: '2026-01-01T11:00:00.00',
      text: '@Ann done',
      parentId: THREAD,
      mentions: [{ personId: ANN, mentionId: MENTION, start: 0, length: 4 }],
    });
    const result = wb.listThreadedComments(0);
    const list = result;
    assert.equal(list.length, 2);
    assert.equal(list[0].id, THREAD);
    assert.equal(list[0].parentId, '');
    assert.equal(list[0].row, 2);
    assert.equal(list[0].col, 3);
    assert.equal(list[0].done, false);
    assert.deepEqual(list[0].mentions, []);
    assert.equal(list[1].id, REPLY);
    assert.equal(list[1].parentId, THREAD);
    assert.equal(list[1].personId, BOB);
    assert.deepEqual(list[1].mentions, [{ personId: ANN, mentionId: MENTION, start: 0, length: 4 }]);

    wb.updateThreadedComment(0, REPLY, '@Ann fixed', [{ personId: ANN, mentionId: MENTION, start: 0, length: 4 }]);
    assert.equal(wb.listThreadedComments(0)[1].text, '@Ann fixed');
    wb.updateThreadedComment(0, REPLY, 'plain', []);
    assert.deepEqual(wb.listThreadedComments(0)[1].mentions, []);

    wb.setThreadResolved(0, THREAD, true);
    assert.equal(wb.listThreadedComments(0)[0].done, true);
    wb.setThreadResolved(0, THREAD, false);
    assert.equal(wb.listThreadedComments(0)[0].done, false);

    assert.throws(() => wb.removePerson(ANN), isFormulonError, 'a person with comments stays');
    assert.throws(
      () => wb.updateThreadedComment(0, '{00000000-0000-0000-0000-000000000000}', 'x', []),
      isFormulonError,
    );
    assert.throws(() => wb.setThreadResolved(0, REPLY, true), isFormulonError, 'a reply is not a thread');
    assert.throws(() => wb.addThreadedComment(0, null), TypeError);
    assert.throws(() => wb.listThreadedComments(99), isFormulonError);

    wb.removeThreadedComment(0, REPLY);
    assert.equal(wb.listThreadedComments(0).length, 1);
    wb.removeThreadedComment(0, THREAD);
    assert.equal(wb.listThreadedComments(0).length, 0);
    wb.removePerson(ANN);
    assert.equal(wb.listPeople().length, 1);
    assert.throws(() => wb.removePerson(ANN), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('the new accessors report a released handle', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.dispose();
  const released = (e) => isFormulonError(e, mod.StatusCode.BindingInvalidHandle);
  assert.throws(() => wb.getAutoFilter(0), released);
  assert.throws(() => wb.evaluateAutoFilter(0), released);
  assert.throws(() => wb.validateValue(0, 0, 0, { kind: 0 }), released);
  assert.throws(() => wb.getInvalidCells(0), released);
  assert.throws(() => wb.listThreadedComments(0), released);
  assert.throws(() => wb.listPeople(), released);
  assert.throws(() => wb.addPerson({ id: ANN, displayName: 'Ann' }), released);
});
