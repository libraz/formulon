import test from 'node:test';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';

const ANN = '{11111111-1111-1111-1111-111111111111}';
const BOB = '{22222222-2222-2222-2222-222222222222}';
const THREAD = '{AAAAAAAA-0000-0000-0000-000000000001}';
const REPLY = '{AAAAAAAA-0000-0000-0000-000000000002}';
const MENTION = '{BBBBBBBB-0000-0000-0000-000000000001}';

function hiddenRows(wb) {
  return wb
    .sheet(0)
    .listRowOverrides()
    .filter((r) => r.hidden)
    .map((r) => r.row);
}

function seedColumn(wb) {
  for (const [row, text] of ['h', 'a', 'b', 'a'].entries()) wb.sheet(0).setText(row, 0, text);
}

const RANGE = { firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 0 };

test('filter, sort and validation enums keep their C ordinals', async () => {
  const mod = await getModule();
  assert.equal(mod.FilterKind.Icon, 6);
  assert.equal(mod.ComparisonOperator.GreaterThan, 6);
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
    const absent = wb.sheet(0).getAutoFilter();
    assert.equal(absent, null);

    const filter = {
      range: RANGE,
      columns: [
        {
          columnOffset: 0,
          kind: mod.FilterKind.Values,
          values: ['a'],
          dateGroups: [{ year: 2026, month: 3, grouping: mod.DateTimeGrouping.Month }],
        },
      ],
    };
    wb.sheet(0).setAutoFilter(filter);
    const got = wb.sheet(0).getAutoFilter();
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
    assert.equal(col.dxfIndex, 0);
    assert.equal(got.sort, null);
    // Setting the unchanged record back is accepted.
    wb.sheet(0).setAutoFilter(got);
    assert.deepEqual(wb.sheet(0).getAutoFilter(), got);

    // Replace the date group with a plain value filter: dateGroups drops out.
    wb.sheet(0).setAutoFilter({
      range: RANGE,
      columns: [{ columnOffset: 0, kind: mod.FilterKind.Values, values: ['a'] }],
    });

    const ev = wb.sheet(0).evaluateAutoFilter();
    assert.equal(ev.firstRow, 1);
    assert.deepEqual(ev.match, [true, false, true]);
    assert.deepEqual(hiddenRows(wb), []);

    wb.sheet(0).applyAutoFilter();
    assert.deepEqual(hiddenRows(wb), [2]);
    wb.sheet(0).clearAutoFilterCriteria();
    assert.deepEqual(hiddenRows(wb), []);
    assert.equal(wb.sheet(0).getAutoFilter().columns.length, 0);

    // A custom criterion and a sort state survive a round trip.
    wb.sheet(0).setAutoFilter({
      range: RANGE,
      columns: [
        {
          columnOffset: 0,
          kind: mod.FilterKind.Custom,
          conditions: [{ operator: mod.ComparisonOperator.NotEqual, value: 'b' }],
        },
      ],
      sort: {
        range: { firstRow: 1, firstCol: 0, lastRow: 3, lastCol: 0 },
        conditions: [{ range: { firstRow: 1, firstCol: 0, lastRow: 3, lastCol: 0 }, descending: true }],
      },
    });
    const custom = wb.sheet(0).getAutoFilter();
    assert.equal(custom.columns[0].kind, 2);
    assert.equal(custom.columns[0].conditions.length, 1);
    assert.equal(custom.columns[0].conditions[0].operator, mod.ComparisonOperator.NotEqual);
    assert.equal(custom.columns[0].conditions[0].value, 'b');
    assert.equal(custom.sort.conditions.length, 1);
    assert.equal(custom.sort.conditions[0].descending, true);
    assert.equal(custom.sort.columnSort, false);
    assert.deepEqual(wb.sheet(0).evaluateAutoFilter().match, [true, false, true]);

    // An omitted `top` reads as true (the schema default).
    wb.sheet(0).setAutoFilter({
      range: RANGE,
      columns: [{ columnOffset: 0, kind: mod.FilterKind.Top10, topValue: 2 }],
    });
    assert.equal(wb.sheet(0).getAutoFilter().columns[0].top, true);

    wb.sheet(0).setAutoFilter(null);
    assert.equal(wb.sheet(0).getAutoFilter(), null);
    assert.throws(() => wb.sheet(0).evaluateAutoFilter(), isFormulonError);

    // An out-of-range column id is rejected and leaves the sheet unchanged.
    assert.throws(
      () =>
        wb.sheet(0).setAutoFilter({
          range: RANGE,
          columns: [{ columnOffset: 5, kind: mod.FilterKind.Values, values: ['a'] }],
        }),
      isFormulonError,
    );
    assert.equal(wb.sheet(0).getAutoFilter(), null);
    // Removing a filter that is not there is a no-op.
    wb.sheet(0).setAutoFilter(null);
    assert.equal(wb.sheet(0).getAutoFilter(), null);
  } finally {
    wb.dispose();
  }
});

test('table AutoFilter entry points reject a table index that does not exist', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.equal(wb.sheet(0).listTables().length, 0);
    assert.throws(() => wb.sheet(0).table(0).getAutoFilter(), isFormulonError);
    assert.throws(() => wb.sheet(0).table(0).evaluateAutoFilter(), isFormulonError);
    assert.throws(() => wb.sheet(0).table(0).setAutoFilter({ range: RANGE, columns: [] }), isFormulonError);
    assert.throws(() => wb.sheet(0).table(0).setAutoFilter(null), isFormulonError);
    assert.throws(() => wb.sheet(0).table(0).applyAutoFilter(), isFormulonError);
    assert.throws(() => wb.sheet(0).table(0).clearAutoFilterCriteria(), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('validateValue and listInvalidCells follow the rule covering a cell', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const free = wb.sheet(0).validateValue(0, 1, { kind: mod.ValueKind.Number, number: 99 });
    assert.equal(free.ruleIndex, null);
    assert.equal(free.valid, true);

    wb.sheet(0).addValidation({
      ranges: [{ firstRow: 0, lastRow: 3, firstCol: 1, lastCol: 1 }],
      type: 1,
      operator: 0,
      errorStyle: mod.ValidationErrorStyle.Warning,
      formula1: '1',
      formula2: '10',
      showErrorMessage: true,
    });
    const ok = wb.sheet(0).validateValue(0, 1, { kind: mod.ValueKind.Number, number: 5 });
    assert.deepEqual(
      { valid: ok.valid, ruleIndex: ok.ruleIndex, errorStyle: ok.errorStyle },
      { valid: true, ruleIndex: 0, errorStyle: mod.ValidationErrorStyle.Warning },
    );
    const bad = wb.sheet(0).validateValue(0, 1, { kind: mod.ValueKind.Number, number: 50 });
    assert.equal(bad.valid, false);
    assert.equal(bad.ruleIndex, 0);
    const text = wb.sheet(0).validateValue(0, 1, { kind: mod.ValueKind.Text, text: '5' });
    assert.equal(text.valid, false, 'text "5" is text, not a number');
    assert.throws(() => wb.sheet(99).validateValue(0, 0, { kind: mod.ValueKind.Number, number: 1 }), isFormulonError);
    assert.equal(wb.sheet(0).validateValue(0, 0, null).valid, true, 'null is a blank value');
    assert.throws(() => wb.sheet(0).validateValue(0, 0, undefined), TypeError);
    assert.throws(
      () => wb.sheet(0).validateValue(0, 0, { number: 50 }),
      TypeError,
      'a value without `kind` is rejected, not read as blank',
    );

    wb.sheet(0).setNumber(0, 1, 5);
    wb.sheet(0).setNumber(1, 1, 50);
    wb.sheet(0).setNumber(2, 1, 70);
    const page = wb.sheet(0).getInvalidCells();
    assert.deepEqual(
      page.cells.map((c) => [c.row, c.col, c.value.number]),
      [
        [1, 1, 50],
        [2, 1, 70],
      ],
    );
    assert.equal(page.nextCursor, null);

    const first = wb.sheet(0).getInvalidCells(0, 1);
    assert.equal(first.cells.length, 1);
    assert.equal(typeof first.nextCursor, 'number');
    const second = wb.sheet(0).getInvalidCells(first.nextCursor, 1);
    assert.deepEqual(
      second.cells.map((c) => c.row),
      [2],
    );
    assert.throws(() => wb.sheet(99).getInvalidCells(), isFormulonError);
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

    wb.sheet(0).addThreadedComment({
      id: THREAD,
      cell: { row: 2, col: 3 },
      personId: ANN,
      created: { year: 2026, month: 1, day: 1, hour: 10, minute: 0, second: 0 },
      text: 'Review this',
    });
    wb.sheet(0).addThreadedComment({
      id: REPLY,
      personId: BOB,
      created: { year: 2026, month: 1, day: 1, hour: 11, minute: 0, second: 0 },
      text: '@Ann done',
      parentId: THREAD,
      mentions: [{ personId: ANN, mentionId: MENTION, start: 0, length: 4 }],
    });
    const result = wb.sheet(0).listThreadedComments();
    const list = result;
    assert.equal(list.length, 2);
    assert.equal(list[0].id, THREAD);
    assert.equal(list[0].parentId, '');
    assert.deepEqual({ ...list[0].cell }, { row: 2, col: 3 });
    assert.equal(list[0].done, false);
    assert.deepEqual(list[0].mentions, []);
    assert.equal(list[1].id, REPLY);
    assert.equal(list[1].parentId, THREAD);
    assert.equal(list[1].personId, BOB);
    assert.deepEqual(list[1].mentions, [{ personId: ANN, mentionId: MENTION, start: 0, length: 4 }]);

    wb.sheet(0).updateThreadedComment(REPLY, {
      text: '@Ann fixed',
      mentions: [{ personId: ANN, mentionId: MENTION, start: 0, length: 4 }],
    });
    assert.equal(wb.sheet(0).listThreadedComments()[1].text, '@Ann fixed');
    wb.sheet(0).updateThreadedComment(REPLY, { text: 'plain', mentions: [] });
    assert.deepEqual(wb.sheet(0).listThreadedComments()[1].mentions, []);

    wb.sheet(0).updateThreadedComment(THREAD, { done: true });
    assert.equal(wb.sheet(0).listThreadedComments()[0].done, true);
    wb.sheet(0).updateThreadedComment(THREAD, { done: false });
    assert.equal(wb.sheet(0).listThreadedComments()[0].done, false);

    assert.throws(() => wb.removePerson(ANN), isFormulonError, 'a person with comments stays');
    assert.throws(
      () => wb.sheet(0).updateThreadedComment('{00000000-0000-0000-0000-000000000000}', { text: 'x', mentions: [] }),
      isFormulonError,
    );
    assert.throws(
      () => wb.sheet(0).updateThreadedComment(REPLY, { done: true }),
      (e) => isFormulonError(e, 9006),
      'a reply is not a thread',
    );
    assert.throws(() => wb.sheet(0).updateThreadedComment(THREAD, { done: null }), TypeError);
    assert.throws(() => wb.sheet(0).addThreadedComment(null), TypeError);
    assert.throws(() => wb.sheet(99).listThreadedComments(), isFormulonError);

    wb.sheet(0).removeThreadedComment(REPLY);
    assert.equal(wb.sheet(0).listThreadedComments().length, 1);
    wb.sheet(0).removeThreadedComment(THREAD);
    assert.equal(wb.sheet(0).listThreadedComments().length, 0);
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
  assert.throws(() => wb.sheet(0).getAutoFilter(), released);
  assert.throws(() => wb.sheet(0).evaluateAutoFilter(), released);
  assert.throws(() => wb.sheet(0).validateValue(0, 0, { kind: 0 }), released);
  assert.throws(() => wb.sheet(0).getInvalidCells(), released);
  assert.throws(() => wb.sheet(0).listThreadedComments(), released);
  assert.throws(() => wb.listPeople(), released);
  assert.throws(() => wb.addPerson({ id: ANN, displayName: 'Ann' }), released);
});
