import test from 'node:test';
import { assert, getModule } from './smoke_support.mjs';

const ANN = '{11111111-1111-1111-1111-111111111111}';
const BOB = '{22222222-2222-2222-2222-222222222222}';
const THREAD = '{AAAAAAAA-0000-0000-0000-000000000001}';
const REPLY = '{AAAAAAAA-0000-0000-0000-000000000002}';
const MENTION = '{BBBBBBBB-0000-0000-0000-000000000001}';

function must(r, what) {
  const status = typeof r.ok === 'boolean' ? r : r.status;
  assert.ok(status.ok, `${what}: ${JSON.stringify(status)}`);
  return r;
}

function hiddenRows(wb) {
  return wb
    .getSheetRowOverrides(0)
    .rows.filter((r) => r.hidden)
    .map((r) => r.row);
}

function seedColumn(wb) {
  for (const [row, text] of ['h', 'a', 'b', 'a'].entries()) must(wb.setText(0, row, 0, text), 'setText');
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
    assert.ok(absent.status.ok);
    assert.equal(absent.autoFilter, null);

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
    must(wb.setAutoFilter(0, filter), 'setAutoFilter');
    const got = wb.getAutoFilter(0);
    must(got, 'getAutoFilter');
    assert.deepEqual(got.autoFilter.range, RANGE);
    assert.equal(got.autoFilter.columns.length, 1);
    const col = got.autoFilter.columns[0];
    assert.equal(col.kind, 1);
    assert.deepEqual(col.values, ['a']);
    assert.equal(col.showButton, true);
    assert.equal(col.hiddenButton, false);
    assert.equal(col.dateGroups.length, 1);
    assert.equal(col.dateGroups[0].year, 2026);
    assert.equal(col.dateGroups[0].month, 3);
    assert.equal(col.dateGroups[0].grouping, mod.DateTimeGrouping.Month);
    assert.equal(col.dxfId, 0);
    assert.equal(got.autoFilter.sort, null);
    // Setting the unchanged record back is accepted.
    must(wb.setAutoFilter(0, got.autoFilter), 'set round trip');
    assert.deepEqual(wb.getAutoFilter(0).autoFilter, got.autoFilter);

    // Replace the date group with a plain value filter: dateGroups drops out.
    must(
      wb.setAutoFilter(0, { range: RANGE, columns: [{ colId: 0, kind: mod.FilterKind.Values, values: ['a'] }] }),
      'set values',
    );

    const ev = wb.evaluateAutoFilter(0);
    must(ev, 'evaluateAutoFilter');
    assert.equal(ev.firstRow, 1);
    assert.deepEqual(ev.match, [true, false, true]);
    assert.deepEqual(hiddenRows(wb), []);

    must(wb.applyAutoFilter(0), 'applyAutoFilter');
    assert.deepEqual(hiddenRows(wb), [2]);
    must(wb.clearAutoFilter(0), 'clearAutoFilter');
    assert.deepEqual(hiddenRows(wb), []);
    assert.equal(wb.getAutoFilter(0).autoFilter.columns.length, 0);

    // A custom criterion and a sort state survive a round trip.
    must(
      wb.setAutoFilter(0, {
        range: RANGE,
        columns: [
          { colId: 0, kind: mod.FilterKind.Custom, customCount: 1, op1: mod.FilterOperator.NotEqual, val1: 'b' },
        ],
        sort: {
          ref: { firstRow: 1, firstCol: 0, lastRow: 3, lastCol: 0 },
          conditions: [{ ref: { firstRow: 1, firstCol: 0, lastRow: 3, lastCol: 0 }, descending: true }],
        },
      }),
      'set custom',
    );
    const custom = wb.getAutoFilter(0).autoFilter;
    assert.equal(custom.columns[0].kind, 2);
    assert.equal(custom.columns[0].op1, mod.FilterOperator.NotEqual);
    assert.equal(custom.columns[0].val1, 'b');
    assert.equal(custom.sort.conditions.length, 1);
    assert.equal(custom.sort.conditions[0].descending, true);
    assert.equal(custom.sort.columnSort, false);
    assert.deepEqual(wb.evaluateAutoFilter(0).match, [true, false, true]);

    must(wb.removeAutoFilter(0), 'removeAutoFilter');
    assert.equal(wb.getAutoFilter(0).autoFilter, null);
    const none = wb.evaluateAutoFilter(0);
    assert.equal(none.status.ok, false);
    assert.deepEqual(none.match, []);

    // An out-of-range column id is rejected and leaves the sheet unchanged.
    const bad = wb.setAutoFilter(0, {
      range: RANGE,
      columns: [{ colId: 5, kind: mod.FilterKind.Values, values: ['a'] }],
    });
    assert.equal(bad.ok, false);
    assert.equal(wb.getAutoFilter(0).autoFilter, null);
    assert.equal(wb.setAutoFilter(0, null).ok, false);
  } finally {
    wb.dispose();
  }
});

test('table AutoFilter entry points reject a table index that does not exist', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.equal(wb.tableCount().value, 0);
    const got = wb.getTableAutoFilter(0);
    assert.equal(got.status.ok, false);
    assert.equal(got.autoFilter, null);
    const ev = wb.evaluateTableAutoFilter(0);
    assert.equal(ev.status.ok, false);
    assert.deepEqual(ev.match, []);
    assert.equal(wb.setTableAutoFilter(0, { range: RANGE, columns: [] }).ok, false);
    assert.equal(wb.removeTableAutoFilter(0).ok, false);
    assert.equal(wb.applyTableAutoFilter(0).ok, false);
    assert.equal(wb.clearTableAutoFilter(0).ok, false);
  } finally {
    wb.dispose();
  }
});

test('validateValue and listInvalidCells follow the rule covering a cell', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const free = wb.validateValue(0, 0, 1, { kind: mod.ValueKind.Number, number: 99 });
    must(free, 'validateValue without rule');
    assert.equal(free.hasRule, false);
    assert.equal(free.valid, true);

    must(
      wb.addValidation(0, {
        ranges: [{ firstRow: 0, lastRow: 3, firstCol: 1, lastCol: 1 }],
        type: 1,
        op: 0,
        errorStyle: mod.ValidationErrorStyle.Warning,
        formula1: '1',
        formula2: '10',
        showErrorMessage: true,
      }),
      'addValidation',
    );
    const ok = wb.validateValue(0, 0, 1, { kind: mod.ValueKind.Number, number: 5 });
    must(ok, 'validateValue in range');
    assert.deepEqual(
      { hasRule: ok.hasRule, valid: ok.valid, ruleIndex: ok.ruleIndex, errorStyle: ok.errorStyle },
      { hasRule: true, valid: true, ruleIndex: 0, errorStyle: mod.ValidationErrorStyle.Warning },
    );
    const bad = wb.validateValue(0, 0, 1, { kind: mod.ValueKind.Number, number: 50 });
    assert.equal(bad.valid, false);
    assert.equal(bad.hasRule, true);
    const text = wb.validateValue(0, 0, 1, { kind: mod.ValueKind.Text, text: '5' });
    assert.equal(text.valid, false, 'text "5" is text, not a number');
    assert.equal(wb.validateValue(99, 0, 0, { kind: mod.ValueKind.Number, number: 1 }).status.ok, false);
    assert.equal(wb.validateValue(0, 0, 0, null).status.ok, false);

    must(wb.setNumber(0, 0, 1, 5), 'setNumber');
    must(wb.setNumber(0, 1, 1, 50), 'setNumber');
    must(wb.setNumber(0, 2, 1, 70), 'setNumber');
    const page = wb.listInvalidCells(0);
    must(page, 'listInvalidCells');
    assert.deepEqual(
      page.cells.map((c) => [c.row, c.col, c.value.number]),
      [
        [1, 1, 50],
        [2, 1, 70],
      ],
    );
    assert.equal(page.nextCursor, null);

    const first = wb.listInvalidCells(0, 0, 1);
    assert.equal(first.cells.length, 1);
    assert.equal(typeof first.nextCursor, 'number');
    const second = wb.listInvalidCells(0, first.nextCursor, 1);
    assert.deepEqual(
      second.cells.map((c) => c.row),
      [2],
    );
    assert.equal(wb.listInvalidCells(99).status.ok, false);
  } finally {
    wb.dispose();
  }
});

test('threaded comments cover a thread, a reply with a mention, resolution and removal', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.equal(wb.getPersons().persons.length, 0);
    assert.ok(wb.getPersons().status.ok);
    must(wb.addPerson({ id: ANN, displayName: 'Ann', userId: 'ann@example.com', providerId: 'AD' }), 'addPerson Ann');
    must(wb.addPerson({ id: BOB, displayName: 'Bob', userId: '', providerId: 'None' }), 'addPerson Bob');
    assert.equal(wb.addPerson({ id: ANN, displayName: 'Ann again' }).ok, false);
    const persons = wb.getPersons();
    assert.deepEqual(
      persons.persons.map((p) => [p.id, p.displayName, p.userId, p.providerId]),
      [
        [ANN, 'Ann', 'ann@example.com', 'AD'],
        [BOB, 'Bob', '', 'None'],
      ],
    );

    must(
      wb.addThreadedComment(0, {
        id: THREAD,
        row: 2,
        col: 3,
        personId: ANN,
        created: '2026-01-01T10:00:00.00',
        text: 'Review this',
      }),
      'open thread',
    );
    must(
      wb.addThreadedComment(0, {
        id: REPLY,
        personId: BOB,
        created: '2026-01-01T11:00:00.00',
        text: '@Ann done',
        parentId: THREAD,
        mentions: [{ personId: ANN, mentionId: MENTION, start: 0, length: 4 }],
      }),
      'reply',
    );
    const result = wb.getThreadedComments(0);
    must(result, 'getThreadedComments');
    const list = result.comments;
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

    must(
      wb.editThreadedComment(0, REPLY, '@Ann fixed', [{ personId: ANN, mentionId: MENTION, start: 0, length: 4 }]),
      'edit',
    );
    assert.equal(wb.getThreadedComments(0).comments[1].text, '@Ann fixed');
    must(wb.editThreadedComment(0, REPLY, 'plain', []), 'edit without mentions');
    assert.deepEqual(wb.getThreadedComments(0).comments[1].mentions, []);

    must(wb.setThreadResolved(0, THREAD, true), 'resolve');
    assert.equal(wb.getThreadedComments(0).comments[0].done, true);
    must(wb.setThreadResolved(0, THREAD, false), 'reopen');
    assert.equal(wb.getThreadedComments(0).comments[0].done, false);

    assert.equal(wb.removePerson(ANN).ok, false, 'a person with comments stays');
    assert.equal(wb.editThreadedComment(0, '{00000000-0000-0000-0000-000000000000}', 'x', []).ok, false);
    assert.equal(wb.setThreadResolved(0, REPLY, true).ok, false, 'a reply is not a thread');
    assert.equal(wb.addThreadedComment(0, null).ok, false);
    assert.equal(wb.getThreadedComments(99).status.ok, false);

    must(wb.removeThreadedComment(0, REPLY), 'remove reply');
    assert.equal(wb.getThreadedComments(0).comments.length, 1);
    must(wb.removeThreadedComment(0, THREAD), 'remove thread');
    assert.equal(wb.getThreadedComments(0).comments.length, 0);
    must(wb.removePerson(ANN), 'removePerson');
    assert.equal(wb.getPersons().persons.length, 1);
    assert.equal(wb.removePerson(ANN).ok, false);
  } finally {
    wb.dispose();
  }
});

test('the new accessors report a released handle', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.dispose();
  assert.equal(wb.getAutoFilter(0).status.status, 7000);
  assert.equal(wb.evaluateAutoFilter(0).status.status, 7000);
  assert.equal(wb.validateValue(0, 0, 0, { kind: 0 }).status.status, 7000);
  assert.equal(wb.listInvalidCells(0).status.status, 7000);
  assert.equal(wb.getThreadedComments(0).status.status, 7000);
  assert.equal(wb.getPersons().status.status, 7000);
  assert.equal(wb.addPerson({ id: ANN, displayName: 'Ann' }).status, 7000);
});
