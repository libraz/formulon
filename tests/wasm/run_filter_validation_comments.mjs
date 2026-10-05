import assert from 'node:assert/strict';
import {
  DateTimeGrouping,
  DynamicFilterType,
  FilterKind,
  FilterOperator,
  SortBy,
  SortMethod,
  ValidationErrorStyle,
} from '../../packages/npm/common.mjs';

const ALICE = '{0A1B2C3D-0000-4000-8000-000000000001}';
const BOB = '{0A1B2C3D-0000-4000-8000-000000000002}';
const THREAD = '{6F1E0D2C-1111-4111-8111-000000000001}';
const REPLY = '{6F1E0D2C-1111-4111-8111-000000000002}';
const MENTION = '{6F1E0D2C-1111-4111-8111-0000000000AA}';
const CREATED = '2026-10-05T09:30:00.00';

function withWorkbook(Module, fn) {
  const wb = Module.Workbook.createDefault();
  try {
    fn(wb);
  } finally {
    wb.delete();
  }
}

// A1 is the header; A2:A6 hold 1..5.
function fillNumbers(wb) {
  assert.ok(wb.setText(0, 0, 0, 'N').ok);
  for (let i = 1; i <= 5; i += 1) assert.ok(wb.setNumber(0, i, 0, i).ok);
}

const range = (firstRow, firstCol, lastRow, lastCol) => ({ firstRow, firstCol, lastRow, lastCol });
const hiddenRows = (wb) =>
  wb
    .getSheetRowOverrides(0)
    .rows.filter((r) => r.hidden)
    .map((r) => r.row);

const customGreaterThanThree = {
  colId: 0,
  kind: FilterKind.Custom,
  customCount: 1,
  op1: FilterOperator.GreaterThan,
  val1: '3',
};

export function registerFilterValidationComments(Module, test) {
  test('filter, sort and validation enums keep their C ordinals', () => {
    assert.equal(FilterKind.Icon, 6);
    assert.equal(FilterOperator.GreaterThan, 5);
    assert.equal(DynamicFilterType.YearToDate, 18);
    assert.equal(DynamicFilterType.M12, 34);
    assert.equal(SortBy.Icon, 3);
    assert.equal(SortMethod.Stroke, 2);
    assert.equal(DateTimeGrouping.Second, 5);
    assert.equal(ValidationErrorStyle.Information, 2);
  });

  test('getAutoFilter / setAutoFilter round-trip every criterion field', () => {
    withWorkbook(Module, (wb) => {
      fillNumbers(wb);
      const absent = wb.getAutoFilter(0);
      assert.ok(absent.status.ok);
      assert.equal(absent.autoFilter, null);

      const filter = {
        range: range(0, 0, 5, 2),
        columns: [
          {
            colId: 0,
            kind: FilterKind.Values,
            filterBlank: true,
            values: ['a', 'b'],
            dateGroups: [
              { year: 2026, month: 10, day: 5, hour: 0, minute: 0, second: 0, grouping: DateTimeGrouping.Day },
            ],
          },
          {
            colId: 1,
            kind: FilterKind.Top10,
            top: true,
            percent: true,
            topVal: 10,
            hasFilterVal: true,
            filterVal: 4.5,
          },
          {
            colId: 2,
            kind: FilterKind.Custom,
            customAnd: true,
            customCount: 2,
            op1: FilterOperator.GreaterThan,
            val1: '1',
            op2: FilterOperator.LessThan,
            val2: '9',
          },
        ],
        sort: {
          ref: range(1, 0, 5, 2),
          columnSort: false,
          caseSensitive: true,
          sortMethod: SortMethod.None,
          conditions: [{ ref: range(1, 0, 5, 0), descending: true, sortBy: SortBy.Value, customList: '' }],
        },
      };
      assert.ok(wb.setAutoFilter(0, filter).ok);
      const got = wb.getAutoFilter(0);
      assert.ok(got.status.ok, JSON.stringify(got.status));
      const af = got.autoFilter;
      assert.deepEqual(af.range, filter.range);
      assert.equal(af.columns.length, 3);
      assert.deepEqual(af.columns[0].values, ['a', 'b']);
      assert.equal(af.columns[0].filterBlank, true);
      assert.deepEqual(af.columns[0].dateGroups, filter.columns[0].dateGroups);
      assert.equal(af.columns[1].top, true);
      assert.equal(af.columns[1].percent, true);
      assert.equal(af.columns[1].topVal, 10);
      assert.equal(af.columns[1].filterVal, 4.5);
      assert.equal(af.columns[2].customAnd, true);
      assert.equal(af.columns[2].val2, '9');
      assert.equal(af.sort.caseSensitive, true);
      assert.equal(af.sort.conditions.length, 1);
      assert.equal(af.sort.conditions[0].descending, true);

      // get -> set of the unchanged record is a no-op.
      assert.ok(wb.setAutoFilter(0, af).ok);
      assert.deepEqual(wb.getAutoFilter(0).autoFilter, af);

      assert.ok(wb.removeAutoFilter(0).ok);
      assert.equal(wb.getAutoFilter(0).autoFilter, null);
    });
  });

  test('evaluate, apply and clear a sheet AutoFilter', () => {
    withWorkbook(Module, (wb) => {
      fillNumbers(wb);
      assert.equal(wb.evaluateAutoFilter(0).status.ok, false);
      assert.equal(wb.applyAutoFilter(0).ok, false);
      assert.ok(wb.setAutoFilter(0, { range: range(0, 0, 5, 0), columns: [customGreaterThanThree], sort: null }).ok);

      const ev = wb.evaluateAutoFilter(0);
      assert.ok(ev.status.ok, JSON.stringify(ev.status));
      assert.equal(ev.firstRow, 1);
      assert.deepEqual(ev.match, [false, false, false, true, true]);
      assert.deepEqual(hiddenRows(wb), []);

      assert.ok(wb.applyAutoFilter(0).ok);
      assert.deepEqual(hiddenRows(wb), [1, 2, 3]);
      assert.ok(wb.clearAutoFilter(0).ok);
      assert.deepEqual(hiddenRows(wb), []);
      assert.equal(wb.getAutoFilter(0).autoFilter.columns.length, 0);
    });
  });

  test('table AutoFilter methods mirror the sheet ones', () => {
    withWorkbook(Module, (wb) => {
      fillNumbers(wb);
      const created = wb.createTable({ sheetIndex: 0, ref: 'A1:A6', name: 'Nums', columns: ['N'] });
      assert.ok(created.status.ok, JSON.stringify(created.status));
      const table = created.index;
      const got = wb.getTableAutoFilter(table);
      assert.ok(got.status.ok);
      assert.equal(got.autoFilter.range.lastRow, 5);

      assert.ok(
        wb.setTableAutoFilter(table, { range: range(0, 0, 5, 0), columns: [customGreaterThanThree], sort: null }).ok,
      );
      const ev = wb.evaluateTableAutoFilter(table);
      assert.deepEqual(ev.match, [false, false, false, true, true]);
      assert.ok(wb.applyTableAutoFilter(table).ok);
      assert.deepEqual(hiddenRows(wb), [1, 2, 3]);
      assert.ok(wb.clearTableAutoFilter(table).ok);
      assert.deepEqual(hiddenRows(wb), []);
      assert.ok(wb.removeTableAutoFilter(table).ok);
      assert.equal(wb.getTableAutoFilter(table).autoFilter, null);
      assert.equal(wb.getTableAutoFilter(99).status.ok, false);
    });
  });

  test('setAutoFilter rejects an out-of-range sheet and unordered columns', () => {
    withWorkbook(Module, (wb) => {
      const f = { range: range(0, 0, 5, 1), columns: [], sort: null };
      assert.equal(wb.setAutoFilter(99, f).ok, false);
      const bad = {
        ...f,
        columns: [
          { colId: 1, kind: FilterKind.None },
          { colId: 0, kind: FilterKind.None },
        ],
      };
      assert.equal(wb.setAutoFilter(0, bad).ok, false);
    });
  });

  test('validateValue reports the covering rule and listInvalidCells pages the failures', () => {
    withWorkbook(Module, (wb) => {
      const rule = {
        ranges: [range(0, 0, 9, 0)],
        type: 1,
        op: 0,
        formula1: '1',
        formula2: '10',
        errorStyle: ValidationErrorStyle.Warning,
      };
      assert.ok(wb.addValidation(0, rule).ok);

      const none = wb.validateValue(0, 0, 3, { kind: 1, number: 99 });
      assert.ok(none.status.ok);
      assert.equal(none.hasRule, false);
      assert.equal(none.valid, true);

      const good = wb.validateValue(0, 0, 0, { kind: 1, number: 5 });
      assert.equal(good.hasRule, true);
      assert.equal(good.valid, true);
      assert.equal(good.ruleIndex, 0);

      const bad = wb.validateValue(0, 0, 0, { kind: 1, number: 50 });
      assert.equal(bad.valid, false);
      assert.equal(bad.errorStyle, ValidationErrorStyle.Warning);
      assert.equal(wb.validateValue(0, 0, 0, { kind: 3, text: 'x' }).valid, false);
      assert.equal(wb.validateValue(99, 0, 0, { kind: 1, number: 1 }).status.ok, false);

      assert.ok(wb.setNumber(0, 0, 0, 5).ok);
      assert.ok(wb.setNumber(0, 1, 0, 50).ok);
      assert.ok(wb.setText(0, 2, 0, 'x').ok);
      assert.ok(wb.setNumber(0, 3, 0, 7).ok);
      assert.ok(wb.setNumber(0, 4, 0, 70).ok);

      const all = wb.listInvalidCells(0);
      assert.ok(all.status.ok, JSON.stringify(all.status));
      assert.deepEqual(
        all.cells.map((c) => [c.row, c.col]),
        [
          [1, 0],
          [2, 0],
          [4, 0],
        ],
      );
      assert.equal(all.nextCursor, null);

      const first = wb.listInvalidCells(0, 0, 2);
      assert.equal(first.cells.length, 2);
      assert.notEqual(first.nextCursor, null);
      const rest = wb.listInvalidCells(0, first.nextCursor, 2);
      assert.deepEqual(
        rest.cells.map((c) => c.row),
        [4],
      );
      assert.equal(rest.nextCursor, null);

      assert.equal(wb.listInvalidCells(99).status.ok, false);
      assert.equal(wb.listInvalidCells(0, -1).status.ok, false);
    });
  });

  test('persons and threaded comments support the full CRUD cycle', () => {
    withWorkbook(Module, (wb) => {
      assert.deepEqual(wb.getPersons().persons, []);
      assert.ok(wb.addPerson({ id: ALICE, displayName: 'Alice', userId: 'alice', providerId: 'None' }).ok);
      assert.ok(wb.addPerson({ id: BOB, displayName: 'Bob', userId: 'bob', providerId: 'None' }).ok);
      assert.equal(wb.addPerson({ id: ALICE, displayName: 'Dup' }).ok, false);
      const persons = wb.getPersons();
      assert.ok(persons.status.ok);
      assert.deepEqual(
        persons.persons.map((p) => p.displayName),
        ['Alice', 'Bob'],
      );
      assert.equal(persons.persons[0].userId, 'alice');

      assert.deepEqual(wb.getThreadedComments(0).comments, []);
      const opened = wb.addThreadedComment(0, {
        id: THREAD,
        row: 2,
        col: 1,
        personId: ALICE,
        created: CREATED,
        text: 'Check this, @Bob',
        mentions: [{ personId: BOB, mentionId: MENTION, start: 12, length: 4 }],
      });
      assert.ok(opened.ok, JSON.stringify(opened));
      assert.ok(
        wb.addThreadedComment(0, {
          id: REPLY,
          personId: BOB,
          created: CREATED,
          text: 'Done',
          parentId: THREAD,
        }).ok,
      );

      const list = wb.getThreadedComments(0);
      assert.ok(list.status.ok);
      assert.equal(list.comments.length, 2);
      const [head, reply] = list.comments;
      assert.equal(head.id, THREAD);
      assert.equal(head.row, 2);
      assert.equal(head.col, 1);
      assert.equal(head.done, false);
      assert.equal(head.parentId, '');
      assert.deepEqual(head.mentions, [{ personId: BOB, mentionId: MENTION, start: 12, length: 4 }]);
      assert.equal(reply.parentId, THREAD);
      assert.deepEqual(reply.mentions, []);

      assert.equal(wb.removePerson(BOB).ok, false);
      assert.ok(wb.editThreadedComment(0, THREAD, 'Edited', []).ok);
      assert.equal(wb.getThreadedComments(0).comments[0].text, 'Edited');
      assert.deepEqual(wb.getThreadedComments(0).comments[0].mentions, []);
      assert.equal(wb.editThreadedComment(0, '{00000000-0000-4000-8000-000000000000}', 'x', []).ok, false);

      assert.ok(wb.setThreadResolved(0, THREAD, true).ok);
      assert.equal(wb.getThreadedComments(0).comments[0].done, true);
      assert.ok(wb.setThreadResolved(0, THREAD, false).ok);
      assert.equal(wb.getThreadedComments(0).comments[0].done, false);

      assert.ok(wb.removeThreadedComment(0, REPLY).ok);
      assert.equal(wb.getThreadedComments(0).comments.length, 1);
      assert.ok(wb.removeThreadedComment(0, THREAD).ok);
      assert.equal(wb.getThreadedComments(0).comments.length, 0);
      assert.equal(wb.removeThreadedComment(0, THREAD).ok, false);

      assert.ok(wb.removePerson(BOB).ok);
      assert.equal(wb.getPersons().persons.length, 1);
      assert.equal(wb.getThreadedComments(99).status.ok, false);
    });
  });
}
