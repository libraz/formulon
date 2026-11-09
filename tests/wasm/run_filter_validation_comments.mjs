import assert from 'node:assert/strict';
import {
  ComparisonOperator,
  DateTimeGrouping,
  DynamicFilterType,
  FilterKind,
  SortBy,
  SortMethod,
  StatusCode,
  ValidationErrorStyle,
} from '../../packages/npm/common.mjs';
import { expectFailure, readAll } from './smoke_support.mjs';

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
  wb.setText(0, 0, 0, 'N');
  for (let i = 1; i <= 5; i += 1) wb.setNumber(0, i, 0, i);
}

const range = (firstRow, firstCol, lastRow, lastCol) => ({ firstRow, firstCol, lastRow, lastCol });
const hiddenRows = (wb) =>
  wb
    .listRowOverrides(0)
    .filter((r) => r.hidden)
    .map((r) => r.row);

const customGreaterThanThree = {
  columnOffset: 0,
  kind: FilterKind.Custom,
  conditions: [{ operator: ComparisonOperator.GreaterThan, value: '3' }],
};

export function registerFilterValidationComments(Module, test) {
  test('filter, sort and validation enums keep their C ordinals', () => {
    assert.equal(FilterKind.Icon, 6);
    assert.equal(ComparisonOperator.GreaterThan, 6);
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
      assert.equal(absent, null);

      const filter = {
        range: range(0, 0, 5, 2),
        columns: [
          {
            columnOffset: 0,
            kind: FilterKind.Values,
            filterBlank: true,
            values: ['a', 'b'],
            dateGroups: [
              { year: 2026, month: 10, day: 5, hour: 0, minute: 0, second: 0, grouping: DateTimeGrouping.Day },
            ],
          },
          {
            columnOffset: 1,
            kind: FilterKind.Top10,
            top: true,
            percent: true,
            topValue: 10,
            filterValue: 4.5,
          },
          {
            columnOffset: 2,
            kind: FilterKind.Custom,
            customAnd: true,
            conditions: [
              { operator: ComparisonOperator.GreaterThan, value: '1' },
              { operator: ComparisonOperator.LessThan, value: '9' },
            ],
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
      wb.setAutoFilter(0, filter);
      const got = wb.getAutoFilter(0);
      const af = got;
      assert.deepEqual(af.range, filter.range);
      assert.equal(af.columns.length, 3);
      assert.deepEqual(af.columns[0].values, ['a', 'b']);
      assert.equal(af.columns[0].filterBlank, true);
      assert.deepEqual(af.columns[0].dateGroups, filter.columns[0].dateGroups);
      assert.equal(af.columns[1].top, true);
      assert.equal(af.columns[1].percent, true);
      assert.equal(af.columns[1].topValue, 10);
      assert.equal(af.columns[1].filterValue, 4.5);
      assert.equal(af.columns[0].filterValue, null);
      assert.equal(af.columns[0].iconIndex, null);
      assert.equal(af.sort.conditions[0].dxfIndex, null);
      assert.equal(af.columns[2].customAnd, true);
      assert.deepEqual(af.columns[2].conditions, filter.columns[2].conditions);
      assert.equal(af.sort.caseSensitive, true);
      assert.equal(af.sort.conditions.length, 1);
      assert.equal(af.sort.conditions[0].descending, true);

      // get -> set of the unchanged record is a no-op.
      wb.setAutoFilter(0, af);
      assert.deepEqual(wb.getAutoFilter(0), af);

      // An omitted `top` reads as true (the schema default).
      wb.setAutoFilter(0, { range: af.range, columns: [{ columnOffset: 0, kind: FilterKind.Top10, topValue: 2 }] });
      assert.equal(wb.getAutoFilter(0).columns[0].top, true);

      wb.setAutoFilter(0, null);
      assert.equal(wb.getAutoFilter(0), null);
    });
  });

  test('evaluate, apply and clear a sheet AutoFilter', () => {
    withWorkbook(Module, (wb) => {
      fillNumbers(wb);
      expectFailure(() => wb.evaluateAutoFilter(0));
      expectFailure(() => wb.applyAutoFilter(0));
      wb.setAutoFilter(0, { range: range(0, 0, 5, 0), columns: [customGreaterThanThree], sort: null });

      const ev = wb.evaluateAutoFilter(0);
      assert.equal(ev.firstRow, 1);
      assert.deepEqual(ev.match, [false, false, false, true, true]);
      assert.deepEqual(hiddenRows(wb), []);

      wb.applyAutoFilter(0);
      assert.deepEqual(hiddenRows(wb), [1, 2, 3]);
      wb.clearAutoFilterCriteria(0);
      assert.deepEqual(hiddenRows(wb), []);
      assert.equal(wb.getAutoFilter(0).columns.length, 0);
    });
  });

  test('table AutoFilter methods mirror the sheet ones', () => {
    withWorkbook(Module, (wb) => {
      fillNumbers(wb);
      const table = wb.addTable(0, {
        range: { firstRow: 0, firstCol: 0, lastRow: 5, lastCol: 0 },
        name: 'Nums',
        columns: ['N'],
      });
      const got = wb.getTableAutoFilter(0, table);
      assert.equal(got.range.lastRow, 5);

      wb.setTableAutoFilter(0, table, { range: range(0, 0, 5, 0), columns: [customGreaterThanThree], sort: null });
      const ev = wb.evaluateTableAutoFilter(0, table);
      assert.deepEqual(ev.match, [false, false, false, true, true]);
      wb.applyTableAutoFilter(0, table);
      assert.deepEqual(hiddenRows(wb), [1, 2, 3]);
      wb.clearTableAutoFilterCriteria(0, table);
      assert.deepEqual(hiddenRows(wb), []);
      wb.setTableAutoFilter(0, table, null);
      assert.equal(wb.getTableAutoFilter(0, table), null);
      expectFailure(() => wb.getTableAutoFilter(0, 99), StatusCode.IndexOutOfRange);
      expectFailure(() => wb.getTableAutoFilter(9, 0), StatusCode.SheetIndexOutOfRange);
    });
  });

  test('setAutoFilter rejects an out-of-range sheet and unordered columns', () => {
    withWorkbook(Module, (wb) => {
      const f = { range: range(0, 0, 5, 1), columns: [], sort: null };
      expectFailure(() => wb.setAutoFilter(99, f));
      const bad = {
        ...f,
        columns: [
          { columnOffset: 1, kind: FilterKind.None },
          { columnOffset: 0, kind: FilterKind.None },
        ],
      };
      expectFailure(() => wb.setAutoFilter(0, bad));
    });
  });

  test('UI numeric readers throw on lossy validation fields before mutation', () => {
    withWorkbook(Module, (wb) => {
      const base = {
        ranges: [range(0, 0, 2, 0)],
        type: 1,
        operator: 0,
        formula1: '1',
        formula2: '10',
      };
      wb.addValidation(0, base);
      const before = readAll(
        () => wb.getValidationCount(0),
        (i) => wb.getValidation(0, i),
      );
      assert.equal(before.length, 1);

      for (const value of [256, Number.NaN, Number.POSITIVE_INFINITY, 1.5, -1, Symbol('type')]) {
        assert.throws(
          () => wb.addValidation(0, { ...base, type: value }),
          typeof value === 'symbol' ? TypeError : RangeError,
          `type=${String(value)}`,
        );
        assert.equal(
          readAll(
            () => wb.getValidationCount(0),
            (i) => wb.getValidation(0, i),
          ).length,
          before.length,
        );
      }

      assert.throws(
        () =>
          wb.addValidation(0, {
            ...base,
            ranges: [{ ...base.ranges[0], firstRow: 2 ** 32 }],
          }),
        RangeError,
      );
      assert.deepEqual(
        readAll(
          () => wb.getValidationCount(0),
          (i) => wb.getValidation(0, i),
        ),
        before,
      );

      // Nullish enum fields keep their documented zero defaults.
      wb.addValidation(0, { ranges: [range(3, 0, 3, 0)] });
      assert.equal(
        readAll(
          () => wb.getValidationCount(0),
          (i) => wb.getValidation(0, i),
        )[1].type,
        0,
      );
    });
  });

  test('AutoFilter date-group integers throw on wraparound before replacing a filter', () => {
    withWorkbook(Module, (wb) => {
      const good = {
        range: range(0, 0, 2, 0),
        columns: [
          {
            columnOffset: 0,
            kind: FilterKind.Values,
            values: ['1'],
            dateGroups: [{ year: 2026, month: 1, day: 0, hour: 0, minute: 0, second: 0, grouping: 1 }],
          },
        ],
        sort: null,
      };
      wb.setAutoFilter(0, good);
      const before = wb.getAutoFilter(0);
      const bad = {
        ...good,
        columns: [
          {
            ...good.columns[0],
            dateGroups: [{ ...good.columns[0].dateGroups[0], month: 257 }],
          },
        ],
      };
      assert.throws(() => wb.setAutoFilter(0, bad), RangeError);
      assert.deepEqual(wb.getAutoFilter(0), before);

      assert.throws(
        () =>
          wb.setAutoFilter(0, {
            ...good,
            columns: [{ ...good.columns[0], columnOffset: Number.NaN }],
          }),
        RangeError,
      );
      assert.deepEqual(wb.getAutoFilter(0), before);

      assert.throws(
        () =>
          wb.setAutoFilter(0, {
            ...good,
            columns: [{ ...good.columns[0], kind: 'bad' }],
          }),
        TypeError,
      );
      assert.deepEqual(wb.getAutoFilter(0), before);

      wb.setAutoFilter(0, {
        range: range(0, 0, 2, 1),
        columns: [
          { columnOffset: null, kind: undefined },
          { columnOffset: 1, kind: 1, dateGroups: [{ month: null }] },
        ],
        sort: null,
      });
      const stored = wb.getAutoFilter(0);
      assert.equal(stored.columns[0].columnOffset, 0);
      assert.equal(stored.columns[0].kind, 0);
      assert.equal(stored.columns[1].dateGroups[0].month, 0);
    });
  });

  test('threaded mention and image integer readers throw before C mutation', () => {
    withWorkbook(Module, (wb) => {
      wb.addPerson({ id: ALICE, displayName: 'Alice' });
      const beforeComments = wb.listThreadedComments(0).length;
      assert.throws(
        () =>
          wb.addThreadedComment(0, {
            id: THREAD,
            row: 0,
            col: 0,
            personId: ALICE,
            created: CREATED,
            text: '@Alice',
            mentions: [{ personId: ALICE, mentionId: MENTION, start: 2 ** 32, length: 1 }],
          }),
        RangeError,
      );
      assert.equal(wb.listThreadedComments(0).length, beforeComments);

      for (const value of [Number.NaN, 1.5, Symbol('fromRowOffsetEmu')]) {
        assert.throws(
          () => wb.addImage(0, new Uint8Array([1, 2, 3]), { anchor: { fromRowOffsetEmu: value } }),
          typeof value === 'symbol' ? TypeError : RangeError,
          `fromRowOffsetEmu=${String(value)}`,
        );
        assert.equal(wb.listDrawingObjects(0).length, 0);
      }

      expectFailure(() =>
        wb.addImage(0, new Uint8Array([1, 2, 3]), {
          anchor: { fromRowOffsetEmu: -(2 ** 63), fromColOffsetEmu: 2 ** 63 - 1024 },
        }),
      );
      assert.equal(wb.listDrawingObjects(0).length, 0);

      assert.throws(() => wb.addImage(0, new Uint8Array([1, 2, 3]), { anchor: { widthEmu: 2 ** 63 } }), RangeError);
      assert.equal(wb.listDrawingObjects(0).length, 0);
    });
  });

  test('nested readers rethrow getter failures and reject hostile string conversions', () => {
    withWorkbook(Module, (wb) => {
      const beforeValidationCount = readAll(
        () => wb.getValidationCount(0),
        (i) => wb.getValidation(0, i),
      ).length;
      const largeText = 'x'.repeat(1024 * 1024);
      for (let i = 0; i < 20; i += 1) {
        const input = { ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }], formula1: largeText };
        Object.defineProperty(input, 'formula2', {
          configurable: true,
          enumerable: true,
          get() {
            throw new Error('formula2 getter should stay inside the JS reader boundary');
          },
        });
        assert.throws(
          () => wb.addValidation(0, input),
          /formula2 getter should stay inside the JS reader boundary/,
          `throwing getter unexpectedly succeeded at iteration ${i}`,
        );
        assert.equal(
          readAll(
            () => wb.getValidationCount(0),
            (i) => wb.getValidation(0, i),
          ).length,
          beforeValidationCount,
        );
      }

      const proxyFilter = {
        range: { firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 },
        columns: new Proxy([], {
          get() {
            throw new Error('columns getter should stay inside the JS reader boundary');
          },
        }),
      };
      assert.throws(() => wb.setAutoFilter(0, proxyFilter), /columns getter should stay inside the JS reader boundary/);

      const badString = {
        ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
        formula1: {
          toString() {
            throw new Error('toString must not run');
          },
        },
      };
      assert.throws(() => wb.addValidation(0, badString), TypeError);
      assert.equal(
        readAll(
          () => wb.getValidationCount(0),
          (i) => wb.getValidation(0, i),
        ).length,
        beforeValidationCount,
      );

      const bytesProxy = new Proxy(new Uint8Array([1, 2, 3]), {
        get() {
          throw new Error('bytes getter should stay inside the JS reader boundary');
        },
      });
      assert.throws(() => wb.addImage(0, bytesProxy, {}), TypeError);
      assert.equal(wb.listDrawingObjects(0).length, 0);
    });
  });

  test('validateValue reports the covering rule and listInvalidCells pages the failures', () => {
    withWorkbook(Module, (wb) => {
      const rule = {
        ranges: [range(0, 0, 9, 0)],
        type: 1,
        operator: 0,
        formula1: '1',
        formula2: '10',
        errorStyle: ValidationErrorStyle.Warning,
      };
      const ruleId = wb.addValidation(0, rule);

      const none = wb.validateValue(0, 0, 3, { kind: 1, number: 99 });
      assert.deepEqual({ ...none }, { valid: true, ruleId: null, errorStyle: null });

      const good = wb.validateValue(0, 0, 0, { kind: 1, number: 5 });
      assert.equal(good.valid, true);
      assert.equal(good.ruleId, ruleId);

      const bad = wb.validateValue(0, 0, 0, { kind: 1, number: 50 });
      assert.equal(bad.valid, false);
      assert.equal(bad.errorStyle, ValidationErrorStyle.Warning);
      assert.equal(wb.validateValue(0, 0, 0, { kind: 3, text: 'x' }).valid, false);
      expectFailure(() => wb.validateValue(99, 0, 0, { kind: 1, number: 1 }));
      assert.throws(() => wb.validateValue(0, 0, 0, { number: 50 }), TypeError);

      wb.setNumber(0, 0, 0, 5);
      wb.setNumber(0, 1, 0, 50);
      wb.setText(0, 2, 0, 'x');
      wb.setNumber(0, 3, 0, 7);
      wb.setNumber(0, 4, 0, 70);

      const all = wb.getInvalidCells(0);
      assert.deepEqual(
        all.cells.map((c) => [c.row, c.col]),
        [
          [1, 0],
          [2, 0],
          [4, 0],
        ],
      );
      assert.equal(all.nextCursor, null);

      const first = wb.getInvalidCells(0, 0, 2);
      assert.equal(first.cells.length, 2);
      assert.notEqual(first.nextCursor, null);
      const rest = wb.getInvalidCells(0, first.nextCursor, 2);
      assert.deepEqual(
        rest.cells.map((c) => c.row),
        [4],
      );
      assert.equal(rest.nextCursor, null);

      expectFailure(() => wb.getInvalidCells(99));
      assert.throws(() => wb.getInvalidCells(0, -1), RangeError);
    });
  });

  test('addThreadedComment returns the id, generating one and stamping the pinned clock when omitted', () => {
    withWorkbook(Module, (wb) => {
      wb.addPerson({ id: ALICE, displayName: 'Alice' });
      wb.setPinnedNow(2026, 10, 5, 9, 30, 0);
      const generated = wb.addThreadedComment(0, { row: 0, col: 0, personId: ALICE, text: 'First' });
      assert.match(generated, /^\{[0-9A-F]{8}(-[0-9A-F]{4}){3}-[0-9A-F]{12}\}$/);
      const read = wb.getThreadedComment(0, generated);
      assert.equal(read.id, generated);
      assert.equal(read.created, '2026-10-05T09:30:00.00');

      const given = wb.addThreadedComment(0, { id: THREAD, row: 1, col: 0, personId: ALICE, text: 'Second' });
      assert.equal(given, THREAD);
      assert.equal(wb.getThreadedComment(0, THREAD).created, '2026-10-05T09:30:00.00');
      const stamped = wb.addThreadedComment(0, {
        row: 2,
        col: 0,
        personId: ALICE,
        text: 'Third',
        created: CREATED,
      });
      assert.equal(wb.getThreadedComment(0, stamped).created, CREATED);
      assert.notEqual(stamped, generated);

      expectFailure(
        () => wb.addThreadedComment(0, { id: THREAD, row: 3, col: 0, personId: ALICE, text: 'Dup' }),
        StatusCode.AlreadyExists,
      );
      expectFailure(
        () => wb.addThreadedComment(0, { row: 4, col: 0, personId: ALICE, text: 'Orphan', parentId: BOB }),
        StatusCode.NotFound,
      );
      expectFailure(
        () => wb.addThreadedComment(0, { row: 5, col: 0, personId: BOB, text: 'Nobody' }),
        StatusCode.NotFound,
      );
    });
  });

  test('persons and threaded comments support the full CRUD cycle', () => {
    withWorkbook(Module, (wb) => {
      assert.deepEqual(wb.listPeople(), []);
      wb.addPerson({ id: ALICE, displayName: 'Alice', userId: 'alice', providerId: 'None' });
      wb.addPerson({ id: BOB, displayName: 'Bob', userId: 'bob', providerId: 'None' });
      expectFailure(() => wb.addPerson({ id: ALICE, displayName: 'Dup' }));
      const persons = wb.listPeople();
      assert.deepEqual(
        persons.map((p) => p.displayName),
        ['Alice', 'Bob'],
      );
      assert.equal(persons[0].userId, 'alice');

      assert.deepEqual(wb.listThreadedComments(0), []);
      wb.addThreadedComment(0, {
        id: THREAD,
        row: 2,
        col: 1,
        personId: ALICE,
        created: CREATED,
        text: 'Check this, @Bob',
        mentions: [{ personId: BOB, mentionId: MENTION, start: 12, length: 4 }],
      });
      wb.addThreadedComment(0, {
        id: REPLY,
        personId: BOB,
        created: CREATED,
        text: 'Done',
        parentId: THREAD,
      });

      const list = wb.listThreadedComments(0);
      assert.equal(list.length, 2);
      const [head, reply] = list;
      assert.equal(head.id, THREAD);
      assert.equal(head.row, 2);
      assert.equal(head.col, 1);
      assert.equal(head.done, false);
      assert.equal(head.parentId, null);
      assert.deepEqual(head.mentions, [{ personId: BOB, mentionId: MENTION, start: 12, length: 4 }]);
      assert.equal(reply.parentId, THREAD);
      assert.deepEqual(reply.mentions, []);

      expectFailure(() => wb.removePerson(BOB));
      wb.updateThreadedComment(0, THREAD, { text: 'Edited', mentions: [] });
      assert.equal(wb.listThreadedComments(0)[0].text, 'Edited');
      assert.deepEqual(wb.listThreadedComments(0)[0].mentions, []);
      expectFailure(() => wb.updateThreadedComment(0, '{00000000-0000-4000-8000-000000000000}', { text: 'x' }));

      wb.updateThreadedComment(0, THREAD, { done: true });
      assert.equal(wb.listThreadedComments(0)[0].done, true);
      assert.equal(wb.listThreadedComments(0)[0].text, 'Edited');
      wb.updateThreadedComment(0, THREAD, { done: false });
      assert.equal(wb.listThreadedComments(0)[0].done, false);
      expectFailure(() => wb.updateThreadedComment(0, REPLY, { done: true }));

      wb.removeThreadedComment(0, REPLY);
      assert.equal(wb.listThreadedComments(0).length, 1);
      wb.removeThreadedComment(0, THREAD);
      assert.equal(wb.listThreadedComments(0).length, 0);
      expectFailure(() => wb.removeThreadedComment(0, THREAD));

      wb.removePerson(BOB);
      assert.equal(wb.listPeople().length, 1);
      expectFailure(() => wb.listThreadedComments(99));
    });
  });
}
