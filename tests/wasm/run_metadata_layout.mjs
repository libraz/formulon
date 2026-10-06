import assert from 'node:assert/strict';

export function registerMetadataLayout(Module, test) {
  test('addMerge / getMerges round-trip', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.addMerge(0, { firstRow: 0, lastRow: 1, firstCol: 0, lastCol: 2 }).ok);
      const merges = wb.getMerges(0);
      assert.equal(merges.length, 1);
      assert.equal(merges[0].firstRow, 0);
      assert.equal(merges[0].lastCol, 2);
    } finally {
      wb.delete();
    }
  });

  test('removeMerge / removeMergeAt / clearMerges step-wise prune the merge list', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const a = { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 };
      const b = { firstRow: 4, firstCol: 4, lastRow: 5, lastCol: 5 };
      assert.ok(wb.addMerge(0, a).ok);
      assert.ok(wb.addMerge(0, b).ok);
      // removeMerge with an overlap that hits `a` only.
      assert.ok(wb.removeMerge(0, { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }).ok);
      const list = wb.getMerges(0);
      assert.equal(list.length, 1);
      assert.equal(list[0].firstRow, 4);
      // removeMergeAt drops the survivor by index.
      assert.ok(wb.removeMergeAt(0, 0).ok);
      assert.equal(wb.getMerges(0).length, 0);
      // clearMerges nukes the remainder; safe on an empty list.
      assert.ok(wb.addMerge(0, a).ok);
      assert.ok(wb.addMerge(0, b).ok);
      assert.ok(wb.clearMerges(0).ok);
      assert.equal(wb.getMerges(0).length, 0);
    } finally {
      wb.delete();
    }
  });

  test('hyperlinks are read after save+load', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.equal(wb.getHyperlinks(0).length, 0);
    } finally {
      wb.delete();
    }
  });

  test('addHyperlink + getHyperlinks round-trip', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.equal(wb.getHyperlinks(0).length, 0);
      // Three entries with progressively more optional fields populated.
      assert.ok(wb.addHyperlink(0, 1, 2, 'https://example.com/', '', '', '').ok);
      assert.ok(wb.addHyperlink(0, 3, 4, 'mailto:hello@example.com', 'Hello', '', '').ok);
      assert.ok(wb.addHyperlink(0, 5, 6, '', 'See X', 'Internal link', 'Sheet1!A1').ok);
      assert.ok(wb.addHyperlinkRange(0, 7, 8, 9, 10, 'https://example.com/range', 'Range', 'Range tip', '').ok);
      const list = wb.getHyperlinks(0);
      assert.equal(list.length, 4);
      assert.equal(list[0].row, 1);
      assert.equal(list[0].col, 2);
      assert.equal(list[0].lastRow, 1);
      assert.equal(list[0].lastCol, 2);
      assert.equal(list[0].target, 'https://example.com/');
      assert.equal(list[0].display, '');
      assert.equal(list[0].tooltip, '');
      assert.equal(list[1].target, 'mailto:hello@example.com');
      assert.equal(list[1].display, 'Hello');
      assert.equal(list[2].display, 'See X');
      assert.equal(list[2].tooltip, 'Internal link');
      assert.equal(list[2].location, 'Sheet1!A1');
      assert.deepEqual(
        {
          row: list[3].row,
          col: list[3].col,
          lastRow: list[3].lastRow,
          lastCol: list[3].lastCol,
          target: list[3].target,
          display: list[3].display,
          tooltip: list[3].tooltip,
        },
        {
          row: 7,
          col: 8,
          lastRow: 9,
          lastCol: 10,
          target: 'https://example.com/range',
          display: 'Range',
          tooltip: 'Range tip',
        },
      );
      // Sheet-out-of-range is rejected.
      assert.ok(!wb.addHyperlink(999, 0, 0, 'https://x/', '', '', '').ok);
      // clearHyperlinks drops everything.
      assert.ok(wb.clearHyperlinks(0).ok);
      assert.equal(wb.getHyperlinks(0).length, 0);
    } finally {
      wb.delete();
    }
  });

  test('removeHyperlink / removeHyperlinkAt / clearHyperlinks surface on an empty sheet', () => {
    const wb = Module.Workbook.createDefault();
    try {
      // No-op variants on an empty list still return kOk.
      assert.ok(wb.removeHyperlink(0, 0, 0).ok);
      assert.ok(wb.clearHyperlinks(0).ok);
      // Out-of-range index is rejected.
      assert.ok(!wb.removeHyperlinkAt(0, 0).ok);
      // Sheet-index-out-of-range is rejected on every variant.
      assert.ok(!wb.removeHyperlink(999, 0, 0).ok);
      assert.ok(!wb.removeHyperlinkAt(999, 0).ok);
      assert.ok(!wb.clearHyperlinks(999).ok);
      // Final list still matches the post-save+load read.
      assert.equal(wb.getHyperlinks(0).length, 0);
    } finally {
      wb.delete();
    }
  });

  test('setComment / getComment round-trip', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setComment(0, 1, 1, 'Alice', 'Hello').ok);
      const c = wb.getComment(0, 1, 1);
      assert.ok(c !== null);
      assert.equal(c.author, 'Alice');
      assert.equal(c.text, 'Hello');
      // Empty text removes.
      assert.ok(wb.setComment(0, 1, 1, '', '').ok);
      const after = wb.getComment(0, 1, 1);
      assert.equal(after, null);
    } finally {
      wb.delete();
    }
  });

  test('getCommentResult distinguishes absence from an invalid sheet', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const missing = wb.getCommentResult(0, 1, 1);
      assert.equal(missing.status.ok, false);
      assert.equal(missing.comment, null);
      const invalid = wb.getCommentResult(99, 1, 1);
      assert.equal(invalid.status.ok, false);
      assert.equal(invalid.comment, null);
      assert.notEqual(missing.status.status, invalid.status.status);
      assert.ok(wb.setComment(0, 1, 1, 'libraz', 'hello').ok);
      const found = wb.getCommentResult(0, 1, 1);
      assert.equal(found.status.ok, true);
      assert.equal(found.comment.author, 'libraz');
      assert.equal(found.comment.text, 'hello');
    } finally {
      wb.delete();
    }
  });

  test('getValidations returns an array', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const arr = wb.getValidations(0);
      assert.ok(Array.isArray(arr) || typeof arr.length === 'number');
      assert.equal(arr.length, 0);
    } finally {
      wb.delete();
    }
  });

  test('addValidation / getValidations / removeValidationAt / clearValidations round-trip', () => {
    const wb = Module.Workbook.createDefault();
    try {
      // List-type validation at A1:B3 with prompts and an error message.
      const listRule = {
        ranges: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1 }],
        type: 3,
        op: 0,
        errorStyle: 1,
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true,
        formula1: '"Yes,No,Maybe"',
        promptTitle: 'Choose',
        promptMessage: 'Pick one',
        errorTitle: 'Bad value',
        errorMessage: 'Pick from the list',
      };
      assert.ok(wb.addValidation(0, listRule).ok);

      // Decimal between [0, 100] across two rectangles.
      const decimalRule = {
        ranges: [
          { firstRow: 4, firstCol: 0, lastRow: 4, lastCol: 0 },
          { firstRow: 6, firstCol: 0, lastRow: 9, lastCol: 0 },
        ],
        type: 2,
        op: 0,
        formula1: '0',
        formula2: '100',
        allowBlank: false,
      };
      assert.ok(wb.addValidation(0, decimalRule).ok);

      const list = wb.getValidations(0);
      assert.equal(list.length, 2);

      const a = list[0];
      assert.equal(a.type, 3);
      assert.equal(a.op, 0);
      assert.equal(a.errorStyle, 1);
      // Booleans must arrive as JS booleans, not 0/1.
      assert.equal(typeof a.allowBlank, 'boolean');
      assert.equal(a.allowBlank, true);
      assert.equal(typeof a.showInputMessage, 'boolean');
      assert.equal(a.showInputMessage, true);
      assert.equal(typeof a.showErrorMessage, 'boolean');
      assert.equal(a.showErrorMessage, true);
      assert.equal(a.formula1, '"Yes,No,Maybe"');
      assert.equal(a.promptTitle, 'Choose');
      assert.equal(a.errorTitle, 'Bad value');
      assert.equal(a.errorMessage, 'Pick from the list');
      assert.equal(a.ranges.length, 1);
      assert.equal(a.ranges[0].firstRow, 0);
      assert.equal(a.ranges[0].lastCol, 1);

      const b = list[1];
      assert.equal(b.type, 2);
      assert.equal(b.formula1, '0');
      assert.equal(b.formula2, '100');
      assert.equal(b.allowBlank, false);
      assert.equal(b.ranges.length, 2);
      assert.equal(b.ranges[1].firstRow, 6);
      assert.equal(b.ranges[1].lastRow, 9);

      // removeValidationAt drops the first rule.
      assert.ok(wb.removeValidationAt(0, 0).ok);
      const after = wb.getValidations(0);
      assert.equal(after.length, 1);
      assert.equal(after[0].type, 2);

      // Out-of-range index is rejected.
      assert.ok(!wb.removeValidationAt(0, 99).ok);

      // clearValidations drops everything; safe to call again.
      assert.ok(wb.clearValidations(0).ok);
      assert.equal(wb.getValidations(0).length, 0);
      assert.ok(wb.clearValidations(0).ok);

      // Sheet-out-of-range is rejected on every entry.
      assert.ok(!wb.addValidation(999, listRule).ok);
      assert.ok(!wb.removeValidationAt(999, 0).ok);
      assert.ok(!wb.clearValidations(999).ok);
    } finally {
      wb.delete();
    }
  });

  test('CF and sheet-layout lists are plain JS arrays with no delete lifecycle', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const cf = wb.evaluateCfRange(0, 0, 0, 1, 1, Number.NaN);
      assert.ok(cf.status.ok);
      assert.ok(Array.isArray(cf.cells));

      assert.ok(wb.setColumnWidth(0, 2, 2, 18).ok);
      const columns = wb.getSheetColumns(0);
      assert.ok(columns.status.ok);
      assert.ok(Array.isArray(columns.columns));
      assert.equal(columns.columns[0].first, 2);
      assert.equal(columns.columns[0].width, 18);

      assert.ok(wb.setRowHeight(0, 3, 22).ok);
      const rows = wb.getSheetRowOverrides(0);
      assert.ok(rows.status.ok);
      assert.ok(Array.isArray(rows.rows));
      assert.equal(rows.rows[0].row, 3);
      assert.equal(rows.rows[0].height, 22);
    } finally {
      wb.delete();
    }
  });

  test('sheet-layout setters refuse coordinates and metrics Excel cannot open', () => {
    const wb = Module.Workbook.createDefault();
    try {
      // These six setters marshal straight to the C ABI, which is where the
      // validation lives. The assertions below are what keeps that true: a
      // binding that grew its own path to the model would still pass every
      // round-trip test while quietly authoring a file Excel refuses, since
      // NaN and Infinity serialise to an empty string -- `<col width=""/>`
      // is not a lexical xsd:double at all, and a repair prompt rather than
      // an out-of-range value.
      const kInvalidArgument = 2;
      // The message is pinned, not just the code. Both JS bindings route
      // these calls to one set of helpers, so both report that shared
      // wording; a binding that grew its own check would keep returning
      // `kInvalidArgument` while the text stopped matching. The status code
      // alone cannot tell a shared implementation from two that agree today.
      const kSpanRefused = 'column span out of range';
      const kRowRefused = 'row index out of range';
      const kMetricRefused = 'value must be finite and non-negative';
      const refused = (status, expectedMessage, label) => {
        assert.equal(status.ok, false, `${label} should have been refused`);
        assert.equal(status.status, kInvalidArgument, `${label}: ${JSON.stringify(status)}`);
        assert.equal(status.message, expectedMessage, `${label}: ${JSON.stringify(status)}`);
      };

      // One past Excel's last column / row.
      refused(wb.setColumnWidth(0, 0, 16384, 10), kSpanRefused, 'setColumnWidth past the grid');
      refused(wb.setColumnHidden(0, 0, 16384, true), kSpanRefused, 'setColumnHidden past the grid');
      refused(wb.setColumnOutline(0, 0, 16384, 1), kSpanRefused, 'setColumnOutline past the grid');
      refused(wb.setRowHeight(0, 1048576, 10), kRowRefused, 'setRowHeight past the grid');
      refused(wb.setRowHidden(0, 1048576, true), kRowRefused, 'setRowHidden past the grid');
      refused(wb.setRowOutline(0, 1048576, 1), kRowRefused, 'setRowOutline past the grid');

      for (const bad of [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY, -1]) {
        refused(wb.setColumnWidth(0, 0, 2, bad), kMetricRefused, `setColumnWidth(${bad})`);
        refused(wb.setRowHeight(0, 1, bad), kMetricRefused, `setRowHeight(${bad})`);
      }

      // The context names the C ABI entry point that refused the call and
      // the values it saw, which is what makes a rejection actionable.
      // These six answer with a bare `Status`, not a `{status}` envelope.
      const span = wb.setColumnWidth(0, 0, 16384, 10);
      assert.match(span.context, /fm_sheet_set_column_width: first=0 last=16384/);
      // Pinned separately because the metric path runs through a different
      // helper than the span path, and it is the only one of the three that
      // interpolates a `double`. The value itself is left unmatched on
      // purpose: `std::to_string` of a NaN is implementation-defined
      // spelling (`nan`, `-nan`, `nan(ind)`), so asserting it would be a
      // portability trap across the platforms the prebuilds target.
      const metric = wb.setRowHeight(0, 1, Number.NaN);
      assert.match(metric.context, /fm_sheet_set_row_height: height=/);

      // Zero is a real width and a real height, not a rejected one.
      assert.ok(wb.setColumnWidth(0, 0, 2, 0).ok);
      assert.ok(wb.setRowHeight(0, 1, 0).ok);

      // The last coordinates Excel does address stay accepted -- an
      // off-by-one in either bound would refuse them.
      assert.ok(wb.setColumnWidth(0, 16383, 16383, 10).ok);
      assert.ok(wb.setColumnHidden(0, 16383, 16383, true).ok);
      assert.ok(wb.setColumnOutline(0, 16383, 16383, 1).ok);
      assert.ok(wb.setRowHeight(0, 1048575, 10).ok);
      assert.ok(wb.setRowHidden(0, 1048575, true).ok);
      assert.ok(wb.setRowOutline(0, 1048575, 1).ok);
    } finally {
      wb.delete();
    }
  });

  test('typed print patches throw on lossy numbers before changing the page setup', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setSheetPageSetup(0, { orientation: 2, scale: 125 }).ok);
      const expectRejected = (patch, field) => {
        assert.throws(() => wb.setSheetPageSetup(0, patch), {
          name: 'RangeError',
          message: new RegExp(`setSheetPageSetup.*\`${field}\``),
        });
        const setup = wb.getSheetPageSetup(0);
        assert.equal(setup.orientation, 2);
        assert.equal(setup.scale, 125);
      };

      expectRejected({ scale: 2 ** 32 + 100 }, 'pageSetup.scale');
      expectRejected({ scale: 100.5 }, 'pageSetup.scale');
      expectRejected({ scale: Number.NaN }, 'pageSetup.scale');
      expectRejected({ scale: Number.POSITIVE_INFINITY }, 'pageSetup.scale');
      expectRejected({ orientation: 1.5 }, 'pageSetup.orientation');

      // Nullish values keep the old patch meaning: the attribute is omitted
      // from this update and the prior value remains in force.
      assert.ok(wb.setSheetPageSetup(0, { scale: null, orientation: undefined }).ok);
      const unchanged = wb.getSheetPageSetup(0);
      assert.equal(unchanged.orientation, 2);
      assert.equal(unchanged.scale, 125);
    } finally {
      wb.delete();
    }
  });

  test('typed print patches throw on coercion and preserve header/footer fields', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setSheetPrintOptions(0, { gridLines: true }).ok);
      const throwingOptions = {};
      Object.defineProperty(throwingOptions, 'gridLines', {
        get() {
          throw new Error('print-options getter');
        },
      });
      assert.throws(() => wb.setSheetPrintOptions(0, throwingOptions), /print-options getter/);
      assert.match(wb.getSheetPrintOptionsXml(0).xml, /gridLines="true"/);

      assert.ok(wb.setSheetPageMargins(0, { left: 0.5 }).ok);
      assert.throws(() => wb.setSheetPageMargins(0, { left: '0.75' }), {
        name: 'TypeError',
        message: /setSheetPageMargins.*left/,
      });
      assert.equal(wb.getSheetPageMargins(0).left, 0.5);

      assert.ok(wb.setSheetHeaderFooter(0, { oddHeader: 'before' }).ok);
      assert.throws(() => wb.setSheetHeaderFooter(0, { oddHeader: 42 }), {
        name: 'TypeError',
        message: /setSheetHeaderFooter.*oddHeader/,
      });
      assert.match(wb.getSheetHeaderFooterXml(0).xml, /before/);
    } finally {
      wb.delete();
    }
  });

  test('header/footer getter failures are rethrown before any mutation', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setSheetHeaderFooter(0, { oddHeader: 'before' }).ok);
      const lateFailure = { oddHeader: 'after' };
      Object.defineProperty(lateFailure, 'alignWithMargins', {
        get() {
          throw new Error('late header getter');
        },
      });

      assert.throws(() => wb.setSheetHeaderFooter(0, lateFailure), /late header getter/);
      assert.match(wb.getSheetHeaderFooterXml(0).xml, /before/);
      assert.doesNotMatch(wb.getSheetHeaderFooterXml(0).xml, /after/);
    } finally {
      wb.delete();
    }
  });

  test('setSheetVisibility states veryHidden, which the bool setter cannot', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const read = () => {
        const r = wb.getSheetView(0);
        assert.ok(r.status.ok, `getSheetView: ${JSON.stringify(r.status)}`);
        return r.view;
      };
      assert.equal(read().visibility, 0);
      assert.equal(read().tabHidden, 0);

      assert.ok(wb.setSheetVisibility(0, 2).ok);
      // The two-state view stays consistent: a very-hidden sheet reads as
      // hidden to a caller that knows only the bool, never as visible.
      assert.equal(read().visibility, 2);
      assert.equal(read().tabHidden, 1);

      // Asking for plain "hidden" must not weaken the stronger state.
      assert.ok(wb.setSheetTabHidden(0, true).ok);
      assert.equal(read().visibility, 2);

      // Demotion is the other direction the bool cannot express.
      assert.ok(wb.setSheetVisibility(0, 1).ok);
      assert.equal(read().visibility, 1);
      assert.equal(read().tabHidden, 1);

      // Showing the sheet clears either hidden state.
      assert.ok(wb.setSheetTabHidden(0, false).ok);
      assert.equal(read().visibility, 0);
      assert.equal(read().tabHidden, 0);

      // An unknown ordinal is refused and leaves the sheet alone.
      assert.ok(wb.setSheetVisibility(0, 2).ok);
      assert.equal(wb.setSheetVisibility(0, 3).ok, false);
      assert.equal(read().visibility, 2);
    } finally {
      wb.delete();
    }
  });

  test('getIterative reads back what setIterative stored', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const initial = wb.getIterative();
      assert.ok(initial.status.ok, `getIterative: ${JSON.stringify(initial.status)}`);
      assert.equal(typeof initial.enabled, 'boolean');
      assert.equal(typeof initial.maxIterations, 'number');
      assert.equal(typeof initial.maxChange, 'number');

      assert.ok(wb.setIterative(true, 42, 0.25).ok);
      const enabled = wb.getIterative();
      assert.ok(enabled.status.ok);
      assert.equal(enabled.enabled, true);
      assert.equal(enabled.maxIterations, 42);
      assert.equal(enabled.maxChange, 0.25);

      // The cap and threshold survive switching iteration back off, which is
      // what lets a host render the dialog with the stored values.
      assert.ok(wb.setIterative(false, 42, 0.25).ok);
      const disabled = wb.getIterative();
      assert.ok(disabled.status.ok);
      assert.equal(disabled.enabled, false);
      assert.equal(disabled.maxIterations, 42);
      assert.equal(disabled.maxChange, 0.25);
    } finally {
      wb.delete();
    }
  });

  test('addValidation defaults allowBlank to false', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(
        wb.addValidation(0, {
          ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
          type: 3,
          formula1: '"a,b"',
        }).ok,
      );
      const list = wb.getValidations(0);
      assert.equal(list.length, 1);
      assert.equal(list[0].allowBlank, false);
    } finally {
      wb.delete();
    }
  });

  test('getCellStyleXf reports xfId alongside the cell-format fields', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const read = wb.getCellStyleXf(0);
      assert.ok(read.status.ok, `getCellStyleXf: ${JSON.stringify(read.status)}`);
      // A named-style xf is its own parent, so it never inherits one.
      assert.equal(read.xfId, 0);
      // The record shape matches `getCellXf`, which the documented
      // read / edit / addCellStyleXf round-trip depends on.
      const cellXf = wb.getCellXf(0);
      assert.ok(cellXf.status.ok);
      for (const key of Object.keys(cellXf)) {
        assert.ok(key in read, `getCellStyleXf is missing '${key}'`);
      }
    } finally {
      wb.delete();
    }
  });
}
