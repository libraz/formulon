import assert from 'node:assert/strict';
import { StatusCode } from '../../packages/npm/common.mjs';
import { expectFailure } from './smoke_support.mjs';

const rect = (firstRow, firstCol, lastRow, lastCol) => ({ firstRow, firstCol, lastRow, lastCol });

function withWorkbook(Module, fn) {
  const wb = Module.Workbook.createDefault();
  try {
    fn(wb);
  } finally {
    wb.delete();
  }
}

export function registerSheetScope(Module, test) {
  test('tables are addressed by sheet and a sheet-local index', () => {
    withWorkbook(Module, (wb) => {
      wb.addSheet('Second');
      const first = wb.addTable(0, { range: rect(0, 0, 2, 1), name: 'Sales', columns: ['Product', 'Amount'] });
      const second = wb.addTable(1, { range: rect(0, 0, 2, 1), name: 'Costs', columns: ['Item', 'Price'] });
      assert.equal(first, 0);
      assert.equal(second, 0);
      assert.equal(wb.getTableCount(0), 1);
      assert.equal(wb.getTableCount(1), 1);

      assert.equal(wb.getTable(1, 0).name, 'Costs');
      assert.equal(wb.getTable(1, 0).range.sheet, 1);
      assert.deepEqual(
        wb.getTable(0, 0).columns.map((c) => c.name),
        ['Product', 'Amount'],
      );

      // The workbook-wide lookups return the sheet and the sheet-local index.
      assert.deepEqual({ ...wb.findTable('costs') }, { sheet: 1, index: 0 });
      assert.equal(wb.findTable('Nothing'), null);
      const id = wb.getTableRuntimeId(1, 0);
      assert.deepEqual({ ...wb.resolveTableRuntimeId(id) }, { sheet: 1, index: 0 });

      // One table name; a taken name is refused.
      expectFailure(
        () => wb.addTable(0, { range: rect(5, 0, 7, 1), name: 'COSTS', columns: ['A', 'B'] }),
        StatusCode.AlreadyExists,
      );
      expectFailure(() => wb.getTable(9, 0), StatusCode.SheetIndexOutOfRange);
      expectFailure(() => wb.getTable(0, 5), StatusCode.IndexOutOfRange);
      expectFailure(() => wb.getTableCount(9), StatusCode.SheetIndexOutOfRange);

      wb.renameTable(0, 0, 'Revenue');
      assert.equal(wb.getTable(0, 0).name, 'Revenue');
      assert.deepEqual({ ...wb.findTable('revenue') }, { sheet: 0, index: 0 });
      assert.equal(wb.findTable('Sales'), null);

      wb.removeTable(0, 0);
      assert.equal(wb.getTableCount(0), 0);
      assert.equal(wb.getTableCount(1), 1);
      expectFailure(() => wb.resolveTableRuntimeId(wb.getTableRuntimeId(1, 0) + 1000), StatusCode.NotFound);
    });
  });

  test('a table takes the style default, and its style info is read and patched apart from the table', () => {
    withWorkbook(Module, (wb) => {
      const styled = wb.addTable(0, { range: rect(0, 0, 2, 1), name: 'Styled', columns: ['A', 'B'] });
      assert.deepEqual(
        { ...wb.getTableStyleOptions(0, styled) },
        {
          styleName: 'TableStyleMedium2',
          showFirstColumn: false,
          showLastColumn: false,
          showRowStripes: true,
          showColumnStripes: false,
        },
      );
      assert.equal('styleName' in wb.getTable(0, styled), false);

      wb.updateTableStyleOptions(0, styled, { styleName: 'TableStyleLight9', showColumnStripes: true });
      const info = wb.getTableStyleOptions(0, styled);
      assert.equal(info.styleName, 'TableStyleLight9');
      assert.equal(info.showColumnStripes, true);
      assert.equal(info.showRowStripes, true);

      // `null` removes the style name and keeps the element and the toggles.
      wb.updateTableStyleOptions(0, styled, { styleName: null });
      assert.equal(wb.getTableStyleOptions(0, styled).styleName, null);
      assert.equal(wb.getTableStyleOptions(0, styled).showColumnStripes, true);

      // An empty style name creates the table without a style element.
      const bare = wb.addTable(0, { range: rect(5, 0, 7, 1), name: 'Bare', columns: ['C', 'D'], styleName: '' });
      assert.equal(wb.getTableStyleOptions(0, bare), null);

      // The workbook default applies when the name is omitted.
      wb.updateTableStyleDefaults({ tableStyle: 'TableStyleLight1' });
      const defaulted = wb.addTable(0, { range: rect(10, 0, 12, 1), name: 'Defaulted', columns: ['E', 'F'] });
      assert.equal(wb.getTableStyleOptions(0, defaulted).styleName, 'TableStyleLight1');
    });
  });

  test('sheet protection: null removes, {} protects with the dialog defaults, update merges', () => {
    withWorkbook(Module, (wb) => {
      assert.equal(wb.getSheetProtection(0), null);

      wb.setSheetProtection(0, {});
      const defaults = wb.getSheetProtection(0);
      assert.equal(defaults.allowSelectLockedCells, true);
      assert.equal(defaults.allowSelectUnlockedCells, true);
      assert.equal(defaults.allowSort, false);
      assert.equal(defaults.allowFormatCells, false);
      assert.equal(defaults.legacyPasswordHash, null);
      assert.equal(defaults.spinCount, 0);

      wb.setSheetProtection(0, null);
      assert.equal(wb.getSheetProtection(0), null);
      // Unprotecting an unprotected sheet does nothing.
      wb.setSheetProtection(0, null);
      assert.equal(wb.getSheetProtection(0), null);

      // A patch on an unprotected sheet protects it, unless it changes nothing.
      wb.updateSheetProtection(0, {});
      assert.equal(wb.getSheetProtection(0), null);
      wb.updateSheetProtection(0, { allowSort: true, legacyPasswordHash: 'CC1A', allowSelectLockedCells: false });
      const patched = wb.getSheetProtection(0);
      assert.equal(patched.allowSort, true);
      assert.equal(patched.allowSelectLockedCells, false);
      assert.equal(patched.allowSelectUnlockedCells, true);
      assert.equal(patched.legacyPasswordHash, 'CC1A');

      wb.updateSheetProtection(0, { legacyPasswordHash: null, spinCount: 10 });
      assert.equal(wb.getSheetProtection(0).legacyPasswordHash, null);
      assert.equal(wb.getSheetProtection(0).spinCount, 10);
      assert.equal(wb.getSheetProtection(0).allowSort, true);

      // A flag has no cleared form.
      assert.throws(() => wb.updateSheetProtection(0, { allowSort: null }), TypeError);
      assert.equal(wb.getSheetProtection(0).allowSort, true);

      // The record read back is accepted as is.
      const read = wb.getSheetProtection(0);
      wb.setSheetProtection(0, null);
      wb.setSheetProtection(0, read);
      assert.deepEqual({ ...wb.getSheetProtection(0) }, { ...read });
    });
  });

  test('row and column breaks upsert with a span and are removed with null', () => {
    withWorkbook(Module, (wb) => {
      wb.setRowBreak(0, 5, {});
      const whole = wb.listRowBreaks(0);
      assert.equal(whole.length, 1);
      assert.equal(whole[0].row, 5);
      assert.equal(whole[0].first, 0);
      assert.ok(whole[0].last > 0, 'a whole-sheet break stores the full span');
      assert.equal(whole[0].manual, true);

      // The same row replaces the break rather than adding a second.
      wb.setRowBreak(0, 5, { first: 2, last: 4, manual: false });
      assert.deepEqual(
        wb.listRowBreaks(0).map((b) => ({ ...b })),
        [{ row: 5, first: 2, last: 4, manual: false }],
      );

      wb.setColumnBreak(0, 3, { manual: true });
      const column = wb.listColumnBreaks(0);
      assert.equal(column.length, 1);
      assert.equal(column[0].col, 3);
      assert.equal(column[0].manual, true);

      wb.setRowBreak(0, 5, null);
      wb.setColumnBreak(0, 3, null);
      assert.equal(wb.listRowBreaks(0).length, 0);
      assert.equal(wb.listColumnBreaks(0).length, 0);
      // Removing a break that is not there does nothing.
      wb.setRowBreak(0, 5, null);
      wb.setColumnBreak(0, 3, null);
      expectFailure(() => wb.setRowBreak(9, 5, {}), StatusCode.SheetIndexOutOfRange);
    });
  });

  test('setHyperlink replaces the link anchored at the cell', () => {
    withWorkbook(Module, (wb) => {
      wb.setHyperlink(0, 1, 1, 'https://example.com/a', '', '', '');
      wb.setHyperlink(0, 1, 1, 'https://example.com/b', 'B', '', '');
      const links = wb.listHyperlinks(0);
      assert.equal(links.length, 1);
      assert.equal(links[0].target, 'https://example.com/b');
      assert.equal(links[0].display, 'B');

      wb.setHyperlinkRange(0, rect(1, 1, 2, 2), 'https://example.com/c', '', '', '');
      assert.equal(wb.listHyperlinks(0).length, 1);
      assert.deepEqual(wb.listHyperlinks(0)[0].rect, rect(1, 1, 2, 2));
    });
  });

  test('a data validation omitting its flags takes the dialog defaults', () => {
    withWorkbook(Module, (wb) => {
      wb.addValidation(0, { ranges: [rect(0, 0, 4, 0)], type: 3, formula1: '"a,b"' });
      const defaults = wb.getValidation(0, 0);
      assert.equal(defaults.allowBlank, true);
      assert.equal(defaults.showInputMessage, true);
      assert.equal(defaults.showErrorMessage, true);
      assert.equal(defaults.showInCellDropdown, true);

      wb.addValidation(0, {
        ranges: [rect(5, 0, 9, 0)],
        type: 3,
        formula1: '"a,b"',
        allowBlank: false,
        showInCellDropdown: false,
        showInputMessage: false,
        showErrorMessage: false,
      });
      const explicit = wb.getValidation(0, 1);
      assert.equal(explicit.allowBlank, false);
      assert.equal(explicit.showInCellDropdown, false);
      assert.equal(explicit.showInputMessage, false);
      assert.equal(explicit.showErrorMessage, false);
    });
  });

  test('setAutoFilter(null) removes a sheet filter and does nothing without one', () => {
    withWorkbook(Module, (wb) => {
      wb.setText(0, 0, 0, 'N');
      wb.setNumber(0, 1, 0, 1);
      wb.setAutoFilter(0, null);
      assert.equal(wb.getAutoFilter(0), null);
      wb.setAutoFilter(0, { range: rect(0, 0, 1, 0), columns: [], sort: null });
      assert.notEqual(wb.getAutoFilter(0), null);
      wb.setAutoFilter(0, null);
      assert.equal(wb.getAutoFilter(0), null);
    });
  });

  test('hasPassthrough reports a part by path', () => {
    withWorkbook(Module, (wb) => {
      assert.equal(wb.hasPassthrough('xl/doesNotExist.xml'), false);
    });
  });

  test('getDefaultFont reads font 0 and activeSheet reads the active index', () => {
    withWorkbook(Module, (wb) => {
      assert.deepEqual({ ...wb.getDefaultFont() }, { ...wb.getFont(0) });
      assert.equal(wb.activeSheet(), 0);
      wb.addSheet('Second');
      wb.setActiveSheet(1);
      assert.equal(wb.activeSheet(), 1);
    });
  });
}
