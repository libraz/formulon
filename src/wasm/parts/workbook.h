//
// `JsWorkbook` is the move-only RAII wrapper around `fm_workbook_t*`
// that backs the JS-facing `Workbook` class. Its method bodies are
// split across `parts/workbook_*.cpp` by API surface area so each TU
// stays manageable; this header is the single declaration the binding
// registrar (`parts/bindings_register.cpp`) keys off.
//
// All entry points reuse the same translation contract: every call returns
// its payload directly, and a failure is recorded through the builders in
// `parts/embind_common.h` while the method returns the type's default; the
// pre-js wrapper throws the recorded `FormulonError` once the native frame
// has returned. Static factories (`createDefault`, `createEmpty`,
// `createForProfile`, `load`) return a heap `JsWorkbook` the JS caller owns.
//
// The contract holds in the callback direction too: the JS the binding
// calls back into runs behind `call_js_callback`, so a callback that
// throws is rethrown by the outer call instead of unwinding C++ frames that
// have no landing pads.

#ifndef FORMULON_WASM_PARTS_WORKBOOK_H_
#define FORMULON_WASM_PARTS_WORKBOOK_H_

#include <emscripten/val.h>

#include <cstddef>
#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"

namespace formulon {
namespace wasm {
namespace parts {

class JsWorkbook {
 public:
  JsWorkbook() = default;
  ~JsWorkbook();

  JsWorkbook(const JsWorkbook&) = delete;
  JsWorkbook& operator=(const JsWorkbook&) = delete;

  JsWorkbook(JsWorkbook&& other) noexcept;
  JsWorkbook& operator=(JsWorkbook&& other) noexcept;

  // ---- Lifecycle / static factories ---------------------------------------

  static emscripten::val createDefault();
  static emscripten::val createEmpty();
  static emscripten::val createForProfile(const std::string& profileId);
  static emscripten::val loadBytes(emscripten::val bytes);

  /// True once `dispose` has released the native handle.
  bool disposed() const { return handle_ == nullptr; }
  /// Releases the native handle and keeps the wrapper; idempotent. Every later
  /// call fails with `kBindingInvalidHandle`.
  void dispose();

  /// Serialises the workbook as a `Uint8Array` in container `format`
  /// (`fm_workbook_format_t` ordinal: 1 = xlsx, 2 = xlsb); `undefined`
  /// means xlsx. An unknown value fails with `kInvalidArgument`.
  emscripten::val save(emscripten::val format) const;
  /// `save` plus the loss counters the write produced.
  JsSaveDiagnostics saveWithDiagnostics(emscripten::val format) const;
  JsReadDiagnostics readDiagnostics() const;

  // ---- Sheet management ---------------------------------------------------

  uint32_t addSheet(const std::string& name);
  void moveSheet(uint32_t fromIndex, uint32_t toIndex);
  uint32_t copySheet(uint32_t sourceIndex, uint32_t insertBefore, emscripten::val newName);
  void removeSheet(uint32_t index);
  void renameSheet(uint32_t index, const std::string& newName);

  /// Creates or wholly replaces the defined name `input` describes
  /// (`{ name, formula, localSheetId?, comment?, hidden? }`); an omitted or
  /// `null` `localSheetId` is workbook scope.
  void setDefinedName(emscripten::val input);

  void insertRows(uint32_t sheet, uint32_t row, uint32_t count);
  void deleteRows(uint32_t sheet, uint32_t row, uint32_t count);
  void insertColumns(uint32_t sheet, uint32_t col, uint32_t count);
  void deleteColumns(uint32_t sheet, uint32_t col, uint32_t count);

  double sheetCount() const;
  std::string sheetName(uint32_t idx) const;
  /// The index of the sheet named `name`, compared case-insensitively, or
  /// `null` when no sheet has that name.
  emscripten::val getSheetIndex(const std::string& name) const;
  /// The active sheet (`activeTab` of the first `<workbookView>`) as its
  /// index, and its setter.
  uint32_t activeSheet() const;
  void setActiveSheet(uint32_t sheet);

  // ---- Cell value / formula ops ------------------------------------------

  void setNumber(uint32_t sheet, uint32_t row, uint32_t col, double value);
  void setBool(uint32_t sheet, uint32_t row, uint32_t col, bool value);
  void setError(uint32_t sheet, uint32_t row, uint32_t col, emscripten::val errorCode);
  void setText(uint32_t sheet, uint32_t row, uint32_t col, const std::string& text);
  void setCellPhonetic(uint32_t sheet, uint32_t row, uint32_t col, const std::string& phonetic);
  /// Stores a cell's furigana as one `<rPh>` block per element of `runs`,
  /// each `{ sb, eb, text }`. Unlike `setCellPhonetic`, which annotates the
  /// whole cell, this preserves which characters each reading covers. See
  /// `fm_sheet_set_cell_phonetic_runs` for the ordering rules.
  void setCellPhoneticRuns(uint32_t sheet, uint32_t row, uint32_t col, emscripten::val runs);

  /// Stores how the cell's guide renders: `fontId`, `type` (0=halfwidth
  /// katakana, 1=fullwidth katakana, 2=hiragana, 3=no conversion) and
  /// `alignment` (0=no control, 1=left, 2=center, 3=distributed).
  /// Independent of the readings in both directions.
  void setCellPhoneticProperties(uint32_t sheet, uint32_t row, uint32_t col, emscripten::val properties);
  /// The cell's "Show Phonetic" display flag; value writes leave it alone.
  void setCellShowPhonetic(uint32_t sheet, uint32_t row, uint32_t col, bool show);
  void setBlank(uint32_t sheet, uint32_t row, uint32_t col);
  void setFormula(uint32_t sheet, uint32_t row, uint32_t col, const std::string& formula);

  /// The stored A1 formula, or `null` for a cell that holds none.
  emscripten::val getFormula(uint32_t sheet, uint32_t row, uint32_t col) const;
  /// R1C1 counterpart of `getFormula`, relative to the cell, without `=`.
  emscripten::val getFormulaR1C1(uint32_t sheet, uint32_t row, uint32_t col) const;
  /// One page of the populated cells inside `range`:
  /// `{ cells: [{ row, col, formula, value }], nextCursor }`.
  /// `cursor` / `limit` may be omitted or `null`; `nextCursor` is `null`
  /// once the range is exhausted.
  emscripten::val getCellsInRange(uint32_t sheet, emscripten::val range, emscripten::val cursor,
                                  emscripten::val limit) const;
  /// `{ text, displayStatus }`: the text Excel displays for a cell.
  emscripten::val getDisplayText(uint32_t sheet, uint32_t row, uint32_t col) const;
  /// Renders a `Value` under a stored-form number format code.
  emscripten::val formatValue(emscripten::val value, const std::string& formatCode) const;

  emscripten::val getValue(uint32_t sheet, uint32_t row, uint32_t col) const;
  /// The cell's furigana; `null` when it has none.
  emscripten::val getCellPhonetic(uint32_t sheet, uint32_t row, uint32_t col) const;
  /// Reads a cell's furigana as `{ runs: [{ sb, eb, text }] }`,
  /// spans included. `getCellPhonetic` returns the same readings
  /// concatenated, without the spans.
  double getCellPhoneticRunCount(uint32_t sheet, uint32_t row, uint32_t col) const;
  emscripten::val getCellPhoneticRun(uint32_t sheet, uint32_t row, uint32_t col, uint32_t phoneticRunIndex) const;

  /// Reads the guide's rendering back as `{ fontId, type, alignment }`.
  /// A cell with no annotation reports the all-zero triple.
  emscripten::val getCellPhoneticProperties(uint32_t sheet, uint32_t row, uint32_t col) const;
  bool getCellShowPhonetic(uint32_t sheet, uint32_t row, uint32_t col) const;
  emscripten::val getLambdaText(uint32_t sheet, uint32_t row, uint32_t col) const;

  /// Evaluates `formula` as if entered at `(sheet, row, col)` and returns a
  /// single scalar result without mutating the workbook. Array results are
  /// reduced to their top-left element. See `fm_sheet_evaluate_formula`.
  emscripten::val evaluateFormulaText(uint32_t sheet, uint32_t row, uint32_t col, const std::string& formula) const;

  /// Evaluates `formula` as if entered at `(sheet, row, col)` and returns the
  /// whole multi-cell result without mutating the workbook. Unlike
  /// `evaluateFormulaText` (which reduces an array to its top-left element),
  /// this yields the full grid as a JS object
  /// `{ rows, cols, cells }`, where `cells` is a `rows` x `cols`
  /// nested array of Value objects (row-major, `cells[r][c]`); a scalar
  /// result is reported as a 1x1 array. See
  /// `fm_sheet_evaluate_formula_array`.
  emscripten::val evaluateFormulaArray(uint32_t sheet, uint32_t row, uint32_t col, const std::string& formula) const;

  /// Evaluates `formula` as a conditional-formatting predicate anchored at
  /// `(sheet, row, col)`, with relative refs written relative to
  /// `(anchorRow, anchorCol)`; returns a coerced boolean result. See
  /// `fm_sheet_evaluate_cf_formula`.
  emscripten::val evaluateConditionalFormula(uint32_t sheet, uint32_t row, uint32_t col, uint32_t anchorRow,
                                             uint32_t anchorCol, const std::string& formula) const;

  // ---- Recalc / calc mode ------------------------------------------------

  void recalc();
  /// Takes the raw JS argument so the binding can reject non-numeric,
  /// fractional, non-finite, and out-of-range values before embind's normal
  /// integer coercion would silently change them.
  JsParallelRecalcStats recalcParallel(emscripten::val threadCount);
  /// True when a change has not been applied by a full recalc yet. See
  /// `fm_workbook_has_dirty_cells`.
  bool hasDirtyCells() const;
  /// Applies `patch` (`{ enabled?, maxIterations?, maxChange? }`); an omitted
  /// key keeps the stored setting.
  void setIterative(emscripten::val patch);
  /// The stored iterative-calculation settings as `{enabled,
  /// maxIterations, maxChange}`. The cap and threshold are the stored
  /// values even while `enabled` is false, so a host can render Excel's
  /// iterative-calculation dialog without having written them first.
  emscripten::val getIterative() const;
  double calcMode() const;
  void setCalcMode(uint32_t mode);
  bool fullPrecision() const;
  void setFullPrecision(bool fullPrecision);
  bool calcOnSave() const;
  void setCalcOnSave(bool calcOnSave);
  double refMode() const;
  void setRefMode(uint32_t mode);
  bool date1904() const;
  void setDate1904(bool date1904);
  /// The pinned wall-clock reading as `{year, month, day, hour, minute,
  /// second}`, or `null` when the workbook follows the host clock. `null`
  /// rather than a flag beside the fields because JS already has a way to
  /// say "no reading", and every field combination is a legal date.
  emscripten::val pinnedNow() const;
  void setPinnedNow(uint32_t year, uint32_t month, uint32_t day, uint32_t hour, uint32_t minute, uint32_t second);
  void clearPinnedNow();
  std::string excelProfileId() const;
  void setExcelProfileId(const std::string& profile_id);
  /// The stored file name; `null` when none is set.
  emscripten::val fileName() const;
  void setFileName(const std::string& name);
  double partialRecalc(uint32_t sheet, emscripten::val viewport);
  /// Installs the JS callback the iterative solver invokes after each
  /// sweep, or clears it when `cb` is null / undefined. A callback that
  /// throws aborts the solve, and the recalc that ran it rethrows that
  /// exception.
  void setIterativeProgress(emscripten::val cb);

  // ---- Iteration / metadata accessors ------------------------------------

  double cellCount(uint32_t sheet) const;
  emscripten::val cellAt(uint32_t sheet, uint32_t idx) const;

  emscripten::val listDefinedNames() const;
  emscripten::val getDefinedName(const std::string& name, int32_t localSheetId) const;
  void removeDefinedName(const std::string& name, int32_t localSheetId);

  double tableCount() const;
  emscripten::val tableAt(uint32_t idx) const;
  uint32_t createTable(emscripten::val spec);
  void updateTable(uint32_t idx, emscripten::val spec);
  void removeTable(uint32_t idx);
  /// Sets column `column`'s totals-row function / label from
  /// `{ function?, label?, formula? }`.
  void setTableColumnTotals(uint32_t idx, uint32_t column, emscripten::val totals);
  /// Renames column `column` of table `idx`; `value` is the final name.
  std::string renameTableColumn(uint32_t idx, uint32_t column, const std::string& name);

  emscripten::val listPassthroughs() const;
  emscripten::val getPassthrough(const std::string& path) const;

  double pivotCount(uint32_t sheet) const;
  emscripten::val pivotLayout(uint32_t sheet, uint32_t pivotIndex) const;

  emscripten::val evaluateCfRange(uint32_t sheet, emscripten::val range, double todaySerial) const;

  /// Computes the printable page grid for one sheet. Returns
  /// `{ printArea, horizontalBreaks, verticalBreaks, pageCount }`.
  emscripten::val paginate(uint32_t sheet) const;

  // ---- Print settings -----------------------------------------------------
  //
  // Three surfaces over the same model, from most to least raw. A bad
  // sheet index fails the call, so it stays distinguishable from an
  // absent setting.

  /// Raw print-settings fragments; `null` when the sheet declares no such
  /// element. The setters take "" to remove the element.
  emscripten::val getSheetPageSetupXml(uint32_t sheet) const;
  void setSheetPageSetupXml(uint32_t sheet, const std::string& xml);
  emscripten::val getSheetPageMarginsXml(uint32_t sheet) const;
  void setSheetPageMarginsXml(uint32_t sheet, const std::string& xml);
  emscripten::val getSheetPrintOptionsXml(uint32_t sheet) const;
  void setSheetPrintOptionsXml(uint32_t sheet, const std::string& xml);
  emscripten::val getSheetHeaderFooterXml(uint32_t sheet) const;
  void setSheetHeaderFooterXml(uint32_t sheet, const std::string& xml);
  emscripten::val getSheetSheetPrXml(uint32_t sheet) const;
  void setSheetSheetPrXml(uint32_t sheet, const std::string& xml);

  /// Toggles `<sheetPr><pageSetUpPr fitToPage>` without touching anything
  /// else `<sheetPr>` carries. Pair with `fitToWidth` / `fitToHeight` on
  /// the page setup to state the target.
  void setSheetFitToPage(uint32_t sheet, bool enabled);

  /// Print area as comma-separated A1 ranges. `{ ranges }`.
  emscripten::val getSheetPrintArea(uint32_t sheet) const;
  void setSheetPrintArea(uint32_t sheet, emscripten::val ranges);

  /// Repeat rows / columns as `{ rows, columns }`, each a `{ first, last }`
  /// span or `null`; `null` when neither axis repeats.
  emscripten::val getSheetPrintTitles(uint32_t sheet) const;
  /// Replaces the print titles with `titles` (same shape); `null` removes them.
  void setSheetPrintTitles(uint32_t sheet, emscripten::val titles);

  /// Manual page breaks. The enumerators return the whole array rather
  /// than a count + getter pair, matching the rest of this surface.
  void addSheetRowBreak(uint32_t sheet, uint32_t row, bool manual);
  void addSheetColBreak(uint32_t sheet, uint32_t col, bool manual);
  void removeSheetRowBreak(uint32_t sheet, uint32_t row);
  void removeSheetColBreak(uint32_t sheet, uint32_t col);
  void clearSheetBreaks(uint32_t sheet);
  emscripten::val getSheetRowBreaks(uint32_t sheet) const;
  emscripten::val getSheetColBreaks(uint32_t sheet) const;

  /// Typed patch setters. Only the keys present on the input object are
  /// applied; every other attribute in the stored XML is left alone.
  void setSheetPageSetup(uint32_t sheet, emscripten::val setup);
  void setSheetPageMargins(uint32_t sheet, emscripten::val margins);
  void setSheetPrintOptions(uint32_t sheet, emscripten::val options);
  void setSheetHeaderFooter(uint32_t sheet, emscripten::val headerFooter);
  emscripten::val getSheetPageSetup(uint32_t sheet) const;
  emscripten::val getSheetPageMargins(uint32_t sheet) const;

  // ---- Sheet view / layout ------------------------------------------------

  JsSheetView getSheetView(uint32_t sheet) const;
  JsSheetProtection getSheetProtection(uint32_t sheet) const;
  void setSheetProtection(uint32_t sheet, emscripten::val in);
  void setSheetZoom(uint32_t sheet, uint32_t zoomScale);
  void setSheetFreeze(uint32_t sheet, uint32_t freezeRows, uint32_t freezeCols);
  /// States the whole tab visibility, including `veryHidden`. Takes a raw
  /// signed ordinal so an unknown value is rejected by the C ABI rather than
  /// coerced by embind.
  void setSheetVisibility(uint32_t sheet, int32_t visibility);
  void setSheetShowGridLines(uint32_t sheet, bool show);
  void setSheetShowRowColHeaders(uint32_t sheet, bool show);
  void setSheetShowZeros(uint32_t sheet, bool show);
  void setSheetShowFormulas(uint32_t sheet, bool show);
  void setSheetShowOutlineSymbols(uint32_t sheet, bool show);
  void setSheetRightToLeft(uint32_t sheet, bool rightToLeft);
  void setSheetTabSelected(uint32_t sheet, bool selected);
  /// Scroll origin of the sheet view, and the first visible cell of a frozen
  /// view's unfrozen pane (which must lie past the frozen rows and columns).
  void setSheetTopLeft(uint32_t sheet, uint32_t row, uint32_t col);
  void setSheetPaneTopLeft(uint32_t sheet, uint32_t row, uint32_t col);
  void setSheetSplit(uint32_t sheet, double xTwips, double yTwips);
  /// The active pane's selection as `{ activeRow, activeCol, ranges }`,
  /// and its setter taking `{ activeRow, activeCol, ranges }`.
  emscripten::val getSheetSelection(uint32_t sheet) const;
  void setSheetSelection(uint32_t sheet, emscripten::val selection);
  /// `mode` is an `fm_sheet_view_mode_t` ordinal; the C ABI rejects others.
  void setSheetViewMode(uint32_t sheet, int32_t mode);

  emscripten::val getSheetColumns(uint32_t sheet) const;
  void setColumnWidth(uint32_t sheet, uint32_t first, uint32_t last, double width);
  void setColumnHidden(uint32_t sheet, uint32_t first, uint32_t last, bool hidden);
  void setColumnOutline(uint32_t sheet, uint32_t first, uint32_t last, uint32_t level);
  void setColumnCollapsed(uint32_t sheet, uint32_t first, uint32_t last, bool collapsed);
  void setColumnStyle(uint32_t sheet, uint32_t first, uint32_t last, uint32_t xfIndex);
  void clearColumnStyle(uint32_t sheet, uint32_t first, uint32_t last);

  emscripten::val getSheetRowOverrides(uint32_t sheet) const;
  void setRowHeight(uint32_t sheet, uint32_t first, uint32_t last, double height);
  void setRowAutoHeight(uint32_t sheet, uint32_t first, uint32_t last, double height);
  void setRowHidden(uint32_t sheet, uint32_t first, uint32_t last, bool hidden);
  void setRowOutline(uint32_t sheet, uint32_t first, uint32_t last, uint32_t level);
  void setRowCollapsed(uint32_t sheet, uint32_t first, uint32_t last, bool collapsed);
  void setRowStyle(uint32_t sheet, uint32_t first, uint32_t last, uint32_t xfIndex);
  void clearRowStyle(uint32_t sheet, uint32_t first, uint32_t last);
  emscripten::val rangeStyleBases(uint32_t sheet, emscripten::val range) const;
  emscripten::val rangeStyleBasesExcluding(uint32_t sheet, emscripten::val range, emscripten::val excludedRows,
                                           emscripten::val excludedCols) const;
  void remapRangeStyle(uint32_t sheet, emscripten::val range, emscripten::val fromXfs, emscripten::val toXfs);
  emscripten::val snapshotRangeStyle(uint32_t sheet, emscripten::val range) const;
  void restoreRangeStyle(uint32_t sheet, emscripten::val bytes);
  /// Drops the height override at `row`, returning it to the sheet default.
  void clearRowHeight(uint32_t sheet, uint32_t first, uint32_t last);
  void clearColumnWidth(uint32_t sheet, uint32_t first, uint32_t last);

  /// Sheet default column / row metrics (`<sheetFormatPr>`).
  emscripten::val getSheetFormatDefaults(uint32_t sheet) const;
  void setSheetFormatDefaults(uint32_t sheet, emscripten::val defaults);

  // ---- Point geometry (`mode`: 0 = display, 1 = print) --------------------

  /// `{ x, y, width, height }` of an inclusive cell range, in points
  /// from the top-left of A1.
  emscripten::val getCellRectPt(uint32_t sheet, emscripten::val range, int32_t mode) const;
  double getColumnWidthPt(uint32_t sheet, uint32_t col, int32_t mode) const;
  double getRowHeightPt(uint32_t sheet, uint32_t row) const;
  emscripten::val getWidthModel(uint32_t sheet, int32_t mode) const;
  double columnCharsToPt(uint32_t sheet, int32_t mode, double chars) const;
  double columnPtToChars(uint32_t sheet, int32_t mode, double pt) const;

  /** Complete worksheet-level `<autoFilter>` fragment. `null` means no
   * filter; a bad sheet fails the call instead. */
  emscripten::val getSheetAutoFilterXml(uint32_t sheet) const;
  void setSheetAutoFilterXml(uint32_t sheet, const std::string& xml);

  // ---- Styles -------------------------------------------------------------

  double getCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col) const;
  void setCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col, uint32_t xf_index);
  void clearCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col);
  /// Applies one xf across an inclusive rectangle, materialising blank
  /// cells so an empty ruled box renders. Saves one ABI crossing per cell.
  void setRangeXfIndex(uint32_t sheet, emscripten::val range, uint32_t xf_index);
  emscripten::val getCellXf(uint32_t xf_index) const;
  emscripten::val getFont(uint32_t font_index) const;
  emscripten::val getFill(uint32_t fill_index) const;
  emscripten::val getBorder(uint32_t border_index) const;
  emscripten::val getNumFmt(uint32_t num_fmt_id) const;
  emscripten::val getDxf(uint32_t dxf_index) const;

  uint32_t addFont(emscripten::val record);
  /// Overwrites an existing font slot; every xf naming it restyles at once.
  /// See `fm_workbook_set_font`.
  void setFont(uint32_t font_index, emscripten::val record);
  /// Declares the workbook's default font -- font 0, the record an unstyled
  /// cell resolves to. See `fm_workbook_set_default_font`.
  void setDefaultFont(emscripten::val record);
  uint32_t addFill(emscripten::val record);
  uint32_t addBorder(emscripten::val record);
  uint32_t addNumFmt(const std::string& format_code);
  uint32_t addXf(emscripten::val record);
  uint32_t addDxf(emscripten::val record);

  double fontCount() const;
  double fillCount() const;
  double borderCount() const;
  double xfCount() const;
  double dxfCount() const;
  emscripten::val listCellStyles() const;
  double cellStyleXfCount() const;
  emscripten::val getCellStyle(const std::string& name) const;
  emscripten::val getCellStyleXf(uint32_t index) const;
  uint32_t addCellStyleXf(emscripten::val record);
  /// Adds or replaces a `<cellStyle>` from `{ name, xfId, builtinId, iLevel, hidden, customBuiltin }`.
  void setCellStyle(emscripten::val record);
  /// Removes the named style; cell formats that used it fall back to Normal.
  void removeCellStyle(const std::string& name);
  /// Custom table / PivotTable styles (`<tableStyles>`) and the default
  /// style names.
  emscripten::val listTableStyles() const;
  emscripten::val getTableStyle(const std::string& name) const;
  void setTableStyle(emscripten::val style);
  void removeTableStyle(const std::string& name);
  emscripten::val getTableStyleDefaults() const;
  void setTableStyleDefaults(emscripten::val defaults);

  // ---- Theme, colour resolution and effective style ----------------------

  /// `{ source, colors }`: the twelve `clrScheme` colours as ARGB.
  emscripten::val getThemeColors() const;
  /// The four theme typefaces `{ majorLatin, majorEastAsian, minorLatin, minorEastAsian }`.
  emscripten::val getThemeFonts() const;
  emscripten::val getThemeScriptFont(bool major, const std::string& script) const;
  void setThemeColors(emscripten::val colors);
  void setThemeFonts(emscripten::val fonts);
  void setThemeScriptFont(bool major, const std::string& script, const std::string& typeface);
  void resetTheme();
  /// `{ argb, resolution }` for a `ColorSpec` used in `context` (`ColorContext.*`).
  emscripten::val resolveColor(emscripten::val spec, int32_t context) const;
  /// `{ xfIndex, source, ..., font, borders, numFmtCode }`: the formatting a cell shows.
  emscripten::val getEffectiveStyle(uint32_t sheet, uint32_t row, uint32_t col) const;

  double getExternalLinkCount() const;
  emscripten::val getExternalLink(uint32_t externalLinkIndex) const;

  // ---- Sheet UI features (merges / hyperlinks / comments / validations) --

  void addMerge(uint32_t sheet, emscripten::val range);
  void removeMerge(uint32_t sheet, emscripten::val range);
  void removeMergeAt(uint32_t sheet, uint32_t index);
  void clearMerges(uint32_t sheet);

  emscripten::val getComment(uint32_t sheet, uint32_t row, uint32_t col) const;
  emscripten::val getComments(uint32_t sheet) const;
  void setComment(uint32_t sheet, uint32_t row, uint32_t col, const std::string& author, const std::string& text);
  /// Whether the note at the cell is shown without hovering, or `null` when there is no note.
  /// Removes the note at `(row, col)`; a cell without one fails with
  /// `NotFound`.
  void removeComment(uint32_t sheet, uint32_t row, uint32_t col);
  emscripten::val getCommentVisible(uint32_t sheet, uint32_t row, uint32_t col) const;
  void setCommentVisible(uint32_t sheet, uint32_t row, uint32_t col, bool visible);

  /// The workbook's `<workbookProtection>`.
  emscripten::val getWorkbookProtection() const;
  void setWorkbookProtection(emscripten::val protection);
  /// The sheet's `<protectedRange>` entries, in file order.
  double getProtectedRangeCount(uint32_t sheet) const;
  emscripten::val getProtectedRange(uint32_t sheet, uint32_t protectedRangeIndex) const;
  uint32_t addProtectedRange(uint32_t sheet, emscripten::val range);
  void updateProtectedRange(uint32_t sheet, uint32_t index, emscripten::val range);
  void removeProtectedRange(uint32_t sheet, uint32_t index);

  void addHyperlink(uint32_t sheet, uint32_t row, uint32_t col, const std::string& target, const std::string& display,
                    const std::string& tooltip, const std::string& location);
  void addHyperlinkRange(uint32_t sheet, emscripten::val range, const std::string& target, const std::string& display,
                         const std::string& tooltip, const std::string& location);
  void removeHyperlink(uint32_t sheet, uint32_t row, uint32_t col);
  void removeHyperlinkAt(uint32_t sheet, uint32_t index);
  void clearHyperlinks(uint32_t sheet);
  emscripten::val getHyperlinks(uint32_t sheet) const;
  emscripten::val getHyperlink(uint32_t sheet, uint32_t row, uint32_t col) const;
  emscripten::val getMerges(uint32_t sheet) const;
  /// The merges intersecting `range`, in storage order.
  emscripten::val getMergesInRange(uint32_t sheet, emscripten::val range) const;

  /// Typed AutoFilter model for a sheet or (workbook-wide table index) a table.
  emscripten::val getAutoFilter(uint32_t sheet) const;
  void setAutoFilter(uint32_t sheet, emscripten::val filter);
  void removeAutoFilter(uint32_t sheet);
  void applyAutoFilter(uint32_t sheet);
  void clearAutoFilter(uint32_t sheet);
  emscripten::val evaluateAutoFilter(uint32_t sheet) const;
  emscripten::val getTableAutoFilter(uint32_t table) const;
  void setTableAutoFilter(uint32_t table, emscripten::val filter);
  void removeTableAutoFilter(uint32_t table);
  void applyTableAutoFilter(uint32_t table);
  void clearTableAutoFilter(uint32_t table);
  emscripten::val evaluateTableAutoFilter(uint32_t table) const;

  emscripten::val validateValue(uint32_t sheet, uint32_t row, uint32_t col, emscripten::val value) const;
  emscripten::val listInvalidCells(uint32_t sheet, emscripten::val cursor, emscripten::val limit) const;

  emscripten::val getThreadedComments(uint32_t sheet) const;
  emscripten::val getThreadedComment(uint32_t sheet, const std::string& id) const;
  void addThreadedComment(uint32_t sheet, emscripten::val comment);
  void editThreadedComment(uint32_t sheet, const std::string& id, const std::string& text, emscripten::val mentions);
  void setThreadResolved(uint32_t sheet, const std::string& threadId, bool done);
  void removeThreadedComment(uint32_t sheet, const std::string& id);
  emscripten::val getPersons() const;
  void addPerson(emscripten::val person);
  void removePerson(const std::string& id);
  emscripten::val getPerson(const std::string& id) const;

  emscripten::val listDrawingObjects(uint32_t sheet) const;
  emscripten::val getDrawingObject(uint32_t sheet, uint32_t objectId) const;
  emscripten::val getImage(uint32_t sheet, uint32_t objectId) const;
  uint32_t insertImage(uint32_t sheet, emscripten::val bytes, emscripten::val input);
  void removeDrawing(uint32_t sheet, uint32_t objectId);
  void setDrawingAnchor(uint32_t sheet, uint32_t objectId, emscripten::val anchor);
  void setDrawingName(uint32_t sheet, uint32_t objectId, const std::string& name);
  void setDrawingZOrder(uint32_t sheet, uint32_t objectId, uint32_t zOrderIndex);
  emscripten::val snapshotDrawing(uint32_t sheet, uint32_t objectId) const;
  uint32_t restoreDrawing(uint32_t sheet, emscripten::val bytes, emscripten::val opts);

  emscripten::val listShapes(uint32_t sheet) const;
  uint32_t insertShape(uint32_t sheet, emscripten::val spec);
  void updateShape(uint32_t sheet, uint32_t objectId, emscripten::val spec);
  void removeShape(uint32_t sheet, uint32_t objectId);
  emscripten::val getShape(uint32_t sheet, uint32_t objectId) const;

  emscripten::val listSlicers(uint32_t sheet) const;
  emscripten::val getSlicer(uint32_t sheet, uint32_t objectId) const;
  uint32_t addSlicer(uint32_t sheet, uint32_t table, uint32_t column, emscripten::val spec);
  void updateSlicer(uint32_t sheet, uint32_t objectId, emscripten::val spec);
  void removeSlicer(uint32_t sheet, uint32_t objectId);
  void addDataTable(uint32_t sheet, emscripten::val spec);
  emscripten::val getDataTable(uint32_t sheet, uint32_t row, uint32_t col) const;
  void removeDataTable(uint32_t sheet, uint32_t row, uint32_t col);
  void recalcDataTables();

  double getSparklineGroupCount(uint32_t sheet) const;
  emscripten::val getSparklineGroup(uint32_t sheet, uint32_t sparklineGroupIndex) const;
  uint32_t addSparklineGroup(uint32_t sheet, emscripten::val spec);
  void updateSparklineGroup(uint32_t sheet, uint32_t index, emscripten::val spec);
  void setSparklineGroupSparklines(uint32_t sheet, uint32_t index, emscripten::val sparklines);
  void removeSparklineGroup(uint32_t sheet, uint32_t index);

  emscripten::val getChart(uint32_t sheet, uint32_t objectId) const;
  emscripten::val getChartData(uint32_t sheet, uint32_t objectId) const;
  emscripten::val previewChartData(uint32_t sheet, uint32_t objectId, emscripten::val overrides) const;
  uint32_t insertChart(uint32_t sheet, emscripten::val options);
  void setChartType(uint32_t sheet, uint32_t objectId, uint32_t plot, int32_t kind, emscripten::val grouping);
  void setChartTitle(uint32_t sheet, uint32_t objectId, emscripten::val title);
  void setChartLegend(uint32_t sheet, uint32_t objectId, int32_t position, bool overlay, emscripten::val hiddenSeries);
  void setChartAxis(uint32_t sheet, uint32_t objectId, uint32_t axisId, emscripten::val patch);
  void setChartDisplay(uint32_t sheet, uint32_t objectId, emscripten::val patch);
  void setChartSeries(uint32_t sheet, uint32_t objectId, uint32_t plot, uint32_t seriesIndex, emscripten::val patch);
  void setChartPoint(uint32_t sheet, uint32_t objectId, uint32_t plot, uint32_t seriesIndex, uint32_t pointIndex,
                     emscripten::val patch);
  void setChartPlot(uint32_t sheet, uint32_t objectId, uint32_t plot, emscripten::val patch);
  void setChartFilter(uint32_t sheet, uint32_t objectId, emscripten::val hiddenSeries,
                      emscripten::val hiddenCategories);
  uint32_t addChartSeries(uint32_t sheet, uint32_t objectId, uint32_t plot, emscripten::val patch);
  void removeChartSeries(uint32_t sheet, uint32_t objectId, uint32_t plot, uint32_t seriesIndex);
  void setChartSource(uint32_t sheet, uint32_t objectId, const std::string& source, int32_t direction,
                      emscripten::val headerRow, emscripten::val headerColumn);

  double getValidationCount(uint32_t sheet) const;
  emscripten::val getValidation(uint32_t sheet, uint32_t validationIndex) const;
  uint32_t addValidation(uint32_t sheet, emscripten::val v);
  void removeValidationAt(uint32_t sheet, uint32_t index);
  void clearValidations(uint32_t sheet);

  // ---- Conditional formats -----------------------------------------------

  double getCfRuleCount(uint32_t sheet) const;
  emscripten::val getCfRule(uint32_t sheet, uint32_t cfRuleIndex) const;
  uint32_t addConditionalFormat(uint32_t sheet, emscripten::val v);
  void removeConditionalFormatAt(uint32_t sheet, uint32_t index);
  void clearConditionalFormats(uint32_t sheet);

  // ---- Misc trace / metadata / spill -------------------------------------

  emscripten::val precedents(uint32_t sheet, uint32_t row, uint32_t col, uint32_t depth) const;
  emscripten::val dependents(uint32_t sheet, uint32_t row, uint32_t col, uint32_t depth) const;
  emscripten::val spillInfo(uint32_t sheet, uint32_t row, uint32_t col) const;
  emscripten::val parseReference(const std::string& text, emscripten::val options) const;
  std::string formatReference(emscripten::val range, emscripten::val options) const;
  emscripten::val getArrayFormula(uint32_t sheet, uint32_t row, uint32_t col) const;
  void setArrayFormula(uint32_t sheet, emscripten::val range, const std::string& formula);

  // ---- PivotCache mutation -----------------------------------------------

  emscripten::val listPivotCaches() const;
  uint32_t pivotCacheCreate(uint32_t requestedId);
  void pivotCacheRemove(uint32_t cacheId);
  /// `{ present, ref, sheet, name }`, the strings `null` when absent; `null`
  /// when the cache has no worksheet source.
  emscripten::val pivotCacheGetWorksheetSource(uint32_t cacheId) const;
  void pivotCacheSetWorksheetSource(uint32_t cacheId, emscripten::val source);

  double pivotCacheFieldCount(uint32_t cacheId) const;
  std::string pivotCacheFieldName(uint32_t cacheId, uint32_t fieldIndex) const;
  uint32_t pivotCacheFieldAdd(uint32_t cacheId, const std::string& name);
  void pivotCacheFieldClear(uint32_t cacheId);

  double pivotCacheFieldSharedItemCount(uint32_t cacheId, uint32_t fieldIndex) const;
  emscripten::val pivotCacheFieldSharedItemValue(uint32_t cacheId, uint32_t fieldIndex, uint32_t itemIndex) const;
  void pivotCacheFieldAddSharedItemNumber(uint32_t cacheId, uint32_t fieldIndex, double value);
  void pivotCacheFieldAddSharedItemDate(uint32_t cacheId, uint32_t fieldIndex, double serial);
  bool pivotCacheFieldSharedItemIsDate(uint32_t cacheId, uint32_t fieldIndex, uint32_t itemIndex) const;
  void pivotCacheFieldAddSharedItemText(uint32_t cacheId, uint32_t fieldIndex, const std::string& utf8);
  void pivotCacheFieldAddSharedItemBool(uint32_t cacheId, uint32_t fieldIndex, bool value);
  void pivotCacheFieldAddSharedItemBlank(uint32_t cacheId, uint32_t fieldIndex);
  void pivotCacheFieldAddSharedItemError(uint32_t cacheId, uint32_t fieldIndex, int32_t errorCode);
  void pivotCacheFieldClearSharedItems(uint32_t cacheId, uint32_t fieldIndex);

  double pivotCacheRecordCount(uint32_t cacheId) const;
  uint32_t pivotCacheRecordAdd(uint32_t cacheId);
  void pivotCacheRecordClear(uint32_t cacheId);
  void pivotCacheRecordSetNumber(uint32_t cacheId, uint32_t recordIndex, uint32_t fieldIndex, double value);
  void pivotCacheRecordSetDate(uint32_t cacheId, uint32_t recordIndex, uint32_t fieldIndex, double serial);
  void pivotCacheRecordSetText(uint32_t cacheId, uint32_t recordIndex, uint32_t fieldIndex, const std::string& utf8);
  void pivotCacheRecordSetBool(uint32_t cacheId, uint32_t recordIndex, uint32_t fieldIndex, bool value);
  void pivotCacheRecordSetBlank(uint32_t cacheId, uint32_t recordIndex, uint32_t fieldIndex);
  void pivotCacheRecordSetError(uint32_t cacheId, uint32_t recordIndex, uint32_t fieldIndex, int32_t errorCode);

  // ---- PivotTable mutation ------------------------------------------------

  uint32_t pivotCreate(uint32_t sheet, const std::string& name, uint32_t cacheId, uint32_t anchorRow,
                       uint32_t anchorCol);
  void pivotRemove(uint32_t sheet, uint32_t pivotIndex);
  void pivotSetName(uint32_t sheet, uint32_t pivotIndex, const std::string& name);
  void pivotSetAnchor(uint32_t sheet, uint32_t pivotIndex, emscripten::val rect);
  void pivotSetGrandTotals(uint32_t sheet, uint32_t pivotIndex, bool rowsEnabled, bool colsEnabled);
  void pivotSetStyleInfo(uint32_t sheet, uint32_t pivotIndex, emscripten::val info);
  double pivotGetLayout(uint32_t sheet, uint32_t pivotIndex) const;
  void pivotSetLayout(uint32_t sheet, uint32_t pivotIndex, uint32_t layout);
  /// Definition read-back; works for pivots created in-session and loaded
  /// from a package alike. `pivotCacheId` returns the cache id.
  uint32_t pivotCacheId(uint32_t sheet, uint32_t pivotIndex) const;
  std::string pivotName(uint32_t sheet, uint32_t pivotIndex) const;
  emscripten::val pivotGetGrandTotals(uint32_t sheet, uint32_t pivotIndex) const;
  emscripten::val pivotGetStyleInfo(uint32_t sheet, uint32_t pivotIndex) const;
  emscripten::val pivotRowFieldOrder(uint32_t sheet, uint32_t pivotIndex) const;
  emscripten::val pivotColFieldOrder(uint32_t sheet, uint32_t pivotIndex) const;

  double pivotFieldCount(uint32_t sheet, uint32_t pivotIndex) const;
  uint32_t pivotFieldAdd(uint32_t sheet, uint32_t pivotIndex, emscripten::val spec);
  emscripten::val pivotFieldAt(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex) const;
  double pivotFieldItemCount(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex) const;
  emscripten::val pivotFieldItemAt(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, uint32_t itemIndex) const;
  void pivotFieldClear(uint32_t sheet, uint32_t pivotIndex);
  void pivotFieldSetAxis(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, int32_t axis);
  void pivotFieldSetSort(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, uint32_t order,
                         const std::string& byField);
  void pivotFieldSetSubtotalTop(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, bool top);
  void pivotFieldSetDefaultSubtotal(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, bool on);
  void pivotFieldSetInsertBlankRow(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, bool on);
  void pivotFieldSetFillDownLabels(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, bool on);
  void pivotFieldSetShowDetail(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, bool on);
  void pivotFieldSetItemShowDetail(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, uint32_t itemIndex,
                                   bool on);
  void pivotFieldSetItemCaption(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, uint32_t itemIndex,
                                const std::string& caption);
  void pivotFieldSetCustomName(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, const std::string& name);
  void pivotFieldSetShowAll(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, bool on);
  emscripten::val pivotGetOptions(uint32_t sheet, uint32_t pivotIndex) const;
  /// Members the object omits keep the pivot's current value.
  void pivotSetOptions(uint32_t sheet, uint32_t pivotIndex, emscripten::val options);
  emscripten::val pivotCacheGetOptions(uint32_t cacheId) const;
  void pivotCacheSetOptions(uint32_t cacheId, emscripten::val options);
  uint32_t pivotCacheCalculatedFieldAdd(uint32_t cacheId, const std::string& name, const std::string& formula);
  void pivotCacheCalculatedFieldSetFormula(uint32_t cacheId, uint32_t fieldIndex, const std::string& formula);
  void pivotCacheCalculatedFieldRemove(uint32_t cacheId, uint32_t fieldIndex);
  /// The formula of a calculated cache field; `null` for a field that is not calculated.
  emscripten::val pivotCacheFieldFormula(uint32_t cacheId, uint32_t fieldIndex) const;
  uint32_t pivotCacheCalculatedItemAdd(uint32_t cacheId, uint32_t fieldIndex, const std::string& name,
                                       const std::string& formula);
  double pivotCacheCalculatedItemCount(uint32_t cacheId) const;
  emscripten::val pivotCacheCalculatedItemAt(uint32_t cacheId, uint32_t idx) const;
  void pivotCacheCalculatedItemSetFormula(uint32_t cacheId, uint32_t idx, const std::string& formula);
  void pivotCacheCalculatedItemMove(uint32_t cacheId, uint32_t idx, uint32_t newIndex);
  void pivotCacheCalculatedItemRemove(uint32_t cacheId, uint32_t idx);
  /// `excludeDataField` is a data field index, or SIZE_MAX for none.
  std::string pivotDataFieldDefaultName(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, int32_t aggregation,
                                        std::size_t excludeDataField) const;
  void pivotProject(uint32_t sheet, uint32_t pivotIndex);
  bool pivotProjectionPending(uint32_t sheet, uint32_t pivotIndex) const;
  uint32_t pivotShowDetails(uint32_t sheet, uint32_t pivotIndex, uint32_t row, uint32_t col);
  uint32_t pivotFieldAddItem(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, const std::string& name,
                             bool visible);
  /// Appends a manual-filter item addressed by its position in the bound
  /// cache field's shared items. The only form that can express the blank
  /// item, which has no label of its own and is matched by its binding.
  uint32_t pivotFieldAddItemAt(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, uint32_t cacheIndex,
                               bool visible);
  void pivotFieldClearItems(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex);
  void pivotFieldSetItemVisible(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, uint32_t itemIndex,
                                bool visible);
  void pivotFieldAddSubtotalFn(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, uint32_t agg);
  void pivotFieldClearSubtotalFns(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex);
  /// Groups a date field by the `PivotDatePeriod` bits of `periods`; NaN
  /// bounds take the field's own.
  void pivotFieldSetDateGroup(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, uint32_t periods,
                              uint32_t calendar, uint32_t intervalDays, double start, double end);
  void pivotFieldClearDateGroup(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex);
  void pivotFieldSetNumberFormat(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex,
                                 const std::string& formatCode);

  void pivotSetRowFieldOrder(uint32_t sheet, uint32_t pivotIndex, emscripten::val indices);
  void pivotSetColFieldOrder(uint32_t sheet, uint32_t pivotIndex, emscripten::val indices);

  double pivotDataFieldCount(uint32_t sheet, uint32_t pivotIndex) const;
  uint32_t pivotDataFieldAdd(uint32_t sheet, uint32_t pivotIndex, emscripten::val spec);
  emscripten::val pivotDataFieldAt(uint32_t sheet, uint32_t pivotIndex, uint32_t dataFieldIndex) const;
  void pivotDataFieldClear(uint32_t sheet, uint32_t pivotIndex);
  void pivotDataFieldSet(uint32_t sheet, uint32_t pivotIndex, uint32_t dataFieldIndex, emscripten::val spec);

  double pivotFilterCount(uint32_t sheet, uint32_t pivotIndex) const;
  /// Reads the active filter at `filterIndex`.
  ///
  /// The active-filter list is **session state**: an entry added through
  /// `pivotFilterAdd` affects evaluation only while this workbook handle
  /// lives and is not written by `save`. Conversely a `<filters>` block
  /// Excel wrote is preserved verbatim on save but does not appear here, so
  /// `pivotFilterCount` reports only what this session added. The filter
  /// surface that does persist is pivot field item visibility.
  emscripten::val pivotFilterAt(uint32_t sheet, uint32_t pivotIndex, uint32_t filterIndex) const;
  uint32_t pivotFilterAdd(uint32_t sheet, uint32_t pivotIndex, emscripten::val spec);
  void pivotFilterClear(uint32_t sheet, uint32_t pivotIndex);
  void pivotFilterRemoveAt(uint32_t sheet, uint32_t pivotIndex, uint32_t filterIndex);

 private:
  using RowColEditFn = fm_status_t (*)(fm_workbook_t*, std::size_t, uint32_t, uint32_t);
  void invoke_row_col_edit(RowColEditFn fn, uint32_t sheet, uint32_t index, uint32_t count);

  using TraceFn = fm_status_t (*)(const fm_workbook_t*, std::size_t, uint32_t, uint32_t, uint32_t, fm_cell_nodes_t**);
  emscripten::val trace_to_val(TraceFn fn, uint32_t sheet, uint32_t row, uint32_t col, uint32_t depth) const;

  /// Style-pool add: `pull(reader, record)` fills the caller's record, and
  /// `add(handle_, &index)` stores it. Returns the new index, or 0 after a
  /// recorded failure.
  template <typename Pull, typename Add>
  uint32_t add_style_record(const char* op, const emscripten::val& record, Pull&& pull, Add&& add);

  /// Style-pool replace: as `add_style_record` without an index.
  template <typename Pull, typename Apply>
  void set_style_record(const char* op, const emscripten::val& record, Pull&& pull, Apply&& apply);

  // Re-points the engine's `user_data` at this wrapper. Called after a
  // move, which leaves the C ABI holding the address of the object the
  // callback was registered from.
  void rebind_progress_callback();

  // C-ABI compatible trampoline that forwards to the JS callback held by
  // the `JsWorkbook` in `user_data`, through `call_js_callback`, so a
  // callback that throws aborts the solve instead of unwinding the
  // engine. Returning `false` from the JS side aborts it too, and the two
  // are told apart on the way out of `recalc`. The return type matches
  // `fm_iterative_progress_cb`: the header-wide wide-POD boolean
  // convention (`int32_t`, `0` = abort), not a C `bool`.
  static int32_t iterativeProgressTrampoline(uint32_t iteration, double max_residual, uint32_t max_iterations,
                                             void* user_data);

  fm_workbook_t* handle_ = nullptr;
  // The JS progress callback installed on this workbook, or null. Held per
  // wrapper rather than per module so two workbooks driven from one page
  // keep their own callbacks; the Node addon has always behaved this way
  // and `packages/npm-native/README.md` claims the two surfaces agree.
  emscripten::val progress_callback_ = emscripten::val::null();
  // Set when `progress_callback_` threw during the pass that is still
  // unwinding. Consumed by the recalc entry points.
  bool progress_callback_threw_ = false;
  // True for the duration of a call into `progress_callback_`. Mirrors the
  // Node addon's `in_iterative_progress_callback_` guard: `~JsWorkbook()`
  // refuses to run while this is set, so `delete()` from inside the
  // callback can't free the handle a still-running `recalc` is using.
  bool in_iterative_progress_callback_ = false;
};

// ---- Free helpers ------------------------------------------------------

/// Convenience entry point that mirrors `formulon_cli eval <formula>`:
/// spin up an empty workbook with the default Sheet1, place the formula
/// at A1, recalc, and return the resulting cell value.
emscripten::val eval_formula(const std::string& formula);

/// Returns the Formulon library version (NUL-terminated UTF-8).
std::string version_string();

/// Returns the static C string for `status`.
std::string status_string(int32_t status);

/// The catalog record of a built-in, or `null`. See `fm_get_function_metadata`.
emscripten::val function_metadata(const std::string& name);

/// Every built-in's canonical name.
emscripten::val function_names();

std::string localize_function_name(const std::string& canonical_name, const std::string& profile_id);
std::string canonicalize_function_name(const std::string& localized_name, const std::string& profile_id);
std::string localize_formula(const std::string& formula, const std::string& profile_id);
std::string canonicalize_formula(const std::string& formula, const std::string& profile_id);

/// Separators, boolean names, date order, currency placement and localized
/// error names of a profile. See `fm_get_locale_facts`.
emscripten::val locale_facts(const std::string& profile_id);

/// The profile's localized error names in error-code order.
emscripten::val list_locale_error_names(const std::string& profile_id);

/// The profile's collation order of `a` and `b`. See `fm_compare_text`.
double compare_text(const std::string& a, const std::string& b, int32_t strength, const std::string& profile_id);

/// Equal-under-collation class ids of `values`. See `fm_get_collation_classes`.
emscripten::val collation_classes(emscripten::val values, int32_t strength, const std::string& profile_id);

/// Format and pixel size sniffed from an image header. See `fm_get_image_info`.
emscripten::val probe_image(emscripten::val bytes);

/// Sets the engine's minimum structured-log severity. Process-wide, not
/// per workbook handle; the default is `FM_LOG_LEVEL_OFF` (4), under which
/// the engine writes nothing anywhere.
void set_log_min_level(int32_t level);

/// Routes structured-log records to `sink`, or restores the (silent at the
/// default threshold) stderr fallback when `sink` is null/undefined.
/// Process-wide, not per workbook handle. The sink receives one complete
/// JSON record as a `Uint8Array` view valid only for the duration of the
/// call; copy or decode it before returning.
void set_log_sink(emscripten::val sink);

/// Reads a `DrawingAnchor` (`anchorKind`, `editAs`, `row`, `col`, the EMU
/// offsets and size) shared by the picture, chart and re-anchor calls; an
/// absent `anchorKind` takes `default_anchor_kind`.
fm_drawing_anchor_t pull_drawing_anchor(JsNarrowNumericReader& reader, const emscripten::val& anchor,
                                        int32_t default_anchor_kind);

/// Reads an `ImageInsertInput` (`name`, `description`, nested `anchor`); an
/// absent `input` is all defaults. The stores own the returned strings.
fm_image_insert_t pull_image_insert(JsNarrowNumericReader& reader, const emscripten::val& input,
                                    int32_t default_anchor_kind, std::string& name_store,
                                    std::string& description_store);

}  // namespace parts
}  // namespace wasm
}  // namespace formulon

#endif  // FORMULON_WASM_PARTS_WORKBOOK_H_
