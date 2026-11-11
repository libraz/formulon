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
  /// Estimated bytes the workbook holds; 0 once disposed or when the
  /// estimate fails. See `fm_workbook_get_memory_usage`.
  double memoryUsage() const;

  /// Serialises the workbook as a `Uint8Array` in container `format`
  /// (`fm_workbook_format_t` ordinal: 1 = xlsx, 2 = xlsb); `undefined`
  /// means xlsx. An unknown value fails with `kInvalidArgument`.
  emscripten::val save(emscripten::val format) const;
  /// `save` plus the loss counters the write produced.
  JsSaveDiagnostics saveWithDiagnostics(emscripten::val format) const;
  JsReadDiagnostics readDiagnostics() const;

  // ---- Sheet management ---------------------------------------------------

  /// Adds a sheet before `insertBefore` (`sheetCount` appends); a `null` `name` takes the engine's default name.
  uint32_t addSheet(emscripten::val name, uint32_t insertBefore);
  /// Moves the sheet at `fromIndex` before the sheet now at `insertBefore` (`sheetCount` appends).
  void moveSheet(uint32_t fromIndex, uint32_t insertBefore);
  uint32_t copySheet(uint32_t sourceIndex, uint32_t insertBefore, emscripten::val newName);
  void removeSheet(uint32_t index);
  void renameSheet(uint32_t index, const std::string& newName);

  /// Creates or wholly replaces the defined name `input` describes
  /// (`{ name, formula, localSheetId?, comment?, hidden? }`); an omitted or
  /// `null` `localSheetId` is workbook scope.
  void setDefinedName(emscripten::val input);

  /// Row and column edits take the inclusive span `[first, last]`.
  void insertRows(uint32_t sheet, uint32_t first, uint32_t last);
  void deleteRows(uint32_t sheet, uint32_t first, uint32_t last);
  void insertColumns(uint32_t sheet, uint32_t first, uint32_t last);
  void deleteColumns(uint32_t sheet, uint32_t first, uint32_t last);
  /// Inserts / deletes the cells of `range` (a `CellRect`), moving the band
  /// by `shift` (an `fm_cell_shift_t` ordinal).
  void insertCells(uint32_t sheet, emscripten::val range, int32_t shift);
  void deleteCells(uint32_t sheet, emscripten::val range, int32_t shift);

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
  /// each `{ start, end, text }`. Unlike `setCellPhonetic`, which annotates the
  /// whole cell, this preserves which characters each reading covers. See
  /// `fm_sheet_set_cell_phonetic_runs` for the ordering rules.
  void setCellPhoneticRuns(uint32_t sheet, uint32_t row, uint32_t col, emscripten::val runs);

  /// Stores how the cell's guide renders: `fontIndex`, `type` (0=halfwidth
  /// katakana, 1=fullwidth katakana, 2=hiragana, 3=no conversion) and
  /// `alignment` (0=no control, 1=left, 2=center, 3=distributed).
  /// Independent of the readings in both directions; `null` resets the
  /// cell to the unstated triple.
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

  /// Reads the guide's rendering back as `{ fontIndex, type, alignment }`.
  /// A cell with no annotation reports the all-zero triple.
  emscripten::val getCellPhoneticProperties(uint32_t sheet, uint32_t row, uint32_t col) const;
  bool getCellShowPhonetic(uint32_t sheet, uint32_t row, uint32_t col) const;
  emscripten::val getLambdaText(uint32_t sheet, uint32_t row, uint32_t col) const;

  /// Evaluates `formula` as if entered at `(sheet, row, col)` and returns the
  /// whole multi-cell result without mutating the workbook. It yields the
  /// full grid as a JS object
  /// `{ rows, cols, cells }`, where `cells` is a `rows` x `cols`
  /// nested array of Value objects (row-major, `cells[r][c]`); a scalar
  /// result is reported as a 1x1 array. See
  /// `fm_sheet_evaluate_formula_array`.
  emscripten::val evaluateFormulaArray(uint32_t sheet, uint32_t row, uint32_t col, const std::string& formula) const;

  /// Evaluates `formula` as a conditional-formatting predicate anchored at
  /// `(sheet, row, col)`, with relative refs written relative to
  /// `(anchorRow, anchorCol)`; returns the coerced result as a JS boolean. See
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
  /// `setPinnedNow(null)`: releases the pin; any other single argument is a TypeError.
  void releasePinnedNow(emscripten::val now);
  std::string profileId() const;
  void setProfileId(const std::string& profile_id);
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

  /// Tables are addressed by their sheet and a sheet-local index.
  double tableCount(uint32_t sheet) const;
  emscripten::val tableAt(uint32_t sheet, uint32_t idx) const;
  /// The table's applied style, `null` when it has no `<tableStyleInfo>`.
  emscripten::val getTableStyleOptions(uint32_t sheet, uint32_t idx) const;
  void updateTableStyleOptions(uint32_t sheet, uint32_t idx, emscripten::val patch);
  uint32_t tableRuntimeId(uint32_t sheet, uint32_t idx) const;
  /// `{ sheet, index }` of the table with runtime id `id`.
  emscripten::val resolveTableRuntimeId(uint32_t id) const;
  /// `{ sheet, index }` of the table named `name` (case-insensitive), `null` when there is none.
  emscripten::val findTable(const std::string& name) const;
  uint32_t createTable(uint32_t sheet, emscripten::val spec);
  void updateTable(uint32_t sheet, uint32_t idx, emscripten::val spec);
  void renameTable(uint32_t sheet, uint32_t idx, const std::string& name);
  void removeTable(uint32_t sheet, uint32_t idx);
  /// Sets column `column`'s totals-row function / label from
  /// `{ function?, label?, formula? }`.
  /// Renames column `column` of table `idx`; the result is the final name.
  std::string renameTableColumn(uint32_t sheet, uint32_t idx, uint32_t column, const std::string& name);
  /// Inserts data rows `[first, last]` into table `idx`; `first` equal to the data-row count appends.
  void insertTableRows(uint32_t sheet, uint32_t idx, uint32_t first, uint32_t last);
  /// Inserts columns `[first, last]` into table `idx`; `first` equal to the column count appends.
  void insertTableColumns(uint32_t sheet, uint32_t idx, uint32_t first, uint32_t last);

  emscripten::val listPassthroughs() const;
  bool hasPassthrough(const std::string& path) const;

  double pivotCount(uint32_t sheet) const;
  emscripten::val pivotLayout(uint32_t sheet, uint32_t pivotIndex) const;

  emscripten::val evaluateCfRange(uint32_t sheet, emscripten::val range, emscripten::val options) const;

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
  /// The `<tabColor>` colour spec, or `null` when the sheet has none.
  emscripten::val getSheetTabColor(uint32_t sheet) const;
  /// Sets the tab colour from a colour spec; `null` clears it.
  void setSheetTabColor(uint32_t sheet, emscripten::val color);

  /// Toggles `<sheetPr><pageSetUpPr fitToPage>` without touching anything
  /// else `<sheetPr>` carries. Pair with `fitToWidth` / `fitToHeight` on
  /// the page setup to state the target.

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
  /// `{ first?, last?, manual? }` upserts the break; `null` removes it.
  void setSheetRowBreak(uint32_t sheet, uint32_t row, emscripten::val input);
  void setSheetColBreak(uint32_t sheet, uint32_t col, emscripten::val input);
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
  emscripten::val getSheetPrintOptions(uint32_t sheet) const;
  emscripten::val getSheetHeaderFooter(uint32_t sheet) const;

  // ---- Sheet view / layout ------------------------------------------------

  JsSheetView getSheetView(uint32_t sheet) const;
  /// The sheet tab's visibility as an `fm_sheet_visibility_t` ordinal, and its setter.
  int32_t getVisibility(uint32_t sheet) const;
  void setVisibility(uint32_t sheet, int32_t visibility);
  /// The sheet protection record, `null` when the sheet is not protected.
  emscripten::val getSheetProtection(uint32_t sheet) const;
  /// Protects the sheet with the record; `null` unprotects it.
  void setSheetProtection(uint32_t sheet, emscripten::val in);
  void updateSheetProtection(uint32_t sheet, emscripten::val patch);
  /// Applies the view patch `{ zoomPct?, freezeRows?, freezeCols?, showGridLines?,
  /// showRowColumnHeaders?, showZeros?, showFormulas?, showOutlineSymbols?, rightToLeft?, tabSelected?, viewMode?,
  /// topLeftRow?, topLeftCol?, paneTopLeftRow?, paneTopLeftCol?, splitXTwips?, splitYTwips? }`; an omitted key
  /// keeps the stored value, and no key can be cleared. The cross-field checks run in the C layer.
  void updateSheetView(uint32_t sheet, emscripten::val patch);
  /// The active pane's selection as `{ activeRow, activeCol, ranges }`,
  /// and its setter taking `{ activeRow, activeCol, ranges }`.
  emscripten::val getSheetSelection(uint32_t sheet) const;
  void setSheetSelection(uint32_t sheet, emscripten::val selection);

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
  emscripten::val listCellXfIndices(uint32_t sheet, emscripten::val range) const;
  emscripten::val listShowPhoneticCells(uint32_t sheet, emscripten::val range) const;
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

  emscripten::val getCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col) const;
  void setCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col, uint32_t xf_index);
  void clearCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col);
  /// Applies one xf across an inclusive rectangle, materialising blank
  /// cells so an empty ruled box renders. Saves one ABI crossing per cell.
  void setRangeXfIndex(uint32_t sheet, emscripten::val range, uint32_t xf_index);
  emscripten::val getCellXf(uint32_t xf_index) const;
  emscripten::val getFont(uint32_t font_index) const;
  /// The workbook default font, font record 0.
  emscripten::val getDefaultFont() const;
  emscripten::val getFill(uint32_t fill_index) const;
  emscripten::val getBorder(uint32_t border_index) const;
  emscripten::val getNumFmt(uint32_t num_fmt_id) const;
  emscripten::val listNumberFormats() const;
  emscripten::val getDxf(uint32_t dxf_index) const;

  uint32_t addFont(emscripten::val record);
  /// Overwrites an existing font slot; every xf naming it restyles at once.
  /// See `fm_workbook_set_font`.
  void setFont(uint32_t font_index, emscripten::val record);
  /// Declares the workbook's default font -- font 0, the record an unstyled
  /// cell resolves to. See `fm_workbook_set_default_font`.
  void setDefaultFont(emscripten::val record);
  uint32_t addFill(emscripten::val record);
  /// Overwrite an existing slot like `setFont`; see `fm_workbook_set_fill`, `_border`, `_cell_xf` and `_dxf`.
  void setFill(uint32_t fill_index, emscripten::val record);
  uint32_t addBorder(emscripten::val record);
  void setBorder(uint32_t border_index, emscripten::val record);
  uint32_t addNumFmt(const std::string& format_code);
  uint32_t addXf(emscripten::val record);
  void setXf(uint32_t xf_index, emscripten::val record);
  uint32_t addDxf(emscripten::val record);
  void setDxf(uint32_t dxf_index, emscripten::val record);

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
  void setCellStyleXf(uint32_t index, emscripten::val record);
  /// Adds or replaces a `<cellStyle>` from `{ name, styleXfIndex, builtinId, iLevel, hidden, customBuiltin }`.
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
  emscripten::val getThemeScriptFont(int32_t scheme, const std::string& script) const;
  /// Every script entry of the `scheme` font set, in the theme part's file order.
  emscripten::val listThemeScriptFonts(int32_t scheme) const;
  void setThemeColors(emscripten::val colors);
  void setThemeFonts(emscripten::val fonts);
  void setThemeScriptFont(int32_t scheme, const std::string& script, const std::string& typeface);
  void resetTheme();
  /// `{ argb, resolution }` for a `ColorSpec` used in `context` (`ColorContext.*`).
  emscripten::val resolveColor(emscripten::val spec, int32_t context) const;
  /// `{ xfIndex, source, ..., font, borders, numFmtCode }`: the formatting a cell shows.
  emscripten::val getEffectiveStyle(uint32_t sheet, uint32_t row, uint32_t col) const;

  double getExternalLinkCount() const;
  emscripten::val getExternalLink(uint32_t externalLinkIndex) const;

  // ---- Sheet UI features (merges / hyperlinks / comments / validations) --

  /// Adds the merge and returns the rectangle as stored.
  emscripten::val addMerge(uint32_t sheet, emscripten::val range);
  void removeMerge(uint32_t sheet, emscripten::val range);
  void removeMergeAt(uint32_t sheet, uint32_t index);
  void clearMerges(uint32_t sheet);

  emscripten::val getComment(uint32_t sheet, uint32_t row, uint32_t col) const;
  emscripten::val getComments(uint32_t sheet) const;
  /// Writes the note `{ text, author? }` at `(row, col)`; `null` removes it and is a no-op when there is none.
  void setComment(uint32_t sheet, uint32_t row, uint32_t col, emscripten::val comment);
  /// Whether the note at the cell is shown without hovering, or `null` when there is no note.
  emscripten::val getCommentVisible(uint32_t sheet, uint32_t row, uint32_t col) const;
  void setCommentVisible(uint32_t sheet, uint32_t row, uint32_t col, bool visible);

  /// The workbook's `<workbookProtection>`, or `null` when it has none; `null` to the setter removes it.
  emscripten::val getWorkbookProtection() const;
  void setWorkbookProtection(emscripten::val protection);
  /// The sheet's `<protectedRange>` entries, in file order.
  double getProtectedRangeCount(uint32_t sheet) const;
  /// A string `protectedRange` is a name (`null` when no range has it); a number is a position.
  emscripten::val getProtectedRange(uint32_t sheet, emscripten::val protectedRange) const;
  /// Adds the range and returns its name, the key the other calls take.
  emscripten::val addProtectedRange(uint32_t sheet, emscripten::val range);
  void updateProtectedRange(uint32_t sheet, const std::string& name, emscripten::val range);
  void removeProtectedRange(uint32_t sheet, const std::string& name);

  /// Sets the hyperlink `{ target, location?, display?, tooltip?, range? }` anchored at `(row, col)`; an omitted
  /// or `null` `range` is the key cell. `null` removes the link anchored there and is a no-op when there is none.
  void setHyperlink(uint32_t sheet, uint32_t row, uint32_t col, emscripten::val link);
  /// Inserts the hyperlink `link` at list position `index`; `link.range` is required.
  void insertHyperlinkAt(uint32_t sheet, uint32_t index, emscripten::val link);
  void removeHyperlinkAt(uint32_t sheet, uint32_t index);
  void clearHyperlinks(uint32_t sheet);
  emscripten::val getHyperlinks(uint32_t sheet) const;
  emscripten::val getHyperlink(uint32_t sheet, uint32_t row, uint32_t col) const;
  emscripten::val getMerges(uint32_t sheet) const;
  /// The merges intersecting `range`, in storage order.
  emscripten::val getMergesInRange(uint32_t sheet, emscripten::val range) const;

  /// Typed AutoFilter model for a sheet or (sheet and sheet-local table index) a table.
  /// `setAutoFilter` / `setTableAutoFilter` take `null` to remove the filter.
  emscripten::val getAutoFilter(uint32_t sheet) const;
  void setAutoFilter(uint32_t sheet, emscripten::val filter);
  void applyAutoFilter(uint32_t sheet);
  void clearAutoFilter(uint32_t sheet);
  emscripten::val evaluateAutoFilter(uint32_t sheet) const;
  emscripten::val getTableAutoFilter(uint32_t sheet, uint32_t table) const;
  void setTableAutoFilter(uint32_t sheet, uint32_t table, emscripten::val filter);
  void applyTableAutoFilter(uint32_t sheet, uint32_t table);
  void clearTableAutoFilter(uint32_t sheet, uint32_t table);
  emscripten::val evaluateTableAutoFilter(uint32_t sheet, uint32_t table) const;

  emscripten::val validateValue(uint32_t sheet, uint32_t row, uint32_t col, emscripten::val value) const;
  emscripten::val listInvalidCells(uint32_t sheet, emscripten::val cursor, emscripten::val limit) const;

  emscripten::val getThreadedComments(uint32_t sheet) const;
  emscripten::val getThreadedComment(uint32_t sheet, const std::string& id) const;
  /// Adds the comment and returns its id (the given one, or the engine's when `id` is omitted).
  emscripten::val addThreadedComment(uint32_t sheet, emscripten::val comment);
  void editThreadedComment(uint32_t sheet, const std::string& id, emscripten::val patch);
  void removeThreadedComment(uint32_t sheet, const std::string& id);
  emscripten::val getPersons() const;
  emscripten::val addPerson(emscripten::val person);
  void removePerson(const std::string& id);
  emscripten::val getPerson(const std::string& id) const;

  emscripten::val listDrawingObjects(uint32_t sheet) const;
  emscripten::val getDrawingObject(uint32_t sheet, uint32_t objectId) const;
  emscripten::val getImage(uint32_t sheet, uint32_t objectId) const;
  uint32_t insertImage(uint32_t sheet, emscripten::val bytes, emscripten::val input);
  void removeDrawing(uint32_t sheet, uint32_t objectId);
  void updateDrawing(uint32_t sheet, uint32_t objectId, emscripten::val patch);
  void setDrawingZOrder(uint32_t sheet, uint32_t objectId, uint32_t zOrderIndex);
  emscripten::val snapshotDrawing(uint32_t sheet, uint32_t objectId) const;
  uint32_t restoreDrawing(uint32_t sheet, emscripten::val bytes, emscripten::val opts);

  uint32_t insertShape(uint32_t sheet, emscripten::val spec);
  void updateShape(uint32_t sheet, uint32_t objectId, emscripten::val spec);
  emscripten::val getShape(uint32_t sheet, uint32_t objectId) const;

  emscripten::val getSlicer(uint32_t sheet, uint32_t objectId) const;
  uint32_t addSlicer(uint32_t sheet, uint32_t tableSheet, uint32_t table, uint32_t column, emscripten::val spec);
  void updateSlicer(uint32_t sheet, uint32_t objectId, emscripten::val spec);
  /// Adds the data table and returns `{ row, col }`, the top-left cell of its result range (the key
  /// `getDataTable` / `removeDataTable` take).
  emscripten::val addDataTable(uint32_t sheet, emscripten::val spec);
  emscripten::val getDataTable(uint32_t sheet, uint32_t row, uint32_t col) const;
  void removeDataTable(uint32_t sheet, uint32_t row, uint32_t col);
  void recalcDataTables();

  double getSparklineGroupCount(uint32_t sheet) const;
  emscripten::val getSparklineGroup(uint32_t sheet, emscripten::val sparklineGroup) const;
  emscripten::val addSparklineGroup(uint32_t sheet, emscripten::val spec);
  void updateSparklineGroup(uint32_t sheet, const std::string& groupId, emscripten::val spec);
  void removeSparklineGroup(uint32_t sheet, const std::string& groupId);

  emscripten::val getChart(uint32_t sheet, uint32_t objectId) const;
  emscripten::val getChartData(uint32_t sheet, uint32_t objectId) const;
  emscripten::val previewChartData(uint32_t sheet, uint32_t objectId, emscripten::val overrides) const;
  uint32_t insertChart(uint32_t sheet, emscripten::val options);
  void setChartType(uint32_t sheet, uint32_t objectId, uint32_t plot, int32_t type, emscripten::val options);
  void setChartTitle(uint32_t sheet, uint32_t objectId, emscripten::val title);
  /// Replaces the legend with `{ position, overlay, hiddenEntries }`; `null` removes it.
  void setChartLegend(uint32_t sheet, uint32_t objectId, emscripten::val legend);
  /// Merges `{ position?, overlay?, hiddenEntries? }` into the legend.
  void updateChartLegend(uint32_t sheet, uint32_t objectId, emscripten::val patch);
  void setChartAxis(uint32_t sheet, uint32_t objectId, uint32_t axisId, emscripten::val patch);
  void setChartDisplay(uint32_t sheet, uint32_t objectId, emscripten::val patch);
  void setChartSeries(uint32_t sheet, uint32_t objectId, uint32_t plot, uint32_t seriesId, emscripten::val patch);
  /// Replaces the point's formatting with the record; `null` removes it.
  void setChartPoint(uint32_t sheet, uint32_t objectId, uint32_t plot, uint32_t seriesId, uint32_t pointIndex,
                     emscripten::val point);
  void updateChartPoint(uint32_t sheet, uint32_t objectId, uint32_t plot, uint32_t seriesId, uint32_t pointIndex,
                        emscripten::val patch);
  void setChartPlot(uint32_t sheet, uint32_t objectId, uint32_t plot, emscripten::val patch);
  void setChartFilter(uint32_t sheet, uint32_t objectId, emscripten::val hiddenSeries,
                      emscripten::val hiddenCategories);
  uint32_t addChartSeries(uint32_t sheet, uint32_t objectId, uint32_t plot, emscripten::val patch);
  void removeChartSeries(uint32_t sheet, uint32_t objectId, uint32_t plot, uint32_t seriesId);
  void setChartSource(uint32_t sheet, uint32_t objectId, const std::string& source, int32_t direction,
                      emscripten::val headerRow, emscripten::val headerColumn);

  double getValidationCount(uint32_t sheet) const;
  /// A string `validation` is an id (`null` when no rule has it); a number is a position.
  emscripten::val getValidation(uint32_t sheet, emscripten::val validation) const;
  /// Adds the rule and returns its id; a rule is addressed by id after that.
  emscripten::val addValidation(uint32_t sheet, emscripten::val v);
  void updateValidation(uint32_t sheet, const std::string& validationId, emscripten::val patch);
  void removeValidationAt(uint32_t sheet, const std::string& validationId);
  void clearValidations(uint32_t sheet);

  // ---- Conditional formats -----------------------------------------------

  double getCfRuleCount(uint32_t sheet) const;
  /// A string `cfRule` is an id (`null` when no rule has it); a number is a position.
  emscripten::val getCfRule(uint32_t sheet, emscripten::val cfRule) const;
  /// Applies the rule patch to the rule with id `cfRuleId`; `colorScale`, `dataBar` and `iconSet` patch their
  /// own keys.
  void updateCfRule(uint32_t sheet, const std::string& cfRuleId, emscripten::val patch);
  /// Adds the rule and returns its id, the key `updateCfRule` / `removeCfRule` take.
  emscripten::val addConditionalFormat(uint32_t sheet, emscripten::val v);
  void removeConditionalFormatAt(uint32_t sheet, const std::string& cfRuleId);
  void clearConditionalFormats(uint32_t sheet);

  // ---- Misc trace / metadata / spill -------------------------------------

  emscripten::val precedents(uint32_t sheet, uint32_t row, uint32_t col, uint32_t depth) const;
  emscripten::val dependents(uint32_t sheet, uint32_t row, uint32_t col, uint32_t depth) const;
  emscripten::val getSpill(uint32_t sheet, uint32_t row, uint32_t col) const;
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
  /// Rebuilds the cache from its worksheet source and re-projects the pivots drawing from it.
  void pivotCacheRefresh(uint32_t cacheId);

  double pivotCacheFieldCount(uint32_t cacheId) const;
  std::string pivotCacheFieldName(uint32_t cacheId, uint32_t fieldIndex) const;
  /// The eight `<sharedItems>` content flags of a cache field as `{ containsNumber, ... }` booleans.
  emscripten::val pivotCacheFieldTypes(uint32_t cacheId, uint32_t fieldIndex) const;
  uint32_t pivotCacheFieldAdd(uint32_t cacheId, const std::string& name);
  void pivotCacheFieldClear(uint32_t cacheId);

  double pivotCacheFieldSharedItemCount(uint32_t cacheId, uint32_t fieldIndex) const;
  emscripten::val pivotCacheFieldSharedItemValue(uint32_t cacheId, uint32_t fieldIndex, uint32_t itemIndex) const;
  double pivotCacheFieldAddSharedItemNumber(uint32_t cacheId, uint32_t fieldIndex, double value);
  double pivotCacheFieldAddSharedItemDate(uint32_t cacheId, uint32_t fieldIndex, double serial);
  bool pivotCacheFieldSharedItemIsDate(uint32_t cacheId, uint32_t fieldIndex, uint32_t itemIndex) const;
  double pivotCacheFieldAddSharedItemText(uint32_t cacheId, uint32_t fieldIndex, const std::string& utf8);
  double pivotCacheFieldAddSharedItemBool(uint32_t cacheId, uint32_t fieldIndex, bool value);
  double pivotCacheFieldAddSharedItemBlank(uint32_t cacheId, uint32_t fieldIndex);
  double pivotCacheFieldAddSharedItemError(uint32_t cacheId, uint32_t fieldIndex, int32_t errorCode);
  void pivotCacheFieldClearSharedItems(uint32_t cacheId, uint32_t fieldIndex);

  double pivotCacheRecordCount(uint32_t cacheId) const;
  uint32_t pivotCacheRecordAdd(uint32_t cacheId);
  void pivotCacheRecordClear(uint32_t cacheId);
  /// The cell as a tagged `Value`; a shared-item cell reads the item's value and a date reads as its serial.
  emscripten::val pivotCacheRecordValue(uint32_t cacheId, uint32_t recordIndex, uint32_t fieldIndex) const;
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
  void pivotSetAnchor(uint32_t sheet, uint32_t pivotIndex, emscripten::val cell);
  void pivotSetGrandTotals(uint32_t sheet, uint32_t pivotIndex, bool rowsEnabled, bool colsEnabled);
  void pivotUpdateStyleOptions(uint32_t sheet, uint32_t pivotIndex, emscripten::val info);
  double pivotGetLayout(uint32_t sheet, uint32_t pivotIndex) const;
  void pivotSetLayout(uint32_t sheet, uint32_t pivotIndex, uint32_t layout);
  /// Definition read-back; works for pivots created in-session and loaded
  /// from a package alike. `pivotCacheId` returns the cache id.
  uint32_t pivotCacheId(uint32_t sheet, uint32_t pivotIndex) const;
  std::string pivotName(uint32_t sheet, uint32_t pivotIndex) const;
  uint32_t pivotRuntimeId(uint32_t sheet, uint32_t pivotIndex) const;
  uint32_t resolvePivotRuntimeId(uint32_t sheet, uint32_t id) const;
  emscripten::val pivotGetGrandTotals(uint32_t sheet, uint32_t pivotIndex) const;
  emscripten::val pivotGetStyleOptions(uint32_t sheet, uint32_t pivotIndex) const;
  emscripten::val pivotRowFieldOrder(uint32_t sheet, uint32_t pivotIndex) const;
  emscripten::val pivotColFieldOrder(uint32_t sheet, uint32_t pivotIndex) const;
  emscripten::val pivotPageFieldOrder(uint32_t sheet, uint32_t pivotIndex) const;

  double pivotFieldCount(uint32_t sheet, uint32_t pivotIndex) const;
  emscripten::val pivotFieldAt(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex) const;
  double pivotFieldItemCount(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex) const;
  emscripten::val pivotFieldItemAt(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, uint32_t itemIndex) const;
  void pivotFieldClear(uint32_t sheet, uint32_t pivotIndex);
  void pivotFieldUpdate(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, emscripten::val patch);
  void pivotFieldItemUpdate(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, uint32_t itemIndex,
                            emscripten::val patch);
  void pivotFieldSetAxis(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, int32_t axis);
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
  emscripten::val getPivotSpill(uint32_t sheet, uint32_t pivotIndex) const;
  uint32_t pivotShowDetails(uint32_t sheet, uint32_t pivotIndex, uint32_t row, uint32_t col);
  /// Appends a manual-filter item: `item` is a new item's name, or `{ cacheIndex }` for the bound cache
  /// field's shared item at that position (the only form that can express the blank item, which has no
  /// label of its own and is matched by its binding).
  uint32_t pivotFieldAddItem(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, emscripten::val item,
                             bool visible);
  void pivotFieldClearItems(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex);
  /// Replaces the field's custom subtotal functions with the `PivotAggregation` list; an empty list clears them.
  void pivotFieldSetSubtotalFns(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, emscripten::val aggregations);
  /// Groups a date field by the `PivotDatePeriod` bits of `periods`; NaN
  /// bounds take the field's own.
  void pivotFieldSetDateGroup(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, uint32_t periods,
                              uint32_t calendar, uint32_t intervalDays, double start, double end);
  /// The `null` form of `pivotFieldSetDateGroup`: removes the grouping; any other `group` is a type error.
  void pivotFieldRemoveDateGroup(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, emscripten::val group);
  /// Groups a number field into ranges of `interval`; NaN bounds follow the field's own values.
  void pivotFieldSetNumericGroup(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, double interval,
                                 double start, double end);
  /// The `null` form of `pivotFieldSetNumericGroup`: removes the grouping; any other `group` is a type error.
  void pivotFieldRemoveNumericGroup(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, emscripten::val group);
  /// The stored date grouping over a field as `{ periods, calendar, intervalDays, start, end, autoStart, autoEnd }`;
  /// `null` when the field is not date-grouped.
  emscripten::val pivotFieldGetDateGroup(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex) const;
  /// The stored numeric grouping over a field as `{ start, end, interval, autoStart, autoEnd }`; `null` when the
  /// field is not numerically grouped.
  emscripten::val pivotFieldGetNumericGroup(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex) const;

  void pivotSetRowFieldOrder(uint32_t sheet, uint32_t pivotIndex, emscripten::val indices);
  void pivotSetColFieldOrder(uint32_t sheet, uint32_t pivotIndex, emscripten::val indices);
  void pivotSetPageFieldOrder(uint32_t sheet, uint32_t pivotIndex, emscripten::val indices);

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
  void invoke_row_col_edit(RowColEditFn fn, uint32_t sheet, uint32_t first, uint32_t last);

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

/// Evaluates `formula` read-only against a fresh single-sheet workbook,
/// anchored at `Sheet1!A1`, and returns the whole result as
/// `{rows, cols, cells}` (a scalar is 1 x 1).
emscripten::val evaluate_formula(const std::string& formula);

/// Returns the Formulon library version (NUL-terminated UTF-8).
std::string version_string();

/// Returns the static C string for `status`.

/// The catalog record of a built-in, or `null`. See `fm_get_function_metadata`.
emscripten::val function_metadata(const std::string& name);

/// Every built-in's canonical name.
emscripten::val function_names();

std::string localize_function_name(const std::string& canonical_name, const std::string& profile_id);
std::string canonicalize_function_name(const std::string& localized_name, const std::string& profile_id);
/// `fm_localize_formula` / `fm_canonicalize_formula`, `options` read by `read_notation_options`.
std::string localize_formula(const std::string& formula, const std::string& profile_id, emscripten::val options);
std::string canonicalize_formula(const std::string& formula, const std::string& profile_id, emscripten::val options);

/// `options` as an object (absent reads as `{}`); false after recording a TypeError, worded for the call `call`
/// (`Workbook.formatReference`), for another type.
bool options_object(const emscripten::val& options, const char* call, emscripten::val* out);

/// The `{ r1c1?, baseRow?, baseCol? }` options of `parseReference`, `localizeFormula` and `canonicalizeFormula`.
struct JsNotationOptions {
  bool r1c1 = false;
  bool has_base = false;
  uint32_t base_row = 0;
  uint32_t base_col = 0;
};

/// Reads `options` of the call `call`; false after recording the argument error: a non-object, a non-boolean
/// `r1c1`, a coordinate that is not a uint32, or one base coordinate without the other.
bool read_notation_options(const emscripten::val& options, const char* call, JsNotationOptions* out);

/// `options` as the C record `Spec` (`fm_parse_reference_spec_t`, `fm_formula_notation_spec_t`).
template <typename Spec>
Spec notation_spec(const JsNotationOptions& options) {
  Spec spec = FM_STRUCT_INIT(Spec);
  spec.r1c1 = options.r1c1 ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
  spec.has_base = options.has_base ? 1 : 0;
  spec.base_row = options.base_row;
  spec.base_col = options.base_col;
  return spec;
}

/// Separators, boolean names, date order, currency placement and localized
/// error names of a profile. See `fm_get_locale_facts`.
emscripten::val locale_facts(const std::string& profile_id);

/// Every Excel profile id the engine accepts. See `fm_get_profile_id_count`.
emscripten::val list_profile_ids();

/// The profile's collation order of `a` and `b`. See `fm_compare_text`.
double compare_text(const std::string& a, const std::string& b, int32_t strength, const std::string& profile_id);

/// The Excel serial of `civil` ({year, month, day, hour, minute, second}). See `fm_convert_civil_to_serial`.
double convert_civil_to_serial(emscripten::val civil, bool date1904);

/// The civil date and time of `serial` as an object. See `fm_convert_serial_to_civil`.
emscripten::val convert_serial_to_civil(double serial, bool date1904);

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

/// Reads a `DrawingPlacementInput` (`name`, `description`, `hidden`, nested `anchor`) from `input`, which is the
/// placement object itself or a chart / shape / slicer input that extends it; an absent `input` is all defaults.
/// `param` names `input` in a rejection. The stores own the returned strings.
fm_drawing_placement_spec_t pull_drawing_placement(JsNarrowNumericReader& reader, const emscripten::val& input,
                                                   const char* param, std::string& name_store,
                                                   std::string& description_store);

/// Reads the `DrawingPatch` keys (`name`, `description`, `hidden`, `anchor`) of `patch`, which may also carry the
/// keys of a kind-specific patch; an omitted key is `KEEP`. The stores own the returned strings.
fm_drawing_patch_t pull_drawing_patch(JsNarrowNumericReader& reader, const emscripten::val& patch, const char* param,
                                      std::string& name_store, std::string& description_store);

/// The `ChartFill` / `ChartLine` record of a shape's or chart element's paint, `null` when it states no element of
/// its own (the all-zero record).
emscripten::val chart_fill_to_val(const fm_chart_fill_t& fill);
emscripten::val chart_line_to_val(const fm_chart_line_t& line);

/// Reads `owner.fill` / `owner.line` as a paint patch: omitted keeps every member, `null` clears every member
/// (the style's paint), an object sets the members it carries.
void pull_chart_fill_patch(JsNarrowNumericReader& reader, const emscripten::val& owner, const char* param,
                           fm_chart_fill_patch_t* out);
void pull_chart_line_patch(JsNarrowNumericReader& reader, const emscripten::val& owner, const char* param,
                           fm_chart_line_patch_t* out);

/// Reads `owner.fill` / `owner.line` as the paint of a new object: omitted or `null` is the style's paint (the
/// all-zero record), and a `null` member is an omitted one.
fm_chart_fill_t pull_chart_fill(JsNarrowNumericReader& reader, const emscripten::val& owner, const char* param);
fm_chart_line_t pull_chart_line(JsNarrowNumericReader& reader, const emscripten::val& owner, const char* param);

}  // namespace parts
}  // namespace wasm
}  // namespace formulon

#endif  // FORMULON_WASM_PARTS_WORKBOOK_H_
