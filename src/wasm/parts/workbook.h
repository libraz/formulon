//
// `JsWorkbook` is the move-only RAII wrapper around `fm_workbook_t*`
// that backs the JS-facing `Workbook` class. Its method bodies are
// split across `parts/workbook_*.cpp` by API surface area so each TU
// stays manageable; this header is the single declaration the binding
// registrar (`parts/bindings_register.cpp`) keys off.
//
// All entry points reuse the same translation contract: every fallible
// call returns a `JsStatus` (or a result envelope wrapping it) instead
// of throwing across the JS boundary. Static factories (`createDefault`,
// `createEmpty`, `loadBytes`) return `JsWorkbook*` allocated on the
// heap; JS callers manage lifetime via embind's `.delete()`.
//
// The contract holds in the callback direction too: the JS the binding
// calls back into runs behind `call_js_callback`, so a callback that
// throws is reported as a status instead of unwinding C++ frames that
// have no landing pads. See `parts/embind_common.h`.

#ifndef FORMULON_WASM_PARTS_WORKBOOK_H_
#define FORMULON_WASM_PARTS_WORKBOOK_H_

#include <emscripten/val.h>

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

  static JsWorkbook* createDefault();
  static JsWorkbook* createEmpty();
  static JsWorkbook* loadBytes(emscripten::val bytes);

  bool isValid() const { return handle_ != nullptr; }
  JsSaveResult save() const;

  /// Serialises the workbook using an explicit container `format`
  /// (`fm_workbook_format_t` ordinal: 1 = xlsx, 2 = xlsb). Unknown /
  /// undocumented values return `kInvalidArgument` via `status`.
  JsSaveResult saveAs(int32_t format) const;
  JsSaveDiagnosticsResult saveWithDiagnostics(int32_t format) const;
  JsReadDiagnosticsResult readDiagnostics() const;

  // ---- Sheet management ---------------------------------------------------

  JsStatus addSheet(const std::string& name);
  JsStatus moveSheet(uint32_t fromIdx, uint32_t toIdx);
  JsStatus removeSheet(uint32_t index);
  JsStatus renameSheet(uint32_t index, const std::string& newName);

  JsStatus setDefinedName(const std::string& name, const std::string& formula);
  JsStatus setDefinedNameScoped(const std::string& name, const std::string& formula, int32_t localSheetId);

  JsStatus insertRows(uint32_t sheet, uint32_t row, uint32_t count);
  JsStatus deleteRows(uint32_t sheet, uint32_t row, uint32_t count);
  JsStatus insertCols(uint32_t sheet, uint32_t col, uint32_t count);
  JsStatus deleteCols(uint32_t sheet, uint32_t col, uint32_t count);

  uint32_t sheetCount() const;
  JsStringResult sheetName(uint32_t idx) const;

  // ---- Cell value / formula ops ------------------------------------------

  JsStatus setNumber(uint32_t sheet, uint32_t row, uint32_t col, double value);
  JsStatus setBool(uint32_t sheet, uint32_t row, uint32_t col, bool value);
  JsStatus setError(uint32_t sheet, uint32_t row, uint32_t col, int32_t errorCode);
  JsStatus setText(uint32_t sheet, uint32_t row, uint32_t col, const std::string& text);
  JsStatus setCellPhonetic(uint32_t sheet, uint32_t row, uint32_t col, const std::string& phonetic);
  /// Stores a cell's furigana as one `<rPh>` block per element of `runs`,
  /// each `{ sb, eb, text }`. Unlike `setCellPhonetic`, which annotates the
  /// whole cell, this preserves which characters each reading covers. See
  /// `fm_workbook_set_cell_phonetic_runs` for the ordering rules.
  JsStatus setCellPhoneticRuns(uint32_t sheet, uint32_t row, uint32_t col, emscripten::val runs);

  /// Stores how the cell's guide renders: `fontId`, `type` (0=halfwidth
  /// katakana, 1=fullwidth katakana, 2=hiragana, 3=no conversion) and
  /// `alignment` (0=no control, 1=left, 2=center, 3=distributed).
  /// Independent of the readings in both directions.
  JsStatus setCellPhoneticProperties(uint32_t sheet, uint32_t row, uint32_t col, emscripten::val properties);
  JsStatus setBlank(uint32_t sheet, uint32_t row, uint32_t col);
  JsStatus setFormula(uint32_t sheet, uint32_t row, uint32_t col, const std::string& formula);

  JsCellResult getValue(uint32_t sheet, uint32_t row, uint32_t col) const;
  emscripten::val getCellPhonetic(uint32_t sheet, uint32_t row, uint32_t col) const;
  /// Reads a cell's furigana as `{ status, runs: [{ sb, eb, text }] }`,
  /// spans included. `getCellPhonetic` returns the same readings
  /// concatenated, without the spans.
  emscripten::val getCellPhoneticRuns(uint32_t sheet, uint32_t row, uint32_t col) const;

  /// Reads the guide's rendering back as `{ fontId, type, alignment }`.
  /// A cell with no annotation reports the all-zero triple.
  emscripten::val getCellPhoneticProperties(uint32_t sheet, uint32_t row, uint32_t col) const;
  emscripten::val getLambdaText(uint32_t sheet, uint32_t row, uint32_t col) const;

  /// Evaluates `formula` as if entered at `(sheet, row, col)` and returns a
  /// single scalar result without mutating the workbook. Array results are
  /// reduced to their top-left element. See `fm_workbook_evaluate_formula`.
  JsEvalResult evaluateFormulaText(uint32_t sheet, uint32_t row, uint32_t col, const std::string& formula) const;

  /// Evaluates `formula` as if entered at `(sheet, row, col)` and returns the
  /// whole multi-cell result without mutating the workbook. Unlike
  /// `evaluateFormulaText` (which reduces an array to its top-left element),
  /// this yields the full grid as a JS object
  /// `{ status, rows, cols, cells }`, where `cells` is a `rows` x `cols`
  /// nested array of Value objects (row-major, `cells[r][c]`); a scalar
  /// result is reported as a 1x1 array. See
  /// `fm_workbook_evaluate_formula_array`.
  emscripten::val evaluateFormulaArray(uint32_t sheet, uint32_t row, uint32_t col, const std::string& formula) const;

  /// Evaluates `formula` as a conditional-formatting predicate anchored at
  /// `(sheet, row, col)`, with relative refs written relative to
  /// `(anchorRow, anchorCol)`; returns a coerced boolean result. See
  /// `fm_workbook_evaluate_cf_formula`.
  JsEvalResult evaluateConditionalFormula(uint32_t sheet, uint32_t row, uint32_t col, uint32_t anchorRow,
                                          uint32_t anchorCol, const std::string& formula) const;

  // ---- Recalc / calc mode ------------------------------------------------

  JsStatus recalc();
  /// Takes the raw JS argument so the binding can reject non-numeric,
  /// fractional, non-finite, and out-of-range values before embind's normal
  /// unsigned-integer coercion would silently change them.
  JsParallelRecalcResult recalcParallel(emscripten::val threadCount);
  JsStatus setIterative(bool enabled, uint32_t max_iterations, double max_change);
  /// The stored iterative-calculation settings as `{status, enabled,
  /// maxIterations, maxChange}`. The cap and threshold are the stored
  /// values even while `enabled` is false, so a host can render Excel's
  /// iterative-calculation dialog without having written them first.
  emscripten::val getIterative() const;
  uint32_t calcMode() const;
  JsStatus setCalcMode(uint32_t mode);
  /// The pinned wall-clock reading as `{year, month, day, hour, minute,
  /// second}`, or `null` when the workbook follows the host clock. `null`
  /// rather than a flag beside the fields because JS already has a way to
  /// say "no reading", and every field combination is a legal date.
  emscripten::val pinnedNow() const;
  JsStatus setPinnedNow(uint32_t year, uint32_t month, uint32_t day, uint32_t hour, uint32_t minute, uint32_t second);
  JsStatus clearPinnedNow();
  std::string excelProfileId() const;
  JsStatus setExcelProfileId(const std::string& profile_id);
  emscripten::val partialRecalc(emscripten::val viewport);
  /// Installs the JS callback the iterative solver invokes after each
  /// sweep, or clears it when `cb` is null / undefined. A callback that
  /// throws aborts the solve, and the recalc that ran it reports
  /// `kBindingCallbackException`.
  JsStatus setIterativeProgress(emscripten::val cb);

  // ---- Iteration / metadata accessors ------------------------------------

  uint32_t cellCount(uint32_t sheet) const;
  emscripten::val cellAt(uint32_t sheet, uint32_t idx) const;

  uint32_t definedNameCount() const;
  emscripten::val definedNameAt(uint32_t idx) const;

  uint32_t tableCount() const;
  emscripten::val tableAt(uint32_t idx) const;
  JsAddStyleResult createTable(emscripten::val spec);
  JsStatus updateTable(uint32_t idx, emscripten::val spec);
  JsStatus removeTable(uint32_t idx);

  uint32_t passthroughCount() const;
  emscripten::val passthroughAt(uint32_t idx) const;

  uint32_t pivotCount(uint32_t sheet) const;
  emscripten::val pivotLayout(uint32_t sheet, uint32_t pivotIndex) const;

  emscripten::val evaluateCfRange(uint32_t sheet, uint32_t firstRow, uint32_t firstCol, uint32_t lastRow,
                                  uint32_t lastCol, double todaySerial) const;

  /// Computes the printable page grid for one sheet. Returns
  /// `{ status, printArea, horizontalBreaks, verticalBreaks, pageCount }`.
  emscripten::val paginate(uint32_t sheet) const;

  // ---- Print settings -----------------------------------------------------
  //
  // Three surfaces over the same model, from most to least raw. Each
  // getter returns `{ status, ... }` so a bad sheet index is
  // distinguishable from an absent setting.

  /// Raw print-settings fragments. `{ status, xml }`; `xml` is the empty
  /// string when the sheet declares no such element.
  emscripten::val getSheetPageSetupXml(uint32_t sheet) const;
  JsStatus setSheetPageSetupXml(uint32_t sheet, const std::string& xml);
  emscripten::val getSheetPageMarginsXml(uint32_t sheet) const;
  JsStatus setSheetPageMarginsXml(uint32_t sheet, const std::string& xml);
  emscripten::val getSheetPrintOptionsXml(uint32_t sheet) const;
  JsStatus setSheetPrintOptionsXml(uint32_t sheet, const std::string& xml);
  emscripten::val getSheetHeaderFooterXml(uint32_t sheet) const;
  JsStatus setSheetHeaderFooterXml(uint32_t sheet, const std::string& xml);
  emscripten::val getSheetSheetPrXml(uint32_t sheet) const;
  JsStatus setSheetSheetPrXml(uint32_t sheet, const std::string& xml);

  /// Toggles `<sheetPr><pageSetUpPr fitToPage>` without touching anything
  /// else `<sheetPr>` carries. Pair with `fitToWidth` / `fitToHeight` on
  /// the page setup to state the target.
  JsStatus setSheetFitToPage(uint32_t sheet, bool enabled);

  /// Print area as comma-separated A1 ranges. `{ status, ranges }`.
  emscripten::val getSheetPrintArea(uint32_t sheet) const;
  JsStatus setSheetPrintArea(uint32_t sheet, const std::string& rangesA1);

  /// Repeat rows / columns. `{ status, repeatRows, repeatCols }`.
  emscripten::val getSheetPrintTitles(uint32_t sheet) const;
  JsStatus setSheetPrintTitles(uint32_t sheet, const std::string& repeatRows, const std::string& repeatCols);

  /// Manual page breaks. The enumerators return the whole array rather
  /// than a count + getter pair, matching the rest of this surface.
  JsStatus addSheetRowBreak(uint32_t sheet, uint32_t row, bool manual);
  JsStatus addSheetColBreak(uint32_t sheet, uint32_t col, bool manual);
  JsStatus removeSheetRowBreak(uint32_t sheet, uint32_t row);
  JsStatus removeSheetColBreak(uint32_t sheet, uint32_t col);
  JsStatus clearSheetBreaks(uint32_t sheet);
  emscripten::val getSheetRowBreaks(uint32_t sheet) const;
  emscripten::val getSheetColBreaks(uint32_t sheet) const;

  /// Typed patch setters. Only the keys present on the input object are
  /// applied; every other attribute in the stored XML is left alone.
  JsStatus setSheetPageSetup(uint32_t sheet, emscripten::val setup);
  JsStatus setSheetPageMargins(uint32_t sheet, emscripten::val margins);
  JsStatus setSheetPrintOptions(uint32_t sheet, emscripten::val options);
  JsStatus setSheetHeaderFooter(uint32_t sheet, emscripten::val headerFooter);
  emscripten::val getSheetPageSetup(uint32_t sheet) const;
  emscripten::val getSheetPageMargins(uint32_t sheet) const;

  // ---- Sheet view / layout ------------------------------------------------

  JsSheetViewResult getSheetView(uint32_t sheet) const;
  JsSheetProtectionResult getSheetProtection(uint32_t sheet) const;
  JsStatus setSheetProtection(uint32_t sheet, JsSheetProtection in);
  JsStatus setSheetZoom(uint32_t sheet, uint32_t zoomScale);
  JsStatus setSheetFreeze(uint32_t sheet, uint32_t freezeRows, uint32_t freezeCols);
  JsStatus setSheetTabHidden(uint32_t sheet, bool hidden);
  /// States the whole tab visibility, including `veryHidden`, which the
  /// bool setter cannot express in either direction. Takes a raw signed
  /// ordinal so an unknown value is rejected by the C ABI rather than
  /// coerced by embind.
  JsStatus setSheetVisibility(uint32_t sheet, int32_t visibility);
  JsStatus setSheetShowGridLines(uint32_t sheet, bool show);
  JsStatus setSheetShowRowColHeaders(uint32_t sheet, bool show);
  JsStatus setSheetShowZeros(uint32_t sheet, bool show);
  JsStatus setSheetRightToLeft(uint32_t sheet, bool rightToLeft);
  JsStatus setSheetTabSelected(uint32_t sheet, bool selected);
  JsStatus setSheetViewMode(uint32_t sheet, std::string mode);

  emscripten::val getSheetColumns(uint32_t sheet) const;
  JsStatus setColumnWidth(uint32_t sheet, uint32_t first, uint32_t last, double width);
  JsStatus setColumnHidden(uint32_t sheet, uint32_t first, uint32_t last, bool hidden);
  JsStatus setColumnOutline(uint32_t sheet, uint32_t first, uint32_t last, uint32_t level);

  emscripten::val getSheetRowOverrides(uint32_t sheet) const;
  JsStatus setRowHeight(uint32_t sheet, uint32_t row, double height);
  JsStatus setRowHidden(uint32_t sheet, uint32_t row, bool hidden);
  JsStatus setRowOutline(uint32_t sheet, uint32_t row, uint32_t level);

  /** Complete worksheet-level `<autoFilter>` fragment. The empty string
   * means no filter; the status form distinguishes that from a bad sheet. */
  emscripten::val getSheetAutoFilterXml(uint32_t sheet) const;
  JsStatus setSheetAutoFilterXml(uint32_t sheet, const std::string& xml);

  // ---- Styles -------------------------------------------------------------

  emscripten::val getCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col) const;
  JsStatus setCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col, uint32_t xf_index);
  /// Applies one xf across an inclusive rectangle, materialising blank
  /// cells so an empty ruled box renders. Saves one ABI crossing per cell.
  JsStatus setRangeXfIndex(uint32_t sheet, uint32_t firstRow, uint32_t firstCol, uint32_t lastRow, uint32_t lastCol,
                           uint32_t xf_index);
  emscripten::val getCellXf(uint32_t xf_index) const;
  emscripten::val getFont(uint32_t font_index) const;
  emscripten::val getFill(uint32_t fill_index) const;
  emscripten::val getBorder(uint32_t border_index) const;
  emscripten::val getNumFmt(uint32_t num_fmt_id) const;
  emscripten::val getDxf(uint32_t dxf_index) const;

  JsAddStyleResult addFont(emscripten::val record);
  /// Overwrites an existing font slot; every xf naming it restyles at once.
  /// See `fm_styles_set_font`.
  JsStatus setFont(uint32_t font_index, emscripten::val record);
  /// Declares the workbook's default font -- font 0, the record an unstyled
  /// cell resolves to. See `fm_workbook_set_default_font`.
  JsStatus setDefaultFont(emscripten::val record);
  JsAddStyleResult addFill(emscripten::val record);
  JsAddStyleResult addBorder(emscripten::val record);
  JsAddNumFmtResult addNumFmt(const std::string& format_code);
  JsAddStyleResult addXf(emscripten::val record);
  JsAddStyleResult addDxf(emscripten::val record);

  uint32_t fontCount() const;
  uint32_t fillCount() const;
  uint32_t borderCount() const;
  uint32_t xfCount() const;
  uint32_t dxfCount() const;
  uint32_t cellStyleCount() const;
  uint32_t cellStyleXfCount() const;
  emscripten::val getCellStyle(uint32_t index) const;
  emscripten::val getCellStyleXf(uint32_t index) const;
  JsAddStyleResult addCellStyleXf(emscripten::val record);
  JsStatus setCellStyle(const std::string& name, uint32_t xfId, uint32_t builtinId);

  emscripten::val getExternalLinks() const;

  // ---- Sheet UI features (merges / hyperlinks / comments / validations) --

  JsStatus addMerge(uint32_t sheet, emscripten::val range);
  JsStatus removeMerge(uint32_t sheet, emscripten::val range);
  JsStatus removeMergeAt(uint32_t sheet, uint32_t index);
  JsStatus clearMerges(uint32_t sheet);

  emscripten::val getComment(uint32_t sheet, uint32_t row, uint32_t col) const;
  emscripten::val getCommentResult(uint32_t sheet, uint32_t row, uint32_t col) const;
  emscripten::val getComments(uint32_t sheet) const;
  JsStatus setComment(uint32_t sheet, uint32_t row, uint32_t col, const std::string& author, const std::string& text);

  JsStatus addHyperlink(uint32_t sheet, uint32_t row, uint32_t col, const std::string& target,
                        const std::string& display, const std::string& tooltip, const std::string& location);
  JsStatus addHyperlinkRange(uint32_t sheet, uint32_t row, uint32_t col, uint32_t lastRow, uint32_t lastCol,
                             const std::string& target, const std::string& display, const std::string& tooltip,
                             const std::string& location);
  JsStatus removeHyperlink(uint32_t sheet, uint32_t row, uint32_t col);
  JsStatus removeHyperlinkAt(uint32_t sheet, uint32_t index);
  JsStatus clearHyperlinks(uint32_t sheet);
  emscripten::val getHyperlinks(uint32_t sheet) const;
  emscripten::val getMerges(uint32_t sheet) const;

  emscripten::val getValidations(uint32_t sheet) const;
  JsStatus addValidation(uint32_t sheet, emscripten::val v);
  JsStatus removeValidationAt(uint32_t sheet, uint32_t index);
  JsStatus clearValidations(uint32_t sheet);

  // ---- Conditional formats -----------------------------------------------

  emscripten::val getConditionalFormats(uint32_t sheet) const;
  JsAddStyleResult addConditionalFormat(uint32_t sheet, emscripten::val v);
  JsStatus removeConditionalFormatAt(uint32_t sheet, uint32_t index);
  JsStatus clearConditionalFormats(uint32_t sheet);

  // ---- Misc trace / metadata / spill -------------------------------------

  emscripten::val precedents(uint32_t sheet, uint32_t row, uint32_t col, uint32_t depth) const;
  emscripten::val dependents(uint32_t sheet, uint32_t row, uint32_t col, uint32_t depth) const;
  emscripten::val functionMetadata(const std::string& name, uint32_t locale) const;
  emscripten::val functionNames() const;
  std::string localizeFunctionName(const std::string& canonical_name, uint32_t locale) const;
  std::string canonicalizeFunctionName(const std::string& localized_name, uint32_t locale) const;
  emscripten::val spillInfo(uint32_t sheet, uint32_t row, uint32_t col) const;

  // ---- PivotCache mutation -----------------------------------------------

  uint32_t pivotCacheCount() const;
  JsAddStyleResult pivotCacheIdAt(uint32_t idx) const;
  JsAddStyleResult pivotCacheCreate(uint32_t requestedId);
  JsStatus pivotCacheRemove(uint32_t cacheId);
  emscripten::val pivotCacheGetWorksheetSource(uint32_t cacheId) const;
  JsStatus pivotCacheSetWorksheetSource(uint32_t cacheId, emscripten::val source);

  uint32_t pivotCacheFieldCount(uint32_t cacheId) const;
  JsStringResult pivotCacheFieldName(uint32_t cacheId, uint32_t fieldIdx) const;
  JsAddStyleResult pivotCacheFieldAdd(uint32_t cacheId, const std::string& name);
  JsStatus pivotCacheFieldClear(uint32_t cacheId);

  uint32_t pivotCacheFieldSharedItemCount(uint32_t cacheId, uint32_t fieldIdx) const;
  JsStatus pivotCacheFieldAddSharedItemNumber(uint32_t cacheId, uint32_t fieldIdx, double value);
  JsStatus pivotCacheFieldAddSharedItemText(uint32_t cacheId, uint32_t fieldIdx, const std::string& utf8);
  JsStatus pivotCacheFieldAddSharedItemBool(uint32_t cacheId, uint32_t fieldIdx, bool value);
  JsStatus pivotCacheFieldAddSharedItemBlank(uint32_t cacheId, uint32_t fieldIdx);
  JsStatus pivotCacheFieldAddSharedItemError(uint32_t cacheId, uint32_t fieldIdx, int32_t errorCode);
  JsStatus pivotCacheFieldClearSharedItems(uint32_t cacheId, uint32_t fieldIdx);

  uint32_t pivotCacheRecordCount(uint32_t cacheId) const;
  JsAddStyleResult pivotCacheRecordAdd(uint32_t cacheId);
  JsStatus pivotCacheRecordClear(uint32_t cacheId);
  JsStatus pivotCacheRecordSetNumber(uint32_t cacheId, uint32_t recordIdx, uint32_t fieldIdx, double value);
  JsStatus pivotCacheRecordSetText(uint32_t cacheId, uint32_t recordIdx, uint32_t fieldIdx, const std::string& utf8);
  JsStatus pivotCacheRecordSetBool(uint32_t cacheId, uint32_t recordIdx, uint32_t fieldIdx, bool value);
  JsStatus pivotCacheRecordSetBlank(uint32_t cacheId, uint32_t recordIdx, uint32_t fieldIdx);
  JsStatus pivotCacheRecordSetError(uint32_t cacheId, uint32_t recordIdx, uint32_t fieldIdx, int32_t errorCode);

  // ---- PivotTable mutation ------------------------------------------------

  JsAddStyleResult pivotCreate(uint32_t sheet, const std::string& name, uint32_t cacheId, uint32_t anchorRow,
                               uint32_t anchorCol);
  JsStatus pivotRemove(uint32_t sheet, uint32_t pivotIdx);
  JsStatus pivotSetName(uint32_t sheet, uint32_t pivotIdx, const std::string& name);
  JsStatus pivotSetAnchor(uint32_t sheet, uint32_t pivotIdx, uint32_t anchorRow, uint32_t anchorCol, uint32_t spanRows,
                          uint32_t spanCols);
  JsStatus pivotSetGrandTotals(uint32_t sheet, uint32_t pivotIdx, bool rowsEnabled, bool colsEnabled);
  emscripten::val pivotGetLayout(uint32_t sheet, uint32_t pivotIdx) const;
  JsStatus pivotSetLayout(uint32_t sheet, uint32_t pivotIdx, uint32_t layout);

  uint32_t pivotFieldCount(uint32_t sheet, uint32_t pivotIdx) const;
  JsAddStyleResult pivotFieldAdd(uint32_t sheet, uint32_t pivotIdx, emscripten::val spec);
  JsStatus pivotFieldClear(uint32_t sheet, uint32_t pivotIdx);
  JsStatus pivotFieldSetAxis(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, uint32_t axis);
  JsStatus pivotFieldSetSort(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, bool ascending,
                             const std::string& byField);
  JsStatus pivotFieldSetSubtotalTop(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, bool top);
  JsStatus pivotFieldAddAggregation(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, uint32_t agg);
  JsStatus pivotFieldClearAggregations(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx);
  JsStatus pivotFieldAddItem(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, const std::string& name,
                             bool visible);
  /// Appends a manual-filter item addressed by its position in the bound
  /// cache field's shared items. The only form that can express the blank
  /// item, which has no label of its own and is matched by its binding.
  JsStatus pivotFieldAddItemAt(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, uint32_t cacheIndex, bool visible);
  JsStatus pivotFieldClearItems(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx);
  JsStatus pivotFieldSetItemVisible(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, uint32_t itemIdx,
                                    bool visible);
  JsStatus pivotFieldAddSubtotalFn(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, uint32_t agg);
  JsStatus pivotFieldClearSubtotalFns(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx);
  JsStatus pivotFieldSetDateGroup(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, uint32_t granularity,
                                  uint32_t calendar, int32_t startYear, int32_t endYear);
  JsStatus pivotFieldClearDateGroup(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx);
  JsStatus pivotFieldSetNumberFormat(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, const std::string& utf8);

  JsStatus pivotSetRowFieldOrder(uint32_t sheet, uint32_t pivotIdx, emscripten::val indices);
  JsStatus pivotSetColFieldOrder(uint32_t sheet, uint32_t pivotIdx, emscripten::val indices);

  uint32_t pivotDataFieldCount(uint32_t sheet, uint32_t pivotIdx) const;
  JsAddStyleResult pivotDataFieldAdd(uint32_t sheet, uint32_t pivotIdx, emscripten::val spec);
  JsStatus pivotDataFieldClear(uint32_t sheet, uint32_t pivotIdx);
  JsStatus pivotDataFieldSet(uint32_t sheet, uint32_t pivotIdx, uint32_t dataFieldIdx, emscripten::val spec);

  uint32_t pivotFilterCount(uint32_t sheet, uint32_t pivotIdx) const;
  /// Reads the active filter at `filterIdx`.
  ///
  /// The active-filter list is **session state**: an entry added through
  /// `pivotFilterAdd` affects evaluation only while this workbook handle
  /// lives and is not written by `save`. Conversely a `<filters>` block
  /// Excel wrote is preserved verbatim on save but does not appear here, so
  /// `pivotFilterCount` reports only what this session added. The filter
  /// surface that does persist is pivot field item visibility.
  emscripten::val pivotFilterAt(uint32_t sheet, uint32_t pivotIdx, uint32_t filterIdx) const;
  JsStatus pivotFilterAdd(uint32_t sheet, uint32_t pivotIdx, emscripten::val spec);
  JsStatus pivotFilterClear(uint32_t sheet, uint32_t pivotIdx);
  JsStatus pivotFilterRemoveAt(uint32_t sheet, uint32_t pivotIdx, uint32_t filterIdx);

 private:
  using RowColEditFn = fm_status_t (*)(fm_workbook_t*, uint32_t, uint32_t, uint32_t);
  JsStatus invoke_row_col_edit(RowColEditFn fn, uint32_t sheet, uint32_t index, uint32_t count);

  using TraceFn = fm_status_t (*)(const fm_workbook_t*, uint32_t, uint32_t, uint32_t, uint32_t, fm_cell_nodes_t**);
  emscripten::val trace_to_val(TraceFn fn, uint32_t sheet, uint32_t row, uint32_t col, uint32_t depth) const;

  // Helpers used by pivot mutator implementations.
  static void build_data_field_spec(emscripten::val spec, fm_pivot_data_field_spec_t& out, std::string& name_buf,
                                    std::string& nfmt_buf, bool& has_nfmt);

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
};

// ---- Free helpers ------------------------------------------------------

/// Convenience entry point that mirrors `formulon_cli eval <formula>`:
/// spin up an empty workbook with the default Sheet1, place the formula
/// at A1, recalc, and return the resulting cell value.
JsEvalResult eval_formula(const std::string& formula);

/// Returns the Formulon library version (NUL-terminated UTF-8).
std::string version_string();

/// Returns the static C string for `status`.
std::string status_string(int32_t status);

/// Returns an Excel literal such as `#DIV/0!` for a cell error code.
std::string error_display_name(int32_t error_code);

/// Returns the most-recent thread-local error message.
std::string last_error_message();

/// Returns the most-recent thread-local error context.
std::string last_error_context();

/// Sets the engine's minimum structured-log severity. Process-wide, not
/// per workbook handle; the default is `FM_LOG_LEVEL_OFF` (4), under which
/// the engine writes nothing anywhere.
JsStatus set_log_min_level(int32_t level);

/// Routes structured-log records to `sink`, or restores the (silent at the
/// default threshold) stderr fallback when `sink` is null/undefined.
/// Process-wide, not per workbook handle. The sink receives one complete
/// JSON record as a `Uint8Array` view valid only for the duration of the
/// call; copy or decode it before returning.
JsStatus set_log_sink(emscripten::val sink);

}  // namespace parts
}  // namespace wasm
}  // namespace formulon

#endif  // FORMULON_WASM_PARTS_WORKBOOK_H_
