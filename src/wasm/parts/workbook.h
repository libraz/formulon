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
  JsSaveResult saveEx(int32_t format) const;

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
  JsStatus setBlank(uint32_t sheet, uint32_t row, uint32_t col);
  JsStatus setFormula(uint32_t sheet, uint32_t row, uint32_t col, const std::string& formula);

  JsCellResult getValue(uint32_t sheet, uint32_t row, uint32_t col) const;
  emscripten::val getCellPhonetic(uint32_t sheet, uint32_t row, uint32_t col) const;
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
  JsStatus setIterative(bool enabled, uint32_t max_iterations, double max_change);
  uint32_t calcMode() const;
  JsStatus setCalcMode(uint32_t mode);
  std::string excelProfileId() const;
  JsStatus setExcelProfileId(const std::string& profile_id);
  emscripten::val partialRecalc(emscripten::val viewport);
  JsStatus setIterativeProgress(emscripten::val cb);

  // ---- Iteration / metadata accessors ------------------------------------

  uint32_t cellCount(uint32_t sheet) const;
  emscripten::val cellAt(uint32_t sheet, uint32_t idx) const;

  uint32_t definedNameCount() const;
  emscripten::val definedNameAt(uint32_t idx) const;

  uint32_t tableCount() const;
  emscripten::val tableAt(uint32_t idx) const;

  uint32_t passthroughCount() const;
  emscripten::val passthroughAt(uint32_t idx) const;

  uint32_t pivotCount(uint32_t sheet) const;
  emscripten::val pivotLayout(uint32_t sheet, uint32_t pivotIndex) const;

  JsCfRangeResult evaluateCfRange(uint32_t sheet, uint32_t firstRow, uint32_t firstCol, uint32_t lastRow,
                                  uint32_t lastCol, double todaySerial) const;

  // ---- Sheet view / layout ------------------------------------------------

  JsSheetViewResult getSheetView(uint32_t sheet) const;
  JsSheetProtectionResult getSheetProtection(uint32_t sheet) const;
  JsStatus setSheetProtection(uint32_t sheet, JsSheetProtection in);
  JsStatus setSheetZoom(uint32_t sheet, uint32_t zoomScale);
  JsStatus setSheetFreeze(uint32_t sheet, uint32_t freezeRows, uint32_t freezeCols);
  JsStatus setSheetTabHidden(uint32_t sheet, bool hidden);
  JsStatus setSheetShowGridLines(uint32_t sheet, bool show);
  JsStatus setSheetShowRowColHeaders(uint32_t sheet, bool show);
  JsStatus setSheetShowZeros(uint32_t sheet, bool show);
  JsStatus setSheetRightToLeft(uint32_t sheet, bool rightToLeft);
  JsStatus setSheetTabSelected(uint32_t sheet, bool selected);
  JsStatus setSheetViewMode(uint32_t sheet, std::string mode);

  JsColumnsResult getSheetColumns(uint32_t sheet) const;
  JsStatus setColumnWidth(uint32_t sheet, uint32_t first, uint32_t last, double width);
  JsStatus setColumnHidden(uint32_t sheet, uint32_t first, uint32_t last, bool hidden);
  JsStatus setColumnOutline(uint32_t sheet, uint32_t first, uint32_t last, uint32_t level);

  JsRowsResult getSheetRowOverrides(uint32_t sheet) const;
  JsStatus setRowHeight(uint32_t sheet, uint32_t row, double height);
  JsStatus setRowHidden(uint32_t sheet, uint32_t row, bool hidden);
  JsStatus setRowOutline(uint32_t sheet, uint32_t row, uint32_t level);

  // ---- Styles -------------------------------------------------------------

  emscripten::val getCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col) const;
  JsStatus setCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col, uint32_t xf_index);
  emscripten::val getCellXf(uint32_t xf_index) const;
  emscripten::val getFont(uint32_t font_index) const;
  emscripten::val getFill(uint32_t fill_index) const;
  emscripten::val getBorder(uint32_t border_index) const;
  emscripten::val getNumFmt(uint32_t num_fmt_id) const;
  emscripten::val getDxf(uint32_t dxf_index) const;

  JsAddStyleResult addFont(emscripten::val record);
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
                        const std::string& display, const std::string& tooltip);
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

  // Holder for the currently-installed JS progress callback. The
  // function-local static keeps the slot alive for the WASM module's
  // lifetime without needing a global variable; embind's `val` type is
  // not safe to default-initialise at static-init time.
  static emscripten::val& js_progress_callback();

  // C-ABI compatible trampoline that forwards to the held JS callback.
  // Returning `false` from the JS side aborts the iterative solve.
  static bool iterativeProgressTrampoline(uint32_t iteration, double max_residual, uint32_t max_iterations,
                                          void* user_data);

  fm_workbook_t* handle_ = nullptr;
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

/// Returns the most-recent thread-local error message.
std::string last_error_message();

/// Returns the most-recent thread-local error context.
std::string last_error_context();

}  // namespace parts
}  // namespace wasm
}  // namespace formulon

#endif  // FORMULON_WASM_PARTS_WORKBOOK_H_
