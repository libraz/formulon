// `Workbook` ObjectWrap declaration for the Node.js N-API addon. The
// class declaration lives in this shared header because its method
// bodies are split across per-area TUs under `src/node_addon/parts/`.
// Each part TU implements one logical surface (cells, sheets, pivot,
// styles, etc.); they all observe the same `handle_` member through
// this class definition.
//
// Method names on the C++ side preserve the original `addon.cc`
// naming (PascalCase) so the registration table in
// `Workbook::GetClass()` (lives in `workbook_class.cc`) can resolve
// every `InstanceMethod<&Workbook::Foo>` template argument.

#ifndef FORMULON_NODE_ADDON_PARTS_WORKBOOK_CLASS_H_
#define FORMULON_NODE_ADDON_PARTS_WORKBOOK_CLASS_H_

#include "node_addon/parts/addon_common.h"

namespace formulon_node {

class Workbook : public Napi::ObjectWrap<Workbook> {
 public:
  static Napi::Function GetClass(Napi::Env env);

  explicit Workbook(const Napi::CallbackInfo& info);
  ~Workbook() override;

  Workbook(const Workbook&) = delete;
  Workbook& operator=(const Workbook&) = delete;
  Workbook(Workbook&&) = delete;
  Workbook& operator=(Workbook&&) = delete;

  // Static factories.
  static Napi::Value CreateDefault(const Napi::CallbackInfo& info);
  static Napi::Value CreateEmpty(const Napi::CallbackInfo& info);
  static Napi::Value LoadBytes(const Napi::CallbackInfo& info);

  // Cell mutation.
  Napi::Value SetNumber(const Napi::CallbackInfo& info);
  Napi::Value SetBool(const Napi::CallbackInfo& info);
  Napi::Value SetError(const Napi::CallbackInfo& info);
  Napi::Value SetText(const Napi::CallbackInfo& info);
  Napi::Value SetBlank(const Napi::CallbackInfo& info);
  Napi::Value SetFormula(const Napi::CallbackInfo& info);

  // Cell read.
  Napi::Value GetValue(const Napi::CallbackInfo& info);

  // Ad-hoc, side-effect-free formula evaluation.
  Napi::Value EvaluateFormulaText(const Napi::CallbackInfo& info);
  Napi::Value EvaluateFormulaArray(const Napi::CallbackInfo& info);
  Napi::Value EvaluateConditionalFormula(const Napi::CallbackInfo& info);

  // Lifecycle.
  Napi::Value Dispose(const Napi::CallbackInfo& info);
  Napi::Value IsValid(const Napi::CallbackInfo& info);
  Napi::Value MemoryUsage(const Napi::CallbackInfo& info);

  // Lambda text read.
  Napi::Value GetLambdaText(const Napi::CallbackInfo& info);

  // Recalc + save.
  Napi::Value Recalc(const Napi::CallbackInfo& info);
  Napi::Value RecalcParallel(const Napi::CallbackInfo& info);
  Napi::Value PartialRecalc(const Napi::CallbackInfo& info);
  Napi::Value SetIterative(const Napi::CallbackInfo& info);
  Napi::Value SetIterativeProgress(const Napi::CallbackInfo& info);
  Napi::Value Save(const Napi::CallbackInfo& info);
  Napi::Value SaveAs(const Napi::CallbackInfo& info);
  Napi::Value SaveWithDiagnostics(const Napi::CallbackInfo& info);
  Napi::Value ReadDiagnostics(const Napi::CallbackInfo& info);

  // Workbook-level calc policy / behaviour profile.
  Napi::Value CalcMode(const Napi::CallbackInfo& info);
  Napi::Value SetCalcMode(const Napi::CallbackInfo& info);
  /// Clock seam: `pinnedNow()` returns `{year, ..., second}` or `null`;
  /// `setPinnedNow(y, mo, d, h, mi, s)` and `clearPinnedNow()` return a
  /// status. See `fm_workbook_set_pinned_now` for the accepted ranges.
  Napi::Value PinnedNow(const Napi::CallbackInfo& info);
  Napi::Value SetPinnedNow(const Napi::CallbackInfo& info);
  Napi::Value ClearPinnedNow(const Napi::CallbackInfo& info);
  Napi::Value ExcelProfileId(const Napi::CallbackInfo& info);
  Napi::Value SetExcelProfileId(const Napi::CallbackInfo& info);

  // Dependency-graph trace and dynamic-array spill.
  Napi::Value Precedents(const Napi::CallbackInfo& info);
  Napi::Value Dependents(const Napi::CallbackInfo& info);
  Napi::Value SpillInfo(const Napi::CallbackInfo& info);

  // External links.
  Napi::Value GetExternalLinks(const Napi::CallbackInfo& info);

  // Function catalog.
  Napi::Value FunctionMetadata(const Napi::CallbackInfo& info);
  Napi::Value FunctionNames(const Napi::CallbackInfo& info);
  Napi::Value LocalizeFunctionName(const Napi::CallbackInfo& info);
  Napi::Value CanonicalizeFunctionName(const Napi::CallbackInfo& info);

  // Sheet operations.
  Napi::Value AddSheet(const Napi::CallbackInfo& info);
  Napi::Value RemoveSheet(const Napi::CallbackInfo& info);
  Napi::Value RenameSheet(const Napi::CallbackInfo& info);
  Napi::Value MoveSheet(const Napi::CallbackInfo& info);
  Napi::Value SheetCount(const Napi::CallbackInfo& info);
  Napi::Value SheetName(const Napi::CallbackInfo& info);
  Napi::Value Paginate(const Napi::CallbackInfo& info);

  // Row / column structural edits.
  Napi::Value InsertRows(const Napi::CallbackInfo& info);
  Napi::Value DeleteRows(const Napi::CallbackInfo& info);
  Napi::Value InsertCols(const Napi::CallbackInfo& info);
  Napi::Value DeleteCols(const Napi::CallbackInfo& info);

  // Iteration / metadata.
  Napi::Value CellCount(const Napi::CallbackInfo& info);
  Napi::Value CellAt(const Napi::CallbackInfo& info);
  Napi::Value DefinedNameCount(const Napi::CallbackInfo& info);
  Napi::Value DefinedNameAt(const Napi::CallbackInfo& info);
  Napi::Value TableCount(const Napi::CallbackInfo& info);
  Napi::Value TableAt(const Napi::CallbackInfo& info);
  Napi::Value PassthroughCount(const Napi::CallbackInfo& info);
  Napi::Value PassthroughAt(const Napi::CallbackInfo& info);
  Napi::Value PivotCount(const Napi::CallbackInfo& info);
  Napi::Value PivotLayout(const Napi::CallbackInfo& info);

  // PivotCache mutation.
  Napi::Value PivotCacheCount(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheIdAt(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheCreate(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRemove(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheGetWorksheetSource(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheSetWorksheetSource(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldCount(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldName(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldAdd(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldClear(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldSharedItemCount(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldAddSharedItemNumber(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldAddSharedItemText(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldAddSharedItemBool(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldAddSharedItemBlank(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldAddSharedItemError(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldClearSharedItems(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordCount(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordAdd(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordClear(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordSetNumber(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordSetText(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordSetBool(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordSetBlank(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordSetError(const Napi::CallbackInfo& info);

  // PivotTable mutation.
  Napi::Value PivotCreate(const Napi::CallbackInfo& info);
  Napi::Value PivotRemove(const Napi::CallbackInfo& info);
  Napi::Value PivotSetName(const Napi::CallbackInfo& info);
  Napi::Value PivotSetAnchor(const Napi::CallbackInfo& info);
  Napi::Value PivotSetGrandTotals(const Napi::CallbackInfo& info);
  Napi::Value PivotGetLayout(const Napi::CallbackInfo& info);
  Napi::Value PivotSetLayout(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldCount(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldAdd(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldClear(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldSetAxis(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldSetSort(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldSetSubtotalTop(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldAddAggregation(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldClearAggregations(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldAddItem(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldClearItems(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldSetItemVisible(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldAddSubtotalFn(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldClearSubtotalFns(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldSetDateGroup(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldClearDateGroup(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldSetNumberFormat(const Napi::CallbackInfo& info);
  Napi::Value PivotSetRowFieldOrder(const Napi::CallbackInfo& info);
  Napi::Value PivotSetColFieldOrder(const Napi::CallbackInfo& info);
  Napi::Value PivotDataFieldCount(const Napi::CallbackInfo& info);
  Napi::Value PivotDataFieldAdd(const Napi::CallbackInfo& info);
  Napi::Value PivotDataFieldClear(const Napi::CallbackInfo& info);
  Napi::Value PivotDataFieldSet(const Napi::CallbackInfo& info);
  Napi::Value PivotFilterCount(const Napi::CallbackInfo& info);
  Napi::Value PivotFilterAdd(const Napi::CallbackInfo& info);
  /// Reads the active filter at `filterIdx`.
  ///
  /// The active-filter list is **session state**: an entry added through
  /// `pivotFilterAdd` affects evaluation only while this workbook handle
  /// lives and is not written by `save`. Conversely a `<filters>` block
  /// Excel wrote is preserved verbatim on save but does not appear here,
  /// so `pivotFilterCount` reports only what this session added. The
  /// filter surface that does persist is pivot field item visibility.
  Napi::Value PivotFilterAt(const Napi::CallbackInfo& info);
  Napi::Value PivotFilterClear(const Napi::CallbackInfo& info);
  Napi::Value PivotFilterRemoveAt(const Napi::CallbackInfo& info);

  // Defined names.
  Napi::Value SetDefinedName(const Napi::CallbackInfo& info);
  Napi::Value SetDefinedNameScoped(const Napi::CallbackInfo& info);

  // Conditional formatting.
  Napi::Value EvaluateCfRange(const Napi::CallbackInfo& info);
  Napi::Value GetConditionalFormats(const Napi::CallbackInfo& info);
  Napi::Value AddConditionalFormat(const Napi::CallbackInfo& info);
  Napi::Value RemoveConditionalFormatAt(const Napi::CallbackInfo& info);
  Napi::Value ClearConditionalFormats(const Napi::CallbackInfo& info);

  // Sheet view / layout.
  Napi::Value GetSheetView(const Napi::CallbackInfo& info);
  Napi::Value GetSheetProtection(const Napi::CallbackInfo& info);
  Napi::Value SetSheetProtection(const Napi::CallbackInfo& info);
  Napi::Value SetSheetZoom(const Napi::CallbackInfo& info);
  Napi::Value SetSheetFreeze(const Napi::CallbackInfo& info);
  Napi::Value SetSheetTabHidden(const Napi::CallbackInfo& info);
  Napi::Value SetSheetShowGridLines(const Napi::CallbackInfo& info);
  Napi::Value SetSheetShowRowColHeaders(const Napi::CallbackInfo& info);
  Napi::Value SetSheetShowZeros(const Napi::CallbackInfo& info);
  Napi::Value SetSheetRightToLeft(const Napi::CallbackInfo& info);
  Napi::Value SetSheetTabSelected(const Napi::CallbackInfo& info);
  Napi::Value SetSheetViewMode(const Napi::CallbackInfo& info);
  Napi::Value GetSheetColumns(const Napi::CallbackInfo& info);
  Napi::Value SetColumnWidth(const Napi::CallbackInfo& info);
  Napi::Value SetColumnHidden(const Napi::CallbackInfo& info);
  Napi::Value SetColumnOutline(const Napi::CallbackInfo& info);
  Napi::Value GetSheetRowOverrides(const Napi::CallbackInfo& info);
  Napi::Value SetRowHeight(const Napi::CallbackInfo& info);
  Napi::Value SetRowHidden(const Napi::CallbackInfo& info);
  Napi::Value SetRowOutline(const Napi::CallbackInfo& info);

  // Styles.
  Napi::Value GetCellXfIndex(const Napi::CallbackInfo& info);
  Napi::Value SetCellXfIndex(const Napi::CallbackInfo& info);
  Napi::Value GetCellXf(const Napi::CallbackInfo& info);
  Napi::Value GetFont(const Napi::CallbackInfo& info);
  Napi::Value GetFill(const Napi::CallbackInfo& info);
  Napi::Value GetBorder(const Napi::CallbackInfo& info);
  Napi::Value GetNumFmt(const Napi::CallbackInfo& info);
  Napi::Value GetDxf(const Napi::CallbackInfo& info);
  Napi::Value AddFont(const Napi::CallbackInfo& info);
  Napi::Value AddFill(const Napi::CallbackInfo& info);
  Napi::Value AddBorder(const Napi::CallbackInfo& info);
  Napi::Value AddNumFmt(const Napi::CallbackInfo& info);
  Napi::Value AddXf(const Napi::CallbackInfo& info);
  Napi::Value AddDxf(const Napi::CallbackInfo& info);
  Napi::Value FontCount(const Napi::CallbackInfo& info);
  Napi::Value FillCount(const Napi::CallbackInfo& info);
  Napi::Value BorderCount(const Napi::CallbackInfo& info);
  Napi::Value XfCount(const Napi::CallbackInfo& info);
  Napi::Value DxfCount(const Napi::CallbackInfo& info);

  // Named cell styles.
  Napi::Value CellStyleCount(const Napi::CallbackInfo& info);
  Napi::Value CellStyleXfCount(const Napi::CallbackInfo& info);
  Napi::Value GetCellStyle(const Napi::CallbackInfo& info);
  Napi::Value GetCellStyleXf(const Napi::CallbackInfo& info);

  // Sheet UI features (merges, comments, hyperlinks, validations).
  Napi::Value AddMerge(const Napi::CallbackInfo& info);
  Napi::Value RemoveMerge(const Napi::CallbackInfo& info);
  Napi::Value RemoveMergeAt(const Napi::CallbackInfo& info);
  Napi::Value ClearMerges(const Napi::CallbackInfo& info);
  Napi::Value GetMerges(const Napi::CallbackInfo& info);
  Napi::Value GetComment(const Napi::CallbackInfo& info);
  Napi::Value GetCommentResult(const Napi::CallbackInfo& info);
  Napi::Value GetComments(const Napi::CallbackInfo& info);
  Napi::Value SetComment(const Napi::CallbackInfo& info);
  Napi::Value AddHyperlink(const Napi::CallbackInfo& info);
  Napi::Value AddHyperlinkRange(const Napi::CallbackInfo& info);
  Napi::Value GetHyperlinks(const Napi::CallbackInfo& info);
  Napi::Value RemoveHyperlink(const Napi::CallbackInfo& info);
  Napi::Value RemoveHyperlinkAt(const Napi::CallbackInfo& info);
  Napi::Value ClearHyperlinks(const Napi::CallbackInfo& info);
  Napi::Value GetValidations(const Napi::CallbackInfo& info);
  Napi::Value AddValidation(const Napi::CallbackInfo& info);
  Napi::Value RemoveValidationAt(const Napi::CallbackInfo& info);
  Napi::Value ClearValidations(const Napi::CallbackInfo& info);

  // ---- Argument helpers (visible to all part TUs) -------------------
  //
  // Made public-static so per-area TUs can reach them without a friend
  // declaration per part. JS callers never touch them.
  static uint32_t ArgU32(const Napi::CallbackInfo& info, size_t idx);
  static double ArgDouble(const Napi::CallbackInfo& info, size_t idx);
  static std::string ArgString(const Napi::CallbackInfo& info, size_t idx);
  static bool ArgBool(const Napi::CallbackInfo& info, size_t idx);

  using RowColEditFn = fm_status_t (*)(fm_workbook_t*, uint32_t, uint32_t, uint32_t);
  Napi::Value InvokeRowColEdit(const Napi::CallbackInfo& info, RowColEditFn fn);

  /// Builds an error-Status envelope when the wrapper has been
  /// finalized / destroyed but JS still holds a reference.
  Napi::Object NullHandleError(Napi::Env env) const { return MakeErrorStatus(env, kBindingInvalidHandle); }

 private:
  /// Matches `fm_iterative_progress_cb`: the header-wide wide-POD boolean
  /// convention (`int32_t`, `0` = abort), not a C `bool`.
  static int32_t IterativeProgressTrampoline(uint32_t iteration, double max_residual, uint32_t max_iterations,
                                             void* user_data);
  void DestroyHandle(Napi::Env env);

  /// Tells V8 how much memory this wrapper is really keeping alive.
  ///
  /// A `Workbook` is one pointer as far as the collector can see, so a
  /// script holding a few hundred loaded workbooks looks cheap and the
  /// heuristics never feel enough pressure to collect them. Reporting the
  /// engine-side estimate as external memory puts the real weight in
  /// front of the collector.
  ///
  /// Reports the delta since the last call, so it is idempotent and
  /// safe to invoke after any operation that may have changed the
  /// footprint. Destroying the handle reports the whole amount back.
  void SyncExternalMemory(Napi::Env env);

  fm_workbook_t* handle_ = nullptr;
  Napi::FunctionReference iterative_progress_callback_;
  bool in_iterative_progress_callback_ = false;
  /// Bytes currently attributed to this wrapper in V8's external-memory
  /// accounting; the baseline `SyncExternalMemory` computes its delta
  /// against.
  int64_t reported_external_bytes_ = 0;
};

}  // namespace formulon_node

#endif  // FORMULON_NODE_ADDON_PARTS_WORKBOOK_CLASS_H_
