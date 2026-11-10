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

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <tuple>
#include <type_traits>
#include <utility>

#include "node_addon/parts/addon_common.h"

namespace formulon_node {

/// Reads a `DrawingAnchorInput` (`kind`, `editAs`, the `from*` / `to*` cells and offsets, `widthEmu`, `heightEmu`);
/// an absent `kind` or `editAs` is the operation's default.
inline fm_drawing_anchor_t PullDrawingAnchor(CheckedSpecReader& reader, const Napi::Object& spec) {
  fm_drawing_anchor_t anchor{};
  anchor.kind = reader.I32(spec, "kind", FM_ANCHOR_KIND_DEFAULT);
  anchor.edit_as = reader.I32(spec, "editAs", FM_ANCHOR_EDIT_AS_DEFAULT);
  anchor.from_row = reader.U32(spec, "fromRow", 0U);
  anchor.from_col = reader.U32(spec, "fromCol", 0U);
  anchor.from_row_offset_emu = reader.I64(spec, "fromRowOffsetEmu", 0);
  anchor.from_col_offset_emu = reader.I64(spec, "fromColOffsetEmu", 0);
  anchor.to_row = reader.U32(spec, "toRow", 0U);
  anchor.to_col = reader.U32(spec, "toCol", 0U);
  anchor.to_row_offset_emu = reader.I64(spec, "toRowOffsetEmu", 0);
  anchor.to_col_offset_emu = reader.I64(spec, "toColOffsetEmu", 0);
  anchor.width_emu = reader.I64(spec, "widthEmu", 0);
  anchor.height_emu = reader.I64(spec, "heightEmu", 0);
  return anchor;
}

/// Reads a `DrawingAnchorPatch`: an omitted key keeps, and no key is clearable.
inline fm_drawing_anchor_patch_t PullDrawingAnchorPatch(CheckedSpecReader& reader, const Napi::Object& spec) {
  fm_drawing_anchor_patch_t a{};
  a.kind_op = PatchI32(reader, spec, "kind", &a.kind);
  a.edit_as_op = PatchI32(reader, spec, "editAs", &a.edit_as);
  a.from_row_op = PatchU32(reader, spec, "fromRow", &a.from_row);
  a.from_col_op = PatchU32(reader, spec, "fromCol", &a.from_col);
  a.from_row_offset_emu_op = PatchI64(reader, spec, "fromRowOffsetEmu", &a.from_row_offset_emu);
  a.from_col_offset_emu_op = PatchI64(reader, spec, "fromColOffsetEmu", &a.from_col_offset_emu);
  a.to_row_op = PatchU32(reader, spec, "toRow", &a.to_row);
  a.to_col_op = PatchU32(reader, spec, "toCol", &a.to_col);
  a.to_row_offset_emu_op = PatchI64(reader, spec, "toRowOffsetEmu", &a.to_row_offset_emu);
  a.to_col_offset_emu_op = PatchI64(reader, spec, "toColOffsetEmu", &a.to_col_offset_emu);
  a.width_emu_op = PatchI64(reader, spec, "widthEmu", &a.width_emu);
  a.height_emu_op = PatchI64(reader, spec, "heightEmu", &a.height_emu);
  return a;
}

/// Reads a `DrawingPlacementInput` (`name`, `description`, `hidden`, nested `anchor`) out of `input`; an absent
/// `input` is all defaults. The stores own the returned strings.
inline fm_drawing_placement_spec_t PullPlacement(CheckedSpecReader& reader, const Napi::Value& input,
                                                 std::string& name_store, std::string& description_store) {
  fm_drawing_placement_spec_t ins = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  // The placement keys sit at the top level of the caller's record, so they share its path.
  const Napi::Object options = input.IsObject() ? input.As<Napi::Object>() : Napi::Object::New(input.Env());
  ins.name = reader.String(options, "name", &name_store) ? name_store.c_str() : nullptr;
  ins.description = reader.String(options, "description", &description_store) ? description_store.c_str() : nullptr;
  ins.hidden = reader.Tristate(options, "hidden");
  Napi::Object anchor = Napi::Object::New(input.Env());
  (void)reader.Object(options, "anchor", &anchor);
  ins.anchor = PullDrawingAnchor(reader, anchor);
  return ins;
}

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
  static Napi::Value CreateForProfile(const Napi::CallbackInfo& info);
  static Napi::Value LoadBytes(const Napi::CallbackInfo& info);

  // Cell mutation.
  Napi::Value SetNumber(const Napi::CallbackInfo& info);
  Napi::Value SetBool(const Napi::CallbackInfo& info);
  Napi::Value SetError(const Napi::CallbackInfo& info);
  Napi::Value SetText(const Napi::CallbackInfo& info);
  Napi::Value SetCellPhonetic(const Napi::CallbackInfo& info);
  Napi::Value SetCellPhoneticRuns(const Napi::CallbackInfo& info);
  Napi::Value SetCellPhoneticProperties(const Napi::CallbackInfo& info);
  Napi::Value SetBlank(const Napi::CallbackInfo& info);
  /// The cell's "Show Phonetic" display flag; value writes leave it alone.
  Napi::Value GetCellShowPhonetic(const Napi::CallbackInfo& info);
  Napi::Value SetCellShowPhonetic(const Napi::CallbackInfo& info);
  Napi::Value SetFormula(const Napi::CallbackInfo& info);

  // Cell read.
  Napi::Value GetValue(const Napi::CallbackInfo& info);
  Napi::Value GetCellPhonetic(const Napi::CallbackInfo& info);
  Napi::Value GetCellPhoneticRunCount(const Napi::CallbackInfo& info);
  Napi::Value GetCellPhoneticRun(const Napi::CallbackInfo& info);
  Napi::Value GetCellPhoneticProperties(const Napi::CallbackInfo& info);

  // Ad-hoc, side-effect-free formula evaluation.
  Napi::Value EvaluateFormulaArray(const Napi::CallbackInfo& info);
  Napi::Value EvaluateConditionalFormula(const Napi::CallbackInfo& info);

  // Lifecycle.
  Napi::Value Dispose(const Napi::CallbackInfo& info);
  /// Getter of the read-only `disposed` property: true once the handle is released.
  Napi::Value Disposed(const Napi::CallbackInfo& info);
  Napi::Value MemoryUsage(const Napi::CallbackInfo& info);

  // Lambda text read.
  Napi::Value GetLambdaText(const Napi::CallbackInfo& info);

  // Recalc + save.
  Napi::Value Recalc(const Napi::CallbackInfo& info);
  Napi::Value RecalcParallel(const Napi::CallbackInfo& info);
  Napi::Value PartialRecalc(const Napi::CallbackInfo& info);
  /// `hasDirtyCells()`: true when a change has not been applied by a full recalc yet.
  Napi::Value HasDirtyCells(const Napi::CallbackInfo& info);
  Napi::Value SetIterative(const Napi::CallbackInfo& info);
  /// The stored iterative-calculation settings as `{enabled,
  /// maxIterations, maxChange}`. The cap and threshold are the stored
  /// values even while `enabled` is false, so a host can render Excel's
  /// iterative-calculation dialog without having written them first.
  Napi::Value GetIterative(const Napi::CallbackInfo& info);
  Napi::Value SetIterativeProgress(const Napi::CallbackInfo& info);
  /// `save(format?)`: the workbook's bytes, as XLSX when `format` is omitted.
  Napi::Value Save(const Napi::CallbackInfo& info);
  Napi::Value SaveWithDiagnostics(const Napi::CallbackInfo& info);
  Napi::Value ReadDiagnostics(const Napi::CallbackInfo& info);

  // Workbook-level calc policy / behaviour profile.
  Napi::Value CalcMode(const Napi::CallbackInfo& info);
  Napi::Value SetCalcMode(const Napi::CallbackInfo& info);
  Napi::Value FullPrecision(const Napi::CallbackInfo& info);
  Napi::Value SetFullPrecision(const Napi::CallbackInfo& info);
  Napi::Value CalcOnSave(const Napi::CallbackInfo& info);
  Napi::Value SetCalcOnSave(const Napi::CallbackInfo& info);
  Napi::Value RefMode(const Napi::CallbackInfo& info);
  Napi::Value SetRefMode(const Napi::CallbackInfo& info);
  Napi::Value Date1904(const Napi::CallbackInfo& info);
  Napi::Value SetDate1904(const Napi::CallbackInfo& info);
  /// Clock seam: `getPinnedNow()` returns `{year, ..., second}` or `null`;
  /// `setPinnedNow(y, mo, d, h, mi, s)` and `clearPinnedNow()` return
  /// nothing. See `fm_workbook_set_pinned_now` for the accepted ranges.
  Napi::Value PinnedNow(const Napi::CallbackInfo& info);
  Napi::Value SetPinnedNow(const Napi::CallbackInfo& info);
  /// New workbooks default to win-365-en_US. The getter returns one of the
  /// 28 ids {mac,win}-365-{ja_JP,en_US,de_DE,fr_FR,zh_CN,ko_KR,th_TH,
  /// es_ES,es_MX,pt_BR,ru_RU,zh_TW,it_IT,nl_NL}.
  Napi::Value ProfileId(const Napi::CallbackInfo& info);
  /// SetProfileId accepts the same 28 ids.
  Napi::Value SetProfileId(const Napi::CallbackInfo& info);
  /// The file name `[0]!Name` is spelled with; see `fm_workbook_get_file_name`.
  Napi::Value FileName(const Napi::CallbackInfo& info);
  Napi::Value SetFileName(const Napi::CallbackInfo& info);

  // Dependency-graph trace and dynamic-array spill.
  Napi::Value Precedents(const Napi::CallbackInfo& info);
  Napi::Value Dependents(const Napi::CallbackInfo& info);
  Napi::Value GetSpill(const Napi::CallbackInfo& info);
  Napi::Value ParseReference(const Napi::CallbackInfo& info);
  Napi::Value FormatReference(const Napi::CallbackInfo& info);
  /// `getSheetIndex(name)`: the index of the sheet named `name` (ignoring
  /// case), or `null` when no sheet has that name.
  Napi::Value GetSheetIndex(const Napi::CallbackInfo& info);
  Napi::Value GetArrayFormula(const Napi::CallbackInfo& info);
  Napi::Value SetArrayFormula(const Napi::CallbackInfo& info);

  // External links.
  Napi::Value GetExternalLinkCount(const Napi::CallbackInfo& info);
  Napi::Value GetExternalLink(const Napi::CallbackInfo& info);

  // Sheet operations.
  Napi::Value AddSheet(const Napi::CallbackInfo& info);
  Napi::Value RemoveSheet(const Napi::CallbackInfo& info);
  Napi::Value CopySheet(const Napi::CallbackInfo& info);
  Napi::Value RenameSheet(const Napi::CallbackInfo& info);
  Napi::Value MoveSheet(const Napi::CallbackInfo& info);
  Napi::Value SheetCount(const Napi::CallbackInfo& info);
  /// The active sheet (`activeTab` of the first `<workbookView>`) as an
  /// index, and its setter.
  Napi::Value ActiveSheet(const Napi::CallbackInfo& info);
  Napi::Value SetActiveSheet(const Napi::CallbackInfo& info);
  Napi::Value SheetName(const Napi::CallbackInfo& info);
  Napi::Value Paginate(const Napi::CallbackInfo& info);

  // Print settings. Method names and JS shapes mirror the embind surface
  // exactly; `tools/dev/check_binding_drift.py` fails the build if the
  // two diverge.
  Napi::Value GetSheetPageSetupXml(const Napi::CallbackInfo& info);
  Napi::Value SetSheetPageSetupXml(const Napi::CallbackInfo& info);
  Napi::Value GetSheetPageMarginsXml(const Napi::CallbackInfo& info);
  Napi::Value SetSheetPageMarginsXml(const Napi::CallbackInfo& info);
  Napi::Value GetSheetPrintOptionsXml(const Napi::CallbackInfo& info);
  Napi::Value SetSheetPrintOptionsXml(const Napi::CallbackInfo& info);
  Napi::Value GetSheetHeaderFooterXml(const Napi::CallbackInfo& info);
  Napi::Value SetSheetHeaderFooterXml(const Napi::CallbackInfo& info);
  Napi::Value GetSheetSheetPrXml(const Napi::CallbackInfo& info);
  Napi::Value SetSheetSheetPrXml(const Napi::CallbackInfo& info);
  /// `getTabColor(sheet)` returns the colour spec or `null`;
  /// `setTabColor(sheet, spec | null)` sets or clears it.
  Napi::Value GetSheetTabColor(const Napi::CallbackInfo& info);
  Napi::Value SetSheetTabColor(const Napi::CallbackInfo& info);
  Napi::Value GetSheetPrintArea(const Napi::CallbackInfo& info);
  Napi::Value SetSheetPrintArea(const Napi::CallbackInfo& info);
  Napi::Value GetSheetPrintTitles(const Napi::CallbackInfo& info);
  Napi::Value SetSheetPrintTitles(const Napi::CallbackInfo& info);
  Napi::Value SetRowBreak(const Napi::CallbackInfo& info);
  Napi::Value SetColumnBreak(const Napi::CallbackInfo& info);
  Napi::Value ClearSheetBreaks(const Napi::CallbackInfo& info);
  Napi::Value GetSheetRowBreaks(const Napi::CallbackInfo& info);
  Napi::Value GetSheetColBreaks(const Napi::CallbackInfo& info);
  Napi::Value UpdateSheetPageSetup(const Napi::CallbackInfo& info);
  Napi::Value UpdateSheetPageMargins(const Napi::CallbackInfo& info);
  Napi::Value UpdateSheetPrintOptions(const Napi::CallbackInfo& info);
  Napi::Value UpdateSheetHeaderFooter(const Napi::CallbackInfo& info);
  Napi::Value GetSheetPageSetup(const Napi::CallbackInfo& info);
  Napi::Value GetSheetPageMargins(const Napi::CallbackInfo& info);
  Napi::Value GetSheetPrintOptions(const Napi::CallbackInfo& info);
  Napi::Value GetSheetHeaderFooter(const Napi::CallbackInfo& info);

  // Row / column structural edits.
  Napi::Value InsertRows(const Napi::CallbackInfo& info);
  Napi::Value DeleteRows(const Napi::CallbackInfo& info);
  Napi::Value InsertCols(const Napi::CallbackInfo& info);
  Napi::Value DeleteCols(const Napi::CallbackInfo& info);
  /// Inserts or deletes the cells of a range, shifting the neighbours by a `CellShift`.
  Napi::Value InsertCells(const Napi::CallbackInfo& info);
  Napi::Value DeleteCells(const Napi::CallbackInfo& info);

  // Iteration / metadata.
  Napi::Value CellCount(const Napi::CallbackInfo& info);
  Napi::Value CellAt(const Napi::CallbackInfo& info);
  Napi::Value GetFormula(const Napi::CallbackInfo& info);
  Napi::Value GetFormulaR1C1(const Napi::CallbackInfo& info);
  Napi::Value GetCellsInRange(const Napi::CallbackInfo& info);
  Napi::Value GetDisplayText(const Napi::CallbackInfo& info);
  Napi::Value FormatValue(const Napi::CallbackInfo& info);
  Napi::Value ListDefinedNames(const Napi::CallbackInfo& info);
  Napi::Value GetDefinedName(const Napi::CallbackInfo& info);
  Napi::Value RemoveDefinedName(const Napi::CallbackInfo& info);
  Napi::Value TableCount(const Napi::CallbackInfo& info);
  Napi::Value TableAt(const Napi::CallbackInfo& info);
  Napi::Value TableRuntimeId(const Napi::CallbackInfo& info);
  Napi::Value ResolveTableRuntimeId(const Napi::CallbackInfo& info);
  Napi::Value FindTable(const Napi::CallbackInfo& info);
  Napi::Value GetTableStyleOptions(const Napi::CallbackInfo& info);
  Napi::Value UpdateTableStyleOptions(const Napi::CallbackInfo& info);
  Napi::Value RenameTable(const Napi::CallbackInfo& info);
  Napi::Value CreateTable(const Napi::CallbackInfo& info);
  Napi::Value UpdateTable(const Napi::CallbackInfo& info);
  Napi::Value RemoveTable(const Napi::CallbackInfo& info);
  Napi::Value SetTableColumnTotals(const Napi::CallbackInfo& info);
  Napi::Value RenameTableColumn(const Napi::CallbackInfo& info);
  Napi::Value ListPassthroughs(const Napi::CallbackInfo& info);
  Napi::Value HasPassthrough(const Napi::CallbackInfo& info);
  Napi::Value PivotCount(const Napi::CallbackInfo& info);
  Napi::Value PivotLayout(const Napi::CallbackInfo& info);

  // PivotCache mutation.
  Napi::Value ListPivotCaches(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheCreate(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRemove(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheGetWorksheetSource(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheSetWorksheetSource(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRefresh(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldCount(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldTypes(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldName(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldAdd(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldClear(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldSharedItemCount(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldSharedItemValue(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldAddSharedItemNumber(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldAddSharedItemDate(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldSharedItemIsDate(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldAddSharedItemText(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldAddSharedItemBool(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldAddSharedItemBlank(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldAddSharedItemError(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldClearSharedItems(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordCount(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordAdd(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordClear(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordGetValue(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordSetNumber(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordSetDate(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordSetText(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordSetBool(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordSetBlank(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordSetError(const Napi::CallbackInfo& info);

  // PivotTable mutation.
  Napi::Value PivotCreate(const Napi::CallbackInfo& info);
  Napi::Value PivotRemove(const Napi::CallbackInfo& info);
  Napi::Value PivotRename(const Napi::CallbackInfo& info);
  Napi::Value PivotSetAnchor(const Napi::CallbackInfo& info);
  Napi::Value PivotUpdateGrandTotals(const Napi::CallbackInfo& info);
  Napi::Value PivotUpdateStyleOptions(const Napi::CallbackInfo& info);
  Napi::Value PivotGetLayout(const Napi::CallbackInfo& info);
  Napi::Value PivotSetLayout(const Napi::CallbackInfo& info);
  /// Definition read-back; works for pivots created in-session and loaded
  /// from a package alike. `pivotCacheId` returns the cache id.
  Napi::Value PivotCacheId(const Napi::CallbackInfo& info);
  Napi::Value PivotName(const Napi::CallbackInfo& info);
  Napi::Value PivotRuntimeId(const Napi::CallbackInfo& info);
  Napi::Value ResolvePivotRuntimeId(const Napi::CallbackInfo& info);
  Napi::Value PivotGetGrandTotals(const Napi::CallbackInfo& info);
  Napi::Value PivotGetStyleOptions(const Napi::CallbackInfo& info);
  Napi::Value PivotRowFieldOrder(const Napi::CallbackInfo& info);
  Napi::Value PivotColFieldOrder(const Napi::CallbackInfo& info);
  Napi::Value PivotPageFieldOrder(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldCount(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldAt(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldItemCount(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldItemAt(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldClear(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldSetAxis(const Napi::CallbackInfo& info);
  Napi::Value PivotUpdateField(const Napi::CallbackInfo& info);
  Napi::Value PivotUpdateFieldItem(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldAddItem(const Napi::CallbackInfo& info);
  /// Appends a manual-filter item addressed by its position in the bound
  /// cache field's shared items. The only form that can express the blank
  /// item, which has no label of its own and is matched by its binding.
  Napi::Value PivotFieldClearItems(const Napi::CallbackInfo& info);
  Napi::Value PivotGetOptions(const Napi::CallbackInfo& info);
  Napi::Value PivotUpdateOptions(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheGetOptions(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheUpdateOptions(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheCalculatedFieldAdd(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheCalculatedFieldSetFormula(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheCalculatedFieldRemove(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldFormula(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheCalculatedItemAdd(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheCalculatedItemCount(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheCalculatedItemAt(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheCalculatedItemSetFormula(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheCalculatedItemMove(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheCalculatedItemRemove(const Napi::CallbackInfo& info);
  Napi::Value PivotDataFieldDefaultName(const Napi::CallbackInfo& info);
  Napi::Value PivotProject(const Napi::CallbackInfo& info);
  Napi::Value PivotProjectionPending(const Napi::CallbackInfo& info);
  Napi::Value PivotShowDetails(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldSetSubtotalFunctions(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldSetDateGroup(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldSetNumericGroup(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldGetDateGroup(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldGetNumericGroup(const Napi::CallbackInfo& info);
  Napi::Value PivotSetRowFieldOrder(const Napi::CallbackInfo& info);
  Napi::Value PivotSetColFieldOrder(const Napi::CallbackInfo& info);
  Napi::Value PivotSetPageFieldOrder(const Napi::CallbackInfo& info);
  Napi::Value PivotDataFieldCount(const Napi::CallbackInfo& info);
  Napi::Value PivotDataFieldAdd(const Napi::CallbackInfo& info);
  Napi::Value PivotDataFieldAt(const Napi::CallbackInfo& info);
  Napi::Value PivotDataFieldClear(const Napi::CallbackInfo& info);
  Napi::Value PivotDataFieldSet(const Napi::CallbackInfo& info);
  Napi::Value PivotFilterCount(const Napi::CallbackInfo& info);
  Napi::Value PivotFilterAdd(const Napi::CallbackInfo& info);
  /// Reads the active filter at `filterIndex`.
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

  // Conditional formatting.
  Napi::Value EvaluateCfRange(const Napi::CallbackInfo& info);
  Napi::Value GetCfRuleCount(const Napi::CallbackInfo& info);
  Napi::Value GetCfRule(const Napi::CallbackInfo& info);
  Napi::Value AddConditionalFormat(const Napi::CallbackInfo& info);
  Napi::Value UpdateCfRule(const Napi::CallbackInfo& info);
  Napi::Value RemoveConditionalFormatAt(const Napi::CallbackInfo& info);
  Napi::Value ClearConditionalFormats(const Napi::CallbackInfo& info);

  // Sheet view / layout.
  Napi::Value GetSheetView(const Napi::CallbackInfo& info);
  Napi::Value GetVisibility(const Napi::CallbackInfo& info);
  Napi::Value SetVisibility(const Napi::CallbackInfo& info);
  Napi::Value GetSheetProtection(const Napi::CallbackInfo& info);
  Napi::Value SetSheetProtection(const Napi::CallbackInfo& info);
  Napi::Value UpdateSheetProtection(const Napi::CallbackInfo& info);
  /// Applies the view fields a patch object names: zoom, freeze, split, scroll origin, tab state and display flags.
  Napi::Value UpdateSheetView(const Napi::CallbackInfo& info);
  /// The active pane's selection as `{ activeRow, activeCol, ranges }`,
  /// and its setter taking `{ activeRow, activeCol, ranges }`.
  Napi::Value GetSheetSelection(const Napi::CallbackInfo& info);
  Napi::Value SetSheetSelection(const Napi::CallbackInfo& info);
  Napi::Value GetSheetColumns(const Napi::CallbackInfo& info);
  Napi::Value SetColumnWidth(const Napi::CallbackInfo& info);
  Napi::Value SetColumnHidden(const Napi::CallbackInfo& info);
  Napi::Value SetColumnOutline(const Napi::CallbackInfo& info);
  Napi::Value SetColumnCollapsed(const Napi::CallbackInfo& info);
  Napi::Value SetColumnStyle(const Napi::CallbackInfo& info);
  Napi::Value GetSheetRowOverrides(const Napi::CallbackInfo& info);
  Napi::Value SetRowHeight(const Napi::CallbackInfo& info);
  Napi::Value SetRowAutoHeight(const Napi::CallbackInfo& info);
  Napi::Value ClearRowHeight(const Napi::CallbackInfo& info);
  Napi::Value ClearColumnWidth(const Napi::CallbackInfo& info);
  Napi::Value ClearColumnStyle(const Napi::CallbackInfo& info);
  Napi::Value GetSheetFormatDefaults(const Napi::CallbackInfo& info);
  Napi::Value UpdateSheetFormatDefaults(const Napi::CallbackInfo& info);
  Napi::Value GetCellRectPt(const Napi::CallbackInfo& info);
  Napi::Value GetColumnWidthPt(const Napi::CallbackInfo& info);
  Napi::Value GetRowHeightPt(const Napi::CallbackInfo& info);
  Napi::Value GetWidthModel(const Napi::CallbackInfo& info);
  Napi::Value ColumnCharsToPt(const Napi::CallbackInfo& info);
  Napi::Value ColumnPtToChars(const Napi::CallbackInfo& info);
  Napi::Value SetRowHidden(const Napi::CallbackInfo& info);
  Napi::Value SetRowOutline(const Napi::CallbackInfo& info);
  Napi::Value SetRowCollapsed(const Napi::CallbackInfo& info);
  Napi::Value SetRowStyle(const Napi::CallbackInfo& info);
  Napi::Value ClearRowStyle(const Napi::CallbackInfo& info);
  Napi::Value RangeStyleBases(const Napi::CallbackInfo& info);
  Napi::Value RangeStyleBasesExcluding(const Napi::CallbackInfo& info);
  Napi::Value RemapRangeStyle(const Napi::CallbackInfo& info);
  Napi::Value SnapshotRangeStyle(const Napi::CallbackInfo& info);
  Napi::Value RestoreRangeStyle(const Napi::CallbackInfo& info);

  // Styles.
  Napi::Value GetCellXfIndex(const Napi::CallbackInfo& info);
  Napi::Value SetCellXfIndex(const Napi::CallbackInfo& info);
  Napi::Value ClearCellXfIndex(const Napi::CallbackInfo& info);
  Napi::Value SetRangeXfIndex(const Napi::CallbackInfo& info);
  Napi::Value GetCellXf(const Napi::CallbackInfo& info);
  Napi::Value GetFont(const Napi::CallbackInfo& info);
  Napi::Value GetFill(const Napi::CallbackInfo& info);
  Napi::Value GetBorder(const Napi::CallbackInfo& info);
  Napi::Value GetNumFmt(const Napi::CallbackInfo& info);
  Napi::Value ListNumberFormats(const Napi::CallbackInfo& info);
  Napi::Value GetDxf(const Napi::CallbackInfo& info);
  Napi::Value AddFont(const Napi::CallbackInfo& info);
  Napi::Value SetFont(const Napi::CallbackInfo& info);
  Napi::Value SetFill(const Napi::CallbackInfo& info);
  Napi::Value SetBorder(const Napi::CallbackInfo& info);
  Napi::Value SetCellXf(const Napi::CallbackInfo& info);
  Napi::Value SetDxf(const Napi::CallbackInfo& info);
  Napi::Value SetDefaultFont(const Napi::CallbackInfo& info);
  Napi::Value GetDefaultFont(const Napi::CallbackInfo& info);
  Napi::Value AddFill(const Napi::CallbackInfo& info);
  Napi::Value AddBorder(const Napi::CallbackInfo& info);
  Napi::Value AddNumFmt(const Napi::CallbackInfo& info);
  Napi::Value AddXf(const Napi::CallbackInfo& info);
  Napi::Value AddCellStyleXf(const Napi::CallbackInfo& info);
  Napi::Value SetCellStyleXf(const Napi::CallbackInfo& info);
  Napi::Value AddDxf(const Napi::CallbackInfo& info);
  Napi::Value FontCount(const Napi::CallbackInfo& info);
  Napi::Value FillCount(const Napi::CallbackInfo& info);
  Napi::Value BorderCount(const Napi::CallbackInfo& info);
  Napi::Value XfCount(const Napi::CallbackInfo& info);
  Napi::Value DxfCount(const Napi::CallbackInfo& info);

  // Named cell styles.
  Napi::Value ListCellStyles(const Napi::CallbackInfo& info);
  Napi::Value CellStyleXfCount(const Napi::CallbackInfo& info);
  Napi::Value GetCellStyle(const Napi::CallbackInfo& info);
  Napi::Value GetCellStyleXf(const Napi::CallbackInfo& info);
  Napi::Value SetCellStyle(const Napi::CallbackInfo& info);
  Napi::Value RemoveCellStyle(const Napi::CallbackInfo& info);
  Napi::Value ListTableStyles(const Napi::CallbackInfo& info);
  Napi::Value GetTableStyle(const Napi::CallbackInfo& info);
  Napi::Value SetTableStyle(const Napi::CallbackInfo& info);
  Napi::Value RemoveTableStyle(const Napi::CallbackInfo& info);
  Napi::Value GetTableStyleDefaults(const Napi::CallbackInfo& info);
  Napi::Value SetTableStyleDefaults(const Napi::CallbackInfo& info);

  // Theme, colour resolution, effective style.
  Napi::Value GetThemeColors(const Napi::CallbackInfo& info);
  Napi::Value GetThemeFonts(const Napi::CallbackInfo& info);
  Napi::Value SetThemeColors(const Napi::CallbackInfo& info);
  Napi::Value SetThemeFonts(const Napi::CallbackInfo& info);
  Napi::Value ListThemeScriptFonts(const Napi::CallbackInfo& info);
  Napi::Value SetThemeScriptFont(const Napi::CallbackInfo& info);
  Napi::Value GetThemeScriptFont(const Napi::CallbackInfo& info);
  Napi::Value ResetTheme(const Napi::CallbackInfo& info);
  Napi::Value ResolveColor(const Napi::CallbackInfo& info);
  Napi::Value GetEffectiveStyle(const Napi::CallbackInfo& info);

  // Sheet UI features (merges, comments, hyperlinks, validations).
  Napi::Value AddMerge(const Napi::CallbackInfo& info);
  Napi::Value RemoveMerge(const Napi::CallbackInfo& info);
  Napi::Value RemoveMergeAt(const Napi::CallbackInfo& info);
  Napi::Value ClearMerges(const Napi::CallbackInfo& info);
  Napi::Value GetMerges(const Napi::CallbackInfo& info);
  Napi::Value GetMergesInRange(const Napi::CallbackInfo& info);
  Napi::Value GetComment(const Napi::CallbackInfo& info);
  Napi::Value GetComments(const Napi::CallbackInfo& info);
  Napi::Value SetComment(const Napi::CallbackInfo& info);
  Napi::Value GetCommentVisible(const Napi::CallbackInfo& info);
  Napi::Value SetCommentVisible(const Napi::CallbackInfo& info);
  Napi::Value GetWorkbookProtection(const Napi::CallbackInfo& info);
  Napi::Value SetWorkbookProtection(const Napi::CallbackInfo& info);
  Napi::Value GetProtectedRangeCount(const Napi::CallbackInfo& info);
  Napi::Value GetProtectedRange(const Napi::CallbackInfo& info);
  Napi::Value AddProtectedRange(const Napi::CallbackInfo& info);
  Napi::Value UpdateProtectedRange(const Napi::CallbackInfo& info);
  Napi::Value RemoveProtectedRange(const Napi::CallbackInfo& info);
  Napi::Value SetHyperlink(const Napi::CallbackInfo& info);
  Napi::Value GetHyperlinks(const Napi::CallbackInfo& info);
  Napi::Value GetHyperlink(const Napi::CallbackInfo& info);
  Napi::Value RemoveHyperlinkAt(const Napi::CallbackInfo& info);
  Napi::Value ClearHyperlinks(const Napi::CallbackInfo& info);
  Napi::Value GetValidationCount(const Napi::CallbackInfo& info);
  Napi::Value GetValidation(const Napi::CallbackInfo& info);
  Napi::Value AddValidation(const Napi::CallbackInfo& info);
  Napi::Value UpdateValidation(const Napi::CallbackInfo& info);
  Napi::Value RemoveValidationAt(const Napi::CallbackInfo& info);
  Napi::Value ClearValidations(const Napi::CallbackInfo& info);
  /// Validates `value` against the rule covering a cell, and lists the
  /// populated cells that fail their rule (paged like `getCellsInRange`).
  Napi::Value ValidateValue(const Napi::CallbackInfo& info);
  Napi::Value ListInvalidCells(const Napi::CallbackInfo& info);

  // Raw AutoFilter XML fragment of a sheet.
  Napi::Value GetAutoFilterXml(const Napi::CallbackInfo& info);
  Napi::Value SetAutoFilterXml(const Napi::CallbackInfo& info);

  // Typed AutoFilter, sheet and table scoped.
  Napi::Value GetAutoFilter(const Napi::CallbackInfo& info);
  Napi::Value SetAutoFilter(const Napi::CallbackInfo& info);
  Napi::Value ApplyAutoFilter(const Napi::CallbackInfo& info);
  Napi::Value ClearAutoFilter(const Napi::CallbackInfo& info);
  Napi::Value EvaluateAutoFilter(const Napi::CallbackInfo& info);
  Napi::Value GetTableAutoFilter(const Napi::CallbackInfo& info);
  Napi::Value SetTableAutoFilter(const Napi::CallbackInfo& info);
  Napi::Value ApplyTableAutoFilter(const Napi::CallbackInfo& info);
  Napi::Value ClearTableAutoFilter(const Napi::CallbackInfo& info);
  Napi::Value EvaluateTableAutoFilter(const Napi::CallbackInfo& info);

  // Threaded comments and the workbook person list.
  Napi::Value GetThreadedComments(const Napi::CallbackInfo& info);
  Napi::Value AddThreadedComment(const Napi::CallbackInfo& info);
  Napi::Value EditThreadedComment(const Napi::CallbackInfo& info);
  Napi::Value RemoveThreadedComment(const Napi::CallbackInfo& info);
  Napi::Value GetThreadedComment(const Napi::CallbackInfo& info);
  Napi::Value GetPersons(const Napi::CallbackInfo& info);
  Napi::Value AddPerson(const Napi::CallbackInfo& info);
  Napi::Value RemovePerson(const Napi::CallbackInfo& info);
  Napi::Value GetPerson(const Napi::CallbackInfo& info);

  // Drawing images and the object operations shared with charts.
  Napi::Value ListDrawingObjects(const Napi::CallbackInfo& info);
  Napi::Value GetDrawingObject(const Napi::CallbackInfo& info);
  Napi::Value GetImage(const Napi::CallbackInfo& info);
  Napi::Value InsertImage(const Napi::CallbackInfo& info);
  Napi::Value RemoveDrawing(const Napi::CallbackInfo& info);
  Napi::Value UpdateDrawing(const Napi::CallbackInfo& info);
  Napi::Value SetDrawingZOrder(const Napi::CallbackInfo& info);
  Napi::Value SnapshotDrawing(const Napi::CallbackInfo& info);
  Napi::Value RestoreDrawing(const Napi::CallbackInfo& info);

  // Drawing shapes and sparkline groups.
  Napi::Value InsertShape(const Napi::CallbackInfo& info);
  Napi::Value UpdateShape(const Napi::CallbackInfo& info);
  Napi::Value GetShape(const Napi::CallbackInfo& info);
  Napi::Value GetSlicer(const Napi::CallbackInfo& info);
  Napi::Value AddSlicer(const Napi::CallbackInfo& info);
  Napi::Value UpdateSlicer(const Napi::CallbackInfo& info);
  Napi::Value AddDataTable(const Napi::CallbackInfo& info);
  Napi::Value GetDataTable(const Napi::CallbackInfo& info);
  Napi::Value RemoveDataTable(const Napi::CallbackInfo& info);
  Napi::Value RecalcDataTables(const Napi::CallbackInfo& info);
  Napi::Value GetSparklineGroupCount(const Napi::CallbackInfo& info);
  Napi::Value GetSparklineGroup(const Napi::CallbackInfo& info);
  Napi::Value AddSparklineGroup(const Napi::CallbackInfo& info);
  Napi::Value UpdateSparklineGroup(const Napi::CallbackInfo& info);
  Napi::Value RemoveSparklineGroup(const Napi::CallbackInfo& info);

  // Charts: typed reads, resolved data, creation and typed edits.
  Napi::Value GetChart(const Napi::CallbackInfo& info);
  Napi::Value GetChartData(const Napi::CallbackInfo& info);
  Napi::Value PreviewChartData(const Napi::CallbackInfo& info);
  Napi::Value InsertChart(const Napi::CallbackInfo& info);
  Napi::Value SetChartType(const Napi::CallbackInfo& info);
  Napi::Value SetChartTitle(const Napi::CallbackInfo& info);
  Napi::Value SetChartLegend(const Napi::CallbackInfo& info);
  Napi::Value UpdateChartLegend(const Napi::CallbackInfo& info);
  Napi::Value SetChartAxis(const Napi::CallbackInfo& info);
  Napi::Value SetChartDisplay(const Napi::CallbackInfo& info);
  Napi::Value SetChartSeries(const Napi::CallbackInfo& info);
  Napi::Value SetChartPoint(const Napi::CallbackInfo& info);
  Napi::Value UpdateChartPoint(const Napi::CallbackInfo& info);
  Napi::Value SetChartPlot(const Napi::CallbackInfo& info);
  Napi::Value SetChartFilter(const Napi::CallbackInfo& info);
  Napi::Value AddChartSeries(const Napi::CallbackInfo& info);
  Napi::Value RemoveChartSeries(const Napi::CallbackInfo& info);
  Napi::Value SetChartSource(const Napi::CallbackInfo& info);

  // ---- Argument helpers (visible to all part TUs) -------------------
  //
  // Made public-static so per-area TUs can reach them without a friend
  // declaration per part. JS callers never touch them.
  static uint32_t ArgU32(const Napi::CallbackInfo& info, size_t idx);
  static int32_t ArgI32(const Napi::CallbackInfo& info, size_t idx);
  static double ArgDouble(const Napi::CallbackInfo& info, size_t idx);
  static std::string ArgString(const Napi::CallbackInfo& info, size_t idx);
  static bool ArgBool(const Napi::CallbackInfo& info, size_t idx);
  /// `info[idx]` for a position the registration's `object` mask guarantees.
  static Napi::Object ArgObject(const Napi::CallbackInfo& info, size_t idx);

  /// Message of the refusal every method but `dispose()` throws while this
  /// workbook's iterative progress callback runs; the WASM pre-js guard
  /// throws the same text.
  static constexpr const char* kProgressCallbackReentryMessage =
      "cannot call a Workbook method from its iterative progress callback (kGraphRecalcReentrant, 4005)";

  /// Runs one JS-visible method with this wrapper's active-call count
  /// raised. A getter or callback the method runs may call back into JS;
  /// `dispose()` from there is refused while the count shows the outer call
  /// still using `handle_`, as the WASM wrapper refuses `delete()`. Every
  /// other method is refused while the progress callback runs.
  template <typename Method>
  Napi::Value RunActive(Method method, const Napi::CallbackInfo& info) {
    if (in_iterative_progress_callback_ && method != static_cast<Method>(&Workbook::Dispose)) {
      // The solver holds this workbook mid-sweep: a write would deadlock or
      // change state under it, and a read would see half-computed values.
      return ThrowReentrant(info.Env(), kProgressCallbackReentryMessage);
    }
    ++active_calls_;
    Napi::Value result = (this->*method)(info);
    --active_calls_;
    return result;
  }

  /// Bool argument as the C ABI's `int32_t` flag.
  static int32_t ArgFlag(const Napi::CallbackInfo& info, size_t idx) { return ArgBool(info, idx) ? 1 : 0; }

  /// `info[idx]` converted to the C parameter type `T`: `double` and `int32_t`
  /// read as numbers, `uint8_t` saturates at 255, other integers read as u32.
  template <typename T>
  static T ArgAs(const Napi::CallbackInfo& info, size_t idx) {
    if constexpr (std::is_same_v<T, double>) {
      return ArgDouble(info, idx);
    } else if constexpr (std::is_same_v<T, int32_t>) {
      return ArgI32(info, idx);
    } else if constexpr (std::is_same_v<T, uint8_t>) {
      return static_cast<uint8_t>(std::min(ArgU32(info, idx), 255U));
    } else {
      return static_cast<T>(ArgU32(info, idx));
    }
  }

  /// Runs `call(handle_)` for a method returning nothing: throws the call's
  /// failure, or the disposed-workbook failure when the wrapper no longer
  /// owns a workbook.
  template <typename Call>
  Napi::Value StatusCall(const Napi::CallbackInfo& info, Call&& call) {
    if (handle_ == nullptr) {
      return ThrowInvalidHandle(info.Env());
    }
    return ValueOrThrow(info.Env(), call(handle_), info.Env().Undefined());
  }

  /// `StatusCall` for an add that returns the allocated index: `call(handle, out)`
  /// writes it, and the method returns it as a number.
  template <typename Call>
  Napi::Value IndexCall(const Napi::CallbackInfo& info, Call&& call) {
    Napi::Env env = info.Env();
    if (handle_ == nullptr) {
      return ThrowInvalidHandle(env);
    }
    std::size_t out = 0;
    const fm_status_t rc = call(handle_, &out);
    return ValueOrThrow(env, rc, Napi::Number::New(env, static_cast<double>(out)));
  }

  /// `IndexCall` for an add whose C function writes a `uint32_t` index.
  template <typename Call>
  Napi::Value U32IndexCall(const Napi::CallbackInfo& info, Call&& call) {
    Napi::Env env = info.Env();
    if (handle_ == nullptr) {
      return ThrowInvalidHandle(env);
    }
    uint32_t out = 0;
    const fm_status_t rc = call(handle_, &out);
    return ValueOrThrow(env, rc, Napi::Number::New(env, out));
  }

  /// Entry returning nothing for a C function `(wb, lead, args...)`: `info[0]`
  /// is the leading index (sheet or cache id) and `info[1...]` convert through
  /// `ArgAs` to the parameter types.
  template <typename Lead, typename... Args>
  Napi::Value InvokeSheetCall(const Napi::CallbackInfo& info, fm_status_t (*fn)(fm_workbook_t*, Lead, Args...)) {
    return InvokeSheetCallImpl(info, fn, std::index_sequence_for<Args...>{});
  }

  /// Entry returning nothing for `(sheet, spec:object)`. A missing or
  /// non-object spec throws a TypeError;
  /// `pull(reader, spec)` fills the caller's record, and `apply(handle_)`
  /// runs only if every field read.
  template <typename Pull, typename Apply>
  Napi::Value InvokeSheetSpec(const Napi::CallbackInfo& info, Pull&& pull, Apply&& apply) {
    Napi::Env env = info.Env();
    if (handle_ == nullptr) {
      return ThrowInvalidHandle(env);
    }
    Napi::Object spec;
    if (!ObjectArg(info, 1, &spec)) {
      ThrowArgumentTypeError(env, "", "must be an object");
      return env.Undefined();
    }
    CheckedSpecReader reader(env);
    pull(reader, spec);
    if (!reader.ok()) {
      // A malformed field left a pending JS exception: stop before the C ABI
      // call commits a default value for it.
      return env.Undefined();
    }
    return ValueOrThrow(env, apply(handle_), env.Undefined());
  }

  /// Entry adding a style-pool record and returning its index:
  /// `pull(reader, record)` fills the caller's record from the `info[0]`
  /// object, and `add(handle_, &index)` stores it.
  template <typename Index = std::size_t, typename Pull, typename Add>
  Napi::Value AddStyleRecord(const Napi::CallbackInfo& info, Pull&& pull, Add&& add) {
    Napi::Env env = info.Env();
    if (handle_ == nullptr) {
      return ThrowInvalidHandle(env);
    }
    const Napi::Object record = ArgObject(info, 0);
    CheckedSpecReader reader(env);
    pull(reader, record);
    if (!reader.ok()) {
      // A malformed field left a pending JS exception: stop before the C ABI
      // call commits a default value for it.
      return env.Undefined();
    }
    Index index = 0;
    const fm_status_t rc = add(handle_, &index);
    return ValueOrThrow(env, rc, Napi::Number::New(env, static_cast<double>(index)));
  }

  /// Entry replacing a style-pool record: `pull(reader, record)` fills the
  /// caller's record from the `info[record_arg]` object, and `apply(handle_)`
  /// stores it.
  template <typename Pull, typename Apply>
  Napi::Value SetStyleRecord(const Napi::CallbackInfo& info, size_t record_arg, Pull&& pull, Apply&& apply) {
    Napi::Env env = info.Env();
    if (handle_ == nullptr) {
      return ThrowInvalidHandle(env);
    }
    const Napi::Object record = ArgObject(info, record_arg);
    CheckedSpecReader reader(env);
    pull(reader, record);
    if (!reader.ok()) {
      // A malformed field left a pending JS exception: stop before the C ABI
      // call commits a default value for it.
      return env.Undefined();
    }
    return ValueOrThrow(env, apply(handle_), env.Undefined());
  }

  /// Record getter: `call(handle_, &record)` fills a value-initialised `Rec`,
  /// and `to_js(env, record)` converts it once the call succeeded.
  template <typename Rec, typename Call, typename ToJs>
  Napi::Value RecordGetter(const Napi::CallbackInfo& info, Call&& call, ToJs&& to_js) {
    Napi::Env env = info.Env();
    if (handle_ == nullptr) {
      return ThrowInvalidHandle(env);
    }
    Rec record{};
    const fm_status_t rc = call(handle_, &record);
    return rc == 0 ? Napi::Value(to_js(env, record)) : ThrowStatus(env, rc);
  }

  /// Record result of a getter: `call(handle_)` fills the captured
  /// out-params, and `fill(out)` adds the record's fields once it succeeded.
  template <typename Call, typename Fill>
  Napi::Value StatusObject(const Napi::CallbackInfo& info, Call&& call, Fill&& fill) {
    Napi::Env env = info.Env();
    if (handle_ == nullptr) {
      return ThrowInvalidHandle(env);
    }
    const fm_status_t rc = call(handle_);
    if (rc != 0) {
      return ThrowStatus(env, rc);
    }
    Napi::Object out = Napi::Object::New(env);
    fill(out);
    return out;
  }

  using RowColEditFn = fm_status_t (*)(fm_workbook_t*, size_t, uint32_t, uint32_t);
  Napi::Value InvokeRowColEdit(const Napi::CallbackInfo& info, RowColEditFn fn);

  using PivotFieldOrderFn = fm_status_t (*)(fm_workbook_t*, size_t, size_t, const uint32_t*, size_t);
  Napi::Value InvokePivotFieldOrder(const Napi::CallbackInfo& info, PivotFieldOrderFn fn);

  // Shared bodies for the per-sheet string accessors (the five raw
  // print-settings fragment pairs) and the two break
  // enumerators. The getter returns the string.
  using SheetStringGetFn = fm_status_t (*)(const fm_workbook_t*, size_t, const char**);
  using SheetStringSetFn = fm_status_t (*)(fm_workbook_t*, size_t, const char*);
  Napi::Value SheetStringGetter(const Napi::CallbackInfo& info, SheetStringGetFn getter);
  Napi::Value SheetStringSetter(const Napi::CallbackInfo& info, SheetStringSetFn setter);
  Napi::Value BreaksArray(const Napi::CallbackInfo& info, bool rows);

 private:
  template <typename Lead, typename... Args, size_t... I>
  Napi::Value InvokeSheetCallImpl(const Napi::CallbackInfo& info, fm_status_t (*fn)(fm_workbook_t*, Lead, Args...),
                                  std::index_sequence<I...>) {
    return StatusCall(info, [&](fm_workbook_t* handle) {
      const Lead lead = static_cast<Lead>(ArgU32(info, 0));
      // Braced init evaluates left to right, matching the JS argument order.
      const std::tuple<Args...> args{ArgAs<Args>(info, I + 1)...};
      (void)args;
      return fn(handle, lead, std::get<I>(args)...);
    });
  }

  using CreateFn = fm_status_t (*)(fm_workbook_t**);
  /// Shared body of the `createDefault` / `createEmpty` factories.
  static Napi::Value CreateWith(const Napi::CallbackInfo& info, CreateFn create);
  /// Returns the wrapper of a factory that succeeded; on failure clears the
  /// wrapper's handle and throws the factory's status.
  static Napi::Value FactoryResult(Napi::Env env, Napi::Object jsobj, Workbook* wb, fm_status_t rc);

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
  /// Applies an already measured footprint to V8's external-memory count.
  void ApplyExternalMemoryEstimate(Napi::Env env, size_t bytes);

  /// The JS property of this wrapper holding the iterative progress callback.
  /// Kept on the wrapper rather than in a persistent reference so that a
  /// callback closing over the workbook forms a cycle the collector can see.
  static Napi::Symbol IterativeProgressKey(Napi::Env env);

  fm_workbook_t* handle_ = nullptr;
  /// JS-visible methods of this wrapper currently on the stack.
  uint32_t active_calls_ = 0;
  /// Set while the iterative progress callback runs.
  bool in_iterative_progress_callback_ = false;
  /// Bytes currently attributed to this wrapper in V8's external-memory
  /// accounting; the baseline `SyncExternalMemory` computes its delta
  /// against.
  int64_t reported_external_bytes_ = 0;
};

/// Reads one element of a per-sheet collection by `info[1]`: a number reads by position, a string by id (`null`
/// when no element has it). `param` names the argument in a rejection.
template <typename T>
Napi::Value SheetKeyedItemResult(const Napi::CallbackInfo& info, const fm_workbook_t* handle,
                                 fm_status_t (*at_fn)(const fm_workbook_t*, std::size_t, std::size_t, T*),
                                 fm_status_t (*get_fn)(const fm_workbook_t*, std::size_t, const char*, int32_t*, T*),
                                 Napi::Object (*to_js)(Napi::Env, const T&), const char* param) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const Napi::Value key = info[1];
  if (key.IsString()) {
    const std::string id = Workbook::ArgString(info, 1);
    T entry = InitOut<T>();
    int32_t present = 0;
    const fm_status_t rc = get_fn(handle, Workbook::ArgU32(info, 0), id.c_str(), &present, &entry);
    return OptionalOrNull(env, rc, present, rc == 0 && present != 0 ? Napi::Value(to_js(env, entry)) : env.Null());
  }
  CheckedSpecReader reader(env);
  if (!key.IsNumber()) {
    reader.FailType(param, "a number or a string");
    return env.Undefined();
  }
  const uint32_t index = reader.U32(key, param, 0U);
  if (!reader.ok()) {
    return env.Undefined();
  }
  T entry = InitOut<T>();
  const fm_status_t rc = at_fn(handle, Workbook::ArgU32(info, 0), index, &entry);
  return ValueOrThrow(env, rc, rc == 0 ? Napi::Value(to_js(env, entry)) : env.Undefined());
}

/// Attaches the module functions that need no workbook (function catalog,
/// locale text, collation, image probe) to `exports`. Their positional
/// arguments are checked as a `Workbook` method's are.
void RegisterModuleFunctions(Napi::Env env, Napi::Object exports);

}  // namespace formulon_node

#endif  // FORMULON_NODE_ADDON_PARTS_WORKBOOK_CLASS_H_
