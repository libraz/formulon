//
// Single `EMSCRIPTEN_BINDINGS(formulon)` registration block: every
// value-object, the `Workbook` class chain, and the free-function
// surface land here. The class chain has to stay in one block because
// embind throws "Cannot register type 'Workbook' twice" on a repeated
// `class_<JsWorkbook>("Workbook")` invocation; the upside is that this
// file is the single place to scan when checking what JS sees.
//
// Method bodies (the bulk of the previous monolithic embind.cpp) are
// out-of-class definitions sprinkled across `parts/workbook_*.cpp`
// files. The member-function-pointer arguments below resolve to the
// right TU at link time.

#include <emscripten/bind.h>

#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

EMSCRIPTEN_BINDINGS(formulon) {
  using emscripten::allow_raw_pointers;
  using emscripten::class_;
  using emscripten::function;
  using emscripten::value_object;

  // ---- Value-object surface ------------------------------------------------
  value_object<JsStatus>("Status")
      .field("ok", &JsStatus::ok)
      .field("status", &JsStatus::status)
      .field("message", &JsStatus::message)
      .field("context", &JsStatus::context);

  value_object<JsParallelRecalcStats>("ParallelRecalcStats")
      .field("cellsEvaluated", &JsParallelRecalcStats::cellsEvaluated)
      .field("sccsProcessed", &JsParallelRecalcStats::sccsProcessed)
      .field("parallelSteps", &JsParallelRecalcStats::parallelSteps)
      .field("serialFallbackSteps", &JsParallelRecalcStats::serialFallbackSteps)
      .field("cycleRecoveries", &JsParallelRecalcStats::cycleRecoveries)
      .field("workerThreadsStarted", &JsParallelRecalcStats::workerThreadsStarted)
      .field("workerThreadsUsed", &JsParallelRecalcStats::workerThreadsUsed);

  value_object<JsParallelRecalcResult>("ParallelRecalcResult")
      .field("status", &JsParallelRecalcResult::status)
      .field("stats", &JsParallelRecalcResult::stats);

  value_object<JsValue>("Value")
      .field("kind", &JsValue::kind)
      .field("number", &JsValue::number)
      .field("boolean", &JsValue::boolean)
      .field("text", &JsValue::text)
      .field("errorCode", &JsValue::errorCode);

  value_object<JsCellResult>("CellResult").field("status", &JsCellResult::status).field("value", &JsCellResult::value);

  value_object<JsEvalResult>("EvalResult").field("status", &JsEvalResult::status).field("value", &JsEvalResult::value);

  value_object<JsSaveResult>("SaveResult").field("status", &JsSaveResult::status).field("bytes", &JsSaveResult::bytes);

  value_object<JsSaveDiagnosticsResult>("SaveDiagnosticsResult")
      .field("status", &JsSaveDiagnosticsResult::status)
      .field("bytes", &JsSaveDiagnosticsResult::bytes)
      .field("downgradedFormulaCount", &JsSaveDiagnosticsResult::downgradedFormulaCount)
      .field("deferredFeatureCount", &JsSaveDiagnosticsResult::deferredFeatureCount)
      .field("droppedPartCount", &JsSaveDiagnosticsResult::droppedPartCount)
      .field("droppedRelationshipCount", &JsSaveDiagnosticsResult::droppedRelationshipCount)
      .field("renumberedPartCount", &JsSaveDiagnosticsResult::renumberedPartCount);

  value_object<JsReadDiagnosticsResult>("ReadDiagnosticsResult")
      .field("status", &JsReadDiagnosticsResult::status)
      .field("undecodedFormulaCount", &JsReadDiagnosticsResult::undecodedFormulaCount)
      .field("undecodedDefinedNameCount", &JsReadDiagnosticsResult::undecodedDefinedNameCount)
      .field("undecodedPartCount", &JsReadDiagnosticsResult::undecodedPartCount)
      .field("skippedFeatureCount", &JsReadDiagnosticsResult::skippedFeatureCount)
      .field("unknownContentTypeCount", &JsReadDiagnosticsResult::unknownContentTypeCount);

  value_object<JsStringResult>("StringResult")
      .field("status", &JsStringResult::status)
      .field("value", &JsStringResult::value);

  // ---- Conditional-format value-objects ------------------------------------
  value_object<JsCfColor>("CfColor")
      .field("r", &JsCfColor::r)
      .field("g", &JsCfColor::g)
      .field("b", &JsCfColor::b)
      .field("a", &JsCfColor::a);

  value_object<JsCfMatch>("CfMatch")
      .field("kind", &JsCfMatch::kind)
      .field("priority", &JsCfMatch::priority)
      .field("dxfIdEngaged", &JsCfMatch::dxfIdEngaged)
      .field("dxfId", &JsCfMatch::dxfId)
      .field("color", &JsCfMatch::color)
      .field("barLengthPct", &JsCfMatch::barLengthPct)
      .field("barAxisPositionPct", &JsCfMatch::barAxisPositionPct)
      .field("barIsNegative", &JsCfMatch::barIsNegative)
      .field("barFill", &JsCfMatch::barFill)
      .field("barBorderEngaged", &JsCfMatch::barBorderEngaged)
      .field("barBorder", &JsCfMatch::barBorder)
      .field("barGradient", &JsCfMatch::barGradient)
      .field("iconSetName", &JsCfMatch::iconSetName)
      .field("iconIndex", &JsCfMatch::iconIndex);

  // ---- Sheet view / layout value-objects -----------------------------------
  value_object<JsSheetView>("SheetView")
      .field("zoomScale", &JsSheetView::zoomScale)
      .field("freezeRows", &JsSheetView::freezeRows)
      .field("freezeCols", &JsSheetView::freezeCols)
      .field("tabHidden", &JsSheetView::tabHidden)
      .field("showGridLines", &JsSheetView::showGridLines)
      .field("showRowColHeaders", &JsSheetView::showRowColHeaders)
      .field("showZeros", &JsSheetView::showZeros)
      .field("rightToLeft", &JsSheetView::rightToLeft)
      .field("tabSelected", &JsSheetView::tabSelected)
      .field("viewMode", &JsSheetView::viewMode);

  value_object<JsSheetViewResult>("SheetViewResult")
      .field("status", &JsSheetViewResult::status)
      .field("view", &JsSheetViewResult::view);

  // @size-budget: 14 KB
  // Covers the JsSheetProtection value_object, JsSheetProtectionResult,
  // and the two bridge methods (getSheetProtection / setSheetProtection).
  value_object<JsSheetProtection>("SheetProtection")
      .field("enabled", &JsSheetProtection::enabled)
      .field("algorithmName", &JsSheetProtection::algorithmName)
      .field("hashValue", &JsSheetProtection::hashValue)
      .field("saltValue", &JsSheetProtection::saltValue)
      .field("spinCount", &JsSheetProtection::spinCount)
      .field("legacyPassword", &JsSheetProtection::legacyPassword)
      .field("sheet", &JsSheetProtection::sheet)
      .field("objects", &JsSheetProtection::objects)
      .field("scenarios", &JsSheetProtection::scenarios)
      .field("formatCells", &JsSheetProtection::formatCells)
      .field("formatColumns", &JsSheetProtection::formatColumns)
      .field("formatRows", &JsSheetProtection::formatRows)
      .field("insertColumns", &JsSheetProtection::insertColumns)
      .field("insertRows", &JsSheetProtection::insertRows)
      .field("insertHyperlinks", &JsSheetProtection::insertHyperlinks)
      .field("deleteColumns", &JsSheetProtection::deleteColumns)
      .field("deleteRows", &JsSheetProtection::deleteRows)
      .field("selectLockedCells", &JsSheetProtection::selectLockedCells)
      .field("selectUnlockedCells", &JsSheetProtection::selectUnlockedCells)
      .field("sort", &JsSheetProtection::sort)
      .field("autoFilter", &JsSheetProtection::autoFilter)
      .field("pivotTables", &JsSheetProtection::pivotTables);

  value_object<JsSheetProtectionResult>("SheetProtectionResult")
      .field("status", &JsSheetProtectionResult::status)
      .field("protection", &JsSheetProtectionResult::protection);

  value_object<JsAddStyleResult>("AddStyleResult")
      .field("status", &JsAddStyleResult::status)
      .field("index", &JsAddStyleResult::index);

  value_object<JsAddNumFmtResult>("AddNumFmtResult")
      .field("status", &JsAddNumFmtResult::status)
      .field("numFmtId", &JsAddNumFmtResult::numFmtId);

  // ---- Workbook class ------------------------------------------------------
  class_<JsWorkbook>("Workbook")
      .class_function("createDefault", &JsWorkbook::createDefault, allow_raw_pointers())
      .class_function("createEmpty", &JsWorkbook::createEmpty, allow_raw_pointers())
      .class_function("loadBytes", &JsWorkbook::loadBytes, allow_raw_pointers())
      .function("addBorder", &JsWorkbook::addBorder)
      .function("addCellStyleXf", &JsWorkbook::addCellStyleXf)
      .function("addConditionalFormat", &JsWorkbook::addConditionalFormat)
      .function("addDxf", &JsWorkbook::addDxf)
      .function("addFill", &JsWorkbook::addFill)
      .function("addFont", &JsWorkbook::addFont)
      .function("addHyperlink", &JsWorkbook::addHyperlink)
      .function("addHyperlinkRange", &JsWorkbook::addHyperlinkRange)
      .function("addMerge", &JsWorkbook::addMerge)
      .function("addNumFmt", &JsWorkbook::addNumFmt)
      .function("addSheet", &JsWorkbook::addSheet)
      .function("addSheetColBreak", &JsWorkbook::addSheetColBreak)
      .function("addSheetRowBreak", &JsWorkbook::addSheetRowBreak)
      .function("addValidation", &JsWorkbook::addValidation)
      .function("addXf", &JsWorkbook::addXf)
      .function("borderCount", &JsWorkbook::borderCount)
      .function("calcMode", &JsWorkbook::calcMode)
      .function("canonicalizeFunctionName", &JsWorkbook::canonicalizeFunctionName)
      .function("cellAt", &JsWorkbook::cellAt)
      .function("cellCount", &JsWorkbook::cellCount)
      .function("cellStyleCount", &JsWorkbook::cellStyleCount)
      .function("cellStyleXfCount", &JsWorkbook::cellStyleXfCount)
      .function("clearConditionalFormats", &JsWorkbook::clearConditionalFormats)
      .function("clearHyperlinks", &JsWorkbook::clearHyperlinks)
      .function("clearMerges", &JsWorkbook::clearMerges)
      .function("clearPinnedNow", &JsWorkbook::clearPinnedNow)
      .function("clearValidations", &JsWorkbook::clearValidations)
      .function("createTable", &JsWorkbook::createTable)
      .function("definedNameAt", &JsWorkbook::definedNameAt)
      .function("definedNameCount", &JsWorkbook::definedNameCount)
      .function("deleteCols", &JsWorkbook::deleteCols)
      .function("deleteRows", &JsWorkbook::deleteRows)
      .function("dependents", &JsWorkbook::dependents)
      .function("dxfCount", &JsWorkbook::dxfCount)
      .function("evaluateCfRange", &JsWorkbook::evaluateCfRange)
      .function("evaluateConditionalFormula", &JsWorkbook::evaluateConditionalFormula)
      .function("evaluateFormulaArray", &JsWorkbook::evaluateFormulaArray)
      .function("evaluateFormulaText", &JsWorkbook::evaluateFormulaText)
      .function("excelProfileId", &JsWorkbook::excelProfileId)
      .function("fillCount", &JsWorkbook::fillCount)
      .function("fontCount", &JsWorkbook::fontCount)
      .function("functionMetadata", &JsWorkbook::functionMetadata)
      .function("functionNames", &JsWorkbook::functionNames)
      .function("getBorder", &JsWorkbook::getBorder)
      .function("getCellStyle", &JsWorkbook::getCellStyle)
      .function("getCellStyleXf", &JsWorkbook::getCellStyleXf)
      .function("getCellXf", &JsWorkbook::getCellXf)
      .function("getCellXfIndex", &JsWorkbook::getCellXfIndex)
      .function("getCellPhonetic", &JsWorkbook::getCellPhonetic)
      .function("getComment", &JsWorkbook::getComment)
      .function("getCommentResult", &JsWorkbook::getCommentResult)
      .function("getComments", &JsWorkbook::getComments)
      .function("getConditionalFormats", &JsWorkbook::getConditionalFormats)
      .function("getDxf", &JsWorkbook::getDxf)
      .function("getExternalLinks", &JsWorkbook::getExternalLinks)
      .function("getFill", &JsWorkbook::getFill)
      .function("getFont", &JsWorkbook::getFont)
      .function("getHyperlinks", &JsWorkbook::getHyperlinks)
      .function("getLambdaText", &JsWorkbook::getLambdaText)
      .function("getMerges", &JsWorkbook::getMerges)
      .function("getNumFmt", &JsWorkbook::getNumFmt)
      .function("getSheetColumns", &JsWorkbook::getSheetColumns)
      .function("getSheetAutoFilterXml", &JsWorkbook::getSheetAutoFilterXml)
      .function("getSheetColBreaks", &JsWorkbook::getSheetColBreaks)
      .function("getSheetHeaderFooterXml", &JsWorkbook::getSheetHeaderFooterXml)
      .function("getSheetPageMargins", &JsWorkbook::getSheetPageMargins)
      .function("getSheetPageMarginsXml", &JsWorkbook::getSheetPageMarginsXml)
      .function("getSheetPageSetup", &JsWorkbook::getSheetPageSetup)
      .function("getSheetPageSetupXml", &JsWorkbook::getSheetPageSetupXml)
      .function("getSheetPrintArea", &JsWorkbook::getSheetPrintArea)
      .function("getSheetPrintOptionsXml", &JsWorkbook::getSheetPrintOptionsXml)
      .function("getSheetPrintTitles", &JsWorkbook::getSheetPrintTitles)
      .function("getSheetProtection", &JsWorkbook::getSheetProtection)
      .function("getSheetRowBreaks", &JsWorkbook::getSheetRowBreaks)
      .function("getSheetRowOverrides", &JsWorkbook::getSheetRowOverrides)
      .function("getSheetSheetPrXml", &JsWorkbook::getSheetSheetPrXml)
      .function("getSheetView", &JsWorkbook::getSheetView)
      .function("getValidations", &JsWorkbook::getValidations)
      .function("getValue", &JsWorkbook::getValue)
      .function("insertCols", &JsWorkbook::insertCols)
      .function("insertRows", &JsWorkbook::insertRows)
      .function("isValid", &JsWorkbook::isValid)
      .function("localizeFunctionName", &JsWorkbook::localizeFunctionName)
      .function("moveSheet", &JsWorkbook::moveSheet)
      .function("partialRecalc", &JsWorkbook::partialRecalc)
      .function("paginate", &JsWorkbook::paginate)
      .function("passthroughAt", &JsWorkbook::passthroughAt)
      .function("passthroughCount", &JsWorkbook::passthroughCount)
      .function("pinnedNow", &JsWorkbook::pinnedNow)
      .function("pivotCacheCount", &JsWorkbook::pivotCacheCount)
      .function("pivotCacheCreate", &JsWorkbook::pivotCacheCreate)
      .function("pivotCacheFieldAdd", &JsWorkbook::pivotCacheFieldAdd)
      .function("pivotCacheFieldAddSharedItemBlank", &JsWorkbook::pivotCacheFieldAddSharedItemBlank)
      .function("pivotCacheFieldAddSharedItemBool", &JsWorkbook::pivotCacheFieldAddSharedItemBool)
      .function("pivotCacheFieldAddSharedItemError", &JsWorkbook::pivotCacheFieldAddSharedItemError)
      .function("pivotCacheFieldAddSharedItemNumber", &JsWorkbook::pivotCacheFieldAddSharedItemNumber)
      .function("pivotCacheFieldAddSharedItemText", &JsWorkbook::pivotCacheFieldAddSharedItemText)
      .function("pivotCacheFieldClear", &JsWorkbook::pivotCacheFieldClear)
      .function("pivotCacheFieldClearSharedItems", &JsWorkbook::pivotCacheFieldClearSharedItems)
      .function("pivotCacheFieldCount", &JsWorkbook::pivotCacheFieldCount)
      .function("pivotCacheFieldName", &JsWorkbook::pivotCacheFieldName)
      .function("pivotCacheFieldSharedItemCount", &JsWorkbook::pivotCacheFieldSharedItemCount)
      .function("pivotCacheIdAt", &JsWorkbook::pivotCacheIdAt)
      .function("pivotCacheGetWorksheetSource", &JsWorkbook::pivotCacheGetWorksheetSource)
      .function("pivotCacheRecordAdd", &JsWorkbook::pivotCacheRecordAdd)
      .function("pivotCacheRecordClear", &JsWorkbook::pivotCacheRecordClear)
      .function("pivotCacheRecordCount", &JsWorkbook::pivotCacheRecordCount)
      .function("pivotCacheRecordSetBlank", &JsWorkbook::pivotCacheRecordSetBlank)
      .function("pivotCacheRecordSetBool", &JsWorkbook::pivotCacheRecordSetBool)
      .function("pivotCacheRecordSetError", &JsWorkbook::pivotCacheRecordSetError)
      .function("pivotCacheRecordSetNumber", &JsWorkbook::pivotCacheRecordSetNumber)
      .function("pivotCacheRecordSetText", &JsWorkbook::pivotCacheRecordSetText)
      .function("pivotCacheRemove", &JsWorkbook::pivotCacheRemove)
      .function("pivotCacheSetWorksheetSource", &JsWorkbook::pivotCacheSetWorksheetSource)
      .function("pivotCount", &JsWorkbook::pivotCount)
      .function("pivotCreate", &JsWorkbook::pivotCreate)
      .function("pivotDataFieldAdd", &JsWorkbook::pivotDataFieldAdd)
      .function("pivotDataFieldClear", &JsWorkbook::pivotDataFieldClear)
      .function("pivotDataFieldCount", &JsWorkbook::pivotDataFieldCount)
      .function("pivotDataFieldSet", &JsWorkbook::pivotDataFieldSet)
      .function("pivotFieldAdd", &JsWorkbook::pivotFieldAdd)
      .function("pivotFieldAddAggregation", &JsWorkbook::pivotFieldAddAggregation)
      .function("pivotFieldAddItem", &JsWorkbook::pivotFieldAddItem)
      .function("pivotFieldAddSubtotalFn", &JsWorkbook::pivotFieldAddSubtotalFn)
      .function("pivotFieldClear", &JsWorkbook::pivotFieldClear)
      .function("pivotFieldClearAggregations", &JsWorkbook::pivotFieldClearAggregations)
      .function("pivotFieldClearDateGroup", &JsWorkbook::pivotFieldClearDateGroup)
      .function("pivotFieldClearItems", &JsWorkbook::pivotFieldClearItems)
      .function("pivotFieldClearSubtotalFns", &JsWorkbook::pivotFieldClearSubtotalFns)
      .function("pivotFieldCount", &JsWorkbook::pivotFieldCount)
      .function("pivotFieldSetAxis", &JsWorkbook::pivotFieldSetAxis)
      .function("pivotFieldSetDateGroup", &JsWorkbook::pivotFieldSetDateGroup)
      .function("pivotFieldSetItemVisible", &JsWorkbook::pivotFieldSetItemVisible)
      .function("pivotFieldSetNumberFormat", &JsWorkbook::pivotFieldSetNumberFormat)
      .function("pivotFieldSetSort", &JsWorkbook::pivotFieldSetSort)
      .function("pivotFieldSetSubtotalTop", &JsWorkbook::pivotFieldSetSubtotalTop)
      .function("pivotFilterAdd", &JsWorkbook::pivotFilterAdd)
      .function("pivotFilterAt", &JsWorkbook::pivotFilterAt)
      .function("pivotFilterClear", &JsWorkbook::pivotFilterClear)
      .function("pivotFilterCount", &JsWorkbook::pivotFilterCount)
      .function("pivotFilterRemoveAt", &JsWorkbook::pivotFilterRemoveAt)
      .function("pivotLayout", &JsWorkbook::pivotLayout)
      .function("pivotRemove", &JsWorkbook::pivotRemove)
      .function("pivotSetAnchor", &JsWorkbook::pivotSetAnchor)
      .function("pivotSetColFieldOrder", &JsWorkbook::pivotSetColFieldOrder)
      .function("pivotSetGrandTotals", &JsWorkbook::pivotSetGrandTotals)
      .function("pivotGetLayout", &JsWorkbook::pivotGetLayout)
      .function("pivotSetLayout", &JsWorkbook::pivotSetLayout)
      .function("pivotSetName", &JsWorkbook::pivotSetName)
      .function("pivotSetRowFieldOrder", &JsWorkbook::pivotSetRowFieldOrder)
      .function("precedents", &JsWorkbook::precedents)
      .function("recalc", &JsWorkbook::recalc)
      .function("recalcParallel", &JsWorkbook::recalcParallel)
      .function("removeConditionalFormatAt", &JsWorkbook::removeConditionalFormatAt)
      .function("removeHyperlink", &JsWorkbook::removeHyperlink)
      .function("removeHyperlinkAt", &JsWorkbook::removeHyperlinkAt)
      .function("removeMerge", &JsWorkbook::removeMerge)
      .function("removeMergeAt", &JsWorkbook::removeMergeAt)
      .function("removeSheet", &JsWorkbook::removeSheet)
      .function("removeSheetColBreak", &JsWorkbook::removeSheetColBreak)
      .function("removeSheetRowBreak", &JsWorkbook::removeSheetRowBreak)
      .function("removeTable", &JsWorkbook::removeTable)
      .function("removeValidationAt", &JsWorkbook::removeValidationAt)
      .function("renameSheet", &JsWorkbook::renameSheet)
      .function("save", &JsWorkbook::save)
      .function("saveAs", &JsWorkbook::saveAs)
      .function("saveWithDiagnostics", &JsWorkbook::saveWithDiagnostics)
      .function("readDiagnostics", &JsWorkbook::readDiagnostics)
      .function("setBlank", &JsWorkbook::setBlank)
      .function("setBool", &JsWorkbook::setBool)
      .function("setCalcMode", &JsWorkbook::setCalcMode)
      .function("setCellXfIndex", &JsWorkbook::setCellXfIndex)
      .function("setRangeXfIndex", &JsWorkbook::setRangeXfIndex)
      .function("clearSheetBreaks", &JsWorkbook::clearSheetBreaks)
      .function("setCellStyle", &JsWorkbook::setCellStyle)
      .function("setCellPhonetic", &JsWorkbook::setCellPhonetic)
      .function("setColumnHidden", &JsWorkbook::setColumnHidden)
      .function("setColumnOutline", &JsWorkbook::setColumnOutline)
      .function("setColumnWidth", &JsWorkbook::setColumnWidth)
      .function("setComment", &JsWorkbook::setComment)
      .function("setDefinedName", &JsWorkbook::setDefinedName)
      .function("setDefinedNameScoped", &JsWorkbook::setDefinedNameScoped)
      .function("setError", &JsWorkbook::setError)
      .function("setExcelProfileId", &JsWorkbook::setExcelProfileId)
      .function("setFormula", &JsWorkbook::setFormula)
      .function("setIterative", &JsWorkbook::setIterative)
      .function("setIterativeProgress", &JsWorkbook::setIterativeProgress)
      .function("setNumber", &JsWorkbook::setNumber)
      .function("setPinnedNow", &JsWorkbook::setPinnedNow)
      .function("setRowHeight", &JsWorkbook::setRowHeight)
      .function("setRowHidden", &JsWorkbook::setRowHidden)
      .function("setRowOutline", &JsWorkbook::setRowOutline)
      .function("setSheetFreeze", &JsWorkbook::setSheetFreeze)
      .function("setSheetAutoFilterXml", &JsWorkbook::setSheetAutoFilterXml)
      .function("setSheetFitToPage", &JsWorkbook::setSheetFitToPage)
      .function("setSheetHeaderFooter", &JsWorkbook::setSheetHeaderFooter)
      .function("setSheetHeaderFooterXml", &JsWorkbook::setSheetHeaderFooterXml)
      .function("setSheetPageMargins", &JsWorkbook::setSheetPageMargins)
      .function("setSheetPageMarginsXml", &JsWorkbook::setSheetPageMarginsXml)
      .function("setSheetPageSetup", &JsWorkbook::setSheetPageSetup)
      .function("setSheetPageSetupXml", &JsWorkbook::setSheetPageSetupXml)
      .function("setSheetPrintArea", &JsWorkbook::setSheetPrintArea)
      .function("setSheetPrintOptions", &JsWorkbook::setSheetPrintOptions)
      .function("setSheetPrintOptionsXml", &JsWorkbook::setSheetPrintOptionsXml)
      .function("setSheetPrintTitles", &JsWorkbook::setSheetPrintTitles)
      .function("setSheetSheetPrXml", &JsWorkbook::setSheetSheetPrXml)
      .function("setSheetProtection", &JsWorkbook::setSheetProtection)
      .function("setSheetRightToLeft", &JsWorkbook::setSheetRightToLeft)
      .function("setSheetShowGridLines", &JsWorkbook::setSheetShowGridLines)
      .function("setSheetShowRowColHeaders", &JsWorkbook::setSheetShowRowColHeaders)
      .function("setSheetShowZeros", &JsWorkbook::setSheetShowZeros)
      .function("setSheetTabHidden", &JsWorkbook::setSheetTabHidden)
      .function("setSheetTabSelected", &JsWorkbook::setSheetTabSelected)
      .function("setSheetViewMode", &JsWorkbook::setSheetViewMode)
      .function("setSheetZoom", &JsWorkbook::setSheetZoom)
      .function("setText", &JsWorkbook::setText)
      .function("sheetCount", &JsWorkbook::sheetCount)
      .function("sheetName", &JsWorkbook::sheetName)
      .function("spillInfo", &JsWorkbook::spillInfo)
      .function("tableAt", &JsWorkbook::tableAt)
      .function("tableCount", &JsWorkbook::tableCount)
      .function("updateTable", &JsWorkbook::updateTable)
      .function("xfCount", &JsWorkbook::xfCount);

  // ---- Free functions ------------------------------------------------------
  function("evalFormula", &eval_formula);
  function("versionString", &version_string);
  function("version", &version_string);
  function("statusString", &status_string);
  function("errorDisplayName", &error_display_name);
  function("lastErrorMessage", &last_error_message);
  function("lastErrorContext", &last_error_context);
  function("setLogMinLevel", &set_log_min_level);
  function("setLogSink", &set_log_sink);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
