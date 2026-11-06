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

#include <limits>

#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

EMSCRIPTEN_BINDINGS(formulon) {
  using emscripten::class_;
  using emscripten::function;
  using emscripten::value_object;

  // ---- Value-object surface ------------------------------------------------
  value_object<JsParallelRecalcStats>("ParallelRecalcStats")
      .field("cellsEvaluated", &JsParallelRecalcStats::cellsEvaluated)
      .field("sccsProcessed", &JsParallelRecalcStats::sccsProcessed)
      .field("parallelSteps", &JsParallelRecalcStats::parallelSteps)
      .field("serialFallbackSteps", &JsParallelRecalcStats::serialFallbackSteps)
      .field("cycleRecoveries", &JsParallelRecalcStats::cycleRecoveries)
      .field("workerThreadsStarted", &JsParallelRecalcStats::workerThreadsStarted)
      .field("workerThreadsUsed", &JsParallelRecalcStats::workerThreadsUsed);

  value_object<JsSaveDiagnostics>("SaveDiagnostics")
      .field("bytes", &JsSaveDiagnostics::bytes)
      .field("downgradedFormulaCount", &JsSaveDiagnostics::downgradedFormulaCount)
      .field("deferredFeatureCount", &JsSaveDiagnostics::deferredFeatureCount)
      .field("droppedPartCount", &JsSaveDiagnostics::droppedPartCount)
      .field("droppedRelationshipCount", &JsSaveDiagnostics::droppedRelationshipCount)
      .field("renumberedPartCount", &JsSaveDiagnostics::renumberedPartCount);

  value_object<JsReadDiagnostics>("ReadDiagnostics")
      .field("undecodedFormulaCount", &JsReadDiagnostics::undecodedFormulaCount)
      .field("undecodedDefinedNameCount", &JsReadDiagnostics::undecodedDefinedNameCount)
      .field("undecodedPartCount", &JsReadDiagnostics::undecodedPartCount)
      .field("skippedFeatureCount", &JsReadDiagnostics::skippedFeatureCount)
      .field("unknownContentTypeCount", &JsReadDiagnostics::unknownContentTypeCount);

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
      .field("barDirection", &JsCfMatch::barDirection)
      .field("iconSetName", &JsCfMatch::iconSetName)
      .field("iconIndex", &JsCfMatch::iconIndex);

  // ---- Sheet view / layout value-objects -----------------------------------
  value_object<JsSheetView>("SheetView")
      .field("zoomScale", &JsSheetView::zoomScale)
      .field("freezeRows", &JsSheetView::freezeRows)
      .field("freezeCols", &JsSheetView::freezeCols)
      .field("tabHidden", &JsSheetView::tabHidden)
      .field("visibility", &JsSheetView::visibility)
      .field("showGridLines", &JsSheetView::showGridLines)
      .field("showRowColHeaders", &JsSheetView::showRowColHeaders)
      .field("showZeros", &JsSheetView::showZeros)
      .field("rightToLeft", &JsSheetView::rightToLeft)
      .field("tabSelected", &JsSheetView::tabSelected)
      .field("viewMode", &JsSheetView::viewMode)
      .field("topLeftRow", &JsSheetView::topLeftRow)
      .field("topLeftCol", &JsSheetView::topLeftCol)
      .field("paneTopLeftRow", &JsSheetView::paneTopLeftRow)
      .field("paneTopLeftCol", &JsSheetView::paneTopLeftCol)
      .field("splitX", &JsSheetView::splitX)
      .field("splitY", &JsSheetView::splitY)
      .field("showFormulas", &JsSheetView::showFormulas)
      .field("showOutlineSymbols", &JsSheetView::showOutlineSymbols);

  // @size-budget: 14 KB
  // Covers the JsSheetProtection value_object and the two bridge methods
  // (getSheetProtection / setSheetProtection).
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

  // ---- Workbook class ------------------------------------------------------
  class_<JsWorkbook>("Workbook")
      .property("disposed", &JsWorkbook::disposed)
      .class_function("createDefault", &JsWorkbook::createDefault)
      .class_function("createEmpty", &JsWorkbook::createEmpty)
      .class_function("createForProfile", &JsWorkbook::createForProfile)
      .class_function("load", &JsWorkbook::loadBytes)
      .function("addBorder", &JsWorkbook::addBorder)
      .function("addCellStyleXf", &JsWorkbook::addCellStyleXf)
      .function("addCfRule", &JsWorkbook::addConditionalFormat)
      .function("addDxf", &JsWorkbook::addDxf)
      .function("addFill", &JsWorkbook::addFill)
      .function("addFont", &JsWorkbook::addFont)
      .function("setFont", &JsWorkbook::setFont)
      .function("setDefaultFont", &JsWorkbook::setDefaultFont)
      .function("addHyperlink", &JsWorkbook::addHyperlink)
      .function("addHyperlinkRange", &JsWorkbook::addHyperlinkRange)
      .function("addMerge", &JsWorkbook::addMerge)
      .function("addNumFmt", &JsWorkbook::addNumFmt)
      .function("addSheet", &JsWorkbook::addSheet)
      .function("addColumnBreak", &JsWorkbook::addSheetColBreak)
      .function("addRowBreak", &JsWorkbook::addSheetRowBreak)
      .function("addValidation", &JsWorkbook::addValidation)
      .function("addCellXf", &JsWorkbook::addXf)
      .function("getBorderCount", &JsWorkbook::borderCount)
      .function("getCalcMode", &JsWorkbook::calcMode)
      .function("canonicalizeFormula", &JsWorkbook::canonicalizeFormula)
      .function("canonicalizeFunctionName", &JsWorkbook::canonicalizeFunctionName)
      .function("getCellAt", &JsWorkbook::cellAt)
      .function("getCellCount", &JsWorkbook::cellCount)
      .function("listCellStyles", &JsWorkbook::listCellStyles)
      .function("getCellStyleXfCount", &JsWorkbook::cellStyleXfCount)
      .function("clearCfRules", &JsWorkbook::clearConditionalFormats)
      .function("clearColumnStyle", &JsWorkbook::clearColumnStyle)
      .function("clearHyperlinks", &JsWorkbook::clearHyperlinks)
      .function("clearMerges", &JsWorkbook::clearMerges)
      .function("clearPinnedNow", &JsWorkbook::clearPinnedNow)
      .function("clearRowStyle", &JsWorkbook::clearRowStyle)
      .function("clearValidations", &JsWorkbook::clearValidations)
      .function("getCollationClasses", &JsWorkbook::collationClasses)
      .function("compareText", &JsWorkbook::compareText)
      .function("addTable", &JsWorkbook::createTable)
      .function("listDefinedNames", &JsWorkbook::listDefinedNames)
      .function("getDefinedName", &JsWorkbook::getDefinedName)
      .function("removeDefinedName", &JsWorkbook::removeDefinedName)
      .function("deleteColumns", &JsWorkbook::deleteColumns)
      .function("deleteRows", &JsWorkbook::deleteRows)
      .function("getDependents", &JsWorkbook::dependents)
      .function("dispose", &JsWorkbook::dispose)
      .function("getDxfCount", &JsWorkbook::dxfCount)
      .function("evaluateCfRange", &JsWorkbook::evaluateCfRange)
      .function("evaluateCfFormula", &JsWorkbook::evaluateConditionalFormula)
      .function("evaluateFormulaArray", &JsWorkbook::evaluateFormulaArray)
      .function("evaluateFormula", &JsWorkbook::evaluateFormulaText)
      // Profile ids are {mac,win}-365-{ja_JP,en_US,de_DE,fr_FR,zh_CN,ko_KR,
      // th_TH,es_ES,es_MX,pt_BR,ru_RU,zh_TW,it_IT,nl_NL}. New workbooks default to win-365-en_US.
      .function("getExcelProfileId", &JsWorkbook::excelProfileId)
      .function("getFileName", &JsWorkbook::fileName)
      .function("getFillCount", &JsWorkbook::fillCount)
      .function("getFontCount", &JsWorkbook::fontCount)
      .function("getFunctionMetadata", &JsWorkbook::functionMetadata)
      .function("listFunctionNames", &JsWorkbook::functionNames)
      .function("getBorder", &JsWorkbook::getBorder)
      .function("getCellStyle", &JsWorkbook::getCellStyle)
      .function("getCellStyleXf", &JsWorkbook::getCellStyleXf)
      .function("getCellXf", &JsWorkbook::getCellXf)
      .function("getCellsInRange", &JsWorkbook::getCellsInRange)
      .function("getCellsInRange", emscripten::optional_override([](const JsWorkbook& self, uint32_t sheet,
                                                                    emscripten::val range, emscripten::val cursor) {
                  return self.getCellsInRange(sheet, range, cursor, emscripten::val::undefined());
                }))
      .function("getCellsInRange",
                emscripten::optional_override([](const JsWorkbook& self, uint32_t sheet, emscripten::val range) {
                  return self.getCellsInRange(sheet, range, emscripten::val::undefined(), emscripten::val::undefined());
                }))
      .function("getAutoFilter", &JsWorkbook::getAutoFilter)
      .function("setAutoFilter", &JsWorkbook::setAutoFilter)
      .function("removeAutoFilter", &JsWorkbook::removeAutoFilter)
      .function("applyAutoFilter", &JsWorkbook::applyAutoFilter)
      .function("clearAutoFilter", &JsWorkbook::clearAutoFilter)
      .function("evaluateAutoFilter", &JsWorkbook::evaluateAutoFilter)
      .function("getTableAutoFilter", &JsWorkbook::getTableAutoFilter)
      .function("setTableAutoFilter", &JsWorkbook::setTableAutoFilter)
      .function("removeTableAutoFilter", &JsWorkbook::removeTableAutoFilter)
      .function("applyTableAutoFilter", &JsWorkbook::applyTableAutoFilter)
      .function("clearTableAutoFilter", &JsWorkbook::clearTableAutoFilter)
      .function("evaluateTableAutoFilter", &JsWorkbook::evaluateTableAutoFilter)
      .function("validateValue", &JsWorkbook::validateValue)
      .function("getInvalidCells", &JsWorkbook::listInvalidCells)
      .function("getInvalidCells",
                emscripten::optional_override([](const JsWorkbook& self, uint32_t sheet, emscripten::val cursor) {
                  return self.listInvalidCells(sheet, cursor, emscripten::val::undefined());
                }))
      .function("getInvalidCells", emscripten::optional_override([](const JsWorkbook& self, uint32_t sheet) {
                  return self.listInvalidCells(sheet, emscripten::val::undefined(), emscripten::val::undefined());
                }))
      .function("listThreadedComments", &JsWorkbook::getThreadedComments)
      .function("getThreadedComment", &JsWorkbook::getThreadedComment)
      .function("addThreadedComment", &JsWorkbook::addThreadedComment)
      .function("updateThreadedComment", &JsWorkbook::editThreadedComment)
      .function("setThreadResolved", &JsWorkbook::setThreadResolved)
      .function("removeThreadedComment", &JsWorkbook::removeThreadedComment)
      .function("listPeople", &JsWorkbook::getPersons)
      .function("addPerson", &JsWorkbook::addPerson)
      .function("removePerson", &JsWorkbook::removePerson)
      .function("getPerson", &JsWorkbook::getPerson)
      .function("getImageInfo", &JsWorkbook::probeImage)
      .function("listDrawingObjects", &JsWorkbook::listDrawingObjects)
      .function("getDrawingObject", &JsWorkbook::getDrawingObject)
      .function("getImage", &JsWorkbook::getImage)
      .function("addImage", &JsWorkbook::insertImage)
      .function("addImage", emscripten::optional_override([](JsWorkbook& self, uint32_t sheet, emscripten::val bytes) {
                  return self.insertImage(sheet, bytes, emscripten::val::object());
                }))
      .function("removeDrawing", &JsWorkbook::removeDrawing)
      .function("setDrawingAnchor", &JsWorkbook::setDrawingAnchor)
      .function("setDrawingName", &JsWorkbook::setDrawingName)
      .function("setDrawingZOrder", &JsWorkbook::setDrawingZOrder)
      .function("snapshotDrawing", &JsWorkbook::snapshotDrawing)
      .function("restoreDrawing", &JsWorkbook::restoreDrawing)
      .function("restoreDrawing",
                emscripten::optional_override([](JsWorkbook& self, uint32_t sheet, emscripten::val bytes) {
                  return self.restoreDrawing(sheet, bytes, emscripten::val::object());
                }))
      .function("listShapes", &JsWorkbook::listShapes)
      .function("addShape", &JsWorkbook::insertShape)
      .function("updateShape", &JsWorkbook::updateShape)
      .function("removeShape", &JsWorkbook::removeShape)
      .function("getShape", &JsWorkbook::getShape)
      .function("listSlicers", &JsWorkbook::listSlicers)
      .function("getSlicer", &JsWorkbook::getSlicer)
      .function("addSlicer", &JsWorkbook::addSlicer)
      .function("updateSlicer", &JsWorkbook::updateSlicer)
      .function("removeSlicer", &JsWorkbook::removeSlicer)
      .function("addDataTable", &JsWorkbook::addDataTable)
      .function("getDataTable", &JsWorkbook::getDataTable)
      .function("removeDataTable", &JsWorkbook::removeDataTable)
      .function("recalcDataTables", &JsWorkbook::recalcDataTables)
      .function("getSparklineGroupCount", &JsWorkbook::getSparklineGroupCount)
      .function("getSparklineGroup", &JsWorkbook::getSparklineGroup)
      .function("addSparklineGroup", &JsWorkbook::addSparklineGroup)
      .function("updateSparklineGroup", &JsWorkbook::updateSparklineGroup)
      .function("setSparklineGroupSparklines", &JsWorkbook::setSparklineGroupSparklines)
      .function("removeSparklineGroup", &JsWorkbook::removeSparklineGroup)
      .function("getChart", &JsWorkbook::getChart)
      .function("getChartData", &JsWorkbook::getChartData)
      .function("getChartPreview", &JsWorkbook::previewChartData)
      .function("addChart", &JsWorkbook::insertChart)
      .function("setChartType", &JsWorkbook::setChartType)
      .function("setChartType", emscripten::optional_override([](JsWorkbook& self, uint32_t sheet, uint32_t objectId,
                                                                 uint32_t plot, int32_t kind) {
                  return self.setChartType(sheet, objectId, plot, kind, emscripten::val::undefined());
                }))
      .function("setChartTitle", &JsWorkbook::setChartTitle)
      .function("setChartLegend", &JsWorkbook::setChartLegend)
      .function("setChartLegend", emscripten::optional_override([](JsWorkbook& self, uint32_t sheet, uint32_t objectId,
                                                                   int32_t position, bool overlay) {
                  return self.setChartLegend(sheet, objectId, position, overlay, emscripten::val::undefined());
                }))
      .function("setChartAxis", &JsWorkbook::setChartAxis)
      .function("setChartDisplay", &JsWorkbook::setChartDisplay)
      .function("setChartSeries", &JsWorkbook::setChartSeries)
      .function("setChartPoint", &JsWorkbook::setChartPoint)
      .function("setChartPlot", &JsWorkbook::setChartPlot)
      .function("setChartFilter", &JsWorkbook::setChartFilter)
      .function("addChartSeries", &JsWorkbook::addChartSeries)
      .function("removeChartSeries", &JsWorkbook::removeChartSeries)
      .function("setChartSource", &JsWorkbook::setChartSource)
      .function("setChartSource", emscripten::optional_override([](JsWorkbook& self, uint32_t sheet, uint32_t objectId,
                                                                   const std::string& source, int32_t direction) {
                  return self.setChartSource(sheet, objectId, source, direction, emscripten::val::undefined(),
                                             emscripten::val::undefined());
                }))
      .function(
          "setChartSource",
          emscripten::optional_override([](JsWorkbook& self, uint32_t sheet, uint32_t objectId,
                                           const std::string& source, int32_t direction, emscripten::val headerRow) {
            return self.setChartSource(sheet, objectId, source, direction, headerRow, emscripten::val::undefined());
          }))
      .function("getCellRectPt", &JsWorkbook::getCellRectPt)
      .function("getColumnWidthPt", &JsWorkbook::getColumnWidthPt)
      .function("getDisplayText", &JsWorkbook::getDisplayText)
      .function("getFormula", &JsWorkbook::getFormula)
      .function("getFormulaR1C1", &JsWorkbook::getFormulaR1C1)
      .function("getRowHeightPt", &JsWorkbook::getRowHeightPt)
      .function("getFormatDefaults", &JsWorkbook::getSheetFormatDefaults)
      .function("getWidthModel", &JsWorkbook::getWidthModel)
      .function("formatValue", &JsWorkbook::formatValue)
      .function("clearRowHeight", &JsWorkbook::clearRowHeight)
      .function("clearColumnWidth", &JsWorkbook::clearColumnWidth)
      .function("setFormatDefaults", &JsWorkbook::setSheetFormatDefaults)
      .function("convertColumnCharsToPt", &JsWorkbook::columnCharsToPt)
      .function("convertColumnPtToChars", &JsWorkbook::columnPtToChars)
      .function("getCellXfIndex", &JsWorkbook::getCellXfIndex)
      .function("getCellPhonetic", &JsWorkbook::getCellPhonetic)
      .function("getCellPhoneticRunCount", &JsWorkbook::getCellPhoneticRunCount)
      .function("getCellPhoneticRun", &JsWorkbook::getCellPhoneticRun)
      .function("getCellPhoneticProperties", &JsWorkbook::getCellPhoneticProperties)
      .function("getCellShowPhonetic", &JsWorkbook::getCellShowPhonetic)
      .function("getComment", &JsWorkbook::getComment)
      .function("listComments", &JsWorkbook::getComments)
      .function("getCommentVisible", &JsWorkbook::getCommentVisible)
      .function("getWorkbookProtection", &JsWorkbook::getWorkbookProtection)
      .function("getProtectedRangeCount", &JsWorkbook::getProtectedRangeCount)
      .function("getProtectedRange", &JsWorkbook::getProtectedRange)
      .function("getCfRuleCount", &JsWorkbook::getCfRuleCount)
      .function("getCfRule", &JsWorkbook::getCfRule)
      .function("getDxf", &JsWorkbook::getDxf)
      .function("getExternalLinkCount", &JsWorkbook::getExternalLinkCount)
      .function("getExternalLink", &JsWorkbook::getExternalLink)
      .function("getFill", &JsWorkbook::getFill)
      .function("getFont", &JsWorkbook::getFont)
      .function("listHyperlinks", &JsWorkbook::getHyperlinks)
      .function("getHyperlink", &JsWorkbook::getHyperlink)
      .function("getIterative", &JsWorkbook::getIterative)
      .function("getLambdaText", &JsWorkbook::getLambdaText)
      .function("listMerges", &JsWorkbook::getMerges)
      .function("getMergesInRange", &JsWorkbook::getMergesInRange)
      .function("getNumFmtString", &JsWorkbook::getNumFmt)
      .function("listColumnOverrides", &JsWorkbook::getSheetColumns)
      .function("getAutoFilterXml", &JsWorkbook::getSheetAutoFilterXml)
      .function("listColumnBreaks", &JsWorkbook::getSheetColBreaks)
      .function("getHeaderFooterXml", &JsWorkbook::getSheetHeaderFooterXml)
      .function("getPageMargins", &JsWorkbook::getSheetPageMargins)
      .function("getPageMarginsXml", &JsWorkbook::getSheetPageMarginsXml)
      .function("getPageSetup", &JsWorkbook::getSheetPageSetup)
      .function("getPageSetupXml", &JsWorkbook::getSheetPageSetupXml)
      .function("getPrintArea", &JsWorkbook::getSheetPrintArea)
      .function("getPrintOptionsXml", &JsWorkbook::getSheetPrintOptionsXml)
      .function("getPrintTitles", &JsWorkbook::getSheetPrintTitles)
      .function("getSheetProtection", &JsWorkbook::getSheetProtection)
      .function("listRowBreaks", &JsWorkbook::getSheetRowBreaks)
      .function("listRowOverrides", &JsWorkbook::getSheetRowOverrides)
      .function("getSheetPrXml", &JsWorkbook::getSheetSheetPrXml)
      .function("getView", &JsWorkbook::getSheetView)
      .function("getValidationCount", &JsWorkbook::getValidationCount)
      .function("getValidation", &JsWorkbook::getValidation)
      .function("getValue", &JsWorkbook::getValue)
      .function("insertColumns", &JsWorkbook::insertColumns)
      .function("insertRows", &JsWorkbook::insertRows)
      .function("getLocaleFacts", &JsWorkbook::localeFacts)
      .function("listLocaleErrorNames", &JsWorkbook::listLocaleErrorNames)
      .function("localizeFormula", &JsWorkbook::localizeFormula)
      .function("localizeFunctionName", &JsWorkbook::localizeFunctionName)
      .function("moveSheet", &JsWorkbook::moveSheet)
      .function("copySheet", &JsWorkbook::copySheet)
      .function("recalcViewport", &JsWorkbook::partialRecalc)
      .function("getPagination", &JsWorkbook::paginate)
      .function("listPassthroughs", &JsWorkbook::listPassthroughs)
      .function("getPassthrough", &JsWorkbook::getPassthrough)
      .function("getPinnedNow", &JsWorkbook::pinnedNow)
      .function("listPivotCaches", &JsWorkbook::listPivotCaches)
      .function("addPivotCache", &JsWorkbook::pivotCacheCreate)
      .function("addPivotCacheField", &JsWorkbook::pivotCacheFieldAdd)
      .function("addPivotCacheFieldSharedItemBlank", &JsWorkbook::pivotCacheFieldAddSharedItemBlank)
      .function("addPivotCacheFieldSharedItemBool", &JsWorkbook::pivotCacheFieldAddSharedItemBool)
      .function("addPivotCacheFieldSharedItemError", &JsWorkbook::pivotCacheFieldAddSharedItemError)
      .function("addPivotCacheFieldSharedItemNumber", &JsWorkbook::pivotCacheFieldAddSharedItemNumber)
      .function("addPivotCacheFieldSharedItemDate", &JsWorkbook::pivotCacheFieldAddSharedItemDate)
      .function("isPivotCacheFieldSharedItemDate", &JsWorkbook::pivotCacheFieldSharedItemIsDate)
      .function("addPivotCacheFieldSharedItemText", &JsWorkbook::pivotCacheFieldAddSharedItemText)
      .function("clearPivotCacheFields", &JsWorkbook::pivotCacheFieldClear)
      .function("clearPivotCacheFieldSharedItems", &JsWorkbook::pivotCacheFieldClearSharedItems)
      .function("getPivotCacheFieldCount", &JsWorkbook::pivotCacheFieldCount)
      .function("getPivotCacheFieldName", &JsWorkbook::pivotCacheFieldName)
      .function("getPivotCacheFieldSharedItemCount", &JsWorkbook::pivotCacheFieldSharedItemCount)
      .function("getPivotCacheFieldSharedItemValue", &JsWorkbook::pivotCacheFieldSharedItemValue)
      .function("getPivotCacheWorksheetSource", &JsWorkbook::pivotCacheGetWorksheetSource)
      .function("addPivotCacheRecord", &JsWorkbook::pivotCacheRecordAdd)
      .function("clearPivotCacheRecords", &JsWorkbook::pivotCacheRecordClear)
      .function("getPivotCacheRecordCount", &JsWorkbook::pivotCacheRecordCount)
      .function("setPivotCacheRecordBlank", &JsWorkbook::pivotCacheRecordSetBlank)
      .function("setPivotCacheRecordBool", &JsWorkbook::pivotCacheRecordSetBool)
      .function("setPivotCacheRecordError", &JsWorkbook::pivotCacheRecordSetError)
      .function("setPivotCacheRecordNumber", &JsWorkbook::pivotCacheRecordSetNumber)
      .function("setPivotCacheRecordDate", &JsWorkbook::pivotCacheRecordSetDate)
      .function("setPivotCacheRecordText", &JsWorkbook::pivotCacheRecordSetText)
      .function("removePivotCache", &JsWorkbook::pivotCacheRemove)
      .function("setPivotCacheWorksheetSource", &JsWorkbook::pivotCacheSetWorksheetSource)
      .function("getPivotCount", &JsWorkbook::pivotCount)
      .function("addPivot", &JsWorkbook::pivotCreate)
      .function("addPivotDataField", &JsWorkbook::pivotDataFieldAdd)
      .function("getPivotDataField", &JsWorkbook::pivotDataFieldAt)
      .function("clearPivotDataFields", &JsWorkbook::pivotDataFieldClear)
      .function("getPivotDataFieldCount", &JsWorkbook::pivotDataFieldCount)
      .function("setPivotDataField", &JsWorkbook::pivotDataFieldSet)
      .function("addPivotField", &JsWorkbook::pivotFieldAdd)
      .function("addPivotFieldItem", &JsWorkbook::pivotFieldAddItem)
      .function("addPivotFieldSharedItem", &JsWorkbook::pivotFieldAddItemAt)
      .function("addPivotFieldSubtotalFn", &JsWorkbook::pivotFieldAddSubtotalFn)
      .function("clearPivotFields", &JsWorkbook::pivotFieldClear)
      .function("clearPivotFieldDateGroup", &JsWorkbook::pivotFieldClearDateGroup)
      .function("clearPivotFieldItems", &JsWorkbook::pivotFieldClearItems)
      .function("clearPivotFieldSubtotalFns", &JsWorkbook::pivotFieldClearSubtotalFns)
      .function("getPivotFieldCount", &JsWorkbook::pivotFieldCount)
      .function("getPivotField", &JsWorkbook::pivotFieldAt)
      .function("getPivotFieldItemCount", &JsWorkbook::pivotFieldItemCount)
      .function("getPivotFieldItem", &JsWorkbook::pivotFieldItemAt)
      .function("setPivotFieldAxis", &JsWorkbook::pivotFieldSetAxis)
      .function("setPivotFieldDateGroup", &JsWorkbook::pivotFieldSetDateGroup)
      // The trailing three default as on the Node surface: one-day intervals
      // and the field's own first and last date (NaN).
      .function(
          "setPivotFieldDateGroup",
          emscripten::optional_override([](JsWorkbook& self, uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex,
                                           uint32_t periods, uint32_t calendar, uint32_t intervalDays, double start) {
            return self.pivotFieldSetDateGroup(sheet, pivotIndex, fieldIndex, periods, calendar, intervalDays, start,
                                               std::numeric_limits<double>::quiet_NaN());
          }))
      .function(
          "setPivotFieldDateGroup",
          emscripten::optional_override([](JsWorkbook& self, uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex,
                                           uint32_t periods, uint32_t calendar, uint32_t intervalDays) {
            return self.pivotFieldSetDateGroup(sheet, pivotIndex, fieldIndex, periods, calendar, intervalDays,
                                               std::numeric_limits<double>::quiet_NaN(),
                                               std::numeric_limits<double>::quiet_NaN());
          }))
      .function("setPivotFieldDateGroup",
                emscripten::optional_override([](JsWorkbook& self, uint32_t sheet, uint32_t pivotIndex,
                                                 uint32_t fieldIndex, uint32_t periods, uint32_t calendar) {
                  return self.pivotFieldSetDateGroup(sheet, pivotIndex, fieldIndex, periods, calendar, 1U,
                                                     std::numeric_limits<double>::quiet_NaN(),
                                                     std::numeric_limits<double>::quiet_NaN());
                }))
      .function("setPivotFieldItemVisible", &JsWorkbook::pivotFieldSetItemVisible)
      .function("setPivotFieldNumberFormat", &JsWorkbook::pivotFieldSetNumberFormat)
      .function("setPivotFieldSort", &JsWorkbook::pivotFieldSetSort)
      .function("setPivotFieldSubtotalTop", &JsWorkbook::pivotFieldSetSubtotalTop)
      .function("setPivotFieldDefaultSubtotal", &JsWorkbook::pivotFieldSetDefaultSubtotal)
      .function("setPivotFieldInsertBlankRow", &JsWorkbook::pivotFieldSetInsertBlankRow)
      .function("setPivotFieldFillDownLabels", &JsWorkbook::pivotFieldSetFillDownLabels)
      .function("setPivotFieldShowDetail", &JsWorkbook::pivotFieldSetShowDetail)
      .function("setPivotFieldItemShowDetail", &JsWorkbook::pivotFieldSetItemShowDetail)
      .function("setPivotFieldItemCaption", &JsWorkbook::pivotFieldSetItemCaption)
      .function("setPivotFieldCustomName", &JsWorkbook::pivotFieldSetCustomName)
      .function("setPivotFieldShowAll", &JsWorkbook::pivotFieldSetShowAll)
      .function("getPivotOptions", &JsWorkbook::pivotGetOptions)
      .function("setPivotOptions", &JsWorkbook::pivotSetOptions)
      .function("getPivotCacheOptions", &JsWorkbook::pivotCacheGetOptions)
      .function("setPivotCacheOptions", &JsWorkbook::pivotCacheSetOptions)
      .function("addPivotCacheCalculatedField", &JsWorkbook::pivotCacheCalculatedFieldAdd)
      .function("setPivotCacheCalculatedFieldFormula", &JsWorkbook::pivotCacheCalculatedFieldSetFormula)
      .function("removePivotCacheCalculatedField", &JsWorkbook::pivotCacheCalculatedFieldRemove)
      .function("getPivotCacheFieldFormula", &JsWorkbook::pivotCacheFieldFormula)
      .function("addPivotCacheCalculatedItem", &JsWorkbook::pivotCacheCalculatedItemAdd)
      .function("getPivotCacheCalculatedItemCount", &JsWorkbook::pivotCacheCalculatedItemCount)
      .function("getPivotCacheCalculatedItem", &JsWorkbook::pivotCacheCalculatedItemAt)
      .function("setPivotCacheCalculatedItemFormula", &JsWorkbook::pivotCacheCalculatedItemSetFormula)
      .function("movePivotCacheCalculatedItem", &JsWorkbook::pivotCacheCalculatedItemMove)
      .function("removePivotCacheCalculatedItem", &JsWorkbook::pivotCacheCalculatedItemRemove)
      .function("projectPivot", &JsWorkbook::pivotProject)
      .function("isPivotProjectionPending", &JsWorkbook::pivotProjectionPending)
      .function("getPivotDataFieldDefaultName",
                emscripten::optional_override([](const JsWorkbook& self, uint32_t sheet, uint32_t pivotIndex,
                                                 uint32_t fieldIndex, int32_t aggregation, uint32_t excludeDataField) {
                  return self.pivotDataFieldDefaultName(sheet, pivotIndex, fieldIndex, aggregation, excludeDataField);
                }))
      // Without `excludeDataField`, no data field is left out of the numbering.
      .function("getPivotDataFieldDefaultName",
                emscripten::optional_override([](const JsWorkbook& self, uint32_t sheet, uint32_t pivotIndex,
                                                 uint32_t fieldIndex, int32_t aggregation) {
                  return self.pivotDataFieldDefaultName(sheet, pivotIndex, fieldIndex, aggregation,
                                                        std::numeric_limits<std::size_t>::max());
                }))
      .function("showPivotDetails", &JsWorkbook::pivotShowDetails)
      .function("addPivotFilter", &JsWorkbook::pivotFilterAdd)
      .function("getPivotFilter", &JsWorkbook::pivotFilterAt)
      .function("clearPivotFilters", &JsWorkbook::pivotFilterClear)
      .function("getPivotFilterCount", &JsWorkbook::pivotFilterCount)
      .function("removePivotFilter", &JsWorkbook::pivotFilterRemoveAt)
      .function("getPivotLayout", &JsWorkbook::pivotLayout)
      .function("removePivot", &JsWorkbook::pivotRemove)
      .function("setPivotAnchor", &JsWorkbook::pivotSetAnchor)
      .function("setPivotColumnFieldOrder", &JsWorkbook::pivotSetColFieldOrder)
      .function("setPivotGrandTotals", &JsWorkbook::pivotSetGrandTotals)
      .function("setPivotStyleInfo", &JsWorkbook::pivotSetStyleInfo)
      .function("getPivotReportLayout", &JsWorkbook::pivotGetLayout)
      .function("getPivotCacheId", &JsWorkbook::pivotCacheId)
      .function("getPivotName", &JsWorkbook::pivotName)
      .function("getPivotGrandTotals", &JsWorkbook::pivotGetGrandTotals)
      .function("getPivotStyleInfo", &JsWorkbook::pivotGetStyleInfo)
      .function("getPivotRowFieldOrder", &JsWorkbook::pivotRowFieldOrder)
      .function("getPivotColumnFieldOrder", &JsWorkbook::pivotColFieldOrder)
      .function("setPivotReportLayout", &JsWorkbook::pivotSetLayout)
      .function("setPivotName", &JsWorkbook::pivotSetName)
      .function("setPivotRowFieldOrder", &JsWorkbook::pivotSetRowFieldOrder)
      .function("getPrecedents", &JsWorkbook::precedents)
      .function("recalc", &JsWorkbook::recalc)
      .function("recalcParallel", &JsWorkbook::recalcParallel)
      .function("removeCfRule", &JsWorkbook::removeConditionalFormatAt)
      .function("removeHyperlink", &JsWorkbook::removeHyperlink)
      .function("removeHyperlinkAt", &JsWorkbook::removeHyperlinkAt)
      .function("removeMerge", &JsWorkbook::removeMerge)
      .function("removeMergeAt", &JsWorkbook::removeMergeAt)
      .function("removeSheet", &JsWorkbook::removeSheet)
      .function("removeColumnBreak", &JsWorkbook::removeSheetColBreak)
      .function("removeRowBreak", &JsWorkbook::removeSheetRowBreak)
      .function("removeTable", &JsWorkbook::removeTable)
      .function("removeValidation", &JsWorkbook::removeValidationAt)
      .function("renameSheet", &JsWorkbook::renameSheet)
      .function("renameTableColumn", &JsWorkbook::renameTableColumn)
      .function("save", &JsWorkbook::save)
      // Without `format`, the workbook saves as xlsx.
      .function("save", emscripten::optional_override(
                            [](const JsWorkbook& self) { return self.save(emscripten::val::undefined()); }))
      .function("saveWithDiagnostics", &JsWorkbook::saveWithDiagnostics)
      .function("saveWithDiagnostics", emscripten::optional_override([](const JsWorkbook& self) {
                  return self.saveWithDiagnostics(emscripten::val::undefined());
                }))
      .function("getDiagnostics", &JsWorkbook::readDiagnostics)
      .function("setBlank", &JsWorkbook::setBlank)
      .function("setBool", &JsWorkbook::setBool)
      .function("setCalcMode", &JsWorkbook::setCalcMode)
      .function("getFullPrecision", &JsWorkbook::fullPrecision)
      .function("setFullPrecision", &JsWorkbook::setFullPrecision)
      .function("getCalcOnSave", &JsWorkbook::calcOnSave)
      .function("setCalcOnSave", &JsWorkbook::setCalcOnSave)
      .function("getRefMode", &JsWorkbook::refMode)
      .function("setRefMode", &JsWorkbook::setRefMode)
      .function("getDate1904", &JsWorkbook::date1904)
      .function("setDate1904", &JsWorkbook::setDate1904)
      .function("setCellXfIndex", &JsWorkbook::setCellXfIndex)
      .function("clearCellXfIndex", &JsWorkbook::clearCellXfIndex)
      .function("setRangeXfIndex", &JsWorkbook::setRangeXfIndex)
      .function("clearBreaks", &JsWorkbook::clearSheetBreaks)
      .function("setCellStyle", &JsWorkbook::setCellStyle)
      .function("removeCellStyle", &JsWorkbook::removeCellStyle)
      .function("listTableStyles", &JsWorkbook::listTableStyles)
      .function("getTableStyle", &JsWorkbook::getTableStyle)
      .function("setTableStyle", &JsWorkbook::setTableStyle)
      .function("removeTableStyle", &JsWorkbook::removeTableStyle)
      .function("getTableStyleDefaults", &JsWorkbook::getTableStyleDefaults)
      .function("setTableStyleDefaults", &JsWorkbook::setTableStyleDefaults)
      .function("getThemeFonts", &JsWorkbook::getTheme)
      .function("getThemeScriptFont", &JsWorkbook::getThemeScriptFont)
      .function("setThemeColors", &JsWorkbook::setThemeColors)
      .function("setThemeFonts", &JsWorkbook::setThemeFonts)
      .function("setThemeScriptFont", &JsWorkbook::setThemeScriptFont)
      .function("clearTheme", &JsWorkbook::resetTheme)
      .function("resolveColor", &JsWorkbook::resolveColor)
      .function("getEffectiveStyle", &JsWorkbook::getEffectiveStyle)
      .function("setCellPhonetic", &JsWorkbook::setCellPhonetic)
      .function("setCellPhoneticRuns", &JsWorkbook::setCellPhoneticRuns)
      .function("setCellPhoneticProperties", &JsWorkbook::setCellPhoneticProperties)
      .function("setCellShowPhonetic", &JsWorkbook::setCellShowPhonetic)
      .function("setActiveSheet", &JsWorkbook::setActiveSheet)
      .function("getActiveSheet", &JsWorkbook::activeSheet)
      .function("setColumnHidden", &JsWorkbook::setColumnHidden)
      .function("setColumnOutline", &JsWorkbook::setColumnOutline)
      .function("setColumnCollapsed", &JsWorkbook::setColumnCollapsed)
      .function("setColumnStyle", &JsWorkbook::setColumnStyle)
      .function("setColumnWidth", &JsWorkbook::setColumnWidth)
      .function("setComment", &JsWorkbook::setComment)
      .function("setCommentVisible", &JsWorkbook::setCommentVisible)
      .function("setWorkbookProtection", &JsWorkbook::setWorkbookProtection)
      .function("addProtectedRange", &JsWorkbook::addProtectedRange)
      .function("updateProtectedRange", &JsWorkbook::updateProtectedRange)
      .function("removeProtectedRange", &JsWorkbook::removeProtectedRange)
      .function("setDefinedName", &JsWorkbook::setDefinedName)
      .function("setDefinedNameScoped", &JsWorkbook::setDefinedNameScoped)
      .function("setDefinedNameAttrs", &JsWorkbook::setDefinedNameAttrs)
      .function("setError", &JsWorkbook::setError)
      // `setExcelProfileId` accepts the fourteen ids documented at `excelProfileId`.
      .function("setExcelProfileId", &JsWorkbook::setExcelProfileId)
      .function("setFileName", &JsWorkbook::setFileName)
      .function("setFormula", &JsWorkbook::setFormula)
      .function("setIterative", &JsWorkbook::setIterative)
      .function("setIterativeProgress", &JsWorkbook::setIterativeProgress)
      .function("setNumber", &JsWorkbook::setNumber)
      .function("setPinnedNow", &JsWorkbook::setPinnedNow)
      .function("setRowAutoHeight", &JsWorkbook::setRowAutoHeight)
      .function("setRowHeight", &JsWorkbook::setRowHeight)
      .function("setRowHidden", &JsWorkbook::setRowHidden)
      .function("setRowOutline", &JsWorkbook::setRowOutline)
      .function("setRowCollapsed", &JsWorkbook::setRowCollapsed)
      .function("setRowStyle", &JsWorkbook::setRowStyle)
      .function("getRangeStyleBases", &JsWorkbook::rangeStyleBases)
      .function("getRangeStyleBasesExcluding", &JsWorkbook::rangeStyleBasesExcluding)
      .function("remapRangeStyle", &JsWorkbook::remapRangeStyle)
      .function("snapshotRangeStyle", &JsWorkbook::snapshotRangeStyle)
      .function("restoreRangeStyle", &JsWorkbook::restoreRangeStyle)
      .function("setFreeze", &JsWorkbook::setSheetFreeze)
      .function("setAutoFilterXml", &JsWorkbook::setSheetAutoFilterXml)
      .function("setFitToPage", &JsWorkbook::setSheetFitToPage)
      .function("setHeaderFooter", &JsWorkbook::setSheetHeaderFooter)
      .function("setHeaderFooterXml", &JsWorkbook::setSheetHeaderFooterXml)
      .function("setPageMargins", &JsWorkbook::setSheetPageMargins)
      .function("setPageMarginsXml", &JsWorkbook::setSheetPageMarginsXml)
      .function("setPageSetup", &JsWorkbook::setSheetPageSetup)
      .function("setPageSetupXml", &JsWorkbook::setSheetPageSetupXml)
      .function("setPrintArea", &JsWorkbook::setSheetPrintArea)
      .function("setPrintOptions", &JsWorkbook::setSheetPrintOptions)
      .function("setPrintOptionsXml", &JsWorkbook::setSheetPrintOptionsXml)
      .function("setPrintTitles", &JsWorkbook::setSheetPrintTitles)
      .function("setSheetPrXml", &JsWorkbook::setSheetSheetPrXml)
      .function("setSheetProtection", &JsWorkbook::setSheetProtection)
      .function("setRightToLeft", &JsWorkbook::setSheetRightToLeft)
      .function("setShowGridLines", &JsWorkbook::setSheetShowGridLines)
      .function("setShowRowColHeaders", &JsWorkbook::setSheetShowRowColHeaders)
      .function("setShowZeros", &JsWorkbook::setSheetShowZeros)
      .function("setShowFormulas", &JsWorkbook::setSheetShowFormulas)
      .function("setShowOutlineSymbols", &JsWorkbook::setSheetShowOutlineSymbols)
      .function("setTabHidden", &JsWorkbook::setSheetTabHidden)
      .function("setTabSelected", &JsWorkbook::setSheetTabSelected)
      .function("setTopLeft", &JsWorkbook::setSheetTopLeft)
      .function("setPaneTopLeft", &JsWorkbook::setSheetPaneTopLeft)
      .function("setSplit", &JsWorkbook::setSheetSplit)
      .function("getSelection", &JsWorkbook::getSheetSelection)
      .function("setSelection", &JsWorkbook::setSheetSelection)
      .function("setViewMode", &JsWorkbook::setSheetViewMode)
      .function("setVisibility", &JsWorkbook::setSheetVisibility)
      .function("setZoom", &JsWorkbook::setSheetZoom)
      .function("setText", &JsWorkbook::setText)
      .function("getSheetCount", &JsWorkbook::sheetCount)
      .function("getSheetName", &JsWorkbook::sheetName)
      .function("getSheetIndex", &JsWorkbook::getSheetIndex)
      .function("getSpillInfo", &JsWorkbook::spillInfo)
      .function("parseReference", &JsWorkbook::parseReference)
      .function("parseReference", emscripten::optional_override([](const JsWorkbook& self, const std::string& text) {
                  return self.parseReference(text, emscripten::val::undefined());
                }))
      .function("formatReference", &JsWorkbook::formatReference)
      .function("formatReference", emscripten::optional_override([](const JsWorkbook& self, emscripten::val range) {
                  return self.formatReference(range, emscripten::val::undefined());
                }))
      .function("getArrayFormula", &JsWorkbook::getArrayFormula)
      .function("setArrayFormula", &JsWorkbook::setArrayFormula)
      .function("getTable", &JsWorkbook::tableAt)
      .function("getTableCount", &JsWorkbook::tableCount)
      .function("updateTable", &JsWorkbook::updateTable)
      .function("setTableColumnTotals", &JsWorkbook::setTableColumnTotals)
      .function("getCellXfCount", &JsWorkbook::xfCount);

  // ---- Free functions ------------------------------------------------------
  function("evalFormula", &eval_formula);
  function("getVersion", &version_string);
  function("getStatusName", &status_string);
  function("getErrorDisplayName", &error_display_name);
  function("setLogMinLevel", &set_log_min_level);
  function("setLogSink", &set_log_sink);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
