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

  // ---- Sheet view / layout value-objects -----------------------------------
  value_object<JsSheetView>("SheetView")
      .field("zoomPct", &JsSheetView::zoomPct)
      .field("freezeRows", &JsSheetView::freezeRows)
      .field("freezeCols", &JsSheetView::freezeCols)
      .field("showGridLines", &JsSheetView::showGridLines)
      .field("showRowColumnHeaders", &JsSheetView::showRowColumnHeaders)
      .field("showZeros", &JsSheetView::showZeros)
      .field("rightToLeft", &JsSheetView::rightToLeft)
      .field("tabSelected", &JsSheetView::tabSelected)
      .field("viewMode", &JsSheetView::viewMode)
      .field("topLeftRow", &JsSheetView::topLeftRow)
      .field("topLeftCol", &JsSheetView::topLeftCol)
      .field("paneTopLeftRow", &JsSheetView::paneTopLeftRow)
      .field("paneTopLeftCol", &JsSheetView::paneTopLeftCol)
      .field("splitXTwips", &JsSheetView::splitXTwips)
      .field("splitYTwips", &JsSheetView::splitYTwips)
      .field("showFormulas", &JsSheetView::showFormulas)
      .field("showOutlineSymbols", &JsSheetView::showOutlineSymbols);

  // @size-budget: 14 KB
  // Covers the sheet protection bridge methods (getSheetProtection /
  // setSheetProtection / updateSheetProtection).
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
      .function("setFill", &JsWorkbook::setFill)
      .function("setBorder", &JsWorkbook::setBorder)
      .function("setCellXf", &JsWorkbook::setXf)
      .function("setDxf", &JsWorkbook::setDxf)
      .function("setDefaultFont", &JsWorkbook::setDefaultFont)
      .function("setHyperlink", &JsWorkbook::setHyperlink)
      .function("addMerge", &JsWorkbook::addMerge)
      .function("addNumberFormat", &JsWorkbook::addNumFmt)
      .function("addSheet", &JsWorkbook::addSheet)
      .function("setColumnBreak", &JsWorkbook::setSheetColBreak)
      .function("setRowBreak", &JsWorkbook::setSheetRowBreak)
      .function("addValidation", &JsWorkbook::addValidation)
      .function("addCellXf", &JsWorkbook::addXf)
      .function("getBorderCount", &JsWorkbook::borderCount)
      .function("getCalcMode", &JsWorkbook::calcMode)
      .function("getCompatibilityVersion", &JsWorkbook::compatibilityVersion)
      .function("listCellStyles", &JsWorkbook::listCellStyles)
      .function("getCellStyleXfCount", &JsWorkbook::cellStyleXfCount)
      .function("setCellStyleXf", &JsWorkbook::setCellStyleXf)
      .function("clearCfRules", &JsWorkbook::clearConditionalFormats)
      .function("resetColumnXfIndex", &JsWorkbook::clearColumnStyle)
      .function("clearHyperlinks", &JsWorkbook::clearHyperlinks)
      .function("clearMerges", &JsWorkbook::clearMerges)
      .function("resetRowXfIndex", &JsWorkbook::clearRowStyle)
      .function("clearValidations", &JsWorkbook::clearValidations)
      .function("addTable", &JsWorkbook::createTable)
      .function("listDefinedNames", &JsWorkbook::listDefinedNames)
      .function("getDefinedName", &JsWorkbook::getDefinedName)
      .function("removeDefinedName", &JsWorkbook::removeDefinedName)
      .function("deleteColumns", &JsWorkbook::deleteColumns)
      .function("deleteRows", &JsWorkbook::deleteRows)
      .function("getDependents", &JsWorkbook::dependents)
      .function("dispose", &JsWorkbook::dispose)
      .function("getMemoryUsage", &JsWorkbook::memoryUsage)
      .function("getDxfCount", &JsWorkbook::dxfCount)
      .function("evaluateCfRange", &JsWorkbook::evaluateCfRange)
      .function("evaluateCfFormula", &JsWorkbook::evaluateConditionalFormula)
      .function("evaluateFormulaArray", &JsWorkbook::evaluateFormulaArray)
      // Profile ids are {mac,win}-365-{ja_JP,en_US,de_DE,fr_FR,zh_CN,ko_KR,
      // th_TH,es_ES,es_MX,pt_BR,ru_RU,zh_TW,it_IT,nl_NL}. New workbooks default to win-365-en_US.
      .function("getProfileId", &JsWorkbook::profileId)
      .function("getVisibility", &JsWorkbook::getVisibility)
      .function("getLocation", &JsWorkbook::getLocation)
      .function("getFillCount", &JsWorkbook::fillCount)
      .function("getFontCount", &JsWorkbook::fontCount)
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
      .function("applyAutoFilter", &JsWorkbook::applyAutoFilter)
      .function("clearAutoFilterCriteria", &JsWorkbook::clearAutoFilter)
      .function("evaluateAutoFilter", &JsWorkbook::evaluateAutoFilter)
      .function("getTableAutoFilter", &JsWorkbook::getTableAutoFilter)
      .function("setTableAutoFilter", &JsWorkbook::setTableAutoFilter)
      .function("applyTableAutoFilter", &JsWorkbook::applyTableAutoFilter)
      .function("clearTableAutoFilterCriteria", &JsWorkbook::clearTableAutoFilter)
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
      .function("updateValidation", &JsWorkbook::updateValidation)
      .function("updateCfRule", &JsWorkbook::updateCfRule)
      .function("updateThreadedComment", &JsWorkbook::editThreadedComment)
      .function("removeThreadedComment", &JsWorkbook::removeThreadedComment)
      .function("listPeople", &JsWorkbook::getPersons)
      .function("addPerson", &JsWorkbook::addPerson)
      .function("removePerson", &JsWorkbook::removePerson)
      .function("getPerson", &JsWorkbook::getPerson)
      .function("listDrawingObjects", &JsWorkbook::listDrawingObjects)
      .function("getDrawingObject", &JsWorkbook::getDrawingObject)
      .function("getPicture", &JsWorkbook::getImage)
      .function("addImage", &JsWorkbook::insertImage)
      .function("addImage", emscripten::optional_override([](JsWorkbook& self, uint32_t sheet, emscripten::val bytes) {
                  return self.insertImage(sheet, bytes, emscripten::val::object());
                }))
      .function("removeDrawingObject", &JsWorkbook::removeDrawing)
      .function("updateDrawing", &JsWorkbook::updateDrawing)
      .function("setDrawingZOrder", &JsWorkbook::setDrawingZOrder)
      .function("snapshotDrawing", &JsWorkbook::snapshotDrawing)
      .function("restoreDrawing", &JsWorkbook::restoreDrawing)
      .function("restoreDrawing",
                emscripten::optional_override([](JsWorkbook& self, uint32_t sheet, emscripten::val bytes) {
                  return self.restoreDrawing(sheet, bytes, emscripten::val::object());
                }))
      .function("addShape", &JsWorkbook::insertShape)
      .function("updateShape", &JsWorkbook::updateShape)
      .function("getShape", &JsWorkbook::getShape)
      .function("getSlicer", &JsWorkbook::getSlicer)
      .function("addSlicer", &JsWorkbook::addSlicer)
      .function("updateSlicer", &JsWorkbook::updateSlicer)
      .function("addDataTable", &JsWorkbook::addDataTable)
      .function("getDataTable", &JsWorkbook::getDataTable)
      .function("removeDataTable", &JsWorkbook::removeDataTable)
      .function("recalcDataTables", &JsWorkbook::recalcDataTables)
      .function("getSparklineGroupCount", &JsWorkbook::getSparklineGroupCount)
      .function("getSparklineGroup", &JsWorkbook::getSparklineGroup)
      .function("addSparklineGroup", &JsWorkbook::addSparklineGroup)
      .function("updateSparklineGroup", &JsWorkbook::updateSparklineGroup)
      .function("removeSparklineGroup", &JsWorkbook::removeSparklineGroup)
      .function("getChart", &JsWorkbook::getChart)
      .function("getChartData", &JsWorkbook::getChartData)
      .function("getChartPreview", &JsWorkbook::previewChartData)
      .function("addChart", &JsWorkbook::insertChart)
      .function("setChartType", &JsWorkbook::setChartType)
      .function("setChartType", emscripten::optional_override([](JsWorkbook& self, uint32_t sheet, uint32_t objectId,
                                                                 uint32_t plot, int32_t type) {
                  return self.setChartType(sheet, objectId, plot, type, emscripten::val::undefined());
                }))
      .function("setChartTitle", &JsWorkbook::setChartTitle)
      .function("setChartLegend", &JsWorkbook::setChartLegend)
      .function("updateChartLegend", &JsWorkbook::updateChartLegend)
      .function("updateChartAxis", &JsWorkbook::setChartAxis)
      .function("updateChartDisplay", &JsWorkbook::setChartDisplay)
      .function("updateChartSeries", &JsWorkbook::setChartSeries)
      .function("setChartPoint", &JsWorkbook::setChartPoint)
      .function("updateChartPoint", &JsWorkbook::updateChartPoint)
      .function("updateChartPlot", &JsWorkbook::setChartPlot)
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
      .function("resetRowHeight", &JsWorkbook::clearRowHeight)
      .function("resetColumnWidth", &JsWorkbook::clearColumnWidth)
      .function("updateFormatDefaults", &JsWorkbook::setSheetFormatDefaults)
      .function("convertColumnCharsToPt", &JsWorkbook::columnCharsToPt)
      .function("convertColumnPtToChars", &JsWorkbook::columnPtToChars)
      .function("getCellXfIndex", &JsWorkbook::getCellXfIndex)
      .function("getCellPhonetic", &JsWorkbook::getCellPhonetic)
      .function("getCellPhoneticRunCount", &JsWorkbook::getCellPhoneticRunCount)
      .function("getCellPhoneticRun", &JsWorkbook::getCellPhoneticRun)
      .function("getCellPhoneticProperties", &JsWorkbook::getCellPhoneticProperties)
      .function("getCellShowPhonetic", &JsWorkbook::getCellShowPhonetic)
      .function("isCellStale", &JsWorkbook::isCellStale)
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
      .function("getDefaultFont", &JsWorkbook::getDefaultFont)
      .function("listHyperlinks", &JsWorkbook::getHyperlinks)
      .function("getHyperlink", &JsWorkbook::getHyperlink)
      .function("getIterative", &JsWorkbook::getIterative)
      .function("getLambdaText", &JsWorkbook::getLambdaText)
      .function("listMerges", &JsWorkbook::getMerges)
      .function("listCircularReferences", &JsWorkbook::getCircularReferences)
      .function("getMergesInRange", &JsWorkbook::getMergesInRange)
      .function("getNumberFormat", &JsWorkbook::getNumFmt)
      .function("listNumberFormats", &JsWorkbook::listNumberFormats)
      .function("listColumnOverrides", &JsWorkbook::getSheetColumns)
      .function("getAutoFilterXml", &JsWorkbook::getSheetAutoFilterXml)
      .function("listColumnBreaks", &JsWorkbook::getSheetColBreaks)
      .function("getHeaderFooterXml", &JsWorkbook::getSheetHeaderFooterXml)
      .function("getPageMargins", &JsWorkbook::getSheetPageMargins)
      .function("getPrintOptions", &JsWorkbook::getSheetPrintOptions)
      .function("getHeaderFooter", &JsWorkbook::getSheetHeaderFooter)
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
      .function("updateView", &JsWorkbook::updateSheetView)
      .function("getValidationCount", &JsWorkbook::getValidationCount)
      .function("getValidation", &JsWorkbook::getValidation)
      .function("getValue", &JsWorkbook::getValue)
      .function("insertColumns", &JsWorkbook::insertColumns)
      .function("insertRows", &JsWorkbook::insertRows)
      .function("insertTableColumns", &JsWorkbook::insertTableColumns)
      .function("insertTableRows", &JsWorkbook::insertTableRows)
      .function("insertCells", &JsWorkbook::insertCells)
      .function("deleteCells", &JsWorkbook::deleteCells)
      .function("moveSheet", &JsWorkbook::moveSheet)
      .function("copySheet", &JsWorkbook::copySheet)
      .function("recalcViewport", &JsWorkbook::partialRecalc)
      .function("getPagination", &JsWorkbook::paginate)
      .function("listPassthroughs", &JsWorkbook::listPassthroughs)
      .function("hasPassthrough", &JsWorkbook::hasPassthrough)
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
      .function("refreshPivotCache", &JsWorkbook::pivotCacheRefresh)
      .function("getPivotCacheFieldTypes", &JsWorkbook::pivotCacheFieldTypes)
      .function("clearPivotCacheFieldSharedItems", &JsWorkbook::pivotCacheFieldClearSharedItems)
      .function("getPivotCacheFieldCount", &JsWorkbook::pivotCacheFieldCount)
      .function("getPivotCacheFieldName", &JsWorkbook::pivotCacheFieldName)
      .function("getPivotCacheFieldSharedItemCount", &JsWorkbook::pivotCacheFieldSharedItemCount)
      .function("getPivotCacheFieldSharedItemValue", &JsWorkbook::pivotCacheFieldSharedItemValue)
      .function("getPivotCacheWorksheetSource", &JsWorkbook::pivotCacheGetWorksheetSource)
      .function("getPivotCacheWorksheetSourceRange", &JsWorkbook::pivotCacheGetWorksheetSourceRange)
      .function("addPivotCacheRecord", &JsWorkbook::pivotCacheRecordAdd)
      .function("clearPivotCacheRecords", &JsWorkbook::pivotCacheRecordClear)
      .function("getPivotCacheRecordCount", &JsWorkbook::pivotCacheRecordCount)
      .function("getPivotCacheRecordValue", &JsWorkbook::pivotCacheRecordValue)
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
      .function("addPivotFieldItem", &JsWorkbook::pivotFieldAddItem)
      .function("clearPivotFields", &JsWorkbook::pivotFieldClear)
      .function("clearPivotFieldItems", &JsWorkbook::pivotFieldClearItems)
      .function("setPivotFieldSubtotalFunctions", &JsWorkbook::pivotFieldSetSubtotalFns)
      .function("getPivotFieldCount", &JsWorkbook::pivotFieldCount)
      .function("getPivotField", &JsWorkbook::pivotFieldAt)
      .function("getPivotFieldItemCount", &JsWorkbook::pivotFieldItemCount)
      .function("getPivotFieldItem", &JsWorkbook::pivotFieldItemAt)
      .function("setPivotFieldAxis", &JsWorkbook::pivotFieldSetAxis)
      .function("updatePivotField", &JsWorkbook::pivotFieldUpdate)
      .function("updatePivotFieldItem", &JsWorkbook::pivotFieldItemUpdate)
      .function("setPivotFieldDateGroup", &JsWorkbook::pivotFieldSetDateGroup)
      // `null` as the fourth argument removes the grouping.
      .function("setPivotFieldDateGroup",
                emscripten::optional_override([](JsWorkbook& self, uint32_t sheet, uint32_t pivotIndex,
                                                 uint32_t fieldIndex, emscripten::val group) {
                  return self.pivotFieldRemoveDateGroup(sheet, pivotIndex, fieldIndex, group);
                }))
      .function("getPivotFieldDateGroup", &JsWorkbook::pivotFieldGetDateGroup)
      .function("getPivotFieldNumericGroup", &JsWorkbook::pivotFieldGetNumericGroup)
      .function("setPivotFieldNumericGroup", &JsWorkbook::pivotFieldSetNumericGroup)
      // `null` as the fourth argument removes the grouping.
      .function("setPivotFieldNumericGroup",
                emscripten::optional_override([](JsWorkbook& self, uint32_t sheet, uint32_t pivotIndex,
                                                 uint32_t fieldIndex, emscripten::val group) {
                  return self.pivotFieldRemoveNumericGroup(sheet, pivotIndex, fieldIndex, group);
                }))
      .function("getPivotOptions", &JsWorkbook::pivotGetOptions)
      .function("updatePivotOptions", &JsWorkbook::pivotSetOptions)
      .function("getPivotCacheOptions", &JsWorkbook::pivotCacheGetOptions)
      .function("updatePivotCacheOptions", &JsWorkbook::pivotCacheSetOptions)
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
      .function("getPivotSpill", &JsWorkbook::getPivotSpill)
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
                                                        FM_PIVOT_DATA_FIELD_NONE);
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
      .function("updatePivotGrandTotals", &JsWorkbook::pivotSetGrandTotals)
      .function("updatePivotStyleOptions", &JsWorkbook::pivotUpdateStyleOptions)
      .function("getPivotReportLayout", &JsWorkbook::pivotGetLayout)
      .function("getPivotCacheId", &JsWorkbook::pivotCacheId)
      .function("getPivotName", &JsWorkbook::pivotName)
      .function("getPivotRuntimeId", &JsWorkbook::pivotRuntimeId)
      .function("resolvePivotRuntimeId", &JsWorkbook::resolvePivotRuntimeId)
      .function("getPivotGrandTotals", &JsWorkbook::pivotGetGrandTotals)
      .function("getPivotStyleOptions", &JsWorkbook::pivotGetStyleOptions)
      .function("getPivotRowFieldOrder", &JsWorkbook::pivotRowFieldOrder)
      .function("getPivotColumnFieldOrder", &JsWorkbook::pivotColFieldOrder)
      .function("getPivotPageFieldOrder", &JsWorkbook::pivotPageFieldOrder)
      .function("setPivotReportLayout", &JsWorkbook::pivotSetLayout)
      .function("renamePivot", &JsWorkbook::pivotSetName)
      .function("setPivotRowFieldOrder", &JsWorkbook::pivotSetRowFieldOrder)
      .function("setPivotPageFieldOrder", &JsWorkbook::pivotSetPageFieldOrder)
      .function("getPrecedents", &JsWorkbook::precedents)
      .function("recalc", &JsWorkbook::recalc)
      .function("recalcParallel", &JsWorkbook::recalcParallel)
      .function("hasDirtyCells", &JsWorkbook::hasDirtyCells)
      .function("removeCfRule", &JsWorkbook::removeConditionalFormatAt)
      .function("insertHyperlinkAt", &JsWorkbook::insertHyperlinkAt)
      .function("removeHyperlink", &JsWorkbook::removeHyperlink)
      .function("removeHyperlinkAt", &JsWorkbook::removeHyperlinkAt)
      .function("removeMerge", &JsWorkbook::removeMerge)
      .function("removeMergeAt", &JsWorkbook::removeMergeAt)
      .function("removeSheet", &JsWorkbook::removeSheet)
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
      .function("setCompatibilityVersion", &JsWorkbook::setCompatibilityVersion)
      .function("getFullPrecision", &JsWorkbook::fullPrecision)
      .function("setFullPrecision", &JsWorkbook::setFullPrecision)
      .function("getCalcOnSave", &JsWorkbook::calcOnSave)
      .function("setCalcOnSave", &JsWorkbook::setCalcOnSave)
      .function("getRefMode", &JsWorkbook::refMode)
      .function("setRefMode", &JsWorkbook::setRefMode)
      .function("getDate1904", &JsWorkbook::date1904)
      .function("setDate1904", &JsWorkbook::setDate1904)
      .function("setCellXfIndex", &JsWorkbook::setCellXfIndex)
      .function("resetCellXfIndex", &JsWorkbook::clearCellXfIndex)
      .function("setRangeXfIndex", &JsWorkbook::setRangeXfIndex)
      .function("clearBreaks", &JsWorkbook::clearSheetBreaks)
      .function("setCellStyle", &JsWorkbook::setCellStyle)
      .function("removeCellStyle", &JsWorkbook::removeCellStyle)
      .function("listTableStyles", &JsWorkbook::listTableStyles)
      .function("getTableStyle", &JsWorkbook::getTableStyle)
      .function("setTableStyle", &JsWorkbook::setTableStyle)
      .function("removeTableStyle", &JsWorkbook::removeTableStyle)
      .function("getTableStyleDefaults", &JsWorkbook::getTableStyleDefaults)
      .function("updateTableStyleDefaults", &JsWorkbook::setTableStyleDefaults)
      .function("getThemeColors", &JsWorkbook::getThemeColors)
      .function("getThemeFonts", &JsWorkbook::getThemeFonts)
      .function("getThemeScriptFont", &JsWorkbook::getThemeScriptFont)
      .function("listThemeScriptFonts", &JsWorkbook::listThemeScriptFonts)
      .function("setThemeColors", &JsWorkbook::setThemeColors)
      .function("setThemeFonts", &JsWorkbook::setThemeFonts)
      .function("setThemeScriptFont", &JsWorkbook::setThemeScriptFont)
      .function("setVisibility", &JsWorkbook::setVisibility)
      .function("resetTheme", &JsWorkbook::resetTheme)
      .function("resolveColor", &JsWorkbook::resolveColor)
      .function("getEffectiveStyle", &JsWorkbook::getEffectiveStyle)
      .function("setCellPhonetic", &JsWorkbook::setCellPhonetic)
      .function("setCellPhoneticRuns", &JsWorkbook::setCellPhoneticRuns)
      .function("setCellPhoneticProperties", &JsWorkbook::setCellPhoneticProperties)
      .function("setCellShowPhonetic", &JsWorkbook::setCellShowPhonetic)
      .function("setActiveSheet", &JsWorkbook::setActiveSheet)
      .function("activeSheet", &JsWorkbook::activeSheet)
      .function("setColumnHidden", &JsWorkbook::setColumnHidden)
      .function("setColumnOutline", &JsWorkbook::setColumnOutline)
      .function("setColumnCollapsed", &JsWorkbook::setColumnCollapsed)
      .function("setColumnXfIndex", &JsWorkbook::setColumnStyle)
      .function("setColumnWidth", &JsWorkbook::setColumnWidth)
      .function("setComment", &JsWorkbook::setComment)
      .function("setCommentVisible", &JsWorkbook::setCommentVisible)
      .function("setWorkbookProtection", &JsWorkbook::setWorkbookProtection)
      .function("addProtectedRange", &JsWorkbook::addProtectedRange)
      .function("updateProtectedRange", &JsWorkbook::updateProtectedRange)
      .function("removeProtectedRange", &JsWorkbook::removeProtectedRange)
      .function("setDefinedName", &JsWorkbook::setDefinedName)
      .function("setError", &JsWorkbook::setError)
      // `setProfileId` accepts the ids documented at `getProfileId`.
      .function("setProfileId", &JsWorkbook::setProfileId)
      .function("setLocation", &JsWorkbook::setLocation)
      .function("setFormula", &JsWorkbook::setFormula)
      .function("updateIterative", &JsWorkbook::setIterative)
      .function("setIterativeProgress", &JsWorkbook::setIterativeProgress)
      .function("setNumber", &JsWorkbook::setNumber)
      .function("setPinnedNow", &JsWorkbook::setPinnedNow)
      .function("setPinnedNow", &JsWorkbook::releasePinnedNow)
      .function("setRowAutoHeight", &JsWorkbook::setRowAutoHeight)
      .function("setRowHeight", &JsWorkbook::setRowHeight)
      .function("setRowHidden", &JsWorkbook::setRowHidden)
      .function("setRowOutline", &JsWorkbook::setRowOutline)
      .function("setRowCollapsed", &JsWorkbook::setRowCollapsed)
      .function("setRowXfIndex", &JsWorkbook::setRowStyle)
      .function("getRangeXfBases", &JsWorkbook::rangeStyleBases)
      .function("getRangeXfBasesExcluding", &JsWorkbook::rangeStyleBasesExcluding)
      .function("listCellXfIndices", &JsWorkbook::listCellXfIndices)
      .function("listShowPhoneticCells", &JsWorkbook::listShowPhoneticCells)
      .function("remapRangeXfs", &JsWorkbook::remapRangeStyle)
      .function("snapshotRangeXfs", &JsWorkbook::snapshotRangeStyle)
      .function("restoreRangeXfs", &JsWorkbook::restoreRangeStyle)
      .function("setAutoFilterXml", &JsWorkbook::setSheetAutoFilterXml)
      .function("updateHeaderFooter", &JsWorkbook::setSheetHeaderFooter)
      .function("setHeaderFooterXml", &JsWorkbook::setSheetHeaderFooterXml)
      .function("updatePageMargins", &JsWorkbook::setSheetPageMargins)
      .function("setPageMarginsXml", &JsWorkbook::setSheetPageMarginsXml)
      .function("updatePageSetup", &JsWorkbook::setSheetPageSetup)
      .function("setPageSetupXml", &JsWorkbook::setSheetPageSetupXml)
      .function("setPrintArea", &JsWorkbook::setSheetPrintArea)
      .function("updatePrintOptions", &JsWorkbook::setSheetPrintOptions)
      .function("setPrintOptionsXml", &JsWorkbook::setSheetPrintOptionsXml)
      .function("setPrintTitles", &JsWorkbook::setSheetPrintTitles)
      .function("setSheetPrXml", &JsWorkbook::setSheetSheetPrXml)
      .function("getTabColor", &JsWorkbook::getSheetTabColor)
      .function("setTabColor", &JsWorkbook::setSheetTabColor)
      .function("setSheetProtection", &JsWorkbook::setSheetProtection)
      .function("updateSheetProtection", &JsWorkbook::updateSheetProtection)
      .function("getSelection", &JsWorkbook::getSheetSelection)
      .function("setSelection", &JsWorkbook::setSheetSelection)
      .function("setText", &JsWorkbook::setText)
      .function("getSheetCount", &JsWorkbook::sheetCount)
      .function("getSheetName", &JsWorkbook::sheetName)
      .function("getSheetIndex", &JsWorkbook::getSheetIndex)
      .function("getSpill", &JsWorkbook::getSpill)
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
      .function("getTableStyleOptions", &JsWorkbook::getTableStyleOptions)
      .function("updateTableStyleOptions", &JsWorkbook::updateTableStyleOptions)
      .function("renameTable", &JsWorkbook::renameTable)
      .function("findTable", &JsWorkbook::findTable)
      .function("getTableRuntimeId", &JsWorkbook::tableRuntimeId)
      .function("resolveTableRuntimeId", &JsWorkbook::resolveTableRuntimeId)
      .function("getTableCount", &JsWorkbook::tableCount)
      .function("updateTable", &JsWorkbook::updateTable)
      .function("getCellXfCount", &JsWorkbook::xfCount);

  // ---- Free functions ------------------------------------------------------
  function("evaluateFormula", &evaluate_formula);
  function("getVersion", &version_string);
  function("setLogMinLevel", &set_log_min_level);
  function("setLogSink", &set_log_sink);
  function("canonicalizeFormula", &canonicalize_formula);
  function("canonicalizeFunctionName", &canonicalize_function_name);
  function("compareText", &compare_text);
  function("convertCivilToSerial", &convert_civil_to_serial);
  function("convertSerialToCivil", &convert_serial_to_civil);
  function("getCollationClasses", &collation_classes);
  function("getFunctionMetadata", &function_metadata);
  function("getImageInfo", &probe_image);
  function("getLocaleFacts", &locale_facts);
  function("listFunctionNames", &function_names);
  function("listProfileIds", &list_profile_ids);
  function("localizeFormula", &localize_formula);
  function("localizeFunctionName", &localize_function_name);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
