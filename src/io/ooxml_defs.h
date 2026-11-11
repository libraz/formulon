//
// Shared OOXML schema constants used by the reader and writer.
// Only the namespace and relationship type URIs that more than one reader / writer
// (OOXML or XLSB) needs are hoisted here, plus the XLSB workbook content types
// shared by format detection and the XLSB reader. Content-type strings needed by
// exactly one side (e.g. the writer-only `kCtComments`) stay in
// the consuming `.cpp` so this header tracks the minimum shared surface.
//
// Header-only, `<string_view>`-only, `inline constexpr` — no
// translation unit needed. Lives under `io/` because no module
// outside `io/` consumes the OOXML relationship vocabulary. Symbols
// live directly in `formulon::io` to match the sibling-header style
// used by `workbook_kind.h` and `passthrough_part.h`, which keeps
// reader / writer call sites unqualified.
//
// Schema reference: ECMA-376 Part 1 §11.3 (Relationships) and
// §15 (SpreadsheetML content types).

#ifndef FORMULON_IO_OOXML_DEFS_H_
#define FORMULON_IO_OOXML_DEFS_H_

#include <string_view>

namespace formulon {
namespace io {

// XML namespace URIs shared by several readers / writers.

inline constexpr std::string_view kNsSpreadsheetMl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
inline constexpr std::string_view kNsOfficeRelationships =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
inline constexpr std::string_view kNsPackageRelationships =
    "http://schemas.openxmlformats.org/package/2006/relationships";
inline constexpr std::string_view kNsMarkupCompat = "http://schemas.openxmlformats.org/markup-compatibility/2006";
inline constexpr std::string_view kNsX14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
inline constexpr std::string_view kNsX15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
inline constexpr std::string_view kNsX15ac = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac";
inline constexpr std::string_view kNsExcelMain = "http://schemas.microsoft.com/office/excel/2006/main";
inline constexpr std::string_view kNsContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
inline constexpr std::string_view kNsXr = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
inline constexpr std::string_view kNsXr2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2";
inline constexpr std::string_view kNsXxl21 = "http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021";
inline constexpr std::string_view kNsDynamicArray =
    "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray";
inline constexpr std::string_view kNsRichData = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";

// Relationship type URIs used by Excel-produced packages. The values
// are the canonical ECMA-376 strings; comparison is case-sensitive and
// must remain byte-identical to what Excel emits.

inline constexpr std::string_view kRelOfficeDocument =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
inline constexpr std::string_view kRelWorksheet =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
inline constexpr std::string_view kRelSharedStrings =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
inline constexpr std::string_view kRelStyles =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
inline constexpr std::string_view kRelTable =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
inline constexpr std::string_view kRelPivotCacheDefinition =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
inline constexpr std::string_view kRelPivotCacheRecords =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
inline constexpr std::string_view kRelPivotTable =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
inline constexpr std::string_view kRelHyperlink =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
inline constexpr std::string_view kRelComments =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
inline constexpr std::string_view kRelVmlDrawing =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
inline constexpr std::string_view kRelDrawing =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
inline constexpr std::string_view kRelChart =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
inline constexpr std::string_view kRelPrinterSettings =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";
inline constexpr std::string_view kRelExternalLink =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
inline constexpr std::string_view kRelExternalLinkPath =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";
inline constexpr std::string_view kRelOleLink =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleLink";
inline constexpr std::string_view kRelDdeLink =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ddeLink";
inline constexpr std::string_view kRelThreadedComment =
    "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";
inline constexpr std::string_view kRelPerson = "http://schemas.microsoft.com/office/2017/10/relationships/person";
inline constexpr std::string_view kRelTheme =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
inline constexpr std::string_view kRelSheetMetadata =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
inline constexpr std::string_view kRelCalcChain =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
inline constexpr std::string_view kRelRdRichValue =
    "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue";
inline constexpr std::string_view kRelRdRichValueStructure =
    "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure";
inline constexpr std::string_view kRelRdRichValueTypes =
    "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes";
inline constexpr std::string_view kRelCoreProperties =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
inline constexpr std::string_view kRelExtendedProperties =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
inline constexpr std::string_view kRelCustomProperties =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";

// Content types of the package-level parts both container writers declare as
// `<Default>` entries.
inline constexpr std::string_view kCtPackageRels = "application/vnd.openxmlformats-package.relationships+xml";
inline constexpr std::string_view kCtXml = "application/xml";

// XLSB workbook content types ([Content_Types].xml). Either form marks the
// package as MS-XLSB.
inline constexpr std::string_view kCtWorkbookXlsb = "application/vnd.ms-excel.sheet.binary.macroEnabled.main";
inline constexpr std::string_view kCtWorkbookXlsbAlt = "application/vnd.ms-excel.sheet.macroEnabled.main";

inline bool is_xlsb_workbook_content_type(std::string_view content_type) {
  return content_type == kCtWorkbookXlsb || content_type == kCtWorkbookXlsbAlt;
}

// Rich-value parts: XML in both package formats, at fixed paths.
inline constexpr std::string_view kRdRichValuePartPath = "xl/richData/rdrichvalue.xml";
inline constexpr std::string_view kRdRichValueStructurePartPath = "xl/richData/rdrichvaluestructure.xml";
inline constexpr std::string_view kRdRichValueTypesPartPath = "xl/richData/rdRichValueTypes.xml";
inline constexpr std::string_view kRichDataDirPrefix = "xl/richData/";
inline constexpr std::string_view kCtRdRichValue = "application/vnd.ms-excel.rdrichvalue+xml";
inline constexpr std::string_view kCtRdRichValueStructure = "application/vnd.ms-excel.rdrichvaluestructure+xml";
inline constexpr std::string_view kCtRdRichValueTypes = "application/vnd.ms-excel.rdrichvaluetypes+xml";

// Worksheet drawing part content type; an XML part in both package formats.
inline constexpr std::string_view kCtDrawing = "application/vnd.openxmlformats-officedocument.drawing+xml";

// Chart part content type.
inline constexpr std::string_view kCtChart = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_DEFS_H_
