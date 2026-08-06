//
// Worksheet-part XML builders for the OOXML writer: the `<worksheet>`
// body itself plus the per-sheet `<sheetViews>`, `<cols>`,
// `<mergeCells>`, `<dataValidations>`, `<hyperlinks>`, `<rowBreaks>` /
// `<colBreaks>`, and `<sheetProtection>` blocks, together with the
// per-sheet `_rels` file emission (table parts, pivot-table parts,
// hyperlink targets, printer-settings, comments / VML).
//
// Internal to `src/io/ooxml/`; not part of the public API. The only
// caller outside this header is `src/io/ooxml_writer.cpp`'s
// orchestrator, which threads the results into the in-memory zip
// archive.

#ifndef FORMULON_IO_OOXML_SHEET_XML_BUILDER_H_
#define FORMULON_IO_OOXML_SHEET_XML_BUILDER_H_

#include <string>
#include <string_view>
#include <vector>

#include "io/ooxml/emission_plan.h"

namespace formulon {
class Sheet;
namespace io {
class SharedStrings;

/// Result of building a single per-sheet `_rels` file: the serialised
/// XML alongside the rId strings (`"rIdN"`) the writer assigned to each
/// hyperlink in document order. The orchestrator threads the rId vector
/// into the sheet part's `<hyperlinks>` block so the two stay in sync.
struct SheetRelsResult {
  std::string xml;
  std::vector<std::string> hyperlink_rids;
  std::string printer_settings_rid;
  // rId minted for the sheet's DrawingML relationship, or empty when the
  // sheet anchors no drawing. Threaded back into the worksheet body so
  // its `<drawing r:id>` element matches the rels entry.
  std::string drawing_rid;
  // rId minted for the comments' legacy VML drawing, or empty when the
  // sheet has no comments.
  std::string legacy_drawing_rid;
};

/// Builds the worksheet part body (`xl/worksheets/sheetN.xml`) for a
/// single sheet. `sheet_tables` is the planner-assigned table list owned
/// by this sheet; `hyperlink_rids` carries the per-hyperlink rId strings
/// minted by the matching `BuildSheetRels` call; `printer_settings_rid`
/// is the rId of the printer-settings rel (empty when the sheet has no
/// printer settings).
std::string BuildWorksheetXml(const Sheet& sheet, const std::vector<EmissionPlan::PerSheetTable>& sheet_tables,
                              const std::vector<std::string>& hyperlink_rids, std::string_view printer_settings_rid,
                              std::string_view drawing_rid, std::string_view legacy_drawing_rid,
                              const SharedStrings* shared_strings);

/// Builds the `_rels` document for a single sheet, covering tables,
/// pivot tables, hyperlinks, printer settings, and comments / VML.
SheetRelsResult BuildSheetRels(const Sheet& sheet, const std::vector<EmissionPlan::PerSheetTable>& sheet_tables,
                               const std::vector<EmissionPlan::PivotTablePlan>& sheet_pivot_tables,
                               const EmissionPlan::CommentsPlan& comments_plan);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_SHEET_XML_BUILDER_H_
