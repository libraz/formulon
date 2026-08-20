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

/// Builds the worksheet part body (`xl/worksheets/sheetN.xml`) for a
/// single sheet. `sheet_tables` is the planner-assigned table list owned
/// by this sheet; `table_rids` carries the per-table rId strings minted
/// by the matching `BuildSheetRels` call, index-aligned with
/// `sheet_tables` so every `<tablePart r:id>` names the id that rels
/// file actually declared; `hyperlink_rids` carries the per-hyperlink
/// rId strings the same way; `printer_settings_rid` is the rId of the
/// printer-settings rel (empty when the sheet has no printer settings).
/// `dxf_count` is the `<dxf>` record count of the package's styles part,
/// which bounds the `dxfId` values the conditional-format block may name
/// (see `write_conditional_formattings`).
std::string BuildWorksheetXml(const Sheet& sheet, const std::vector<EmissionPlan::PerSheetTable>& sheet_tables,
                              const std::vector<std::string>& table_rids,
                              const std::vector<std::string>& hyperlink_rids, std::string_view printer_settings_rid,
                              std::string_view drawing_rid, std::string_view legacy_drawing_rid,
                              const SharedStrings* shared_strings, std::size_t dxf_count);

/// Builds the `_rels` document for a single sheet, covering tables,
/// pivot tables, hyperlinks, printer settings, comments / VML, and
/// surviving unknown relationships.
///
/// Every relationship with an internal target is emitted only when that
/// target is present in `plan.passthrough_kept`; external relationships
/// remain eligible without a local payload. A relationship dropped by
/// that rule bumps `diagnostics->dropped_relationship_count`, which is
/// the only signal a caller gets that the saved sheet references less
/// than the source did. `diagnostics` may be NULL, which discards the
/// counts and changes nothing else.
SheetRelsResult BuildSheetRels(const Sheet& sheet, const std::vector<EmissionPlan::PerSheetTable>& sheet_tables,
                               const std::vector<EmissionPlan::PivotTablePlan>& sheet_pivot_tables,
                               const EmissionPlan::CommentsPlan& comments_plan, const EmissionPlan& plan,
                               WriteDiagnostics* diagnostics);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_SHEET_XML_BUILDER_H_
