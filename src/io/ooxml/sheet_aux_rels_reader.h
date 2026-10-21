//
// `xl/worksheets/_rels/sheetN.xml.rels` collectors. Each traverses a retained
// document once for one pipeline phase:
//
//   * `collect_sheet_table_targets`   — `kRelTable` entries (table
//                                        definition parts)
//   * `collect_sheet_pivot_table_targets` — `kRelPivotTable` entries
//                                        (pivot-table parts anchored
//                                        on this sheet)
//   * `collect_sheet_aux_rels`         — hyperlink / comments / VML /
//                                        printer-settings entries that
//                                        the per-sheet consumer needs
//                                        in one pass.
//
// The orchestrator loads the ZIP entry once and keeps these phase-specific
// passes separate so relationship resolution errors retain their order.
//
// Internal helper for the OOXML reader. Not exposed beyond `src/io/`.

#ifndef FORMULON_IO_OOXML_SHEET_AUX_RELS_READER_H_
#define FORMULON_IO_OOXML_SHEET_AUX_RELS_READER_H_

#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include "pugixml.hpp"
#include "unknown_relationship.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace ooxml {

/// Aggregated lookup for the per-sheet auxiliary parts that are not
/// covered by the table and pivot target collectors:
///
///   * `hyperlink_rid_to_target` — every `rId -> Target` pair under the
///     hyperlink relationship type. Hyperlink targets are external
///     URLs (`http://...`, `mailto:...`, `file:...`); the writer
///     re-emits them verbatim, so we keep them as the raw string the
///     OOXML producer emitted (no resolution).
///   * `comments_path` — resolved path of the `kRelComments` target, or
///     empty when the sheet has none.
///   * `vml_path` — resolved path of the `kRelVmlDrawing` target bound
///     to comment geometry (matched to the worksheet body's
///     `<legacyDrawing r:id>`, see `collect_sheet_aux_rels`), or empty. Used
///     to detect the legacy bounding-box stub so it gets marked consumed
///     (passthrough re-emits the bytes). A sheet may carry a *second*
///     `kRelVmlDrawing` relationship for `<legacyDrawingHF>` (header /
///     footer image); that one is never bound here and instead surfaces
///     through `unknown_rels` so its id and target round-trip.
///   * `printer_settings_path` — resolved path of the binary printer
///     settings part referenced by `<pageSetup r:id="...">`, or empty.
///   * `threaded_comments_path` — resolved path of the
///     `kRelThreadedComment` target, or empty when the sheet has none.
///   * `drawing_path` — resolved path of the DrawingML part referenced
///     by the worksheet's `<drawing r:id="...">` element (charts,
///     images, shapes), or empty. The drawing part, its own rels, and
///     any media it anchors round-trip through the passthrough
///     mechanism; the writer re-emits the `<drawing>` reference and the
///     matching sheet-rels relationship so the part stays reachable.
struct SheetAuxRels {
  std::unordered_map<std::string, std::string> hyperlink_rid_to_target;
  std::string comments_path;
  std::string threaded_comments_path;
  std::string vml_path;
  std::string printer_settings_rid;
  std::string printer_settings_path;
  std::string drawing_path;
  // Any relationship type not consumed by a worksheet-specific reader.
  // Internal targets are package-relative; external targets are verbatim.
  std::vector<UnknownRelationship> unknown_rels;
};

/// Collects `kRelTable` targets from a parsed relationship document in
/// document order, resolving each against `sheet_dir`.
Expected<std::vector<std::string>, Error> collect_sheet_table_targets(const pugi::xml_document& document,
                                                                      std::string_view sheet_dir);

/// Collects `kRelPivotTable` targets from a parsed relationship document
/// in document order, resolving each against `sheet_dir`.
Expected<std::vector<std::string>, Error> collect_sheet_pivot_table_targets(const pugi::xml_document& document,
                                                                            std::string_view sheet_dir);

/// Collects the auxiliary relationships from a parsed document. Unrelated
/// types are ignored and VML selection uses the worksheet body's legacy id.
///
/// `legacy_drawing_body_rid` is the `r:id` value of the worksheet body's
/// `<legacyDrawing>` element (empty when the sheet has none). A sheet
/// may carry up to two `kRelVmlDrawing` relationships — one for comment
/// geometry (referenced by `<legacyDrawing>`) and one for a
/// header/footer image (referenced by `<legacyDrawingHF>`, which is
/// captured separately as raw XML and is not this function's concern).
/// The relationship whose `Id` matches `legacy_drawing_body_rid` becomes
/// `vml_path`; when no candidate matches (including when the body
/// carries no `<legacyDrawing>` element at all) the first `kRelVmlDrawing`
/// relationship found is used as a best-effort fallback whenever the
/// sheet also has a `kRelComments` relationship, since a sheet with
/// comments needs exactly one candidate to serve as its comment VML.
/// Every `kRelVmlDrawing` relationship not selected this way is added to
/// `unknown_rels` instead, preserving its id and target verbatim.
Expected<SheetAuxRels, Error> collect_sheet_aux_rels(const pugi::xml_document& document, std::string_view sheet_dir,
                                                     std::string_view legacy_drawing_body_rid);

}  // namespace ooxml
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_SHEET_AUX_RELS_READER_H_
