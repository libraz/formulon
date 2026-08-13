//
// `xl/worksheets/_rels/sheetN.xml.rels` reader. Splits the per-sheet
// relationship file into three thematic lookups:
//
//   * `load_sheet_table_targets`       — `kRelTable` entries (table
//                                        definition parts)
//   * `load_sheet_pivot_table_targets` — `kRelPivotTable` entries
//                                        (pivot-table parts anchored
//                                        on this sheet)
//   * `load_sheet_aux_rels`            — hyperlink / comments / VML /
//                                        printer-settings entries that
//                                        the per-sheet consumer needs
//                                        in a single pass.
//
// Each helper walks the same rels file once. Keeping them as separate
// entry points lets each consumer site read linearly without folding
// in unrelated types.
//
// Internal helper for the OOXML reader. Not exposed beyond `src/io/`.

#ifndef FORMULON_IO_OOXML_SHEET_AUX_RELS_READER_H_
#define FORMULON_IO_OOXML_SHEET_AUX_RELS_READER_H_

#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include "io/unknown_relationship.h"
#include "io/zip_reader.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace ooxml {

/// Aggregated lookup for the per-sheet auxiliary parts that are not
/// already covered by `load_sheet_table_targets` /
/// `load_sheet_pivot_table_targets`:
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
///     `<legacyDrawing r:id>`, see `load_sheet_aux_rels`), or empty. Used
///     to detect the legacy bounding-box stub so it gets marked consumed
///     (passthrough re-emits the bytes). A sheet may carry a *second*
///     `kRelVmlDrawing` relationship for `<legacyDrawingHF>` (header /
///     footer image); that one is never bound here and instead surfaces
///     through `unknown_rels` so its id and target round-trip.
///   * `printer_settings_path` — resolved path of the binary printer
///     settings part referenced by `<pageSetup r:id="...">`, or empty.
///   * `drawing_path` — resolved path of the DrawingML part referenced
///     by the worksheet's `<drawing r:id="...">` element (charts,
///     images, shapes), or empty. The drawing part, its own rels, and
///     any media it anchors round-trip through the passthrough
///     mechanism; the writer re-emits the `<drawing>` reference and the
///     matching sheet-rels relationship so the part stays reachable.
struct SheetAuxRels {
  std::unordered_map<std::string, std::string> hyperlink_rid_to_target;
  std::string comments_path;
  std::string vml_path;
  std::string printer_settings_rid;
  std::string printer_settings_path;
  std::string drawing_path;
  // Any relationship type not consumed by a worksheet-specific reader.
  // Internal targets are package-relative; external targets are verbatim.
  std::vector<UnknownRelationship> unknown_rels;
};

/// Walks `sheet_rels_path` for `kRelTable` entries and returns the
/// resolved table-part paths, each resolved relative to `sheet_dir`.
/// Non-table relationships are silently ignored at this layer.
Expected<std::vector<std::string>, Error> load_sheet_table_targets(const ZipReader& zip,
                                                                   std::string_view sheet_rels_path,
                                                                   std::string_view sheet_dir);

/// Walks `sheet_rels_path` for `kRelPivotTable` entries and returns the
/// resolved part paths in document order, each resolved relative to
/// `sheet_dir`. Non-pivot relationships are silently ignored.
Expected<std::vector<std::string>, Error> load_sheet_pivot_table_targets(const ZipReader& zip,
                                                                         std::string_view sheet_rels_path,
                                                                         std::string_view sheet_dir);

/// Walks `sheet_rels_path` once for hyperlink, comments, VML and
/// printer-settings entries; returns the aggregated lookup. The walker
/// silently ignores unrelated relationship types so each consumer site
/// reads only the slice it cares about.
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
Expected<SheetAuxRels, Error> load_sheet_aux_rels(const ZipReader& zip, std::string_view sheet_rels_path,
                                                  std::string_view sheet_dir, std::string_view legacy_drawing_body_rid);

}  // namespace ooxml
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_SHEET_AUX_RELS_READER_H_
