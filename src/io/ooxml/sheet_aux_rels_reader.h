// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
///   * `vml_path` — resolved path of the `kRelVmlDrawing` target, or
///     empty. Used to detect the legacy bounding-box stub so it gets
///     marked consumed (passthrough re-emits the bytes).
///   * `printer_settings_path` — resolved path of the binary printer
///     settings part referenced by `<pageSetup r:id="...">`, or empty.
struct SheetAuxRels {
  std::unordered_map<std::string, std::string> hyperlink_rid_to_target;
  std::string comments_path;
  std::string vml_path;
  std::string printer_settings_rid;
  std::string printer_settings_path;
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
Expected<SheetAuxRels, Error> load_sheet_aux_rels(const ZipReader& zip, std::string_view sheet_rels_path,
                                                  std::string_view sheet_dir);

}  // namespace ooxml
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_SHEET_AUX_RELS_READER_H_
