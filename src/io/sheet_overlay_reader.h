//
// Per-sheet presentation overlays: `<mergeCells>`, `<hyperlinks>`,
// `<dataValidations>` and `<sheetProtection>`. These carry no cell value, so
// they are read independently of the `<sheetData>` walker in `sheet_reader.h`.

#ifndef FORMULON_IO_SHEET_OVERLAY_READER_H_
#define FORMULON_IO_SHEET_OVERLAY_READER_H_

#include <string>
#include <unordered_map>
#include <vector>

#include "io/package_diagnostics.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {

/// Walks `<mergeCells>` inside `worksheet` and returns the parsed
/// `MergeRange` list in document order. Returns an empty vector when
/// the sheet has no merges.
///
/// Invalid-reference policy — shared by every presentation-overlay
/// reader on this sheet (merges, hyperlinks, data validations, and
/// `read_conditional_formats`): an entry whose reference is missing,
/// empty or unparseable is dropped with a WARN diagnostic
/// (`io.sheet.overlay.skip`) and the walk continues. None of this
/// metadata carries a cell value, so one malformed entry must not cost
/// the caller the whole workbook. Corruption inside `<sheetData>` still
/// fails the sheet with `kIoSheetCorrupt`.
///
/// Every such drop bumps `diagnostics->skipped_feature_count` when
/// `diagnostics` is non-NULL. The structured log is off by default in
/// shipped builds, so that counter is what surfaces the loss to a caller.
/// Passing NULL discards the count.
Expected<std::vector<MergeRange>, Error> read_merges(const pugi::xml_node& worksheet,
                                                     ReadDiagnostics* diagnostics = nullptr);

/// Walks `<hyperlinks>` inside `worksheet` and returns the parsed
/// `Hyperlink` list in document order. The reader populates
/// `Hyperlink::rid` from the `r:id` attribute as observed; the caller
/// is responsible for joining each rid against the sheet's rels file
/// to fill in `target` (external URL) — see `apply_hyperlink_rels`.
/// A malformed `ref=` drops that hyperlink under the invalid-reference
/// policy documented on `read_merges`, `diagnostics` included.
Expected<std::vector<Hyperlink>, Error> read_hyperlinks(const pugi::xml_node& worksheet,
                                                        ReadDiagnostics* diagnostics = nullptr);

/// Joins `hyperlinks` against the parsed sheet rels file (`rid -> target`
/// map) populating `Hyperlink::target` in place for entries whose `rid`
/// matches an entry in `rid_to_target`. Entries with empty `rid` or no
/// matching rid are left untouched (their `target` stays whatever the
/// reader picked up from any inline `location` attribute).
void apply_hyperlink_rels(std::vector<Hyperlink>& hyperlinks,
                          const std::unordered_map<std::string, std::string>& rid_to_target);

/// Walks `<dataValidations>` inside `worksheet` and returns the parsed
/// `DataValidation` list in document order. Returns an empty vector
/// when the sheet has no validations. A malformed `sqref=` drops that
/// validation under the invalid-reference policy documented on
/// `read_merges`, `diagnostics` included.
Expected<std::vector<DataValidation>, Error> read_data_validations(const pugi::xml_node& worksheet,
                                                                   ReadDiagnostics* diagnostics = nullptr);

/// Reads `<sheetProtection>` from the sheet and returns the parsed
/// `SheetProtection`. When the element is absent the result has
/// `enabled = false` and every other field at its default. The reader
/// never fails: malformed attribute values fall back to their default
/// (matching pugixml's xsd:boolean conversion behaviour).
SheetProtection read_sheet_protection(const pugi::xml_node& worksheet);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_SHEET_OVERLAY_READER_H_
