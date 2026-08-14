//
// Reader for `<conditionalFormatting>` blocks inside an OOXML sheet
// (`xl/worksheets/sheet*.xml`). Walks the worksheet DOM and decodes
// every CF block into `cf::ConditionalFormat` records. Consults the
// worksheet-level `<extLst><ext><x14:conditionalFormattings>` overlay
// (Excel 2010+) for `dataBar` rules only: negative-fill / negative-
// border / axis colour+position / gradient-vs-solid have no
// representation in the legacy `<dataBar>` schema and are only ever
// expressed there, cross-referenced to the legacy `<cfRule id="...">`
// GUID. Other x14 rule types (icon-set custom criteria, data-bar
// exponential scaling beyond what's modelled here) are not consulted.
//
// Design references:
//   * src/io/pivot_table_reader.h (sister reader, similar style)

#ifndef FORMULON_IO_CF_READER_H_
#define FORMULON_IO_CF_READER_H_

#include <vector>

#include "cf/cf_types.h"
#include "io/package_diagnostics.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon::io {

/// Walks a `<worksheet>` node and returns every `<conditionalFormatting>`
/// block parsed into a `cf::ConditionalFormat`. Returns an empty vector
/// when the sheet has no CF blocks.
///
/// Behaviour:
///   * `<conditionalFormatting sqref="A1:A10 D5:D15" pivot="0">` —
///     `sqref` is split on whitespace and each token decoded into a
///     `CFCellRange`. Absolute markers (`$A$1:$A$10`) are accepted.
///     Column letters must be upper case in every token shape — cell
///     (`A1:A10`), whole-column (`A:A`) and whole-row (`1:3`) alike —
///     matching what Excel emits. `pivot="1"` sets `pivot_scope`.
///   * A block with a missing / empty / unparseable `sqref` is skipped
///     with a WARN diagnostic and load continues — CF is a presentation
///     overlay and one bad block must not reject the workbook.
///   * `<cfRule type="..." priority="N" stopIfTrue="..." dxfId="...">`
///     attributes feed the corresponding `CFRule` fields. Unknown
///     `type` attributes fold the rule to `Expression` and continue
///     (forward-compat with future Excel additions); the `formula1` is
///     captured if present.
///   * `<formula>` children populate `formula1` / `formula2` in
///     document order; the leading `=` is stripped if present.
///   * `<colorScale>`, `<dataBar>`, `<iconSet>` children populate the
///     respective payload optionals. `<cfvo>` thresholds and `<color
///     rgb="...">` entries are walked in document order.
///   * Visual rule subtree errors fold to defaults (no hard fail) so a
///     malformed colour scale does not reject the entire sheet.
///
/// This reader does not return `kIoSheetCorrupt` for CF-local problems:
/// malformed blocks are skipped, never fatal.
///
/// Each skipped block bumps `diagnostics->skipped_feature_count` when
/// `diagnostics` is non-NULL, so a caller that never reads the structured
/// log still learns the load was lossy. Passing NULL discards the count.
Expected<std::vector<cf::ConditionalFormat>, Error> read_conditional_formats(const pugi::xml_node& worksheet,
                                                                             ReadDiagnostics* diagnostics = nullptr);

}  // namespace formulon::io

#endif  // FORMULON_IO_CF_READER_H_
