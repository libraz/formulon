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

#include <cstddef>
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

/// Disengages every `CFRule::dxf_id` in `formats` that does not address an
/// entry of a `dxf_count`-sized `<dxfs>` table.
///
/// The attribute is read verbatim, so a third-party writer's stale index
/// -- or a `dxfId="-1"`, which arrives as `0xFFFFFFFF` -- reaches the
/// model naming a record that is not there. Callers are entitled to
/// assume a stored index resolves: that is the same promise
/// `NormalizeStyleIndices` keeps for the `<xf>` tables, and without it a
/// workbook that loaded cleanly hands out a `dxf_id` its own styles getter
/// then rejects.
///
/// Unlike an `<xf>` index this disengages rather than clamping to record
/// 0. A CF rule with no `dxfId` is an ordinary, meaningful state (every
/// data-bar, colour-scale and icon-set rule is one), whereas record 0 is
/// a formatting choice the author never made. `dxf_count` of 0 therefore
/// disengages every rule rather than being a special case.
///
/// The styles table has to be loaded before this runs; the OOXML reader
/// sequences it that way.
void normalize_cf_dxf_ids(std::vector<cf::ConditionalFormat>& formats, std::size_t dxf_count);

}  // namespace formulon::io

#endif  // FORMULON_IO_CF_READER_H_
