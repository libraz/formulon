//
// Writer for `<conditionalFormatting>` blocks inside an OOXML sheet
// part. Symmetric counterpart of `src/io/cf_reader.{h,cpp}`: feeding
// the bytes produced here back into the reader must reproduce the
// input list of `cf::ConditionalFormat` records.
//
// Coverage parity with the reader:
//   * All 17 `<cfRule type=...>` spellings.
//   * `cellIs` (8 operators), `top10`, `aboveAverage`, `containsText` /
//     `beginsWith` / `endsWith` / `notContainsText` (with the `text`
//     attribute), `timePeriod` (10 buckets).
//   * Visual rule subtrees: `<colorScale>`, `<dataBar>`, `<iconSet>`.
//   * `<color rgb="AARRGGBB">` always emits the alpha byte (writer
//     prefers the canonical 8-hex form).
//   * `<cfvo>` thresholds with `gte` boundary attribute.
//
// A `DataBarSpec` field that the legacy `<dataBar>` element cannot
// express (negative fill / borders / axis position / axis colour /
// solid-vs-gradient) lives only in the Excel 2010+ `x14` extension, in
// two places at once: a link on the rule itself and a payload element at
// worksheet level. This writer emits the link; the payload is built by
// `build_x14_cf_overlay_entries` and merged into the worksheet
// `<extLst>` by `merge_x14_cf_entries` (see `src/io/cf_overlay.h`),
// because that block is a sibling of `<conditionalFormatting>` rather
// than a child. For a rule loaded from an x14-bearing file both halves
// are already present as captured raw XML and are re-emitted verbatim
// instead of being rebuilt.
//
// Design references:
//   * src/io/cf_reader.h (sister reader; canonical grammar)
//   * src/io/cf_overlay.h (worksheet-level overlay reconciliation)

#ifndef FORMULON_IO_CF_WRITER_H_
#define FORMULON_IO_CF_WRITER_H_

#include <cstddef>
#include <string>
#include <vector>

#include "cf/cf_types.h"

namespace formulon::io {

/// Emits all `<conditionalFormatting>` blocks for one sheet as a single
/// concatenated XML chunk. The output has neither outer XML declaration
/// nor `<worksheet>` wrapper — it is meant to be inlined into
/// `BuildWorksheetXml` between `<sheetData>` and `<tableParts>` (the
/// document-order slot ECMA-376 reserves for CF; see §18.3 of the
/// spec).
///
/// Returns an empty string when `formats` is empty so the caller can
/// concatenate unconditionally without a wrapper element.
///
/// `dxf_count` is the number of `<dxf>` records the same package's
/// `xl/styles.xml` will contain. A rule whose `dxf_id` is not below it
/// names no differential format, so the `dxfId` attribute is omitted
/// rather than written dangling: Excel treats an unresolvable `dxfId`
/// as package corruption and repairs the sheet by discarding *all* of
/// its conditional formatting, which costs far more than the one rule's
/// formatting. The rule itself is still emitted.
std::string write_conditional_formattings(const std::vector<cf::ConditionalFormat>& formats, std::size_t dxf_count);

/// Builds the `<x14:conditionalFormatting>` entries that carry the
/// data-bar settings the legacy `<dataBar>` element cannot express, for
/// every rule in `formats` that needs one. Returns the entries
/// concatenated, with no enclosing `<x14:conditionalFormattings>` /
/// `<ext>` / `<extLst>` wrapper — `merge_x14_cf_entries` owns the
/// wrapping, since it also has to fold them into whatever overlay the
/// source file already had.
///
/// Returns an empty string when no rule needs an entry, which is the
/// common case: a rule whose data bar is expressible in the legacy
/// element alone produces no extension bytes, so a file that never had
/// an x14 overlay does not grow one.
///
/// A rule loaded from an x14-bearing file is included here too. Its
/// entry is a rebuild of what the source file carried, and
/// `merge_x14_cf_entries` drops it in favour of the captured original;
/// deciding that here would need the worksheet `<extLst>`, which this
/// writer does not see.
std::string build_x14_cf_overlay_entries(const std::vector<cf::ConditionalFormat>& formats);

}  // namespace formulon::io

#endif  // FORMULON_IO_CF_WRITER_H_
