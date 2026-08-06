//
// Writer for `<conditionalFormatting>` blocks inside an OOXML sheet
// part. Symmetric counterpart of `src/io/cf_reader.{h,cpp}`: feeding
// the bytes produced here back into the reader must reproduce the
// input list of `cf::ConditionalFormat` records.
//
// Coverage parity with the reader (PR2):
//   * All 17 `<cfRule type=...>` spellings.
//   * `cellIs` (8 operators), `top10`, `aboveAverage`, `containsText` /
//     `beginsWith` / `endsWith` / `notContainsText` (with the `text`
//     attribute), `timePeriod` (10 buckets).
//   * Visual rule subtrees: `<colorScale>`, `<dataBar>`, `<iconSet>`.
//   * `<color rgb="AARRGGBB">` always emits the alpha byte (writer
//     prefers the canonical 8-hex form).
//   * `<cfvo>` thresholds with `gte` boundary attribute.
//
// Out of scope:
//   * Synthesising a *new* worksheet-level `<x14:conditionalFormattings>`
//     overlay for a `DataBarSpec` whose negative-fill / axis / gradient
//     fields were set programmatically (never loaded from an x14-bearing
//     file). `<cf_reader.h>` decodes an existing overlay into
//     `DataBarSpec`, but this writer only re-emits the legacy `<dataBar>`
//     element here; the worksheet-level overlay text for a *loaded* file
//     survives a save cycle unchanged via `Sheet::ext_lst_xml()`'s raw
//     passthrough (see `BuildWorksheetXml`), not through this writer.
//
// Design references:
//   * src/io/cf_reader.h (sister reader; canonical grammar)

#ifndef FORMULON_IO_CF_WRITER_H_
#define FORMULON_IO_CF_WRITER_H_

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
std::string write_conditional_formattings(const std::vector<cf::ConditionalFormat>& formats);

}  // namespace formulon::io

#endif  // FORMULON_IO_CF_WRITER_H_
