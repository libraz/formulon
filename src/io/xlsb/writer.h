// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// MS-XLSB (.xlsb) package writer. The write-side counterpart to
// `io/xlsb/reader.h`: walks a `Workbook` and produces a complete OPC
// package (ZIP container of `[Content_Types].xml`, `_rels/.rels`,
// `xl/workbook.bin`, `xl/worksheets/sheet*.bin`, optional
// `xl/sharedStrings.bin`, plus any passthrough parts the workbook
// preserved from a previous read).
//
// Round-trip with the reader is the primary correctness contract:
// `read_xlsb(write_xlsb(wb))` reproduces every cell value `wb` carried
// in for the in-scope literal kinds (number, boolean, text, error,
// blank). Formula cells round-trip via the Bundle 4.1 stub: a
// `formula_text` of the form `=__FORMULON_XLSB_PTG__(...)` is
// recognised here and the captured Ptg bytes are spliced back into a
// `BrtFmla*` record. Authored-in-engine formulas (no Ptg bytes
// available) currently fall back to writing the cached value as a
// literal and emit a `xlsb.writer.formula_lost` warning; full
// AST→Ptg encoding lands in a later bundle.
//
// Defined names, tables, conditional formats, styles, and other
// sheet-level metadata are out of scope for v1: if the workbook
// carries any, we log `xlsb.writer.deferred` and skip them. The OOXML
// writer remains the canonical round-trip path for those features.
//
// Design references:
//   * backup/plans/04-xlsx-io.md §4.6 (XLSB)
//   * backup/plans/21-xlsb-ptg.md §21.5 (reader/writer symmetry)
//   * backup/plans/26-implementation-plan.md (Phase 4 sequencing)

#ifndef FORMULON_IO_XLSB_WRITER_H_
#define FORMULON_IO_XLSB_WRITER_H_

#include <cstdint>
#include <vector>

#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Serialises `workbook` into an in-memory `.xlsb` byte stream.
///
/// Always-emitted parts:
///   * `[Content_Types].xml`           — Default extensions for `rels`,
///                                       `xml`, `bin`; per-part Override
///                                       entries for the workbook,
///                                       worksheets, and SST when
///                                       present.
///   * `_rels/.rels`                   — package-level rels pointing at
///                                       `xl/workbook.bin`.
///   * `xl/_rels/workbook.bin.rels`    — workbook-level rels (worksheets
///                                       + optional SST).
///   * `xl/workbook.bin`               — `BrtBeginBook | BrtBeginBundleShs
///                                       | BrtBundleSh* | BrtEndBundleShs
///                                       | BrtEndBook`.
///   * `xl/worksheets/sheet<N>.bin`    — one per sheet, 1-based.
///
/// Conditionally emitted parts:
///   * `xl/sharedStrings.bin`          — only when at least one cell
///                                       interns a text value.
///   * Passthrough parts                — every entry in
///                                       `workbook.passthrough_parts()`
///                                       that does not collide with a
///                                       generated path. Collisions are
///                                       resolved in favour of the
///                                       generated copy with a
///                                       `xlsb.writer.passthrough_collision`
///                                       warning.
///
/// Returns `FormulonErrorCode::kIoWriteFailed` on any miniz failure;
/// the error context identifies the offending part. Returns
/// `kInvalidArgument` when `workbook.sheet_count() == 0`.
Expected<std::vector<std::uint8_t>, Error> write_xlsb(const Workbook& workbook);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_WRITER_H_
