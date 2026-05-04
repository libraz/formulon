// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// `CalcMode`: workbook-level calculation policy mirrored from Excel's
// `<calcPr calcMode="...">` attribute. Lives in `formulon::io` so the
// type can be referenced by `Workbook`, the OOXML reader/writer, and
// the C ABI without any of those layers having to include the full
// `workbook.h`. This mirrors the prior split-out of `WorkbookKind`
// (see `workbook_kind.h`) and keeps `Workbook` from being a giant
// dependency root for header consumers that only need the enum.
//
// Design references:
//   * backup/plans/04-xlsx-io.md (`<calcPr>` round-trip)
//   * `workbook_kind.h` (sibling enum split-out)

#ifndef FORMULON_IO_CALC_MODE_H_
#define FORMULON_IO_CALC_MODE_H_

#include <cstdint>

namespace formulon {
namespace io {

/// Excel calc-mode enum. Mirrors the `calcMode` attribute on `<calcPr>`
/// (`auto` / `manual` / `autoNoTable`).
///
/// Plain metadata: the engine itself does not gate evaluation on this
/// setting (every `recalc()` call honours all dirty cells). The value
/// is preserved as round-trip metadata and surfaced through the
/// bindings so a host UI can mirror Excel's user-visible state.
enum class CalcMode : std::uint8_t {
  kAuto = 0,
  kManual = 1,
  kAutoNoTable = 2,
};

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_CALC_MODE_H_
