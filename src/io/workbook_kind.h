// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// `WorkbookKind`: discriminator for the four OOXML workbook variants the
// reader/writer pipeline rounds-trips end-to-end. The engine treats all
// four exactly like a plain `.xlsx` for cell content — they share the
// same workbook / worksheet / sharedStrings / styles schemas — and only
// differs at the `[Content_Types].xml` boundary, where the workbook
// part's content-type string and the package's default extension change.
//
// Macro-enabled variants (`.xlsm` / `.xltm`) additionally carry a
// `xl/vbaProject.bin` payload. The engine NEVER executes VBA; the
// payload is preserved verbatim via the existing passthrough mechanism
// (see `passthrough_part.h`) so Excel can re-open the file with macros
// intact.
//
// Design references:
//   * [OPC] / [ECMA-376] for the canonical content-type strings

#ifndef FORMULON_IO_WORKBOOK_KIND_H_
#define FORMULON_IO_WORKBOOK_KIND_H_

#include <cstdint>

namespace formulon {
namespace io {

/// Discriminator for the OOXML workbook variants the reader/writer
/// pipeline understands. The default is `kXlsx`; the macro-enabled and
/// template variants are detected at read time from the workbook part's
/// content-type string in `[Content_Types].xml` and re-emitted on write.
enum class WorkbookKind : std::uint8_t {
  kXlsx = 0,  ///< Plain `.xlsx` workbook.
  kXlsm = 1,  ///< Macro-enabled workbook (`.xlsm`). Carries `xl/vbaProject.bin`.
  kXltx = 2,  ///< Template (`.xltx`). No macros.
  kXltm = 3,  ///< Macro-enabled template (`.xltm`). Carries `xl/vbaProject.bin`.
};

/// Returns the OOXML content-type string used by the workbook part for
/// `kind`. The returned pointer references a static string literal with
/// program lifetime.
///
/// These are the four canonical strings declared in [OPC] part 1 §10 /
/// [ECMA-376]; any other content type the reader encounters is treated
/// as `kXlsx` and surfaces a structured-log warning rather than failing.
inline const char* workbook_kind_content_type(WorkbookKind kind) {
  switch (kind) {
    case WorkbookKind::kXlsx:
      return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    case WorkbookKind::kXlsm:
      return "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
    case WorkbookKind::kXltx:
      return "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml";
    case WorkbookKind::kXltm:
      return "application/vnd.ms-excel.template.macroEnabled.main+xml";
  }
  // Fallback: should be unreachable, but the engine builds with
  // -fno-exceptions and warnings-as-errors, so we return the safe
  // default rather than UB.
  return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
}

/// Returns the canonical file extension associated with `kind` (no
/// leading dot). Useful for callers that derive an output file name from
/// the workbook kind. The returned pointer references a static string
/// literal with program lifetime.
inline const char* workbook_kind_default_extension(WorkbookKind kind) {
  switch (kind) {
    case WorkbookKind::kXlsx:
      return "xlsx";
    case WorkbookKind::kXlsm:
      return "xlsm";
    case WorkbookKind::kXltx:
      return "xltx";
    case WorkbookKind::kXltm:
      return "xltm";
  }
  return "xlsx";
}

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_WORKBOOK_KIND_H_
