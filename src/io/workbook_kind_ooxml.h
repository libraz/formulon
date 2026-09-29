//
// OOXML serialisation helpers for `WorkbookKind` (declared in
// `workbook_kind.h`): the kind only differs from a plain `.xlsx` at the
// `[Content_Types].xml` boundary, where the workbook part's content-type
// string and the package's default extension change.
//
// Design references:
//   * [OPC] / [ECMA-376] for the canonical content-type strings

#ifndef FORMULON_IO_WORKBOOK_KIND_OOXML_H_
#define FORMULON_IO_WORKBOOK_KIND_OOXML_H_

#include "workbook_kind.h"

namespace formulon {
namespace io {

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

#endif  // FORMULON_IO_WORKBOOK_KIND_OOXML_H_
