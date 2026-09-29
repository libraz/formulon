//
// `WorkbookKind`: discriminator for the four OOXML workbook variants the
// reader/writer pipeline rounds-trips end-to-end. The engine treats all
// four exactly like a plain `.xlsx` for cell content — they share the
// same workbook / worksheet / sharedStrings / styles schemas.
//
// Macro-enabled variants (`.xlsm` / `.xltm`) additionally carry a
// `xl/vbaProject.bin` payload. The engine NEVER executes VBA; the
// payload is preserved verbatim via the existing passthrough mechanism
// (see `passthrough_part.h`) so Excel can re-open the file with macros
// intact.

#ifndef FORMULON_WORKBOOK_KIND_H_
#define FORMULON_WORKBOOK_KIND_H_

#include <cstdint>

namespace formulon {

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

}  // namespace formulon

#endif  // FORMULON_WORKBOOK_KIND_H_
