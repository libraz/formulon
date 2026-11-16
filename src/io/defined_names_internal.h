//
// Internal-only entry point for the defined-names reader. Kept in a
// separate header so the public `defined_name.h` (which carries the
// `DefinedName` struct itself) does not have to expose `pugixml.hpp`.
// Only the OOXML reader and other `io/`-internal pipelines should
// include this header.
//
// The public `DefinedName` struct lives in `defined_name.h` and
// remains the round-trip carrier for the workbook layer.

#ifndef FORMULON_IO_DEFINED_NAMES_INTERNAL_H_
#define FORMULON_IO_DEFINED_NAMES_INTERNAL_H_

#include <vector>

#include "defined_name.h"
#include "io/package_diagnostics.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {

/// Parses every `<definedName>` child of `<workbook>/<definedNames>` in
/// the supplied workbook document. See `defined_name.h` for the
/// `DefinedName` shape and the per-attribute semantics; this is the
/// pugi-typed entry point intended for `io/`-internal callers.
///
/// Without a macro sheet, Excel refuses to open a book holding a name flagged
/// `xlm="1"` and drops a lone `function="1"` or `vbProcedure="1"` on save, so
/// without `book_has_macro_sheet` those flags leave the name, one count per
/// name in `diagnostics->skipped_feature_count`.
Expected<std::vector<DefinedName>, Error> read_defined_names(const pugi::xml_document& workbook_doc,
                                                             bool book_has_macro_sheet = true,
                                                             ReadDiagnostics* diagnostics = nullptr);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_DEFINED_NAMES_INTERNAL_H_
