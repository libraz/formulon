//
// Defined-name reader. Walks `<workbook>/<definedNames>/<definedName>`
// in the parsed `xl/workbook.xml` document and produces a flat list of
// metadata entries suitable for round-trip preservation. The defined
// formulas are NOT validated or evaluated at this layer — that is the
// parser/evaluator's job at use-site. This reader exists so the OOXML
// reader can stash the metadata on the workbook for the writer slice
// (Bundle 2.5) to emit it back unchanged.
//
// Structured-reference resolution (e.g. `Table[@col]`) and named-range
// resolution at evaluation time are explicitly out of scope at this
// layer; both arrive in a follow-up.

#ifndef FORMULON_IO_DEFINED_NAMES_H_
#define FORMULON_IO_DEFINED_NAMES_H_

#include <cstdint>
#include <string>

namespace formulon {
namespace io {

/// In-memory representation of one `<definedName>` element from
/// `xl/workbook.xml`. `name` is preserved with its authored case
/// (although Excel matches names case-insensitively at use-site);
/// `formula` is the raw formula text (Excel sometimes wraps the body in
/// surrounding whitespace, which the reader strips). `local_sheet_id`
/// follows the OOXML attribute semantics: `-1` means workbook scope,
/// `>= 0` is a 0-based sheet index for sheet-scoped names. `comment`
/// is the optional human-readable note Excel exposes in the Name
/// Manager dialog.
struct DefinedName {
  std::string name;
  std::string formula;
  std::int32_t local_sheet_id = -1;
  bool hidden = false;
  std::string comment;
};

// The pugi-typed `read_defined_names(...)` reader entry point lives in
// `io/defined_names_internal.h`. That header is `io/`-internal and is
// the only place pugixml types appear in the defined-name pipeline; the
// public surface above is just the round-trip carrier struct.

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_DEFINED_NAMES_H_
