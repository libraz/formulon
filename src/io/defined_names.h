// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
// layer; both arrive in Phase 4.
//
// Design references:
//   * backup/plans/04-xlsx-io.md (defined names section)
//   * backup/plans/26-implementation-plan.md (Phase 2.4)

#ifndef FORMULON_IO_DEFINED_NAMES_H_
#define FORMULON_IO_DEFINED_NAMES_H_

#include <cstdint>
#include <string>
#include <vector>

#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

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

/// Parses every `<definedName>` child of `<workbook>/<definedNames>` in
/// the supplied workbook document.
///
/// Behaviour:
///   * A workbook with no `<definedNames>` block — or with an empty
///     block — yields an empty vector. This is not an error: most
///     workbooks have no defined names at all.
///   * Declaration order is preserved: callers (and the eventual
///     writer) can rely on `result[i]` matching the i-th `<definedName>`
///     in the source document.
///   * The `name` attribute is required. A `<definedName>` element
///     without a `name=` is a corruption (Excel rejects such input), so
///     the reader surfaces `kIoSheetCorrupt`.
///   * `localSheetId="N"` sets `local_sheet_id` to the parsed integer
///     (any non-numeric or negative value is treated as workbook scope,
///     i.e. `-1`).
///   * `hidden="1"` / `hidden="true"` (case-insensitive) sets
///     `hidden = true`. Any other value (or absent attribute) leaves it
///     `false`.
///   * `comment="..."` is captured verbatim; absent/empty becomes "".
///   * Formula text comes from the element's text payload. Leading and
///     trailing ASCII whitespace is stripped (writers occasionally pad
///     the body for readability); interior whitespace is preserved as-
///     authored.
///
/// Errors:
///   * `kIoXmlParse` — should not occur in this function (the caller
///     hands in an already-parsed `xml_document`); reserved for future
///     paths that might re-parse a sub-document.
///   * `kIoSheetCorrupt` — a `<definedName>` lacks the required `name=`
///     attribute.
Expected<std::vector<DefinedName>, Error> read_defined_names(const pugi::xml_document& workbook_doc);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_DEFINED_NAMES_H_
