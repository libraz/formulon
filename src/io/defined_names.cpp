// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the defined-names reader. See defined_names.h for the
// public contract.

#include "io/defined_names.h"

#include <cctype>
#include <cstdint>
#include <cstdlib>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace {

/// Returns `true` when `value` is the canonical OOXML truthy form for a
/// boolean attribute: `"1"` or `"true"` (case-insensitive). The OOXML
/// boolean lexicon also nominally accepts `"yes"` / `"on"` but Excel
/// itself never emits those for `<definedName hidden=...>`, so we
/// stay consistent with what the writer round-trips.
bool IsXmlBoolTrue(std::string_view value) {
  if (value == "1") {
    return true;
  }
  if (value.size() != 4) {
    return false;
  }
  // Compare against "true" case-insensitively without dragging in
  // <algorithm> for a single 4-byte check.
  return (value[0] == 't' || value[0] == 'T') && (value[1] == 'r' || value[1] == 'R') &&
         (value[2] == 'u' || value[2] == 'U') && (value[3] == 'e' || value[3] == 'E');
}

/// Strips leading and trailing ASCII whitespace from `s` in place. We
/// intentionally do NOT touch interior whitespace: a defined name like
/// `"=SUM( A1, A2 )"` should round-trip unchanged.
void TrimAsciiWhitespace(std::string& s) {
  std::size_t start = 0;
  while (start < s.size() && std::isspace(static_cast<unsigned char>(s[start])) != 0) {
    ++start;
  }
  std::size_t end = s.size();
  while (end > start && std::isspace(static_cast<unsigned char>(s[end - 1])) != 0) {
    --end;
  }
  if (start > 0 || end < s.size()) {
    s = s.substr(start, end - start);
  }
}

/// Parses `localSheetId` if present. OOXML guarantees this is a non-
/// negative integer when emitted; we still defend against malformed
/// input by returning `-1` (workbook scope) on any parse failure rather
/// than rejecting the entire workbook for what is purely passive
/// metadata at this layer.
std::int32_t ParseLocalSheetId(const pugi::xml_attribute& attr) {
  if (!attr) {
    return -1;
  }
  const char* raw = attr.value();
  if (raw == nullptr || *raw == '\0') {
    return -1;
  }
  char* end = nullptr;
  // Use long here so we can range-check without depending on errno.
  long parsed = std::strtol(raw, &end, 10);
  if (end == raw || *end != '\0' || parsed < 0) {
    return -1;
  }
  // Excel caps workbooks well below INT32_MAX sheets; clamp to be safe.
  if (parsed > static_cast<long>(INT32_MAX)) {
    return -1;
  }
  return static_cast<std::int32_t>(parsed);
}

}  // namespace

Expected<std::vector<DefinedName>, Error> read_defined_names(const pugi::xml_document& workbook_doc) {
  std::vector<DefinedName> out;

  pugi::xml_node wb = workbook_doc.child("workbook");
  if (!wb) {
    // Caller should have validated the root before calling us, but a
    // missing root here is "no defined names", not an error: this layer
    // is called purely for metadata extraction and treats absence as
    // empty.
    return out;
  }
  pugi::xml_node names = wb.child("definedNames");
  if (!names) {
    return out;
  }

  for (pugi::xml_node dn = names.child("definedName"); dn; dn = dn.next_sibling("definedName")) {
    pugi::xml_attribute name_attr = dn.attribute("name");
    if (!name_attr || name_attr.value() == nullptr || *name_attr.value() == '\0') {
      return make_error(FormulonErrorCode::kIoSheetCorrupt,
                        "workbook.xml: <definedName> missing required name attribute", "context=defined_names_reader");
    }

    DefinedName entry;
    entry.name = name_attr.value();
    entry.local_sheet_id = ParseLocalSheetId(dn.attribute("localSheetId"));
    if (pugi::xml_attribute hidden_attr = dn.attribute("hidden"); hidden_attr) {
      entry.hidden = IsXmlBoolTrue(hidden_attr.value());
    }
    if (pugi::xml_attribute comment_attr = dn.attribute("comment"); comment_attr) {
      entry.comment = comment_attr.value();
    }

    // Element text payload — pugixml exposes child text nodes via
    // `.child_value()` for the first text child, which is what writers
    // emit for `<definedName>...</definedName>`. Strip framing
    // whitespace; interior whitespace stays put so round-tripping a
    // formula that genuinely contains spaces is faithful.
    entry.formula = dn.child_value();
    TrimAsciiWhitespace(entry.formula);

    out.push_back(std::move(entry));
  }

  return out;
}

}  // namespace io
}  // namespace formulon
