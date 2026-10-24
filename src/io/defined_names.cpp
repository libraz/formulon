//
// Implementation of the defined-names reader. See defined_name.h for the
// public contract.

#include <cctype>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "io/defined_names_internal.h"
#include "io/future_functions.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "io/xsd_bool.h"
#include "parser/ast_format.h"
#include "parser/formula_prefix.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace {

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
  const auto parsed = parse_xml_u32_attr_strict(attr);
  if (!parsed || *parsed > static_cast<std::uint32_t>(INT32_MAX)) {
    return -1;
  }
  return static_cast<std::int32_t>(*parsed);
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
    entry.hidden = read_xsd_bool(dn, "hidden", false);
    if (pugi::xml_attribute comment_attr = dn.attribute("comment"); comment_attr) {
      entry.comment = comment_attr.value();
    }
    {
      std::vector<std::pair<std::string, std::string>> extra;
      capture_unknown_attrs(dn, {"name", "localSheetId", "hidden", "comment"}, extra);
      append_raw_attrs(entry.extra_attrs, extra);
    }

    // Element text payload — pugixml exposes child text nodes via
    // `.child_value()` for the first text child, which is what writers
    // emit for `<definedName>...</definedName>`. `child_value()` has
    // already undergone XML entity decoding; `AppendOoxmlTextUnescaped`
    // additionally decodes the OOXML `_xHHHH_` control-character escapes
    // the writer emits for this slot, matching every other element-text
    // reader. Strip framing whitespace after that; interior whitespace
    // stays put so round-tripping a formula that genuinely contains
    // spaces is faithful.
    entry.formula.clear();
    AppendOoxmlTextUnescaped(entry.formula, dn.child_value());
    TrimAsciiWhitespace(entry.formula);
    // Canonicalise the storage prefixes Excel tags post-2007 / LET-LAMBDA
    // constructs with (`_xlfn.`, `_xlfn._xlws.`, `_xlws.`, `_xlpm.`) so a
    // defined name holding `_xlfn.LAMBDA(_xlpm.x, ...)` is stored and
    // surfaced in `formula_text` the way Excel's formula bar shows it, and
    // so downstream consumers that re-parse `entry.formula` do not need to
    // special-case the prefixed spelling. A stored `SINGLE(x)` /
    // `ANCHORARRAY(x)` reads back as the `@x` / `x#` the formula bar shows.
    entry.formula = parser::spell_storage_operators(parser::strip_storage_prefixes(entry.formula, &has_storage_prefix));

    out.push_back(std::move(entry));
  }

  return out;
}

}  // namespace io
}  // namespace formulon
