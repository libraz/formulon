//
// Tri-state helpers for XSD `boolean` attributes whose schema default is
// a fixed value (usually `true` or `false`). OOXML relies heavily on the
// "absence means the schema default, presence means the explicit value"
// rule: `sheetProtection`'s eleven lock flags, page-break `man`
// attributes, and several pivot hints all follow it. Reading such an
// attribute with a plain `attr.as_bool()` loses the distinction between
// "absent" (→ default) and "present as 0" (→ explicit false), and naive
// writers that only emit when `value == true` silently flip any attribute
// whose default is `true` back to its default when the source set it to
// `false`.
//
// `read_xsd_bool` resolves the read side (absent → caller-supplied
// default); `emit_xsd_bool_attr` resolves the write side (always emit an
// explicit `0` / `1` so the round-trip is lossless regardless of the
// schema default). The `std::string_view` overload of the parser is
// pugixml-free so the lexical rules can be unit-tested in isolation.

#ifndef FORMULON_IO_XSD_BOOL_H_
#define FORMULON_IO_XSD_BOOL_H_

#include <string>
#include <string_view>

#include "pugixml.hpp"

namespace formulon {
namespace io {

/// Parses one XSD `boolean` lexical value. The XSD lexical space is
/// `{true, false, 1, 0}`; comparison is case-insensitive for the
/// alphabetic forms to tolerate producers that emit `True` / `FALSE`.
/// Any value outside the lexical space (including the empty string)
/// yields `default_value` rather than guessing.
inline bool parse_xsd_bool(std::string_view text, bool default_value) {
  if (text == "1") {
    return true;
  }
  if (text == "0") {
    return false;
  }
  auto iequals = [](std::string_view a, std::string_view b) {
    if (a.size() != b.size()) {
      return false;
    }
    for (std::size_t i = 0; i < a.size(); ++i) {
      char ca = a[i];
      if (ca >= 'A' && ca <= 'Z') {
        ca = static_cast<char>(ca - 'A' + 'a');
      }
      if (ca != b[i]) {
        return false;
      }
    }
    return true;
  };
  if (iequals(text, "true")) {
    return true;
  }
  if (iequals(text, "false")) {
    return false;
  }
  return default_value;
}

/// Reads attribute `attr` off `node` as an XSD boolean, returning
/// `default_value` when the attribute is absent. A present-but-malformed
/// value also falls back to `default_value` (see `parse_xsd_bool`).
inline bool read_xsd_bool(const pugi::xml_node& node, const char* attr, bool default_value) {
  const pugi::xml_attribute a = node.attribute(attr);
  if (!a) {
    return default_value;
  }
  return parse_xsd_bool(a.value(), default_value);
}

/// Appends ` name="1"` or ` name="0"` (leading space included) to `out`.
/// Always emits an explicit value so the attribute round-trips losslessly
/// no matter what the schema default is; callers that want to omit the
/// attribute when it equals the default should guard the call themselves.
inline void emit_xsd_bool_attr(std::string& out, std::string_view name, bool value) {
  out.push_back(' ');
  out.append(name.data(), name.size());
  out.append(value ? "=\"1\"" : "=\"0\"");
}

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XSD_BOOL_H_
