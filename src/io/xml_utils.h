// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Tiny shared helper for OOXML rich-text traversal. The `<si>`, `<is>`,
// and `<text>` elements that appear in `xl/sharedStrings.xml`,
// `<c><is>...</is></c>` inline strings, and `xl/comments*.xml` all share
// the same plain-text-extraction rules: walk every direct `<t>` child and
// every nested `<r><t>` rich-text run in document order, optionally
// skipping `<rPh>` phonetic-guide subtrees so the surface text never
// inherits kana annotations.
//
// Consolidates the previously triplicate walkers in `cell_parser.cpp`,
// `sst_reader.cpp`, and `comments_reader.cpp`.

#ifndef FORMULON_IO_XML_UTILS_H_
#define FORMULON_IO_XML_UTILS_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {

/// XML 1.0 declaration prepended to every OOXML part Formulon emits.
/// Centralised so `ooxml_writer.cpp`, `comments_writer.cpp`, and any
/// future part writer share a single byte-for-byte source of truth.
inline constexpr std::string_view kXmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

/// Appends ` name="value"` to `out`, escaping `value` for XML attribute
/// context. `name` is assumed to be a writer-controlled token in the
/// `[A-Za-z][A-Za-z0-9_:-]*` range and is emitted verbatim. The leading
/// space is included so callers can chain attributes onto a tag prefix
/// without worrying about delimiters.
void append_xml_attr(std::string& out, std::string_view name, std::string_view value);

/// Numeric variant of `append_xml_attr`: appends ` name="value"` with
/// `value` rendered via `std::to_string`. No escaping pass is required
/// because the rendered digits are always XML-safe.
void append_xml_attr_uint(std::string& out, std::string_view name, std::uint32_t value);

/// Parses a non-negative decimal integer attribute body. Missing, empty,
/// malformed, signed, or out-of-range input returns `default_value` so a
/// stray optional OOXML attribute does not reject the whole part.
std::uint32_t parse_xml_u32_attr(const pugi::xml_attribute& attr, std::uint32_t default_value);

/// Parses a signed 32-bit decimal attribute body. Missing, empty,
/// malformed, or out-of-range input returns `default_value`.
std::int32_t parse_xml_i32_attr(const pugi::xml_attribute& attr, std::int32_t default_value);

/// Parses an OOXML boolean attribute body. Returns true for "1" or
/// "true", false for everything else (including a missing attribute).
/// Matches the lexicon Excel emits for the majority of `xs:boolean`
/// attributes. For the case-insensitive variant Excel uses on a small
/// subset of sheet flags, see `sheet_reader.cpp`'s local helper.
bool parse_xml_bool_attr(const pugi::xml_attribute& attr);

/// String-view variant of `parse_xml_bool_attr`. Returns true for
/// "1" or "true", false otherwise. Use when the caller already has
/// the attribute value as `std::string_view`.
bool parse_xml_bool(std::string_view value);

// ---------------------------------------------------------------------------
// Node + attribute-name typed accessors.
//
// These complement the `pugi::xml_attribute` overloads above and remove the
// `node.attribute("foo").value()` / `as_uint(0)` boilerplate that is repeated
// in 100+ sites across the OOXML readers. All five overloads share the same
// shape:
//
//   T attr_<type>(const pugi::xml_node& n, const char* name, T def = ...);
//
// On a missing or empty attribute they return `def`. On a malformed attribute
// they return `def` as well — none of these is a structural validator. The
// caller still has the option to call `node.attribute(name)` directly when
// it must distinguish "absent" from "present but empty" (e.g. for OOXML
// flags whose presence alone is meaningful).
// ---------------------------------------------------------------------------

/// Returns the attribute value as a string_view, or `def` if absent or
/// empty. The returned view points into pugixml-owned memory and is
/// valid for the lifetime of the parent `xml_document`.
inline std::string_view attr_str(const pugi::xml_node& n, const char* name, std::string_view def = {}) {
  const pugi::xml_attribute a = n.attribute(name);
  if (!a) {
    return def;
  }
  const char* v = a.value();
  if (v == nullptr || *v == '\0') {
    return def;
  }
  return v;
}

/// Returns the attribute value as `uint32_t`, or `def` on missing /
/// empty / malformed / negative / out-of-range input. Behaviour
/// matches the legacy `parse_xml_u32_attr(node.attribute(name), def)`
/// shape so migration is mechanical.
std::uint32_t attr_u32(const pugi::xml_node& n, const char* name, std::uint32_t def = 0);

/// Returns the attribute value as `int32_t`, or `def` on missing /
/// empty / malformed / out-of-range input. Behaviour matches the
/// legacy `parse_xml_i32_attr(node.attribute(name), def)` shape.
std::int32_t attr_i32(const pugi::xml_node& n, const char* name, std::int32_t def = 0);

/// Returns the attribute value as `double`, or `def` on missing /
/// empty / malformed input. Delegates to pugi's `as_double(def)`.
inline double attr_f64(const pugi::xml_node& n, const char* name, double def = 0.0) {
  return n.attribute(name).as_double(def);
}

/// Returns the attribute value as `bool`, or `def` if absent. "1" and
/// any ASCII-case-insensitive spelling of "true" map to true; everything
/// else (including "0" / "false" / unknown text) maps to false. Matches
/// the OOXML `xs:boolean` lexicon Excel emits across sheet flags, CF
/// rule attributes, and pivot toggles.
bool attr_bool(const pugi::xml_node& n, const char* name, bool def = false);

/// Parses a UTF-8 OOXML part body into `doc`. On failure returns a
/// `kIoXmlParse` error whose context records `<reader_module>` and
/// `<part_name>` and includes pugixml's description. The standard
/// message body is `"<part_name>: pugixml parse failed"`.
///
/// Centralises the parse-error envelope reproduced verbatim by every
/// reader. Callers wishing to use a non-default parse mode or build a
/// bespoke error message must continue to call `load_buffer` directly.
Expected<void, Error> load_xml_buffer(pugi::xml_document& doc, const std::vector<std::uint8_t>& bytes,
                                      std::string_view reader_module, std::string_view part_name);

/// Appends every `<t>` text descendant of `node` (whether a direct child or
/// inside a nested `<r>` rich-text run) to `out`, in document order.
/// `<rPh>` phonetic-guide subtrees are always skipped: kana annotations
/// belong to a separate surface (`PHONETIC()`) and would silently leak
/// into plain comment / SST text otherwise.
///
/// Returns the number of `<t>` elements consumed; callers that treat
/// "no text payload at all" as a structural error (sst_reader) inspect
/// this count rather than `out.empty()`, since a `<t></t>` legitimately
/// contributes zero bytes but is still a valid payload.
std::size_t append_rich_text(const pugi::xml_node& node, std::string& out);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XML_UTILS_H_
