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

#include "pugixml.hpp"

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
