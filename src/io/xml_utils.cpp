// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of `append_rich_text`. See header for the contract.

#include "io/xml_utils.h"

#include <cerrno>
#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <limits>
#include <string>
#include <string_view>

#include "io/xml_escape.h"
#include "pugixml.hpp"

namespace formulon {
namespace io {

void append_xml_attr(std::string& out, std::string_view name, std::string_view value) {
  out.push_back(' ');
  out.append(name.data(), name.size());
  out.append("=\"");
  AppendXmlEscaped(out, value);
  out.append("\"");
}

void append_xml_attr_uint(std::string& out, std::string_view name, std::uint32_t value) {
  out.push_back(' ');
  out.append(name.data(), name.size());
  out.append("=\"");
  out.append(std::to_string(value));
  out.append("\"");
}

std::uint32_t parse_xml_u32_attr(const pugi::xml_attribute& attr, std::uint32_t default_value) {
  if (!attr) {
    return default_value;
  }
  const char* raw = attr.value();
  if (raw == nullptr || *raw == '\0') {
    return default_value;
  }
  if (*raw == '-' || *raw == '+') {
    return default_value;
  }
  errno = 0;
  char* end = nullptr;
  const unsigned long parsed = std::strtoul(raw, &end, 10);
  if (end == raw || *end != '\0' || errno != 0 ||
      parsed > static_cast<unsigned long>(std::numeric_limits<std::uint32_t>::max())) {
    return default_value;
  }
  return static_cast<std::uint32_t>(parsed);
}

std::int32_t parse_xml_i32_attr(const pugi::xml_attribute& attr, std::int32_t default_value) {
  if (!attr) {
    return default_value;
  }
  const char* raw = attr.value();
  if (raw == nullptr || *raw == '\0') {
    return default_value;
  }
  errno = 0;
  char* end = nullptr;
  const long parsed = std::strtol(raw, &end, 10);
  if (end == raw || *end != '\0' || errno != 0 ||
      parsed > static_cast<long>(std::numeric_limits<std::int32_t>::max()) ||
      parsed < static_cast<long>(std::numeric_limits<std::int32_t>::min())) {
    return default_value;
  }
  return static_cast<std::int32_t>(parsed);
}

std::size_t append_rich_text(const pugi::xml_node& node, std::string& out) {
  std::size_t count = 0;

  // Direct `<t>` children: the simple inlineStr / shared-string / comment
  // shape (`<is><t>...</t></is>`).
  for (pugi::xml_node t = node.child("t"); t; t = t.next_sibling("t")) {
    out.append(t.text().get());
    ++count;
  }

  // Rich-text runs `<r><rPr/><t>...</t></r>` in document order. Formatting
  // attributes on `<rPr>` are intentionally not preserved — this layer is
  // plain-text only.
  for (pugi::xml_node r = node.child("r"); r; r = r.next_sibling("r")) {
    for (pugi::xml_node t = r.child("t"); t; t = t.next_sibling("t")) {
      out.append(t.text().get());
      ++count;
    }
  }

  // `<rPh>` phonetic-guide subtrees are deliberately not walked here.
  // Their `<t>` payload is the kana annotation surfaced by `PHONETIC()`
  // through a separate channel; concatenating it into the surface text
  // would silently leak kana into plain comment / SST output.

  return count;
}

}  // namespace io
}  // namespace formulon
