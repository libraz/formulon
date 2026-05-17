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
#include <utility>
#include <vector>

#include "io/xml_escape.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

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

bool parse_xml_bool(std::string_view value) {
  return value == "1" || value == "true";
}

std::uint32_t attr_u32(const pugi::xml_node& n, const char* name, std::uint32_t def) {
  return parse_xml_u32_attr(n.attribute(name), def);
}

std::int32_t attr_i32(const pugi::xml_node& n, const char* name, std::int32_t def) {
  return parse_xml_i32_attr(n.attribute(name), def);
}

bool attr_bool(const pugi::xml_node& n, const char* name, bool def) {
  const pugi::xml_attribute a = n.attribute(name);
  if (!a) {
    return def;
  }
  const char* v = a.value();
  if (v == nullptr || *v == '\0') {
    return def;
  }
  // Fast path for the canonical "1" / "0" Excel emits.
  if (v[0] == '1' && v[1] == '\0') {
    return true;
  }
  if (v[0] == '0' && v[1] == '\0') {
    return false;
  }
  // Case-insensitive 4-char "true". Anything else (including "false",
  // "yes", garbage) maps to false — matches the lenient legacy
  // `parse_xml_bool_attr` lexicon.
  if (v[0] != '\0' && v[1] != '\0' && v[2] != '\0' && v[3] != '\0' && v[4] == '\0') {
    const auto lower = [](char c) -> char { return (c >= 'A' && c <= 'Z') ? static_cast<char>(c + ('a' - 'A')) : c; };
    if (lower(v[0]) == 't' && lower(v[1]) == 'r' && lower(v[2]) == 'u' && lower(v[3]) == 'e') {
      return true;
    }
  }
  return false;
}

bool parse_xml_bool_attr(const pugi::xml_attribute& attr) {
  if (!attr) {
    return false;
  }
  return parse_xml_bool(attr.value());
}

Expected<void, Error> load_xml_buffer(pugi::xml_document& doc, const std::vector<std::uint8_t>& bytes,
                                      std::string_view reader_module, std::string_view part_name) {
  pugi::xml_parse_result parse = doc.load_buffer(bytes.data(), bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    std::string ctx("context=");
    ctx.append(reader_module);
    ctx.append(" part=");
    ctx.append(part_name);
    ctx.append(" desc=");
    ctx.append(parse.description());
    std::string msg(part_name);
    msg.append(": pugixml parse failed");
    return make_error(FormulonErrorCode::kIoXmlParse, std::move(msg), std::move(ctx));
  }
  return Expected<void, Error>::Ok();
}

std::uint32_t parse_rgb_hex(std::string_view hex, std::uint32_t fallback) noexcept {
  if (hex.size() != 6 && hex.size() != 8) {
    return fallback;
  }
  std::uint32_t out = 0;
  for (char c : hex) {
    std::uint32_t digit = 0;
    if (c >= '0' && c <= '9') {
      digit = static_cast<std::uint32_t>(c - '0');
    } else if (c >= 'a' && c <= 'f') {
      digit = static_cast<std::uint32_t>(c - 'a' + 10);
    } else if (c >= 'A' && c <= 'F') {
      digit = static_cast<std::uint32_t>(c - 'A' + 10);
    } else {
      return fallback;
    }
    out = (out << 4U) | digit;
  }
  if (hex.size() == 6) {
    out |= 0xFF000000U;
  }
  return out;
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
