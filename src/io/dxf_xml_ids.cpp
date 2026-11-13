#include "io/dxf_xml_ids.h"

#include <cstddef>
#include <optional>

namespace formulon::io {
namespace {

constexpr std::string_view kDxfAttrNames[] = {"dxfId",
                                              "headerRowDxfId",
                                              "dataDxfId",
                                              "totalsRowDxfId",
                                              "headerRowBorderDxfId",
                                              "tableBorderDxfId",
                                              "totalsRowBorderDxfId"};
constexpr std::string_view kIgnoredElements[] = {"slicerStyleElement", "timelineStyleElement"};

/// One dxf attribute with the byte ranges a rewrite splices.
struct Found {
  DxfXmlId id;
  std::size_t lead_begin = 0;  ///< Start of the whitespace before the name.
  std::size_t value_begin = 0;
  std::size_t value_end = 0;
  std::size_t attr_end = 0;  ///< One past the closing quote.
};

Error Malformed(const char* message) {
  return make_error(FormulonErrorCode::kIoXmlParse, message, "context=dxf_xml_ids");
}

bool IsSpace(char c) {
  return c == ' ' || c == '\t' || c == '\r' || c == '\n';
}

bool IsDxfAttr(std::string_view name) {
  for (const std::string_view candidate : kDxfAttrNames) {
    if (name == candidate) {
      return true;
    }
  }
  return false;
}

bool IsIgnoredElement(std::string_view name) {
  const std::size_t colon = name.rfind(':');
  const std::string_view local = colon == std::string_view::npos ? name : name.substr(colon + 1U);
  for (const std::string_view candidate : kIgnoredElements) {
    if (local == candidate) {
      return true;
    }
  }
  return false;
}

std::optional<std::uint32_t> ParseUnsigned(std::string_view text) {
  if (text.empty()) {
    return std::nullopt;
  }
  std::uint64_t value = 0;
  for (const char c : text) {
    if (c < '0' || c > '9') {
      return std::nullopt;
    }
    value = value * 10U + static_cast<std::uint64_t>(c - '0');
    if (value > 0xFFFFFFFFULL) {
      return std::nullopt;
    }
  }
  return static_cast<std::uint32_t>(value);
}

/// Reads the attributes from `pos` up to the tag's `>` / `/>` (`in_tag`) or
/// the end of `text`, recording dxf attributes unless `element` is ignored.
/// Returns the position after the attributes.
Expected<std::size_t, Error> ScanAttributes(std::string_view text, std::size_t pos, std::string_view element,
                                            bool in_tag, std::vector<Found>& out) {
  const bool ignored = in_tag && IsIgnoredElement(element);
  for (;;) {
    const std::size_t lead_begin = pos;
    while (pos < text.size() && IsSpace(text[pos])) {
      ++pos;
    }
    if (pos == text.size()) {
      if (in_tag) {
        return Malformed("unterminated tag in XML fragment");
      }
      return pos;
    }
    const char c = text[pos];
    if (in_tag && c == '>') {
      return pos + 1U;
    }
    if (in_tag && c == '/') {
      if (pos + 1U < text.size() && text[pos + 1U] == '>') {
        return pos + 2U;
      }
      return Malformed("stray '/' in XML tag");
    }
    const std::size_t name_begin = pos;
    while (pos < text.size() && !IsSpace(text[pos]) && text[pos] != '=' && text[pos] != '>' && text[pos] != '/' &&
           text[pos] != '<' && text[pos] != '"' && text[pos] != '\'') {
      ++pos;
    }
    if (pos == name_begin) {
      return Malformed("XML attribute has no name");
    }
    const std::string_view name = text.substr(name_begin, pos - name_begin);
    while (pos < text.size() && IsSpace(text[pos])) {
      ++pos;
    }
    if (pos == text.size() || text[pos] != '=') {
      return Malformed("XML attribute has no value");
    }
    ++pos;
    while (pos < text.size() && IsSpace(text[pos])) {
      ++pos;
    }
    if (pos == text.size() || (text[pos] != '"' && text[pos] != '\'')) {
      return Malformed("XML attribute value is not quoted");
    }
    const std::size_t value_begin = pos + 1U;
    const std::size_t close = text.find(text[pos], value_begin);
    if (close == std::string_view::npos) {
      return Malformed("unterminated XML attribute value");
    }
    pos = close + 1U;
    if (ignored || !IsDxfAttr(name)) {
      continue;
    }
    if (const auto value = ParseUnsigned(text.substr(value_begin, close - value_begin))) {
      out.push_back(Found{DxfXmlId{element, name, *value}, lead_begin, value_begin, close, pos});
    }
  }
}

/// Skips from `pos` past `terminator`, or fails when it never appears.
Expected<std::size_t, Error> SkipPast(std::string_view text, std::size_t pos, std::string_view terminator) {
  const std::size_t end = text.find(terminator, pos);
  if (end == std::string_view::npos) {
    return Malformed("unterminated construct in XML fragment");
  }
  return end + terminator.size();
}

Expected<std::vector<Found>, Error> ScanFragment(std::string_view text) {
  std::vector<Found> out;
  std::size_t pos = 0;
  while (pos < text.size()) {
    const std::size_t open = text.find('<', pos);
    if (open == std::string_view::npos) {
      break;
    }
    const std::string_view rest = text.substr(open);
    Expected<std::size_t, Error> next = open;
    if (rest.substr(0, 4) == "<!--") {
      next = SkipPast(text, open + 4U, "-->");
    } else if (rest.substr(0, 9) == "<![CDATA[") {
      next = SkipPast(text, open + 9U, "]]>");
    } else if (rest.substr(0, 2) == "<?") {
      next = SkipPast(text, open + 2U, "?>");
    } else if (rest.substr(0, 2) == "<!" || rest.substr(0, 2) == "</") {
      next = SkipPast(text, open + 2U, ">");
    } else {
      std::size_t name_end = open + 1U;
      while (name_end < text.size() && !IsSpace(text[name_end]) && text[name_end] != '/' && text[name_end] != '>') {
        ++name_end;
      }
      if (name_end == open + 1U) {
        return Malformed("XML tag has no name");
      }
      next = ScanAttributes(text, name_end, text.substr(open + 1U, name_end - open - 1U), true, out);
    }
    if (!next) {
      return std::move(next.error());
    }
    pos = next.value();
  }
  return out;
}

Expected<std::vector<Found>, Error> ScanAttrList(std::string_view text) {
  std::vector<Found> out;
  auto end = ScanAttributes(text, 0, std::string_view(), false, out);
  if (!end) {
    return std::move(end.error());
  }
  return out;
}

std::vector<DxfXmlId> Ids(const std::vector<Found>& found) {
  std::vector<DxfXmlId> out;
  out.reserve(found.size());
  for (const Found& f : found) {
    out.push_back(f.id);
  }
  return out;
}

std::string Rewrite(std::string_view text, const std::vector<Found>& found, DxfPartCursor& ids) {
  std::string out;
  out.reserve(text.size());
  std::size_t cursor = 0;
  for (const Found& f : found) {
    if (const std::optional<std::uint32_t> saved = ids.next(f.id.value)) {
      out.append(text.substr(cursor, f.value_begin - cursor));
      out.append(std::to_string(*saved));
      cursor = f.value_end;
    } else {
      out.append(text.substr(cursor, f.lead_begin - cursor));
      cursor = f.attr_end;
    }
  }
  out.append(text.substr(cursor));
  return out;
}

}  // namespace

Expected<std::vector<DxfXmlId>, Error> visit_dxf_xml_ids(std::string_view fragment) {
  auto found = ScanFragment(fragment);
  if (!found) {
    return std::move(found.error());
  }
  return Ids(found.value());
}

Expected<std::vector<DxfXmlId>, Error> visit_dxf_attr_ids(std::string_view attrs) {
  auto found = ScanAttrList(attrs);
  if (!found) {
    return std::move(found.error());
  }
  return Ids(found.value());
}

Expected<std::string, Error> rewrite_dxf_xml_ids(std::string_view fragment, DxfPartCursor& cursor) {
  auto found = ScanFragment(fragment);
  if (!found) {
    return std::move(found.error());
  }
  return Rewrite(fragment, found.value(), cursor);
}

Expected<std::string, Error> rewrite_dxf_attr_ids(std::string_view attrs, DxfPartCursor& cursor) {
  auto found = ScanAttrList(attrs);
  if (!found) {
    return std::move(found.error());
  }
  return Rewrite(attrs, found.value(), cursor);
}

}  // namespace formulon::io
