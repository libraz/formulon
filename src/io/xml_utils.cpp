//
// Implementation of the shared OOXML XML helpers. See header for the
// contracts.

#include "io/xml_utils.h"

#include <algorithm>
#include <cerrno>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <cstring>
#include <limits>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/phonetic_pr.h"
#include "io/xml_escape.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/number_text.h"

namespace formulon {
namespace io {

namespace {

void AppendAttrOpen(std::string& out, std::string_view name) {
  out.push_back(' ');
  out.append(name.data(), name.size());
  out.append("=\"");
}

}  // namespace

void append_xml_attr(std::string& out, std::string_view name, std::string_view value) {
  AppendAttrOpen(out, name);
  AppendXmlAttrEscaped(out, value);
  out.push_back('"');
}

void append_xml_attr_uint(std::string& out, std::string_view name, std::uint64_t value) {
  AppendAttrOpen(out, name);
  out.append(std::to_string(value));
  out.push_back('"');
}

void append_xml_attr_int(std::string& out, std::string_view name, std::int64_t value) {
  AppendAttrOpen(out, name);
  out.append(std::to_string(value));
  out.push_back('"');
}

void append_xml_attr_number(std::string& out, std::string_view name, double value) {
  AppendAttrOpen(out, name);
  append_xml_number(out, value);
  out.push_back('"');
}

namespace {

// Decimal digits (no trailing zeros) and exponent of the leading digit.
struct Sci {
  char digits[24];
  int n = 0;
  int exp10 = 0;
};

// Rounds |value| to `sig` significant digits through format_exponential.
Sci SplitSci(double value, int sig) {
  char buf[48];
  const int len = format_exponential(buf, sizeof(buf), value, sig - 1);
  Sci r;
  int i = buf[0] == '-' ? 1 : 0;
  for (; i < len && buf[i] != 'e'; ++i) {
    if (buf[i] != '.') {
      r.digits[r.n++] = buf[i];
    }
  }
  const bool neg = buf[i + 1] == '-';
  for (i += 2; i < len; ++i) {
    r.exp10 = r.exp10 * 10 + (buf[i] - '0');
  }
  if (neg) {
    r.exp10 = -r.exp10;
  }
  while (r.n > 1 && r.digits[r.n - 1] == '0') {
    --r.n;
  }
  return r;
}

}  // namespace

void append_xml_number(std::string& out, double value) {
  // Zero and denormals read back as 0 in Excel.
  if (!(std::fabs(value) >= std::numeric_limits<double>::min())) {
    out.push_back('0');
    return;
  }
  if (value < 0) {
    out.push_back('-');
  }
  const double mag = std::fabs(value);
  Sci s = SplitSci(mag, 17);
  if (s.exp10 < -1 || s.exp10 > 15) {
    out.push_back(s.digits[0]);
    if (s.n > 1) {
      out.push_back('.');
      out.append(s.digits + 1, static_cast<std::size_t>(s.n - 1));
    }
    out.push_back('E');
    out.push_back(s.exp10 < 0 ? '-' : '+');
    char eb[8];
    const int el = format_signed(eb, sizeof(eb), s.exp10 < 0 ? -s.exp10 : s.exp10);
    out.append(eb, static_cast<std::size_t>(el));
    return;
  }
  // Fixed notation: 15 significant digits when they round-trip, else 17.
  char buf[48];
  format_exponential(buf, sizeof(buf), mag, 14);
  if (std::strtod(buf, nullptr) == mag) {
    s = SplitSci(mag, 15);
  }
  if (s.exp10 < 0) {
    out.append("0.");
    out.append(static_cast<std::size_t>(-s.exp10 - 1), '0');
    out.append(s.digits, static_cast<std::size_t>(s.n));
  } else if (s.n <= s.exp10 + 1) {
    out.append(s.digits, static_cast<std::size_t>(s.n));
    out.append(static_cast<std::size_t>(s.exp10 + 1 - s.n), '0');
  } else {
    out.append(s.digits, static_cast<std::size_t>(s.exp10 + 1));
    out.push_back('.');
    out.append(s.digits + s.exp10 + 1, static_cast<std::size_t>(s.n - s.exp10 - 1));
  }
}

bool has_element_child(const pugi::xml_node& node) {
  for (pugi::xml_node child = node.first_child(); child; child = child.next_sibling()) {
    if (child.type() == pugi::node_element) {
      return true;
    }
  }
  return false;
}

std::uint32_t parse_xml_u32_attr(const pugi::xml_attribute& attr, std::uint32_t default_value) {
  return parse_xml_u32_attr_strict(attr).value_or(default_value);
}

std::optional<std::uint32_t> parse_xml_u32_attr_strict(const pugi::xml_attribute& attr) {
  if (!attr) {
    return std::nullopt;
  }
  const char* raw = attr.value();
  if (raw == nullptr || *raw == '\0' || *raw == '-' || *raw == '+') {
    return std::nullopt;
  }
  errno = 0;
  char* end = nullptr;
  const unsigned long parsed = std::strtoul(raw, &end, 10);
  if (end == raw || *end != '\0' || errno != 0 ||
      parsed > static_cast<unsigned long>(std::numeric_limits<std::uint32_t>::max())) {
    return std::nullopt;
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

std::string_view find_namespace_binding(const pugi::xml_node& node, std::string_view prefix) {
  std::string declaration;
  if (prefix.empty()) {
    declaration = "xmlns";
  } else {
    declaration = "xmlns:";
    declaration.append(prefix);
  }
  for (pugi::xml_node scope = node; scope; scope = scope.parent()) {
    if (pugi::xml_attribute decl = scope.attribute(declaration.c_str()); decl) {
      return decl.value();
    }
  }
  return {};
}

void capture_unknown_attrs(const pugi::xml_node& node, std::initializer_list<std::string_view> known,
                           std::vector<std::pair<std::string, std::string>>& out) {
  std::vector<std::pair<std::string, std::string>> captured;
  std::vector<std::pair<std::string, std::string>> bindings;

  for (pugi::xml_attribute attr = node.first_attribute(); attr; attr = attr.next_attribute()) {
    const std::string_view name = attr.name();
    // The writer emits the declarations for the elements it generates;
    // re-emitting the source's copies would duplicate them.
    if (name == "xmlns" || name.rfind("xmlns:", 0) == 0) {
      continue;
    }
    bool is_known = false;
    for (const std::string_view k : known) {
      if (k == name) {
        is_known = true;
        break;
      }
    }
    if (is_known) {
      continue;
    }
    captured.emplace_back(std::string(name), std::string(attr.value()));

    // Carry the prefix's binding across with it. Emitting the attribute
    // without one leaves the prefix unbound, which no XML parser will
    // accept -- so the loss is not confined to the attribute itself.
    const std::size_t colon = name.find(':');
    if (colon == std::string_view::npos) {
      continue;
    }
    const std::string_view prefix = name.substr(0, colon);
    std::string declaration("xmlns:");
    declaration.append(prefix);
    const bool already_bound = std::any_of(bindings.begin(), bindings.end(),
                                           [&declaration](const auto& entry) { return entry.first == declaration; });
    if (already_bound) {
      continue;
    }
    const std::string_view uri = find_namespace_binding(node, prefix);
    if (!uri.empty()) {
      bindings.emplace_back(std::move(declaration), std::string(uri));
    }
  }

  // Declarations first, so the re-emitted element reads the way Excel
  // writes it and a prefix is bound before the eye reaches its user.
  for (auto& entry : bindings) {
    out.emplace_back(std::move(entry.first), std::move(entry.second));
  }
  for (auto& entry : captured) {
    out.emplace_back(std::move(entry.first), std::move(entry.second));
  }
}

namespace {

/// pugixml writer sink that appends straight into a caller-owned string.
struct StringXmlWriter : pugi::xml_writer {
  std::string* dst = nullptr;
  bool saw_cr = false;
  bool escape_pcdata_cr = false;

  void write(const void* data, std::size_t size) override {
    if (dst == nullptr || size == 0U) {
      return;
    }
    const char* bytes = static_cast<const char*>(data);
    const bool contains_cr = std::memchr(bytes, '\r', size) != nullptr;
    saw_cr = saw_cr || contains_cr;
    if (!escape_pcdata_cr || !contains_cr) {
      dst->append(bytes, size);
      return;
    }

    std::size_t start = 0;
    for (std::size_t i = 0; i < size; ++i) {
      if (bytes[i] == '\r') {
        dst->append(bytes + start, i - start);
        dst->append("&#13;");
        start = i + 1U;
      }
    }
    dst->append(bytes + start, size - start);
  }
};

bool HasCarriageReturn(const char* value) {
  return value != nullptr && std::string_view(value).find('\r') != std::string_view::npos;
}

/// Checks this node and its descendants without visiting a sibling of `root`.
bool HasPcdataCarriageReturn(const pugi::xml_node& root) {
  for (pugi::xml_node current = root; current;) {
    if ((current.type() == pugi::node_pcdata || current.type() == pugi::node_element) &&
        HasCarriageReturn(current.value())) {
      return true;
    }
    if (pugi::xml_node child = current.first_child(); child) {
      current = child;
      continue;
    }
    while (current && current != root && !current.next_sibling()) {
      current = current.parent();
    }
    if (!current || current == root) {
      break;
    }
    current = current.next_sibling();
  }
  return false;
}

void AppendPcdataWithCarriageReturnReferences(StringXmlWriter& sink, const pugi::xml_node& pcdata) {
  const bool previous_mode = sink.escape_pcdata_cr;
  sink.escape_pcdata_cr = true;
  pcdata.print(sink, /*indent=*/"", pugi::format_raw);
  sink.escape_pcdata_cr = previous_mode;
}

void AppendElementName(std::string& out, const pugi::xml_node& element) {
  const char* name = element.name();
  if (name == nullptr || *name == '\0') {
    out.append(":anonymous");
  } else {
    out.append(name);
  }
}

void AppendElementStart(StringXmlWriter& sink, const pugi::xml_node& element, bool empty) {
  std::string& out = *sink.dst;
  out.push_back('<');
  AppendElementName(out, element);
  for (pugi::xml_attribute attribute = element.first_attribute(); attribute; attribute = attribute.next_attribute()) {
    append_xml_attr(out, attribute.name(), attribute.value());
  }
  if (empty) {
    out.append("/>");
  } else {
    out.push_back('>');
  }
}

void AppendElementEnd(std::string& out, const pugi::xml_node& element) {
  out.append("</");
  AppendElementName(out, element);
  out.push_back('>');
}

bool IsEmptyElement(const pugi::xml_node& element) {
  if (element.value() != nullptr && *element.value() != '\0') {
    return false;
  }
  const pugi::xml_node first_child = element.first_child();
  if (!first_child) {
    return true;
  }
  return first_child.type() == pugi::node_pcdata && first_child.value() != nullptr && *first_child.value() == '\0' &&
         !first_child.next_sibling();
}

void AppendSlowRawXml(std::string& out, const pugi::xml_node& root) {
  StringXmlWriter sink;
  sink.dst = &out;
  pugi::xml_node current = root;
  while (current) {
    if (current.type() == pugi::node_element) {
      const bool empty = IsEmptyElement(current);
      AppendElementStart(sink, current, empty);
      if (!empty) {
        if (current.value() != nullptr && *current.value() != '\0') {
          // Embedded pcdata; attribute escaping keeps CR as `&#13;` in text content too.
          AppendXmlAttrEscaped(out, current.value());
        }
        if (pugi::xml_node child = current.first_child(); child) {
          current = child;
          continue;
        }
        AppendElementEnd(out, current);
      }
    } else if (current.type() == pugi::node_document) {
      if (pugi::xml_node child = current.first_child(); child) {
        current = child;
        continue;
      }
    } else if (current.type() == pugi::node_pcdata) {
      AppendPcdataWithCarriageReturnReferences(sink, current);
    } else {
      current.print(sink, /*indent=*/"", pugi::format_raw);
    }

    if (current == root) {
      break;
    }
    while (current != root) {
      if (pugi::xml_node sibling = current.next_sibling(); sibling) {
        current = sibling;
        break;
      }
      current = current.parent();
      if (current.type() == pugi::node_element) {
        AppendElementEnd(out, current);
      }
    }
    if (current == root) {
      break;
    }
  }
}

}  // namespace

void append_raw_xml(std::string& out, const pugi::xml_node& node) {
  const std::size_t output_start = out.size();
  StringXmlWriter sink;
  sink.dst = &out;
  node.print(sink, /*indent=*/"", pugi::format_raw);
  if (sink.saw_cr && HasPcdataCarriageReturn(node)) {
    out.resize(output_start);
    AppendSlowRawXml(out, node);
  }
}

std::string raw_xml(const pugi::xml_node& node) {
  std::string out;
  append_raw_xml(out, node);
  return out;
}

namespace {

/// True when `node` is an element whose name is absent from `known`.
bool IsUnknownElementChild(const pugi::xml_node& node, std::initializer_list<std::string_view> known) {
  if (node.type() != pugi::node_element) {
    return false;
  }
  const std::string_view name = node.name();
  for (const std::string_view k : known) {
    if (k == name) {
      return false;
    }
  }
  return true;
}

}  // namespace

void capture_unknown_children(const pugi::xml_node& parent, std::initializer_list<std::string_view> known,
                              std::string& out) {
  for (pugi::xml_node child = parent.first_child(); child; child = child.next_sibling()) {
    if (IsUnknownElementChild(child, known)) {
      append_raw_xml(out, child);
    }
  }
}

void capture_unknown_children(const pugi::xml_node& parent, std::initializer_list<std::string_view> known,
                              std::vector<std::string>& out) {
  for (pugi::xml_node child = parent.first_child(); child; child = child.next_sibling()) {
    if (IsUnknownElementChild(child, known)) {
      out.push_back(raw_xml(child));
    }
  }
}

void append_raw_attrs(std::string& out, const std::vector<std::pair<std::string, std::string>>& attrs) {
  for (const auto& [name, value] : attrs) {
    out.push_back(' ');
    out.append(name);
    out.append("=\"");
    AppendXmlAttrEscaped(out, value);
    out.push_back('"');
  }
}

namespace {

// `parse_ws_pcdata_single` retains whitespace-only text inside leaf
// elements (those with no element children). This is required so a
// whitespace-only string body — e.g. `<t> </t>` in a shared-string,
// inline-string, or comment part — reads back as the literal
// whitespace rather than being silently dropped to "". Because it only
// affects childless elements, it is safe for every reader that shares
// this loader: structural parts read their data from attributes and
// child elements, never from whitespace-only PCDATA in a leaf.
constexpr unsigned int kPartParseFlags = pugi::parse_default | pugi::parse_ws_pcdata_single;

/// Builds the `kIoXmlParse` envelope both part loaders share, so the
/// copying and in-place paths stay indistinguishable to callers that
/// match on the message or context.
Expected<void, Error> PartParseResult(const pugi::xml_parse_result& parse, std::string_view reader_module,
                                      std::string_view part_name) {
  if (parse) {
    return Expected<void, Error>::Ok();
  }
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

}  // namespace

Expected<void, Error> load_xml_buffer(pugi::xml_document& doc, const std::vector<std::uint8_t>& bytes,
                                      std::string_view reader_module, std::string_view part_name) {
  const pugi::xml_parse_result parse =
      doc.load_buffer(bytes.data(), bytes.size(), kPartParseFlags, pugi::encoding_utf8);
  return PartParseResult(parse, reader_module, part_name);
}

Expected<void, Error> load_xml_buffer_inplace(pugi::xml_document& doc, std::vector<std::uint8_t>& bytes,
                                              std::string_view reader_module, std::string_view part_name) {
  // `load_buffer_inplace` (as opposed to `..._own`) leaves ownership of
  // the storage with the caller: pugixml tokenises within `bytes` and
  // never frees it. Pinning the encoding to UTF-8 matches the copying
  // overload and keeps pugixml out of the transcoding path, which is
  // what would silently reintroduce a full-size private copy.
  const pugi::xml_parse_result parse =
      doc.load_buffer_inplace(bytes.data(), bytes.size(), kPartParseFlags, pugi::encoding_utf8);
  return PartParseResult(parse, reader_module, part_name);
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
    AppendOoxmlTextUnescaped(out, t.text().get());
    ++count;
  }

  // Rich-text runs `<r><rPr/><t>...</t></r>` in document order. Formatting
  // attributes on `<rPr>` are intentionally not preserved — this layer is
  // plain-text only.
  for (pugi::xml_node r = node.child("r"); r; r = r.next_sibling("r")) {
    for (pugi::xml_node t = r.child("t"); t; t = t.next_sibling("t")) {
      AppendOoxmlTextUnescaped(out, t.text().get());
      ++count;
    }
  }

  // `<rPh>` phonetic-guide subtrees are deliberately not walked here.
  // Their `<t>` payload is the kana annotation surfaced by `PHONETIC()`
  // through a separate channel; concatenating it into the surface text
  // would silently leak kana into plain comment / SST output.

  return count;
}

void append_phonetic_runs(const pugi::xml_node& node, std::vector<PhoneticRun>& out) {
  for (pugi::xml_node rph = node.child("rPh"); rph; rph = rph.next_sibling("rPh")) {
    PhoneticRun run;
    run.sb = static_cast<std::uint32_t>(rph.attribute("sb").as_uint(0U));
    run.eb = static_cast<std::uint32_t>(rph.attribute("eb").as_uint(0U));
    for (pugi::xml_node t = rph.child("t"); t; t = t.next_sibling("t")) {
      AppendOoxmlTextUnescaped(run.text, t.text().get());
    }
    out.push_back(std::move(run));
  }
}

PhoneticProperties read_phonetic_properties(const pugi::xml_node& node) {
  PhoneticProperties props;
  if (pugi::xml_node pr = node.child("phoneticPr")) {
    props.font_id = static_cast<std::uint16_t>(pr.attribute("fontId").as_uint(0U));
    props.type = parse_phonetic_type(pr.attribute("type").value());
    props.alignment = parse_phonetic_alignment(pr.attribute("alignment").value());
  }
  return props;
}

std::string capture_root_extra_ns_attrs(const pugi::xml_node& root) {
  std::string out;
  for (pugi::xml_attribute attr = root.first_attribute(); attr; attr = attr.next_attribute()) {
    const std::string_view name = attr.name();
    if (name == "xmlns" || name == "xmlns:r") {
      continue;  // the writer always emits these two itself.
    }
    append_xml_attr(out, name, attr.value());
  }
  return out;
}

}  // namespace io
}  // namespace formulon
