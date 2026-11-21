#include "io/xlsb/scenario_records.h"

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/sqref_text.h"
#include "io/xlsb/feature_formula.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "utils/a1_ref.h"
#include "utils/structured_log.h"

namespace formulon::io::xlsb {
namespace {

constexpr std::string_view kScenarios = "scenarios";
constexpr std::uint16_t kAbsentIndex = 0xFFFFU;
constexpr std::uint32_t kAbsentSqref = 0xFFFFFFFFU;
constexpr std::size_t kSctHead = 10U;
constexpr std::size_t kSctLocked = 2U;
constexpr std::size_t kSctHidden = 6U;
constexpr std::size_t kSlcHead = 18U;

bool IsScenarios(const WorksheetRawChild& child) {
  return child.xml.compare(0, kScenarios.size() + 1U, "<scenarios") == 0;
}

/// A u32 that holds 0 or 1, as the record's flag fields do.
std::optional<bool> FlagU32(const std::uint8_t* p) {
  const std::uint32_t v = load_u32(p);
  if (v > 1U) {
    return std::nullopt;
  }
  return v == 1U;
}

bool DecodeManager(ByteSpan payload, std::string& xml) {
  if (payload.size < 8U) {
    return false;
  }
  const std::uint16_t current = load_u16(payload.data);
  const std::uint16_t show = load_u16(payload.data + 2);
  ByteSpan cursor{payload.data + 4, payload.size - 4U};
  xml.append("<scenarios");
  if (current != kAbsentIndex) {
    append_xml_attr_uint(xml, "current", current);
  }
  if (show != kAbsentIndex) {
    append_xml_attr_uint(xml, "show", show);
  }
  if (load_u32(cursor.data) == kAbsentSqref) {
    if (cursor.size != 4U) {
      return false;
    }
  } else {
    std::vector<MergeRange> ranges;
    if (!read_sqref(cursor, ranges) || cursor.size != 0U) {
      return false;
    }
    append_xml_attr(xml, "sqref", format_sqref_text(ranges));
  }
  xml.append(">");
  return true;
}

/// Appends `<scenario ...>` for a BrtBeginSct payload; the input-cell count it states goes to `cells`.
bool DecodeScenario(ByteSpan payload, std::string& xml, std::uint16_t& cells) {
  if (payload.size < kSctHead) {
    return false;
  }
  const std::uint8_t* p = payload.data;
  const std::uint8_t locked = p[kSctLocked];
  const std::uint8_t hidden = p[kSctHidden];
  if (locked > 1U || hidden > 1U || p[3] != 0U || p[4] != 0U || p[5] != 0U || p[7] != 0U || p[8] != 0U || p[9] != 0U) {
    return false;
  }
  cells = load_u16(p);
  ByteSpan cursor{p + kSctHead, payload.size - kSctHead};
  auto name = read_xlwidestring(cursor);
  if (!name || cursor.size < 4U) {
    return false;
  }
  std::optional<std::string> comment;
  if (load_u32(cursor.data) == 0xFFFFFFFFU) {
    cursor.data += 4;
    cursor.size -= 4U;
  } else {
    auto text = read_xlwidestring(cursor);
    if (!text) {
      return false;
    }
    comment = std::move(text.value());
  }
  auto user = read_xlwidestring(cursor);
  if (!user || cursor.size != 0U) {
    return false;
  }
  xml.append("<scenario");
  append_xml_attr(xml, "name", name.value());
  if (locked == 1U) {
    append_xml_attr(xml, "locked", "1");
  }
  if (hidden == 1U) {
    append_xml_attr(xml, "hidden", "1");
  }
  append_xml_attr_uint(xml, "count", cells);
  if (!user.value().empty()) {
    append_xml_attr(xml, "user", user.value());
  }
  if (comment) {
    append_xml_attr(xml, "comment", *comment);
  }
  xml.append(">");
  return true;
}

bool DecodeInputCell(ByteSpan payload, std::string& xml) {
  if (payload.size < kSlcHead) {
    return false;
  }
  const std::uint8_t* p = payload.data;
  const std::uint32_t row = load_u32(p);
  const std::uint32_t col = load_u32(p + 4);
  const std::optional<bool> deleted = FlagU32(p + 8);
  const std::optional<bool> undone = FlagU32(p + 12);
  const std::uint16_t num_fmt = load_u16(p + 16);
  if (row >= a1::kMaxRows || col >= a1::kMaxCols || !deleted || !undone) {
    return false;
  }
  ByteSpan cursor{p + kSlcHead, payload.size - kSlcHead};
  auto val = read_xlwidestring(cursor);
  if (!val || cursor.size != 0U) {
    return false;
  }
  xml.append("<inputCells");
  append_xml_attr(xml, "r", a1::encode_a1(row, col));
  if (*deleted) {
    append_xml_attr(xml, "deleted", "1");
  }
  if (*undone) {
    append_xml_attr(xml, "undone", "1");
  }
  append_xml_attr(xml, "val", val.value());
  if (num_fmt != 0U) {
    append_xml_attr_uint(xml, "numFmtId", num_fmt);
  }
  xml.append("/>");
  return true;
}

bool KnownAttrs(const pugi::xml_node& node, std::initializer_list<std::string_view> names) {
  for (const pugi::xml_attribute attr : node.attributes()) {
    bool known = false;
    for (std::string_view name : names) {
      known = known || name == attr.name();
    }
    if (!known) {
      return false;
    }
  }
  return true;
}

/// A 0/1 xsd boolean attribute; nullopt for any other spelling.
std::optional<bool> BoolAttr(const pugi::xml_node& node, const char* name) {
  const std::string_view v = attr_str(node, name, "0");
  if (v == "1" || v == "true") {
    return true;
  }
  if (v == "0" || v == "false") {
    return false;
  }
  return std::nullopt;
}

/// An optional u16 index attribute: kAbsentIndex when absent, nullopt when it does not fit below it.
std::optional<std::uint16_t> IndexAttr(const pugi::xml_node& node, const char* name) {
  const pugi::xml_attribute attr = node.attribute(name);
  if (!attr) {
    return kAbsentIndex;
  }
  const std::optional<std::uint32_t> v = parse_xml_u32_attr_strict(attr);
  if (!v || *v >= kAbsentIndex) {
    return std::nullopt;
  }
  return static_cast<std::uint16_t>(*v);
}

bool EncodeInputCell(const pugi::xml_node& cell, std::vector<std::uint8_t>& out) {
  if (!KnownAttrs(cell, {"r", "deleted", "undone", "val", "numFmtId"}) || !cell.attribute("val")) {
    return false;
  }
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  const std::optional<bool> deleted = BoolAttr(cell, "deleted");
  const std::optional<bool> undone = BoolAttr(cell, "undone");
  std::uint32_t num_fmt = 0;
  if (const pugi::xml_attribute attr = cell.attribute("numFmtId")) {
    const std::optional<std::uint32_t> v = parse_xml_u32_attr_strict(attr);
    if (!v || *v > 0xFFFFU) {
      return false;
    }
    num_fmt = *v;
  }
  if (!a1::parse_a1_ref(attr_str(cell, "r"), &row, &col) || !deleted || !undone) {
    return false;
  }
  std::vector<std::uint8_t> payload;
  emit_u32(payload, row);
  emit_u32(payload, col);
  emit_u32(payload, *deleted ? 1U : 0U);
  emit_u32(payload, *undone ? 1U : 0U);
  emit_u16(payload, static_cast<std::uint16_t>(num_fmt));
  emit_xlwidestring(payload, attr_str(cell, "val"));
  emit_record(out, kBrtSlc, payload);
  return true;
}

bool EncodeScenario(const pugi::xml_node& scenario, std::vector<std::uint8_t>& out) {
  if (!KnownAttrs(scenario, {"name", "locked", "hidden", "count", "user", "comment"}) || !scenario.attribute("name")) {
    return false;
  }
  const std::optional<bool> locked = BoolAttr(scenario, "locked");
  const std::optional<bool> hidden = BoolAttr(scenario, "hidden");
  std::vector<pugi::xml_node> cells;
  for (const pugi::xml_node child : scenario.children()) {
    if (child.type() != pugi::node_element || std::string_view(child.name()) != "inputCells") {
      return false;
    }
    cells.push_back(child);
  }
  const std::optional<std::uint32_t> count = parse_xml_u32_attr_strict(scenario.attribute("count"));
  if (!locked || !hidden || cells.size() > 0xFFFFU || (count && *count != cells.size())) {
    return false;
  }
  std::vector<std::uint8_t> payload;
  emit_u16(payload, static_cast<std::uint16_t>(cells.size()));
  const std::uint8_t flags[kSctHead - 2U] = {*locked ? std::uint8_t{1} : std::uint8_t{0}, 0, 0, 0,
                                             *hidden ? std::uint8_t{1} : std::uint8_t{0}, 0, 0, 0};
  payload.insert(payload.end(), flags, flags + sizeof(flags));
  emit_xlwidestring(payload, attr_str(scenario, "name"));
  emit_xlnullablewidestring(payload, scenario.attribute("comment")
                                         ? std::optional<std::string_view>(attr_str(scenario, "comment"))
                                         : std::nullopt);
  emit_xlwidestring(payload, attr_str(scenario, "user"));
  emit_record(out, kBrtBeginSct, payload);
  for (const pugi::xml_node& cell : cells) {
    if (!EncodeInputCell(cell, out)) {
      return false;
    }
  }
  emit_record(out, kBrtEndSct, ByteSpan{});
  return true;
}

bool EncodeScenarios(const pugi::xml_node& root, std::vector<std::uint8_t>& out) {
  if (!KnownAttrs(root, {"current", "show", "sqref"})) {
    return false;
  }
  const std::optional<std::uint16_t> current = IndexAttr(root, "current");
  const std::optional<std::uint16_t> show = IndexAttr(root, "show");
  if (!current || !show) {
    return false;
  }
  std::vector<std::uint8_t> payload;
  emit_u16(payload, *current);
  emit_u16(payload, *show);
  if (root.attribute("sqref")) {
    std::vector<MergeRange> ranges;
    if (!parse_sqref_text(attr_str(root, "sqref"), ranges)) {
      return false;
    }
    emit_sqref(payload, ranges);
  } else {
    emit_u32(payload, kAbsentSqref);
  }
  emit_record(out, kBrtBeginScenMan, payload);
  for (const pugi::xml_node child : root.children()) {
    if (child.type() != pugi::node_element || std::string_view(child.name()) != "scenario" ||
        !EncodeScenario(child, out)) {
      return false;
    }
  }
  emit_record(out, kBrtEndScenMan, ByteSpan{});
  return true;
}

}  // namespace

std::optional<std::string> decode_scenarios(const std::vector<XlsbRecord>& records) {
  if (records.empty() || records.front().type != kBrtBeginScenMan) {
    return std::nullopt;
  }
  std::string xml;
  if (!DecodeManager(records.front().payload, xml)) {
    return std::nullopt;
  }
  std::size_t i = 1;
  while (i < records.size()) {
    std::uint16_t cells = 0;
    if (records[i].type != kBrtBeginSct || !DecodeScenario(records[i].payload, xml, cells)) {
      return std::nullopt;
    }
    ++i;
    for (std::uint16_t c = 0; c < cells; ++c, ++i) {
      if (i >= records.size() || records[i].type != kBrtSlc || !DecodeInputCell(records[i].payload, xml)) {
        return std::nullopt;
      }
    }
    if (i >= records.size() || records[i].type != kBrtEndSct || records[i].payload.size != 0U) {
      return std::nullopt;
    }
    ++i;
    xml.append("</scenario>");
  }
  xml.append("</scenarios>");
  return xml;
}

void write_scenarios(WorksheetRawExtensions& children, std::string xml) {
  for (std::size_t i = children.size(); i-- > 0;) {
    if (IsScenarios(children[i])) {
      children.erase(children.begin() + static_cast<std::ptrdiff_t>(i));
    }
  }
  insert_raw_child(children,
                   WorksheetRawChild{static_cast<std::uint32_t>(worksheet_child::slot_of(kScenarios)), std::move(xml)});
}

void splice_scenarios(std::vector<std::uint8_t>& before_merges, const WorksheetRawExtensions& children) {
  std::vector<std::uint8_t> block;
  for (const WorksheetRawChild& child : children) {
    pugi::xml_document doc;
    if (!IsScenarios(child) || !doc.load_buffer(child.xml.data(), child.xml.size())) {
      continue;
    }
    if (!EncodeScenarios(doc.child(kScenarios.data()), block)) {
      StructuredLog("xlsb.scenarios.skipped").warn();
      return;
    }
  }
  if (block.empty()) {
    return;
  }
  for (const FramedRecord& rec : split_records(before_merges)) {
    if (rec.type == kBrtBeginScenMan) {
      return;
    }
  }
  before_merges.insert(before_merges.begin(), block.begin(), block.end());
}

}  // namespace formulon::io::xlsb
