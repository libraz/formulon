#include "io/xlsb/ignored_error_records.h"

#include <array>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/sqref_text.h"
#include "io/xlsb/cell_watch_records.h"
#include "io/xlsb/feature_formula.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "utils/structured_log.h"

namespace formulon::io::xlsb {
namespace {

constexpr std::string_view kIgnoredErrors = "ignoredErrors";

struct IgnoredErrorFlag {
  const char* attr;
  std::uint32_t bit;
};

/// The BrtCellIgnoreEC flag bits, in the attribute order Excel writes on `<ignoredError>`.
constexpr std::array<IgnoredErrorFlag, 9> kFlags{{
    {"evalError", 0x001U},
    {"twoDigitTextYear", 0x020U},
    {"numberStoredAsText", 0x004U},
    {"formula", 0x010U},
    {"formulaRange", 0x008U},
    {"unlockedFormula", 0x040U},
    {"emptyCellReference", 0x002U},
    {"listDataValidation", 0x080U},
    {"calculatedColumn", 0x100U},
}};
constexpr std::uint32_t kMappedBits = 0x1FFU;

bool IsIgnoredErrors(const WorksheetRawChild& child) {
  return child.xml.compare(0, kIgnoredErrors.size() + 1U, "<ignoredErrors") == 0;
}

bool PrecedesIgnoredErrors(std::uint16_t type) {
  return precedes_cell_watches(type) || type == kBrtBeginCellWatches || type == kBrtCellWatch ||
         type == kBrtEndCellWatches;
}

/// The BrtCellIgnoreEC payload of one `<ignoredError>`; nullopt when the element is not one the records can carry.
std::optional<std::vector<std::uint8_t>> EncodeIgnoredError(const pugi::xml_node& element) {
  std::vector<MergeRange> ranges;
  if (!parse_sqref_text(attr_str(element, "sqref"), ranges)) {
    return std::nullopt;
  }
  std::uint32_t flags = 0;
  for (const pugi::xml_attribute attr : element.attributes()) {
    const std::string_view name = attr.name();
    if (name == "sqref") {
      continue;
    }
    const IgnoredErrorFlag* match = nullptr;
    for (const IgnoredErrorFlag& flag : kFlags) {
      if (name == flag.attr) {
        match = &flag;
      }
    }
    if (match == nullptr) {
      return std::nullopt;
    }
    if (parse_xml_bool_attr(attr)) {
      flags |= match->bit;
    }
  }
  if (flags == 0U) {
    return std::nullopt;
  }
  std::vector<std::uint8_t> payload;
  emit_u32(payload, flags);
  emit_sqref(payload, ranges);
  return payload;
}

}  // namespace

std::optional<std::string> decode_ignored_error(ByteSpan payload) {
  if (payload.size < 4U) {
    return std::nullopt;
  }
  const std::uint32_t flags = load_u32(payload.data);
  if (flags == 0U || (flags & ~kMappedBits) != 0U) {
    return std::nullopt;
  }
  ByteSpan cursor{payload.data + 4, payload.size - 4U};
  std::vector<MergeRange> ranges;
  if (!read_sqref(cursor, ranges) || cursor.size != 0U) {
    return std::nullopt;
  }
  std::string xml = "<ignoredError";
  append_xml_attr(xml, "sqref", format_sqref_text(ranges));
  for (const IgnoredErrorFlag& flag : kFlags) {
    if ((flags & flag.bit) != 0U) {
      append_xml_attr(xml, flag.attr, "1");
    }
  }
  xml.append("/>");
  return xml;
}

void write_ignored_errors(WorksheetRawExtensions& children, const std::vector<std::string>& elements) {
  for (std::size_t i = children.size(); i-- > 0;) {
    if (IsIgnoredErrors(children[i])) {
      children.erase(children.begin() + static_cast<std::ptrdiff_t>(i));
    }
  }
  if (elements.empty()) {
    return;
  }
  std::string xml = "<ignoredErrors>";
  for (const std::string& element : elements) {
    xml.append(element);
  }
  xml.append("</ignoredErrors>");
  insert_raw_child(children, WorksheetRawChild{static_cast<std::uint32_t>(worksheet_child::slot_of(kIgnoredErrors)),
                                               std::move(xml)});
}

void splice_ignored_errors(std::vector<std::uint8_t>& after_hyperlinks, const WorksheetRawExtensions& children) {
  std::vector<std::uint8_t> block;
  for (const WorksheetRawChild& child : children) {
    pugi::xml_document doc;
    if (!IsIgnoredErrors(child) || !doc.load_buffer(child.xml.data(), child.xml.size())) {
      continue;
    }
    for (const pugi::xml_node element : doc.child(kIgnoredErrors.data()).children("ignoredError")) {
      const std::optional<std::vector<std::uint8_t>> payload = EncodeIgnoredError(element);
      if (!payload) {
        StructuredLog("xlsb.ignored_error.skipped").field("sqref", attr_str(element, "sqref")).warn();
        continue;
      }
      emit_record(block, kBrtCellIgnoreEC, *payload);
    }
  }
  insert_framed_block(after_hyperlinks, block, PrecedesIgnoredErrors, kBrtBeginCellIgnoreECs, kBrtEndCellIgnoreECs);
}

}  // namespace formulon::io::xlsb
