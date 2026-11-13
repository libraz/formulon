// Implementation of the worksheet print records. See
// `io/xlsb/print_records.h` for the contract.

#include "io/xlsb/print_records.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "pugixml.hpp"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

// BrtPrintOptions bits.
constexpr std::uint32_t kHorizontalCentered = 0x01U;
constexpr std::uint32_t kVerticalCentered = 0x02U;
constexpr std::uint32_t kHeadings = 0x04U;
constexpr std::uint32_t kGridLines = 0x08U;
constexpr std::uint32_t kGridLinesSet = 0x10U;

/// The root element of a raw print fragment; empty when it does not parse.
pugi::xml_node ParseRoot(pugi::xml_document& doc, const std::string& xml, const char* name) {
  if (xml.empty() || !doc.load_buffer(xml.data(), xml.size())) {
    return pugi::xml_node();
  }
  return doc.child(name);
}

std::uint16_t PrintOptionsFlags(const std::string& xml) {
  pugi::xml_document doc;
  const pugi::xml_node node = ParseRoot(doc, xml, "printOptions");
  const auto bit = [&node](const char* name, bool fallback, std::uint32_t mask) {
    return node.attribute(name).as_bool(fallback) ? mask : 0U;
  };
  const std::uint32_t flags = bit("gridLinesSet", true, kGridLinesSet) |
                              bit("horizontalCentered", false, kHorizontalCentered) |
                              bit("verticalCentered", false, kVerticalCentered) | bit("headings", false, kHeadings) |
                              bit("gridLines", false, kGridLines);
  // Five bits, within u16.
  return static_cast<std::uint16_t>(flags);
}

void EmitMargins(std::vector<std::uint8_t>& dst, const std::string& xml) {
  pugi::xml_document doc;
  const pugi::xml_node node = ParseRoot(doc, xml, "pageMargins");
  const PageMargins defaults;
  std::vector<std::uint8_t> payload;
  emit_double(payload, node.attribute("left").as_double(defaults.left));
  emit_double(payload, node.attribute("right").as_double(defaults.right));
  emit_double(payload, node.attribute("top").as_double(defaults.top));
  emit_double(payload, node.attribute("bottom").as_double(defaults.bottom));
  emit_double(payload, node.attribute("header").as_double(defaults.header));
  emit_double(payload, node.attribute("footer").as_double(defaults.footer));
  emit_record(dst, kBrtMargins, payload);
}

}  // namespace

bool model_owns_print_records(const SheetPrintSettings& print) {
  return !print.print_options_xml.empty() || !print.page_margins_xml.empty();
}

void splice_print_records(std::vector<std::uint8_t>& after_hyperlinks, const SheetPrintSettings& print) {
  if (!model_owns_print_records(print)) {
    return;
  }
  const std::vector<FramedRecord> recs = split_records(after_hyperlinks);
  if (!after_hyperlinks.empty() && recs.empty()) {
    return;
  }
  std::vector<std::uint8_t> out;
  std::vector<std::uint8_t> options;
  emit_u16(options, PrintOptionsFlags(print.print_options_xml));
  emit_record(out, kBrtPrintOptions, options);
  if (!print.page_margins_xml.empty()) {
    EmitMargins(out, print.page_margins_xml);
  }
  for (const FramedRecord& rec : recs) {
    if (rec.type != kBrtPrintOptions && rec.type != kBrtMargins) {
      out.insert(out.end(), after_hyperlinks.begin() + static_cast<std::ptrdiff_t>(rec.begin),
                 after_hyperlinks.begin() + static_cast<std::ptrdiff_t>(rec.end));
    }
  }
  after_hyperlinks = std::move(out);
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
