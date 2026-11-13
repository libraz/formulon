// Implementation of the worksheet print records. See
// `io/xlsb/print_records.h` for the contract.

#include "io/xlsb/print_records.h"

#include <array>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
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

// BrtPageSetup flag bits, and the values Excel writes for an absent attribute.
constexpr std::uint32_t kOverThenDown = 0x0001U;
constexpr std::uint32_t kLandscape = 0x0002U;
constexpr std::uint32_t kBlackAndWhite = 0x0008U;
constexpr std::uint32_t kDraft = 0x0010U;
constexpr std::uint32_t kCellComments = 0x0020U;
constexpr std::uint32_t kUseFirstPageNumber = 0x0080U;
constexpr std::uint32_t kCommentsAtEnd = 0x0100U;
constexpr unsigned kErrorsShift = 9U;
constexpr unsigned kDefaultPaperSize = 1U;
constexpr unsigned kDefaultScale = 100U;
constexpr unsigned kDefaultDpi = 600U;

// BrtBeginHeaderFooter flag bits.
constexpr std::uint32_t kDifferentOddEven = 0x01U;
constexpr std::uint32_t kDifferentFirst = 0x02U;
constexpr std::uint32_t kScaleWithDoc = 0x04U;
constexpr std::uint32_t kAlignWithMargins = 0x08U;

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

std::uint32_t ErrorsOrdinal(std::string_view errors) {
  if (errors == "blank") {
    return 1U;
  }
  if (errors == "dash") {
    return 2U;
  }
  if (errors == "NA") {
    return 3U;
  }
  return 0U;
}

/// True for a `<pageSetup usePrinterDefaults="0"/>` stating nothing else, which Excel saves without a BrtPageSetup.
bool OnlyDeclinesPrinterDefaults(const pugi::xml_node& node) {
  const pugi::xml_attribute first = node.first_attribute();
  return first && !first.next_attribute() && std::string_view(first.name()) == "usePrinterDefaults" &&
         !first.as_bool(true);
}

void EmitPageSetup(std::vector<std::uint8_t>& dst, const std::string& xml) {
  pugi::xml_document doc;
  const pugi::xml_node node = ParseRoot(doc, xml, "pageSetup");
  if (OnlyDeclinesPrinterDefaults(node)) {
    return;
  }
  const auto bit = [&node](const char* name, std::uint32_t mask) { return node.attribute(name).as_bool() ? mask : 0U; };
  const std::string_view comments = node.attribute("cellComments").value();
  std::uint32_t flags = bit("blackAndWhite", kBlackAndWhite) | bit("draft", kDraft) |
                        bit("useFirstPageNumber", kUseFirstPageNumber) |
                        (ErrorsOrdinal(node.attribute("errors").value()) << kErrorsShift);
  if (std::string_view(node.attribute("pageOrder").value()) == "overThenDown") {
    flags |= kOverThenDown;
  }
  if (std::string_view(node.attribute("orientation").value()) == "landscape") {
    flags |= kLandscape;
  }
  if (comments == "asDisplayed") {
    flags |= kCellComments;
  } else if (comments == "atEnd") {
    flags |= kCellComments | kCommentsAtEnd;
  }
  std::vector<std::uint8_t> payload;
  emit_u32(payload, node.attribute("paperSize").as_uint(kDefaultPaperSize));
  emit_u32(payload, node.attribute("scale").as_uint(kDefaultScale));
  emit_u32(payload, node.attribute("horizontalDpi").as_uint(kDefaultDpi));
  emit_u32(payload, node.attribute("verticalDpi").as_uint(kDefaultDpi));
  emit_u32(payload, node.attribute("copies").as_uint(1U));
  emit_u32(payload, node.attribute("firstPageNumber").as_uint(1U));
  emit_u32(payload, node.attribute("fitToWidth").as_uint(1U));
  emit_u32(payload, node.attribute("fitToHeight").as_uint(1U));
  // Eight flag bits and the two-bit errors ordinal, within u16.
  emit_u16(payload, static_cast<std::uint16_t>(flags));
  emit_xlnullablewidestring(payload, std::nullopt);
  emit_record(dst, kBrtPageSetup, payload);
}

/// BrtBeginHeaderFooter (flags, then the six strings, an empty one written as null) and its end record.
void EmitHeaderFooter(std::vector<std::uint8_t>& dst, const std::string& xml) {
  pugi::xml_document doc;
  const pugi::xml_node node = ParseRoot(doc, xml, "headerFooter");
  const auto bit = [&node](const char* name, bool fallback, std::uint32_t mask) {
    return node.attribute(name).as_bool(fallback) ? mask : 0U;
  };
  const std::uint32_t flags = bit("differentOddEven", false, kDifferentOddEven) |
                              bit("differentFirst", false, kDifferentFirst) | bit("scaleWithDoc", true, kScaleWithDoc) |
                              bit("alignWithMargins", true, kAlignWithMargins);
  std::vector<std::uint8_t> payload;
  // Four bits, within u16.
  emit_u16(payload, static_cast<std::uint16_t>(flags));
  for (const char* part : {"oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter"}) {
    const std::string_view text = node.child(part).text().get();
    emit_xlnullablewidestring(payload, text.empty() ? std::nullopt : std::optional<std::string_view>(text));
  }
  emit_record(dst, kBrtBeginHeaderFooter, payload);
  emit_record(dst, kBrtEndHeaderFooter, ByteSpan{});
}

/// The manual breaks of one axis between their begin and end records; Excel saves no automatic break.
void EmitBreaks(std::vector<std::uint8_t>& dst, const std::vector<ManualBreak>& breaks, std::uint16_t begin,
                std::uint16_t end) {
  std::uint32_t manual = 0U;
  for (const ManualBreak& brk : breaks) {
    manual += brk.manual ? 1U : 0U;
  }
  if (manual == 0U) {
    return;
  }
  std::vector<std::uint8_t> head;
  emit_u32(head, manual);
  emit_u32(head, manual);
  emit_record(dst, begin, head);
  for (const ManualBreak& brk : breaks) {
    if (!brk.manual) {
      continue;
    }
    std::vector<std::uint8_t> payload;
    emit_u32(payload, brk.id);
    emit_u32(payload, brk.min);
    emit_u32(payload, brk.max);
    emit_u32(payload, 1U);
    emit_u32(payload, 0U);
    emit_record(dst, kBrtBrk, payload);
  }
  emit_record(dst, end, ByteSpan{});
}

/// The print-record families in stream order.
enum class Family : std::uint8_t { kOptions, kMargins, kPageSetup, kHeaderFooter, kRowBreaks, kColBreaks, kNone };
constexpr std::size_t kFamilyCount = static_cast<std::size_t>(Family::kNone);

/// The family of a retained record; `in_row_breaks` tracks which axis a BrtBrk sits in.
Family FamilyOf(std::uint16_t type, bool& in_row_breaks) {
  switch (type) {
    case kBrtPrintOptions:
      return Family::kOptions;
    case kBrtMargins:
      return Family::kMargins;
    case kBrtPageSetup:
      return Family::kPageSetup;
    case kBrtBeginHeaderFooter:
    case kBrtEndHeaderFooter:
      return Family::kHeaderFooter;
    case kBrtBeginRwBrk:
      in_row_breaks = true;
      return Family::kRowBreaks;
    case kBrtEndRwBrk:
      return Family::kRowBreaks;
    case kBrtBeginColBrk:
      in_row_breaks = false;
      return Family::kColBreaks;
    case kBrtEndColBrk:
      return Family::kColBreaks;
    case kBrtBrk:
      return in_row_breaks ? Family::kRowBreaks : Family::kColBreaks;
    default:
      return Family::kNone;
  }
}

/// True when the model holds `family`, which the writer then emits from it.
bool ModelHolds(Family family, const SheetPrintSettings& print) {
  switch (family) {
    case Family::kOptions:
      return !print.print_options_xml.empty();
    case Family::kMargins:
      return !print.page_margins_xml.empty();
    case Family::kPageSetup:
      return !print.page_setup_xml.empty();
    case Family::kHeaderFooter:
      return !print.header_footer_xml.empty();
    case Family::kRowBreaks:
      return !print.manual_row_breaks.empty();
    case Family::kColBreaks:
      return !print.manual_col_breaks.empty();
    case Family::kNone:
      break;
  }
  return false;
}

void EmitFamily(std::vector<std::uint8_t>& dst, Family family, const SheetPrintSettings& print) {
  switch (family) {
    case Family::kOptions: {
      std::vector<std::uint8_t> options;
      emit_u16(options, PrintOptionsFlags(print.print_options_xml));
      emit_record(dst, kBrtPrintOptions, options);
      break;
    }
    case Family::kMargins:
      EmitMargins(dst, print.page_margins_xml);
      break;
    case Family::kPageSetup:
      EmitPageSetup(dst, print.page_setup_xml);
      break;
    case Family::kHeaderFooter:
      EmitHeaderFooter(dst, print.header_footer_xml);
      break;
    case Family::kRowBreaks:
      EmitBreaks(dst, print.manual_row_breaks, kBrtBeginRwBrk, kBrtEndRwBrk);
      break;
    case Family::kColBreaks:
      EmitBreaks(dst, print.manual_col_breaks, kBrtBeginColBrk, kBrtEndColBrk);
      break;
    case Family::kNone:
      break;
  }
}

}  // namespace

bool model_owns_print_records(const SheetPrintSettings& print) {
  for (std::size_t i = 0; i < kFamilyCount; ++i) {
    if (ModelHolds(static_cast<Family>(i), print)) {
      return true;
    }
  }
  return false;
}

bool has_deferred_print_settings(const SheetPrintSettings& print) {
  // The xlsb save writes no printer-settings part, so a sheet linking one loses it.
  return !print.sheet_pr_xml.empty() || !print.printer_settings_rid.empty();
}

void splice_print_records(std::vector<std::uint8_t>& after_hyperlinks, const SheetPrintSettings& print) {
  if (!model_owns_print_records(print)) {
    return;
  }
  const std::vector<FramedRecord> recs = split_records(after_hyperlinks);
  if (!after_hyperlinks.empty() && recs.empty()) {
    return;
  }
  std::array<std::vector<std::uint8_t>, kFamilyCount> retained;
  std::vector<std::uint8_t> rest;
  bool in_row_breaks = true;
  for (const FramedRecord& rec : recs) {
    const Family family = FamilyOf(rec.type, in_row_breaks);
    std::vector<std::uint8_t>& slot = family == Family::kNone ? rest : retained[static_cast<std::size_t>(family)];
    slot.insert(slot.end(), after_hyperlinks.begin() + static_cast<std::ptrdiff_t>(rec.begin),
                after_hyperlinks.begin() + static_cast<std::ptrdiff_t>(rec.end));
  }
  std::vector<std::uint8_t> out;
  for (std::size_t i = 0; i < kFamilyCount; ++i) {
    const auto family = static_cast<Family>(i);
    // Every sheet carries BrtPrintOptions, so one neither stated nor retained is written with the defaults.
    if (ModelHolds(family, print) || (family == Family::kOptions && retained[i].empty())) {
      EmitFamily(out, family, print);
    } else {
      out.insert(out.end(), retained[i].begin(), retained[i].end());
    }
  }
  out.insert(out.end(), rest.begin(), rest.end());
  after_hyperlinks = std::move(out);
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
