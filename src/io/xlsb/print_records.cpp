// Implementation of the worksheet print records. See
// `io/xlsb/print_records.h` for the contract.

#include "io/xlsb/print_records.h"

#include <array>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/ooxml/print_settings_parse.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/sheet_properties_records.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "utils/resource_budget.h"

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

/// The BrtPageSetup bits the reader states in `<pageSetup>`; a record carrying another bit is left retained.
constexpr std::uint32_t kDecodedPageSetupFlags = kOverThenDown | kLandscape | kBlackAndWhite | kDraft | kCellComments |
                                                 kUseFirstPageNumber | kCommentsAtEnd | (3U << kErrorsShift);
/// Eight u32 fields and the u16 flags, then a null `szRelID` (no printer settings part).
constexpr std::size_t kPageSetupFieldsSize = 34U;
constexpr std::size_t kPageSetupSize = kPageSetupFieldsSize + 4U;
constexpr std::uint32_t kNullStringLength = 0xFFFFFFFFU;

/// The `<pageSetup>` Excel writes in `.xlsx` for a BrtPageSetup, attributes in schema order: values equal to the
/// record's defaults are left out and `orientation` is always stated. Empty for a record with a printer-settings
/// relationship, an unmeasured flag bit, or comments at the end without the comments bit.
std::string DecodePageSetup(ByteSpan payload) {
  if (payload.size != kPageSetupSize || load_u32(payload.data + kPageSetupFieldsSize) != kNullStringLength) {
    return std::string();
  }
  const std::uint8_t* p = payload.data;
  const std::uint32_t flags = static_cast<std::uint32_t>(p[32]) | (static_cast<std::uint32_t>(p[33]) << 8U);
  if ((flags & ~kDecodedPageSetupFlags) != 0U || ((flags & kCommentsAtEnd) != 0U && (flags & kCellComments) == 0U)) {
    return std::string();
  }
  std::string xml = "<pageSetup";
  append_xml_attr_uint_if(xml, "paperSize", load_u32(p), kDefaultPaperSize);
  append_xml_attr_uint_if(xml, "scale", load_u32(p + 4U), kDefaultScale);
  append_xml_attr_uint_if(xml, "firstPageNumber", load_u32(p + 20U), 1U);
  append_xml_attr_uint_if(xml, "fitToWidth", load_u32(p + 24U), 1U);
  append_xml_attr_uint_if(xml, "fitToHeight", load_u32(p + 28U), 1U);
  if ((flags & kOverThenDown) != 0U) {
    xml.append(" pageOrder=\"overThenDown\"");
  }
  xml.append((flags & kLandscape) != 0U ? " orientation=\"landscape\"" : " orientation=\"portrait\"");
  if ((flags & kBlackAndWhite) != 0U) {
    xml.append(" blackAndWhite=\"1\"");
  }
  if ((flags & kDraft) != 0U) {
    xml.append(" draft=\"1\"");
  }
  if ((flags & kCellComments) != 0U) {
    xml.append((flags & kCommentsAtEnd) != 0U ? " cellComments=\"atEnd\"" : " cellComments=\"asDisplayed\"");
  }
  if ((flags & kUseFirstPageNumber) != 0U) {
    xml.append(" useFirstPageNumber=\"1\"");
  }
  static constexpr std::array<const char*, 4> kErrors = {nullptr, "blank", "dash", "NA"};
  if (const char* errors = kErrors[(flags >> kErrorsShift) & 3U]; errors != nullptr) {
    append_xml_attr(xml, "errors", errors);
  }
  append_xml_attr_uint_if(xml, "horizontalDpi", load_u32(p + 8U), kDefaultDpi);
  append_xml_attr_uint_if(xml, "verticalDpi", load_u32(p + 12U), kDefaultDpi);
  append_xml_attr_uint_if(xml, "copies", load_u32(p + 16U), 1U);
  xml.append("/>");
  return xml;
}

/// BrtMargins: six doubles, left, right, top, bottom, header, footer.
constexpr std::size_t kMarginsSize = 48U;

/// The `<pageMargins>` for a BrtMargins, every side stated as Excel does; empty for a non-finite or negative side.
std::string DecodeMargins(ByteSpan payload) {
  if (payload.size != kMarginsSize) {
    return std::string();
  }
  std::string xml = "<pageMargins";
  for (const char* side : {"left", "right", "top", "bottom", "header", "footer"}) {
    double value = 0.0;
    std::memcpy(&value, payload.data, sizeof(value));
    payload.data += sizeof(value);
    if (!std::isfinite(value) || value < 0.0) {
      return std::string();
    }
    append_xml_attr_number(xml, side, value);
  }
  xml.append("/>");
  return xml;
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

/// The retained records of `family`, framed, as the splice keeps them.
std::vector<std::uint8_t> RetainedFamilyBytes(const std::vector<std::uint8_t>& stream,
                                              const std::vector<FramedRecord>& recs, Family family) {
  std::vector<std::uint8_t> out;
  bool in_row_breaks = true;
  for (const FramedRecord& rec : recs) {
    if (FamilyOf(rec.type, in_row_breaks) == family) {
      out.insert(out.end(), stream.begin() + static_cast<std::ptrdiff_t>(rec.begin),
                 stream.begin() + static_cast<std::ptrdiff_t>(rec.end));
    }
  }
  return out;
}

/// The `<headerFooter>` Excel writes in `.xlsx` for a BrtBeginHeaderFooter payload: an attribute only when it differs
/// from its default, the strings in schema order. Empty when the payload does not parse.
std::string DecodeHeaderFooter(ByteSpan payload) {
  if (payload.size < 2U) {
    return std::string();
  }
  const std::uint32_t flags = load_u16(payload.data);
  payload.data += 2U;
  payload.size -= 2U;
  pugi::xml_document doc;
  pugi::xml_node root = doc.append_child("headerFooter");
  if ((flags & kDifferentOddEven) != 0U) {
    root.append_attribute("differentOddEven") = "1";
  }
  if ((flags & kDifferentFirst) != 0U) {
    root.append_attribute("differentFirst") = "1";
  }
  if ((flags & kScaleWithDoc) == 0U) {
    root.append_attribute("scaleWithDoc") = "0";
  }
  if ((flags & kAlignWithMargins) == 0U) {
    root.append_attribute("alignWithMargins") = "0";
  }
  for (const char* part : {"oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter"}) {
    auto text = read_xlnullablewidestring(payload);
    if (!text) {
      return std::string();
    }
    if (!text.value().empty()) {
      root.append_child(part).text().set(text.value().c_str());
    }
  }
  return raw_xml(root);
}

/// Sets `print.header_footer_xml` when the record set re-encodes to exactly the retained bytes.
void DecodeHeaderFooterFamily(const std::vector<std::uint8_t>& stream, const std::vector<FramedRecord>& recs,
                              SheetPrintSettings& print) {
  for (const FramedRecord& rec : recs) {
    if (rec.type != kBrtBeginHeaderFooter) {
      continue;
    }
    std::string xml = DecodeHeaderFooter(rec.payload);
    std::vector<std::uint8_t> again;
    EmitHeaderFooter(again, xml);
    if (!xml.empty() && again == RetainedFamilyBytes(stream, recs, Family::kHeaderFooter)) {
      print.header_footer_xml = std::move(xml);
    }
    return;
  }
}

/// BrtBeginRwBrk / BrtBeginColBrk: the break count and the manual break count.
constexpr std::size_t kBreakHeadSize = 8U;
/// BrtBrk: id, min, max, then the manual and page-break flags.
constexpr std::size_t kBreakSize = 20U;

/// Sets the break vectors the retained records state exactly: every break manual, ids strictly increasing, and the
/// record set equal to what the writer emits for them.
void DecodeBreakFamilies(const std::vector<std::uint8_t>& stream, const std::vector<FramedRecord>& recs,
                         SheetPrintSettings& print) {
  std::vector<ManualBreak> rows;
  std::vector<ManualBreak> cols;
  bool rows_ok = true;
  bool cols_ok = true;
  std::vector<ManualBreak>* axis = nullptr;
  bool* axis_ok = nullptr;
  for (const FramedRecord& rec : recs) {
    if (rec.type == kBrtBeginRwBrk || rec.type == kBrtBeginColBrk) {
      const bool is_row = rec.type == kBrtBeginRwBrk;
      axis = is_row ? &rows : &cols;
      axis_ok = is_row ? &rows_ok : &cols_ok;
      *axis_ok = *axis_ok && rec.payload.size == kBreakHeadSize;
    } else if (rec.type == kBrtEndRwBrk || rec.type == kBrtEndColBrk) {
      axis = nullptr;
      axis_ok = nullptr;
    } else if (rec.type == kBrtBrk && axis != nullptr && *axis_ok) {
      const std::uint8_t* p = rec.payload.data;
      const bool exact = rec.payload.size == kBreakSize && load_u32(p + 12U) == 1U && load_u32(p + 16U) == 0U &&
                         axis->size() < kMaxManualBreaksPerAxis && (axis->empty() || axis->back().id < load_u32(p));
      if (!exact) {
        *axis_ok = false;
        continue;
      }
      axis->push_back(ManualBreak{load_u32(p), load_u32(p + 4U), load_u32(p + 8U), true});
    }
  }
  const auto adopt = [&](bool ok, std::vector<ManualBreak>& found, std::vector<ManualBreak>& model, Family family,
                         std::uint16_t begin, std::uint16_t end) {
    if (!ok || found.empty() || !model.empty()) {
      return;
    }
    std::vector<std::uint8_t> again;
    EmitBreaks(again, found, begin, end);
    if (again == RetainedFamilyBytes(stream, recs, family)) {
      model = std::move(found);
    }
  };
  adopt(rows_ok, rows, print.manual_row_breaks, Family::kRowBreaks, kBrtBeginRwBrk, kBrtEndRwBrk);
  adopt(cols_ok, cols, print.manual_col_breaks, Family::kColBreaks, kBrtBeginColBrk, kBrtEndColBrk);
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

void decode_print_records(const std::vector<std::uint8_t>& after_hyperlinks, SheetPrintSettings& print) {
  const std::vector<FramedRecord> recs = split_records(after_hyperlinks);
  for (const FramedRecord& rec : recs) {
    if (rec.type == kBrtPageSetup && print.page_setup_xml.empty()) {
      if (std::string xml = DecodePageSetup(rec.payload); !xml.empty()) {
        print.page_setup_xml = std::move(xml);
        ooxml::refresh_structured_views("pageSetup", print.page_setup_xml, print);
      }
    } else if (rec.type == kBrtMargins && print.page_margins_xml.empty()) {
      if (std::string xml = DecodeMargins(rec.payload); !xml.empty()) {
        print.page_margins_xml = std::move(xml);
        ooxml::refresh_structured_views("pageMargins", print.page_margins_xml, print);
      }
    } else if (rec.type == kBrtPrintOptions && print.print_options_xml.empty()) {
      // Adopted only when the writer re-emits exactly the retained payload.
      std::optional<std::string> xml = decode_print_options(rec.payload);
      if (xml && !xml->empty() && rec.payload.size == 2U && load_u16(rec.payload.data) == PrintOptionsFlags(*xml)) {
        print.print_options_xml = std::move(*xml);
      }
    }
  }
  if (print.header_footer_xml.empty()) {
    DecodeHeaderFooterFamily(after_hyperlinks, recs, print);
  }
  DecodeBreakFamilies(after_hyperlinks, recs, print);
}

std::optional<std::string> decode_print_options(ByteSpan payload) {
  if (payload.size != 2U) {
    return std::nullopt;
  }
  const std::uint32_t flags =
      static_cast<std::uint32_t>(payload.data[0]) | (static_cast<std::uint32_t>(payload.data[1]) << 8U);
  if ((flags & ~(kGridLinesSet | kHorizontalCentered | kVerticalCentered | kHeadings | kGridLines)) != 0U ||
      (flags & kGridLinesSet) == 0U) {
    return std::nullopt;
  }
  std::string xml;
  for (const auto& [name, mask] : {std::pair<const char*, std::uint32_t>{"horizontalCentered", kHorizontalCentered},
                                   {"verticalCentered", kVerticalCentered},
                                   {"headings", kHeadings},
                                   {"gridLines", kGridLines}}) {
    if ((flags & mask) != 0U) {
      xml.append(" ").append(name).append("=\"1\"");
    }
  }
  return xml.empty() ? std::string() : "<printOptions" + xml + "/>";
}

std::uint32_t unwritten_print_families(const std::vector<std::uint8_t>& after_hyperlinks,
                                       const SheetPrintSettings& print) {
  std::array<bool, kFamilyCount> retained{};
  bool in_row_breaks = true;
  for (const FramedRecord& rec : split_records(after_hyperlinks)) {
    const Family family = FamilyOf(rec.type, in_row_breaks);
    if (family == Family::kNone) {
      continue;
    }
    // Print options all at their defaults state nothing the model could miss.
    if (family == Family::kOptions) {
      const std::optional<std::string> xml = decode_print_options(rec.payload);
      if (xml && xml->empty()) {
        continue;
      }
    }
    retained[static_cast<std::size_t>(family)] = true;
  }
  std::uint32_t count = 0;
  for (std::size_t i = 0; i < kFamilyCount; ++i) {
    count += retained[i] && !ModelHolds(static_cast<Family>(i), print) ? 1U : 0U;
  }
  return count;
}

void emit_view_print_records(std::vector<std::uint8_t>& dst, const SheetPrintSettings& print) {
  for (const Family family : {Family::kRowBreaks, Family::kColBreaks, Family::kMargins, Family::kOptions,
                              Family::kPageSetup, Family::kHeaderFooter}) {
    if (ModelHolds(family, print) || family == Family::kOptions) {
      EmitFamily(dst, family, print);
    }
  }
}

void emit_chartsheet_print_records(std::vector<std::uint8_t>& dst, const SheetPrintSettings& print) {
  for (const Family family : {Family::kMargins, Family::kHeaderFooter}) {
    if (ModelHolds(family, print)) {
      EmitFamily(dst, family, print);
    }
  }
}

bool has_deferred_print_settings(const SheetPrintSettings& print) {
  // The xlsb save writes no printer-settings part, so a sheet linking one loses it.
  return !sheet_pr_writes_all(print.sheet_pr_xml) || !print.printer_settings_rid.empty();
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
