#include "io/xlsb/hyperlink_records.h"

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::uint32_t kMaxHyperlinkRelIdUnits = 32767U;
constexpr std::uint32_t kMaxHyperlinkLocationUnits = 2083U;
constexpr std::uint32_t kMaxHyperlinkTooltipUnits = 255U;
constexpr std::uint32_t kMaxHyperlinkDisplayUnits = 32767U;

Expected<std::string, Error> ReadHyperlinkWideString(ByteSpan& cursor, const char* field, std::uint32_t max_units) {
  if (cursor.size < 4U) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated,
                      std::string("xlsb BrtHLink ") + field + " length truncated", "context=xlsb_reader");
  }
  const std::uint32_t cch =
      static_cast<std::uint32_t>(cursor.data[0]) | (static_cast<std::uint32_t>(cursor.data[1]) << 8U) |
      (static_cast<std::uint32_t>(cursor.data[2]) << 16U) | (static_cast<std::uint32_t>(cursor.data[3]) << 24U);
  if (cch > max_units) {
    return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt,
                      std::string("xlsb BrtHLink ") + field + " exceeds string limit",
                      "context=xlsb_reader cch=" + std::to_string(cch));
  }
  return read_xlwidestring(cursor);
}

Expected<void, Error> EmitBoundedHyperlinkWideString(std::vector<std::uint8_t>& dst, std::string_view text,
                                                     std::uint32_t max_units, const char* field) {
  const std::size_t start = dst.size();
  emit_xlwidestring(dst, text);
  const std::uint32_t cch =
      static_cast<std::uint32_t>(dst[start]) | (static_cast<std::uint32_t>(dst[start + 1U]) << 8U) |
      (static_cast<std::uint32_t>(dst[start + 2U]) << 16U) | (static_cast<std::uint32_t>(dst[start + 3U]) << 24U);
  if (cch > max_units) {
    dst.resize(start);
    return make_error(FormulonErrorCode::kInvalidArgument,
                      std::string("xlsb BrtHLink ") + field + " exceeds string limit",
                      "context=xlsb_sheet_writer cch=" + std::to_string(cch));
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<void, Error> decode_hyperlink(const XlsbRecord& rec, Sheet& sheet) {
  // BrtHLink ([MS-XLSB] §2.4.494): RfX followed by a non-null
  // relationship id and three XLWideStrings (location, tooltip,
  // display). A present-but-empty relationship id is the internal
  // hyperlink form; external links must resolve that id through the
  // sheet relationship part after the record stream is decoded.
  ByteSpan p = rec.payload;
  auto rfx_or = read_rfx(p);
  if (!rfx_or) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtHLink range truncated",
                      "context=xlsb_reader");
  }
  const MergeRange& rfx = rfx_or.value();
  if (!Sheet::rect_in_grid(rfx.first_row, rfx.first_col, rfx.last_row, rfx.last_col)) {
    return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtHLink range out of bounds",
                      "context=xlsb_reader");
  }
  // `read_xlnullablewidestring` intentionally maps both a null value
  // and a present empty value to `""`; inspect the sentinel first so a
  // null RelID cannot be mistaken for the required internal form.
  if (p.size < 4U) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtHLink relationship id truncated",
                      "context=xlsb_reader");
  }
  const std::uint32_t rel_len = static_cast<std::uint32_t>(p.data[0]) | (static_cast<std::uint32_t>(p.data[1]) << 8U) |
                                (static_cast<std::uint32_t>(p.data[2]) << 16U) |
                                (static_cast<std::uint32_t>(p.data[3]) << 24U);
  if (rel_len == 0xFFFFFFFFU) {
    return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtHLink relationship id is null",
                      "context=xlsb_reader");
  }
  if (rel_len > kMaxHyperlinkRelIdUnits) {
    return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtHLink relationship id exceeds string limit",
                      "context=xlsb_reader cch=" + std::to_string(rel_len));
  }
  ASSIGN_OR_RETURN(auto rid, read_xlnullablewidestring(p));
  auto location_or = ReadHyperlinkWideString(p, "location", kMaxHyperlinkLocationUnits);
  if (!location_or) {
    return std::move(location_or.error());
  }
  auto tooltip_or = ReadHyperlinkWideString(p, "tooltip", kMaxHyperlinkTooltipUnits);
  if (!tooltip_or) {
    return std::move(tooltip_or.error());
  }
  auto display_or = ReadHyperlinkWideString(p, "display", kMaxHyperlinkDisplayUnits);
  if (!display_or) {
    return std::move(display_or.error());
  }
  if (p.size != 0U) {
    return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtHLink has trailing bytes",
                      "context=xlsb_reader trailing=" + std::to_string(p.size));
  }
  Hyperlink hyperlink;
  hyperlink.row = rfx.first_row;
  hyperlink.col = rfx.first_col;
  hyperlink.last_row = rfx.last_row;
  hyperlink.last_col = rfx.last_col;
  hyperlink.rid = std::move(rid);
  hyperlink.location = std::move(location_or.value());
  hyperlink.tooltip = std::move(tooltip_or.value());
  hyperlink.display = std::move(display_or.value());
  sheet.mutable_hyperlinks().push_back(std::move(hyperlink));
  return Expected<void, Error>::Ok();
}

Expected<void, Error> emit_hyperlink(std::vector<std::uint8_t>& dst, const Hyperlink& hyperlink, std::string_view rid) {
  if (!Sheet::rect_in_grid(hyperlink.row, hyperlink.col, hyperlink.last_row, hyperlink.last_col)) {
    return make_error(FormulonErrorCode::kInvalidArgument, "xlsb hyperlink rectangle out of grid",
                      "context=xlsb_sheet_writer row=" + std::to_string(hyperlink.row) +
                          " col=" + std::to_string(hyperlink.col) + " last_row=" + std::to_string(hyperlink.last_row) +
                          " last_col=" + std::to_string(hyperlink.last_col));
  }
  std::vector<std::uint8_t> payload;
  emit_rfx(payload, MergeRange{hyperlink.row, hyperlink.col, hyperlink.last_row, hyperlink.last_col});
  // RelID is always present in BrtHLink. Empty-but-present is the internal
  // hyperlink form; null would be malformed on read.
  emit_xlnullablewidestring(payload, std::optional<std::string_view>(rid));
  if (!rid.empty()) {
    const std::size_t rid_start = 16U;
    const std::uint32_t rid_units = static_cast<std::uint32_t>(payload[rid_start]) |
                                    (static_cast<std::uint32_t>(payload[rid_start + 1U]) << 8U) |
                                    (static_cast<std::uint32_t>(payload[rid_start + 2U]) << 16U) |
                                    (static_cast<std::uint32_t>(payload[rid_start + 3U]) << 24U);
    if (rid_units > kMaxHyperlinkRelIdUnits) {
      return make_error(FormulonErrorCode::kInvalidArgument, "xlsb BrtHLink relationship id exceeds string limit",
                        "context=xlsb_sheet_writer cch=" + std::to_string(rid_units));
    }
  }
  if (auto r = EmitBoundedHyperlinkWideString(payload, hyperlink.location, kMaxHyperlinkLocationUnits, "location");
      !r) {
    return std::move(r.error());
  }
  if (auto r = EmitBoundedHyperlinkWideString(payload, hyperlink.tooltip, kMaxHyperlinkTooltipUnits, "tooltip"); !r) {
    return std::move(r.error());
  }
  if (auto r = EmitBoundedHyperlinkWideString(payload, hyperlink.display, kMaxHyperlinkDisplayUnits, "display"); !r) {
    return std::move(r.error());
  }
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtHLink), payload);
  return Expected<void, Error>::Ok();
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
