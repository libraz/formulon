//
// Implementation of the per-sheet XLSB stream emitter. See
// `io/xlsb/sheet_writer.h` for the contract.

#include "io/xlsb/sheet_writer.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <cstdlib>
#include <map>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "cell.h"
#include "io/xlsb/cell_writer.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/sst_writer.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// Returns `true` when a `Cell` is "empty" in the sense that nothing
/// useful would be carried by emitting a `BrtCellBlank` for it. Used
/// to skip implicitly-default-constructed columns produced by sheet
/// row growth (see `Sheet::set_cell_value` docs).
bool IsEmptySlot(const Cell& cell) {
  return cell.formula_text.empty() && cell.cached_value.is_blank();
}

/// Emits a `BrtRowHdr` for `row`. [MS-XLSB] §2.4.770 layout:
///   * rw         : u32 (0-based row index)
///   * iStyleRef  : u32 (row style xf index, or 0)
///   * miyRw      : u16 (custom row height in twips; 0 = use default)
///   * flags1     : u8
///   * flags2     : u8  (outline / hidden / customHeight / fGhostDirty bits)
///   * fPhShow    : u8  (phonetic-guide default)
///   * ccolspan   : u32 (number of following BrtColSpan records)
///   * rgBrtColspan: repeated (colMic, colLast) u32 pairs
///
/// A row with cell records must describe every 1,024-column segment which
/// contains one.  Excel repairs a stream that claims `ccolspan == 0` while
/// following it with cells, even though our own permissive reader can decode
/// it.  An empty layout-only row legitimately has no spans.
void EmitRowHeader(std::vector<std::uint8_t>& dst, std::uint32_t row, const RowLayout* layout,
                   const std::map<std::uint32_t, std::pair<std::uint32_t, std::uint32_t>>& spans) {
  std::vector<std::uint8_t> p;
  emit_u32(p, row);
  emit_u32(p, layout != nullptr && layout->has_style ? layout->style_xf : 0U);  // iStyleRef
  const bool has_height = layout != nullptr && (layout->has_height || layout->height > 0.0);
  const double twips = has_height ? std::round(layout->height * 20.0) : 0.0;
  emit_u16(p, static_cast<std::uint16_t>(std::clamp(twips, 0.0, 65535.0)));  // miyRw
  emit_u8(p, 0);                                                             // flags1
  std::uint8_t flags2 = layout == nullptr ? 0U : static_cast<std::uint8_t>(layout->outline_level & 0x07U);
  if (layout != nullptr && layout->hidden) {
    flags2 |= 0x10U;  // fDyZero
  }
  if (has_height) {
    flags2 |= 0x20U;  // fUnsynced
  }
  if (layout != nullptr && layout->has_style) {
    flags2 |= 0x40U;  // fGhostDirty: iStyleRef is an effective row style
  }
  emit_u8(p, flags2);
  // fPhShow is a distinct byte, not part of flags2.  It is required even
  // when the phonetic guide is disabled; otherwise ccolspan is misaligned.
  emit_u8(p, 0U);
  emit_u32(p, static_cast<std::uint32_t>(spans.size()));  // ccolspan
  for (const auto& [segment, span] : spans) {
    (void)segment;
    emit_u32(p, span.first);   // colMic
    emit_u32(p, span.second);  // colLast
  }
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtRowHdr), p);
}

void EmitColumnInfos(std::vector<std::uint8_t>& dst, const SheetLayout& layout) {
  if (layout.columns.empty()) {
    return;
  }
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginColInfos), ByteSpan{});
  for (const ColumnLayout& column : layout.columns) {
    if (column.first > column.last || column.last >= Sheet::kMaxCols) {
      continue;
    }
    std::vector<std::uint8_t> p;
    emit_u32(p, column.first);
    emit_u32(p, column.last);
    const bool has_width = HasExplicitColumnWidth(column);
    const double width256 = has_width ? std::floor(std::max(0.0, column.width) * 256.0) : 0.0;
    emit_u32(p, static_cast<std::uint32_t>(std::clamp(width256, 0.0, 65535.0)));
    // BrtColInfo has no style-presence bit; ixfe is mandatory. OOXML spans
    // without a style therefore canonicalize to style 0 on an XLSB reload.
    emit_u32(p, column.has_style ? column.style_xf : 0U);  // ixfe
    std::uint16_t flags = has_width ? 0x0002U : 0U;        // fUserSet
    if (column.hidden) {
      flags |= 0x0001U;
    }
    flags |= static_cast<std::uint16_t>((column.outline_level & 0x07U) << 8U);
    emit_u16(p, flags);
    emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtColInfo), p);
  }
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtEndColInfos), ByteSpan{});
}

/// The two `<sheetPr>` members `BrtWsProp` carries: the VBA code name and
/// the tab colour, in the wire form the record wants them.
///
/// `color_type` is the `XColorType` selector; `0` is automatic, which
/// together with the automatic palette index is the "no tab colour" state
/// Excel writes for an untinted tab.
struct WorksheetProperties {
  std::string code_name;
  std::uint8_t color_type = 0U;
  std::uint8_t color_index = 0x40U;
  std::int16_t color_tint = 0;
  std::uint32_t color_argb = 0U;
};

/// Reads the sheet's retained `<sheetPr>` fragment back into the fields
/// `BrtWsProp` can express. The fragment is the same string the OOXML
/// writer emits and the XLSB reader synthesises, so this is the inverse of
/// `reader.cpp`'s `DecodeWorksheetProperties` and the two containers agree
/// on a sheet's code name and tab colour whichever one it was loaded from.
///
/// Members `<sheetPr>` can carry that `BrtWsProp` has no field for --
/// `<outlinePr>`, `<pageSetUpPr>`, `filterMode` -- stay in the fragment
/// and reach an `.xlsx` save unchanged; they are simply not part of what
/// this record emits.
WorksheetProperties ParseSheetProperties(std::string_view raw) {
  WorksheetProperties out;
  if (raw.empty()) {
    return out;
  }
  pugi::xml_document doc;
  if (!doc.load_buffer(raw.data(), raw.size())) {
    return out;
  }
  const pugi::xml_node sheet_pr = doc.document_element();
  if (!sheet_pr || std::string_view(sheet_pr.name()) != "sheetPr") {
    return out;
  }
  out.code_name = sheet_pr.attribute("codeName").value();
  const pugi::xml_node tab = sheet_pr.child("tabColor");
  if (!tab) {
    return out;
  }
  if (const pugi::xml_attribute rgb = tab.attribute("rgb")) {
    out.color_type = 2U;
    out.color_index = 0U;
    out.color_argb = static_cast<std::uint32_t>(std::strtoul(rgb.value(), nullptr, 16));
    return out;
  }
  if (const pugi::xml_attribute theme = tab.attribute("theme")) {
    out.color_type = 3U;
    out.color_index = static_cast<std::uint8_t>(std::min<unsigned>(theme.as_uint(0U), 0xFFU));
    const double tint = std::clamp(std::round(tab.attribute("tint").as_double(0.0) * 32767.0), -32767.0, 32767.0);
    out.color_tint = static_cast<std::int16_t>(tint);
    return out;
  }
  if (const pugi::xml_attribute indexed = tab.attribute("indexed")) {
    out.color_type = 1U;
    out.color_index = static_cast<std::uint8_t>(std::min<unsigned>(indexed.as_uint(0U), 0xFFU));
    return out;
  }
  return out;
}

/// Emits the worksheet-properties record which starts the mandatory worksheet
/// prefix. `BrtWsDim` must follow this record before the view collection.
void EmitWorksheetProperties(std::vector<std::uint8_t>& dst, const Sheet& sheet) {
  const WorksheetProperties properties_model = ParseSheetProperties(sheet.print_settings().sheet_pr_xml);
  std::vector<std::uint8_t> properties;
  emit_u16(properties, 0x04C9U);  // page breaks, publish, outline defaults
  emit_u8(properties, 0x02U);     // evaluate conditional formatting
  // BrtColor. An automatic type with the automatic palette index is the
  // untinted tab; any other selector sets fValidRGB alongside it, matching
  // how the styles writer spells the same structure.
  emit_u8(properties,
          properties_model.color_type == 0U
              ? std::uint8_t{0}
              : static_cast<std::uint8_t>((static_cast<unsigned>(properties_model.color_type) << 1U) | 0x01U));
  emit_u8(properties, properties_model.color_index);
  emit_u16(properties, static_cast<std::uint16_t>(properties_model.color_tint));
  emit_u8(properties, static_cast<std::uint8_t>((properties_model.color_argb >> 16U) & 0xFFU));
  emit_u8(properties, static_cast<std::uint8_t>((properties_model.color_argb >> 8U) & 0xFFU));
  emit_u8(properties, static_cast<std::uint8_t>(properties_model.color_argb & 0xFFU));
  emit_u8(properties, static_cast<std::uint8_t>((properties_model.color_argb >> 24U) & 0xFFU));
  emit_u32(properties, 0xFFFFFFFFU);  // rwSync: unused
  emit_u32(properties, 0xFFFFFFFFU);  // colSync: unused
  emit_xlwidestring(properties, properties_model.code_name);
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtWsProp), properties);
}

/// Emits the mandatory worksheet-view and default-formatting suffix. Omitting
/// these records leaves a technically decodable stream that desktop Excel
/// repairs before opening. The defaults match an Excel-created normal
/// worksheet: visible grid/headings and 100% zoom. The default metrics come
/// from the sheet's modelled `<sheetFormatPr>` values.
void EmitWorksheetViewsAndFormatting(std::vector<std::uint8_t>& dst, const SheetView& sheet_view,
                                     const SheetFormatDefaults& defaults) {
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginWsViews), ByteSpan{});
  std::vector<std::uint8_t> view;
  std::uint16_t flags = 0x0380U;  // ruler, outline symbols, default gridline color
  if (sheet_view.show_grid_lines) {
    flags |= 0x0004U;
  }
  if (sheet_view.show_row_col_headers) {
    flags |= 0x0008U;
  }
  if (sheet_view.show_zeros) {
    flags |= 0x0010U;
  }
  if (sheet_view.right_to_left) {
    flags |= 0x0020U;
  }
  if (sheet_view.tab_selected) {
    flags |= 0x0040U;
  }
  emit_u16(view, flags);
  emit_u32(view, 0U);    // XLVNORMAL
  emit_u32(view, 0U);    // rwTop
  emit_u32(view, 0U);    // colLeft
  emit_u8(view, 0x40U);  // default gridline color
  emit_u8(view, 0U);
  emit_u16(view, 0U);
  const std::uint32_t zoom = std::clamp(sheet_view.zoom_scale, 10U, 400U);
  emit_u16(view, static_cast<std::uint16_t>(zoom));
  emit_u16(view, 0U);
  emit_u16(view, 0U);
  emit_u16(view, 0U);
  emit_u32(view, 0U);  // workbook view index
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginWsView), view);
  if (sheet_view.freeze_rows != 0U || sheet_view.freeze_cols != 0U) {
    // BrtPane uses the frozen row and column counts in the two Xnum fields,
    // and identifies the bottom-right pane as the active pane. fFrozenNoSplit
    // represents Excel's ordinary "Freeze Panes" state.
    std::vector<std::uint8_t> pane;
    emit_double(pane, static_cast<double>(sheet_view.freeze_rows));
    emit_double(pane, static_cast<double>(sheet_view.freeze_cols));
    emit_u32(pane, sheet_view.freeze_rows);
    emit_u32(pane, sheet_view.freeze_cols);
    emit_u32(pane, 0U);  // PNNBOTRIGHT
    emit_u8(pane, 0x02U);
    emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtPane), pane);
  }
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtEndWsView), ByteSpan{});
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtEndWsViews), ByteSpan{});

  // BrtWsFmtInfo stores the default column width as 1/256 character units,
  // while the OOXML model stores it in character units. The sentinel is the
  // only absent marker; zero is a valid explicit width.
  constexpr std::uint32_t kAbsentDefaultColumnWidth = 0xFFFFFFFFU;
  constexpr std::uint16_t kCanonicalDefaultColumnWidth = 8U;
  constexpr std::uint16_t kCanonicalDefaultRowHeightTwips = 300U;
  constexpr double kMaxDefaultColumnWidth = 65535.0 / 256.0;
  constexpr double kMaxDefaultRowHeight = 65535.0 / 20.0;

  const bool valid_base_col_width = std::isfinite(defaults.base_col_width) && defaults.base_col_width >= 0.0 &&
                                    defaults.base_col_width <= 255.0 &&
                                    std::floor(defaults.base_col_width) == defaults.base_col_width;
  const bool valid_default_col_width =
      !defaults.has_default_col_width ||
      (std::isfinite(defaults.default_col_width) && defaults.default_col_width >= 0.0 &&
       defaults.default_col_width <= kMaxDefaultColumnWidth);
  const bool valid_default_row_height =
      !defaults.has_default_row_height ||
      (std::isfinite(defaults.default_row_height) && defaults.default_row_height >= 0.0 &&
       defaults.default_row_height <= kMaxDefaultRowHeight &&
       std::isfinite(std::round(defaults.default_row_height * 20.0)));

  const std::uint16_t base_col_width =
      valid_base_col_width ? static_cast<std::uint16_t>(defaults.base_col_width) : kCanonicalDefaultColumnWidth;
  const bool emit_default_col_width = defaults.has_default_col_width && valid_default_col_width;
  const std::uint32_t dx_g_col = emit_default_col_width
                                     ? static_cast<std::uint32_t>(std::floor(defaults.default_col_width * 256.0))
                                     : kAbsentDefaultColumnWidth;

  std::uint16_t miy_default_row_height = kCanonicalDefaultRowHeightTwips;
  std::uint32_t format_flags = 0U;
  if (defaults.has_default_row_height && valid_default_row_height) {
    miy_default_row_height = static_cast<std::uint16_t>(std::round(defaults.default_row_height * 20.0));
    if (miy_default_row_height == 0U) {
      format_flags |= 0x00000002U;  // fDyZero: explicit zero (including quantized-zero).
    } else {
      format_flags |= 0x00000001U;  // fUnsynced: explicit positive default row height.
    }
  }

  std::vector<std::uint8_t> formatting;
  emit_u32(formatting, dx_g_col);
  emit_u16(formatting, base_col_width);
  emit_u16(formatting, miy_default_row_height);
  emit_u32(formatting, format_flags);  // thick/outline metadata is not authored by this model
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtWsFmtInfo), formatting);
}

/// Emits the bounding RfX of actual emitted cells.  Although BrtWsDim is
/// optional in the grammar, Excel itself writes it for every non-empty sheet
/// and uses it to establish the worksheet's used range before reading the
/// cell table.
void EmitWorksheetDimensions(std::vector<std::uint8_t>& dst, const Sheet& sheet) {
  std::uint32_t first_row = Sheet::kMaxRows;
  std::uint32_t last_row = 0U;
  std::uint32_t first_col = Sheet::kMaxCols;
  std::uint32_t last_col = 0U;
  bool any = false;
  auto include = [&](std::uint32_t row, std::uint32_t col) {
    first_row = std::min(first_row, row);
    last_row = std::max(last_row, row);
    first_col = std::min(first_col, col);
    last_col = std::max(last_col, col);
    any = true;
  };

  for (const auto& [row, cells] : sheet.rows()) {
    for (std::uint32_t col = 0; col < cells.size(); ++col) {
      if (!IsEmptySlot(cells[col])) {
        include(row, col);
      }
    }
  }
  if (!any) {
    return;
  }
  std::vector<std::uint8_t> payload;
  emit_u32(payload, first_row);
  emit_u32(payload, last_row);
  emit_u32(payload, first_col);
  emit_u32(payload, last_col);
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtWsDim), payload);
}

void EmitMerges(std::vector<std::uint8_t>& dst, const Sheet& sheet) {
  if (sheet.merges().empty()) {
    return;
  }
  std::size_t valid_count = 0;
  for (const MergeRange& merge : sheet.merges()) {
    if (merge.first_row <= merge.last_row && merge.first_col <= merge.last_col && merge.last_row < Sheet::kMaxRows &&
        merge.last_col < Sheet::kMaxCols) {
      ++valid_count;
    }
  }
  if (valid_count == 0U) {
    return;
  }
  std::vector<std::uint8_t> count;
  emit_u32(count, static_cast<std::uint32_t>(valid_count));
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginMergeCells), count);
  for (const MergeRange& merge : sheet.merges()) {
    if (merge.first_row > merge.last_row || merge.first_col > merge.last_col || merge.last_row >= Sheet::kMaxRows ||
        merge.last_col >= Sheet::kMaxCols) {
      continue;
    }
    std::vector<std::uint8_t> payload;
    emit_u32(payload, merge.first_row);
    emit_u32(payload, merge.last_row);
    emit_u32(payload, merge.first_col);
    emit_u32(payload, merge.last_col);
    emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtMergeCell), payload);
  }
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtEndMergeCells), ByteSpan{});
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

Expected<void, Error> EmitHyperlinks(std::vector<std::uint8_t>& dst, const Sheet& sheet,
                                     const std::vector<std::string>& relationship_ids) {
  constexpr std::uint32_t kMaxRelIdUnits = 32767U;
  constexpr std::uint32_t kMaxLocationUnits = 2083U;
  constexpr std::uint32_t kMaxTooltipUnits = 255U;
  constexpr std::uint32_t kMaxDisplayUnits = 32767U;
  for (std::size_t i = 0; i < sheet.hyperlinks().size(); ++i) {
    const Hyperlink& hyperlink = sheet.hyperlinks()[i];
    if (!Sheet::rect_in_grid(hyperlink.row, hyperlink.col, hyperlink.last_row, hyperlink.last_col)) {
      return make_error(
          FormulonErrorCode::kInvalidArgument, "xlsb hyperlink rectangle out of grid",
          "context=xlsb_sheet_writer row=" + std::to_string(hyperlink.row) + " col=" + std::to_string(hyperlink.col) +
              " last_row=" + std::to_string(hyperlink.last_row) + " last_col=" + std::to_string(hyperlink.last_col));
    }
    std::vector<std::uint8_t> payload;
    emit_u32(payload, hyperlink.row);
    emit_u32(payload, hyperlink.last_row);
    emit_u32(payload, hyperlink.col);
    emit_u32(payload, hyperlink.last_col);
    const std::string_view rid =
        i < relationship_ids.size() ? std::string_view(relationship_ids[i]) : std::string_view{};
    // RelID is always present in BrtHLink. Empty-but-present is the internal
    // hyperlink form; null would be malformed on read.
    emit_xlnullablewidestring(payload, std::optional<std::string_view>(rid));
    if (!rid.empty()) {
      const std::size_t rid_start = 16U;
      const std::uint32_t rid_units = static_cast<std::uint32_t>(payload[rid_start]) |
                                      (static_cast<std::uint32_t>(payload[rid_start + 1U]) << 8U) |
                                      (static_cast<std::uint32_t>(payload[rid_start + 2U]) << 16U) |
                                      (static_cast<std::uint32_t>(payload[rid_start + 3U]) << 24U);
      if (rid_units > kMaxRelIdUnits) {
        return make_error(FormulonErrorCode::kInvalidArgument, "xlsb BrtHLink relationship id exceeds string limit",
                          "context=xlsb_sheet_writer cch=" + std::to_string(rid_units));
      }
    }
    if (auto r = EmitBoundedHyperlinkWideString(payload, hyperlink.location, kMaxLocationUnits, "location"); !r) {
      return r.error();
    }
    if (auto r = EmitBoundedHyperlinkWideString(payload, hyperlink.tooltip, kMaxTooltipUnits, "tooltip"); !r) {
      return r.error();
    }
    if (auto r = EmitBoundedHyperlinkWideString(payload, hyperlink.display, kMaxDisplayUnits, "display"); !r) {
      return r.error();
    }
    emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtHLink), payload);
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

std::vector<std::string> hyperlink_relationship_ids(const Sheet& sheet) {
  std::unordered_set<std::string> used;
  std::unordered_map<std::string, std::string> assigned_targets;
  used.reserve(sheet.unknown_relationships().size() + sheet.hyperlinks().size());
  assigned_targets.reserve(sheet.hyperlinks().size());
  for (const io::UnknownRelationship& relationship : sheet.unknown_relationships()) {
    if (!relationship.id.empty()) {
      used.insert(relationship.id);
    }
  }
  std::size_t next_id = 1U;
  std::vector<std::string> ids;
  ids.reserve(sheet.hyperlinks().size());
  auto fresh_id = [&]() {
    std::string id;
    do {
      id = "rId" + std::to_string(next_id++);
    } while (used.count(id) != 0U);
    used.insert(id);
    return id;
  };
  for (const Hyperlink& hyperlink : sheet.hyperlinks()) {
    if (hyperlink.target.empty()) {
      ids.emplace_back();
      continue;
    }
    if (!hyperlink.rid.empty()) {
      const auto assigned = assigned_targets.find(hyperlink.rid);
      if (assigned != assigned_targets.end() && assigned->second == hyperlink.target) {
        // Multiple BrtHLink records may legitimately share one relationship
        // when they point at the same external target. Preserve that source
        // sharing instead of minting a semantically equivalent duplicate.
        ids.push_back(hyperlink.rid);
        continue;
      }
      if (used.count(hyperlink.rid) == 0U) {
        ids.push_back(hyperlink.rid);
        used.insert(hyperlink.rid);
        assigned_targets.emplace(hyperlink.rid, hyperlink.target);
        continue;
      }
    }
    const std::string fresh = fresh_id();
    assigned_targets.emplace(fresh, hyperlink.target);
    ids.push_back(fresh);
  }
  return ids;
}

Expected<std::vector<std::uint8_t>, Error> emit_sheet(const Sheet& sheet, SstBuilder& sst,
                                                      const std::vector<std::string>& sheet_names,
                                                      const SheetRangeTable& sheet_ranges, const NameTable& name_table,
                                                      std::uint32_t* downgraded_formula_count,
                                                      std::uint32_t dynamic_array_ifmd) {
  std::vector<std::uint8_t> body;

  // Frame: BrtBeginSheet | BrtBeginSheetData | ... | BrtEndSheetData |
  // BrtEndSheet.
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginSheet), ByteSpan{});
  EmitWorksheetProperties(body, sheet);
  EmitWorksheetDimensions(body, sheet);
  EmitWorksheetViewsAndFormatting(body, sheet.view(), sheet.format_defaults());
  EmitColumnInfos(body, sheet.layout());
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginSheetData), ByteSpan{});

  // Spilled dynamic-array formulas need special treatment: Excel stores
  // every cell of the spill footprint (the anchor and its phantoms) as a
  // PtgExp "shell" cell record, with the real tokens carried once in the
  // anchor's `BrtArrFmla`. Phantom cells are not stored in `Sheet::rows()`
  // (they live in the spill table), so gather them here keyed by their
  // absolute (row, col) before walking rows. Only genuine spills (a
  // footprint larger than one cell) are treated this way; a scalar result
  // stays an ordinary formula cell.
  struct PhantomShell {
    std::uint32_t anchor_row = 0;
    std::uint32_t anchor_col = 0;
    Value value = Value::blank();
  };
  std::map<std::uint32_t, std::map<std::uint32_t, PhantomShell>> phantoms;
  for (const auto& [anchor_row, cells] : sheet.rows()) {
    for (std::uint32_t col = 0; col < cells.size(); ++col) {
      const Cell& cell = cells[col];
      if (cell.formula_text.empty()) {
        continue;
      }
      const SpillRegion* region = sheet.spill_region_at_anchor(anchor_row, col);
      if (region == nullptr || static_cast<std::uint64_t>(region->rows) * region->cols <= 1U) {
        continue;
      }
      for (std::uint32_t r = 0; r < region->rows; ++r) {
        for (std::uint32_t c = 0; c < region->cols; ++c) {
          if (r == 0 && c == 0) {
            continue;  // anchor: emitted as an anchor record, not a phantom
          }
          const std::size_t idx = static_cast<std::size_t>(r) * region->cols + c;
          Value value = idx < region->cells.size() ? region->cells[idx] : Value::blank();
          phantoms[anchor_row + r][col + c] = PhantomShell{anchor_row, col, std::move(value)};
        }
      }
    }
  }

  // Walk rows in ascending order. The Sheet's row map is unordered, so we
  // collect indices first (union of stored rows and phantom-only rows) and
  // sort. Workbooks rarely reach more than a few thousand populated rows,
  // so this is comfortably cheap.
  std::vector<std::uint32_t> row_indices;
  row_indices.reserve(sheet.rows().size() + phantoms.size() + sheet.layout().row_overrides.size());
  for (const auto& kv : sheet.rows()) {
    row_indices.push_back(kv.first);
  }
  for (const auto& kv : phantoms) {
    row_indices.push_back(kv.first);
  }
  for (const RowLayout& layout : sheet.layout().row_overrides) {
    if (layout.row < Sheet::kMaxRows) {
      row_indices.push_back(layout.row);
    }
  }
  std::sort(row_indices.begin(), row_indices.end());
  row_indices.erase(std::unique(row_indices.begin(), row_indices.end()), row_indices.end());

  std::unordered_set<std::uint64_t> downgraded_array_anchors;
  auto anchor_key = [](std::uint32_t row, std::uint32_t col) { return (static_cast<std::uint64_t>(row) << 32U) | col; };

  for (const std::uint32_t row : row_indices) {
    const RowLayout* row_layout = nullptr;
    for (const RowLayout& candidate : sheet.layout().row_overrides) {
      if (candidate.row == row) {
        row_layout = &candidate;
        break;
      }
    }
    const auto stored_it = sheet.rows().find(row);
    const RowCells* row_cells = stored_it != sheet.rows().end() ? &stored_it->second : nullptr;
    const auto phantom_it = phantoms.find(row);
    const std::map<std::uint32_t, PhantomShell>* row_phantoms =
        phantom_it != phantoms.end() ? &phantom_it->second : nullptr;

    // Highest column carrying anything in this row (stored non-empty cell
    // or phantom). A row with nothing to emit is skipped so we do not emit
    // a bare BrtRowHdr.
    std::uint32_t max_col = 0;
    bool any = false;
    if (row_cells != nullptr) {
      for (std::uint32_t col = 0; col < row_cells->size(); ++col) {
        if (!IsEmptySlot((*row_cells)[col])) {
          max_col = std::max(max_col, col);
          any = true;
        }
      }
    }
    if (row_phantoms != nullptr) {
      for (const auto& kv : *row_phantoms) {
        max_col = std::max(max_col, kv.first);
        any = true;
      }
    }
    if (!any && row_layout == nullptr) {
      continue;
    }

    // BrtRowHdr represents cells in 1,024-column segments.  Include only
    // cells actually emitted below; leading default-constructed slots must
    // not enlarge a span.
    std::map<std::uint32_t, std::pair<std::uint32_t, std::uint32_t>> spans;
    auto add_to_span = [&spans](std::uint32_t col) {
      const std::uint32_t segment = col / 1024U;
      const auto [it, inserted] = spans.emplace(segment, std::make_pair(col, col));
      if (!inserted) {
        it->second.first = std::min(it->second.first, col);
        it->second.second = std::max(it->second.second, col);
      }
    };
    if (row_cells != nullptr) {
      for (std::uint32_t col = 0; col < row_cells->size(); ++col) {
        if (!IsEmptySlot((*row_cells)[col])) {
          add_to_span(col);
        }
      }
    }
    if (row_phantoms != nullptr) {
      for (const auto& [col, phantom] : *row_phantoms) {
        (void)phantom;
        add_to_span(col);
      }
    }

    EmitRowHeader(body, row, row_layout, spans);
    for (std::uint32_t col = 0; col <= max_col; ++col) {
      const Cell* cell = row_cells != nullptr && col < row_cells->size() ? &(*row_cells)[col] : nullptr;
      if (cell != nullptr && !IsEmptySlot(*cell)) {
        const SpillRegion* region = cell->formula_text.empty() ? nullptr : sheet.spill_region_at_anchor(row, col);
        if (region != nullptr) {
          const std::uint32_t last_row = row + region->rows - 1U;
          const std::uint32_t last_col = col + region->cols - 1U;
          const Value& anchor_value = !region->cells.empty() ? region->cells.front() : cell->cached_value;
          bool downgraded_to_literal = false;
          if (dynamic_array_ifmd != 0U) {
            std::vector<std::uint8_t> metadata_index;
            emit_u32(metadata_index, dynamic_array_ifmd);  // XLDAPR dynamic-array metadata entry
            emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtCellMeta), metadata_index);
          }
          if (auto r =
                  emit_array_anchor(body, *cell, anchor_value, col, row, last_row, last_col, sheet_names, sheet_ranges,
                                    name_table, sst, downgraded_formula_count, &downgraded_to_literal);
              !r) {
            return r.error();
          }
          if (downgraded_to_literal) {
            downgraded_array_anchors.insert(anchor_key(row, col));
          }
        } else if (auto r = emit_cell(body, *cell, row, col, sst, sheet_names, sheet_ranges, name_table,
                                      downgraded_formula_count);
                   !r) {
          return r.error();
        }
        continue;
      }
      if (row_phantoms != nullptr) {
        const auto ph = row_phantoms->find(col);
        if (ph != row_phantoms->end()) {
          if (downgraded_array_anchors.count(anchor_key(ph->second.anchor_row, ph->second.anchor_col)) != 0U) {
            continue;
          }
          const std::uint32_t xf_index = cell != nullptr ? cell->xf_index : 0U;
          emit_array_phantom(body, col, xf_index, ph->second.value, ph->second.anchor_row, ph->second.anchor_col);
        }
      }
    }
  }

  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtEndSheetData), ByteSpan{});
  // Retained tail records bracket the merged-cell block so a sheet read from
  // an .xlsb keeps its conditional formatting, data validation, hyperlinks,
  // auto-filter and print setup in their original stream positions. The
  // buffers hold already-framed records, so this is a byte append.
  const XlsbSheetTail& tail = sheet.xlsb_tail();
  const std::vector<std::string> hyperlink_rids = hyperlink_relationship_ids(sheet);
  body.insert(body.end(), tail.before_merges.begin(), tail.before_merges.end());
  EmitMerges(body, sheet);
  body.insert(body.end(), tail.after_merges_before_hyperlinks.begin(), tail.after_merges_before_hyperlinks.end());
  if (auto hyperlinks = EmitHyperlinks(body, sheet, hyperlink_rids); !hyperlinks) {
    return hyperlinks.error();
  }
  body.insert(body.end(), tail.after_hyperlinks.begin(), tail.after_hyperlinks.end());
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtEndSheet), ByteSpan{});
  return body;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
