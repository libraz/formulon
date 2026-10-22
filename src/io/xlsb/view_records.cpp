#include "io/xlsb/view_records.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/sheet_view_xml.h"
#include "io/sqref_text.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/a1_ref.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

// `pnn` order shared by BrtPane's active pane and BrtSel's pane.
constexpr std::array<std::string_view, 4> kPaneNames = {"bottomRight", "topRight", "bottomLeft", "topLeft"};
constexpr std::uint32_t kTopLeftPane = 3U;

// BrtPane flags: fFrozen, and fFrozenNoSplit for a freeze with no split
// behind it (a `frozen` rather than `frozenSplit` pane).
constexpr std::uint8_t kPaneFrozen = 0x01U;
constexpr std::uint8_t kPaneFrozenNoSplit = 0x02U;

// BrtSel payload: pnn, rwAct, colAct, the active range's index and the range
// count, then that many four-u32 RfX ranges.
constexpr std::size_t kRfxBytes = 16U;

// `kPaneNames.size()` for a name that is not a pane.
std::uint32_t pane_index(std::string_view pane) {
  for (std::size_t i = 0; i < kPaneNames.size(); ++i) {
    if (kPaneNames[i] == SheetView::pane_key(pane)) {
      return static_cast<std::uint32_t>(i);
    }
  }
  return static_cast<std::uint32_t>(kPaneNames.size());
}

void emit_selection(std::vector<std::uint8_t>& dst, const ResolvedSelection& selection) {
  std::vector<std::uint8_t> payload;
  emit_u32(payload, pane_index(selection.pane));
  emit_u32(payload, selection.active_row);
  emit_u32(payload, selection.active_col);
  emit_u32(payload, selection.active_cell_id);  // dwRfxAct
  emit_u32(payload, static_cast<std::uint32_t>(selection.ranges.size()));
  for (const MergeRange& range : selection.ranges) {
    emit_u32(payload, range.first_row);
    emit_u32(payload, range.last_row);
    emit_u32(payload, range.first_col);
    emit_u32(payload, range.last_col);
  }
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtSel), payload);
}

// BrtBeginWsView flag bits, fWnProt (bit 0) through fWhitespaceHidden (bit 10).
constexpr std::uint16_t kWnProt = 0x0001U;
constexpr std::uint16_t kDspFmla = 0x0002U;
constexpr std::uint16_t kDspGrid = 0x0004U;
constexpr std::uint16_t kDspRwCol = 0x0008U;
constexpr std::uint16_t kDspZeros = 0x0010U;
constexpr std::uint16_t kRightToLeft = 0x0020U;
constexpr std::uint16_t kSelected = 0x0040U;
constexpr std::uint16_t kDspRuler = 0x0080U;
constexpr std::uint16_t kDspGuts = 0x0100U;
constexpr std::uint16_t kDefaultHdr = 0x0200U;
constexpr std::uint16_t kWhitespaceHidden = 0x0400U;

// XLView: XLVNORMAL 0, XLVSHEETLAYOUTVIEW 1 (page break preview), XLVPAGELAYOUTVIEW 2.
constexpr std::array<std::string_view, 3> kViewModes = {"", "pageBreakPreview", "pageLayout"};

// `kViewModes.size()` for a mode XLView cannot name.
std::uint32_t view_mode_index(std::string_view mode) {
  if (mode == "normal") {
    return 0U;
  }
  for (std::size_t i = 0; i < kViewModes.size(); ++i) {
    if (kViewModes[i] == mode) {
      return static_cast<std::uint32_t>(i);
    }
  }
  return static_cast<std::uint32_t>(kViewModes.size());
}

std::uint16_t zoom_or_absent(std::uint32_t zoom) {
  return static_cast<std::uint16_t>(zoom >= 10U && zoom <= 400U ? zoom : 0U);
}

/// Calls `fn(view, workbook_view, written)` for each `<sheetView>` of
/// `extra_views_xml`, `written` when its window is one of `workbook_views`;
/// false when the fragment does not parse.
template <typename Fn>
bool for_each_extra_view(const std::string& extra_views_xml, std::uint32_t workbook_views, Fn&& fn) {
  if (extra_views_xml.empty()) {
    return true;
  }
  pugi::xml_document doc;
  if (!doc.load_buffer(extra_views_xml.data(), extra_views_xml.size(), pugi::parse_default | pugi::parse_fragment)) {
    return false;
  }
  for (pugi::xml_node node = doc.child("sheetView"); node; node = node.next_sibling("sheetView")) {
    SheetView view;
    read_sheet_view_element(node, view);
    const std::uint32_t workbook_view = attr_u32(node, "workbookViewId", 0U);
    fn(view, workbook_view, workbook_view != 0U && workbook_view < workbook_views);
  }
  return true;
}

}  // namespace

Expected<bool, Error> decode_pane(const XlsbRecord& rec, SheetView& view) {
  // BrtPane: Xnum column and row splits, the lower-right pane's top-left
  // cell, the active pane and the flags. A frozen pane's splits are counts,
  // a split's are twips.
  ByteSpan p = rec.payload;
  if (p.size < 29U) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtPane truncated", "context=xlsb_reader");
  }
  double cols = 0.0;
  double rows = 0.0;
  std::memcpy(&cols, p.data, sizeof(cols));
  std::memcpy(&rows, p.data + sizeof(cols), sizeof(rows));
  p.data += 16U;
  p.size -= 16U;
  auto top_row_or = read_u32(p);
  auto left_col_or = read_u32(p);
  auto active_pane_or = read_u32(p);
  auto flags_or = read_u8(p);
  if (!top_row_or || !left_col_or || !active_pane_or || !flags_or) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtPane fields truncated",
                      "context=xlsb_reader");
  }
  const bool valid = std::isfinite(rows) && std::isfinite(cols) && rows >= 0.0 && cols >= 0.0;
  if ((flags_or.value() & kPaneFrozen) == 0U) {
    if (!valid || (rows == 0.0 && cols == 0.0)) {
      return false;
    }
    view.split_x = cols;
    view.split_y = rows;
  } else {
    if (!valid || std::trunc(rows) != rows || std::trunc(cols) != cols || rows >= Sheet::kMaxRows ||
        cols >= Sheet::kMaxCols) {
      return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtPane frozen counts invalid",
                        "context=xlsb_reader");
    }
    view.freeze_rows = static_cast<std::uint32_t>(rows);
    view.freeze_cols = static_cast<std::uint32_t>(cols);
    view.frozen_split = view.is_frozen() && (flags_or.value() & kPaneFrozenNoSplit) == 0U;
  }
  if (top_row_or.value() < Sheet::kMaxRows && left_col_or.value() < Sheet::kMaxCols) {
    view.pane_top_left_cell = a1::encode_a1(top_row_or.value(), left_col_or.value());
  }
  if (active_pane_or.value() < kPaneNames.size()) {
    view.active_pane = std::string(kPaneNames[active_pane_or.value()]);
  }
  return true;
}

bool decode_selection(const XlsbRecord& rec, SheetView& view) {
  ByteSpan p = rec.payload;
  auto pane_or = read_u32(p);
  auto row_or = read_u32(p);
  auto col_or = read_u32(p);
  auto active_range_or = read_u32(p);
  auto count_or = read_u32(p);
  // Sized by division so a huge count cannot wrap a 32-bit size_t.
  if (!pane_or || !row_or || !col_or || !active_range_or || !count_or || count_or.value() == 0U ||
      pane_or.value() >= kPaneNames.size() || !Sheet::coord_in_grid(row_or.value(), col_or.value()) ||
      p.size % kRfxBytes != 0U || p.size / kRfxBytes != count_or.value()) {
    return false;
  }
  std::vector<MergeRange> ranges;
  for (std::uint32_t i = 0; i < count_or.value(); ++i) {
    MergeRange range;
    range.first_row = read_u32(p).value();
    range.last_row = read_u32(p).value();
    range.first_col = read_u32(p).value();
    range.last_col = read_u32(p).value();
    if (!Sheet::coord_in_grid(range.first_row, range.first_col) ||
        !Sheet::coord_in_grid(range.last_row, range.last_col) || range.first_row > range.last_row ||
        range.first_col > range.last_col) {
      return false;
    }
    ranges.push_back(range);
  }
  // Excel's XLSX omits an A1 active cell and an A1 sqref, and the top-left
  // pane's selection when both are A1; the model keeps that spelling.
  const bool a1_active = row_or.value() == 0U && col_or.value() == 0U;
  const bool a1_ranges = ranges.size() == 1U && ranges.front() == MergeRange{};
  if (a1_active && a1_ranges && pane_or.value() == kTopLeftPane) {
    return true;
  }
  view.selections.push_back(SheetSelection{std::string(kPaneNames[pane_or.value()]),
                                           a1_active ? std::string() : a1::encode_a1(row_or.value(), col_or.value()),
                                           a1_ranges ? std::string() : format_sqref_text(ranges),
                                           active_range_or.value() < ranges.size() ? active_range_or.value() : 0U});
  return true;
}

void emit_pane_and_selections(std::vector<std::uint8_t>& dst, const SheetView& view) {
  if (view.divides_rows() || view.divides_cols()) {
    std::vector<std::uint8_t> pane;
    const bool frozen = view.is_frozen();
    emit_double(pane, frozen ? static_cast<double>(view.freeze_cols) : view.split_x);
    emit_double(pane, frozen ? static_cast<double>(view.freeze_rows) : view.split_y);
    std::uint32_t pane_row = 0;
    std::uint32_t pane_col = 0;
    pane_origin(view, &pane_row, &pane_col);
    emit_u32(pane, pane_row);
    emit_u32(pane, pane_col);
    emit_u32(pane, pane_index(view.effective_active_pane()));
    std::uint8_t flags = 0U;
    if (frozen) {
      flags = view.frozen_split ? kPaneFrozen : static_cast<std::uint8_t>(kPaneFrozen | kPaneFrozenNoSplit);
    }
    emit_u8(pane, flags);
    emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtPane), pane);
  }
  for (const ResolvedSelection& selection : resolve_view_selections(view)) {
    emit_selection(dst, selection);
  }
}

Expected<void, Error> decode_ws_view(const XlsbRecord& rec, SheetView& view, std::uint32_t* workbook_view) {
  // BrtBeginWsView: flags u16, xlView u32, rwTop u32, colLeft u32, icvHdr u8,
  // two reserved fields, wScale u16, then wScaleNormal / wScaleSLV / wScalePLV
  // u16 and iWbkView u32.
  ByteSpan p = rec.payload;
  auto flags_or = read_u16(p);
  auto mode_or = read_u32(p);
  auto top_row_or = read_u32(p);
  auto left_col_or = read_u32(p);
  auto color_or = read_u8(p);
  auto skip_reserved8 = read_u8(p);
  auto skip_reserved16 = read_u16(p);
  auto zoom_or = read_u16(p);
  if (!flags_or || !mode_or || !top_row_or || !left_col_or || !color_or || !skip_reserved8 || !skip_reserved16 ||
      !zoom_or) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtBeginWsView truncated",
                      "context=xlsb_reader");
  }
  if ((top_row_or.value() != 0U || left_col_or.value() != 0U) && top_row_or.value() < Sheet::kMaxRows &&
      left_col_or.value() < Sheet::kMaxCols) {
    view.top_left_cell = a1::encode_a1(top_row_or.value(), left_col_or.value());
  }
  const std::uint16_t flags = flags_or.value();
  view.window_protection = (flags & kWnProt) != 0U;
  view.show_formulas = (flags & kDspFmla) != 0U;
  view.show_grid_lines = (flags & kDspGrid) != 0U;
  view.show_row_col_headers = (flags & kDspRwCol) != 0U;
  view.show_zeros = (flags & kDspZeros) != 0U;
  view.right_to_left = (flags & kRightToLeft) != 0U;
  view.tab_selected = (flags & kSelected) != 0U;
  view.show_ruler = (flags & kDspRuler) != 0U;
  view.show_outline_symbols = (flags & kDspGuts) != 0U;
  view.default_grid_color = (flags & kDefaultHdr) != 0U;
  view.show_white_space = (flags & kWhitespaceHidden) == 0U;
  view.color_id = color_or.value();
  view.view_mode = mode_or.value() < kViewModes.size() ? std::string(kViewModes[mode_or.value()]) : std::string();
  if (zoom_or.value() >= 10U && zoom_or.value() <= 400U) {
    view.zoom_scale = zoom_or.value();
  }
  // wScaleNormal, wScaleSLV and wScalePLV follow wScale when the record carries them.
  for (std::uint32_t* target :
       {&view.zoom_scale_normal, &view.zoom_scale_sheet_layout_view, &view.zoom_scale_page_layout_view}) {
    auto scale_or = read_u16(p);
    if (!scale_or) {
      break;
    }
    *target = zoom_or_absent(scale_or.value());
  }
  auto workbook_view_or = read_u32(p);
  *workbook_view = workbook_view_or ? workbook_view_or.value() : 0U;
  return Expected<void, Error>::Ok();
}

void emit_ws_view(std::vector<std::uint8_t>& dst, const SheetView& view, std::uint32_t workbook_view) {
  std::vector<std::uint8_t> payload;
  std::uint16_t flags = 0U;
  const std::pair<bool, std::uint16_t> bits[] = {
      {view.window_protection, kWnProt},
      {view.show_formulas, kDspFmla},
      {view.show_grid_lines, kDspGrid},
      {view.show_row_col_headers, kDspRwCol},
      {view.show_zeros, kDspZeros},
      {view.right_to_left, kRightToLeft},
      {view.tab_selected, kSelected},
      {view.show_ruler, kDspRuler},
      {view.show_outline_symbols, kDspGuts},
      {view.default_grid_color, kDefaultHdr},
      {!view.show_white_space, kWhitespaceHidden},
  };
  for (const auto& [on, bit] : bits) {
    if (on) {
      flags |= bit;
    }
  }
  emit_u16(payload, flags);
  const std::uint32_t mode = view_mode_index(view.view_mode);
  emit_u32(payload, mode < kViewModes.size() ? mode : 0U);  // xlView
  std::uint32_t top_row = 0U;
  std::uint32_t left_col = 0U;
  (void)a1::parse_a1_ref(view.top_left_cell, &top_row, &left_col);
  emit_u32(payload, top_row);                                                                  // rwTop
  emit_u32(payload, left_col);                                                                 // colLeft
  emit_u8(payload, static_cast<std::uint8_t>(std::min<std::uint32_t>(view.color_id, 0xFFU)));  // icvHdr
  emit_u8(payload, 0U);
  emit_u16(payload, 0U);
  emit_u16(payload, static_cast<std::uint16_t>(std::clamp(view.zoom_scale, 10U, 400U)));
  emit_u16(payload, zoom_or_absent(view.zoom_scale_normal));
  emit_u16(payload, zoom_or_absent(view.zoom_scale_sheet_layout_view));
  emit_u16(payload, zoom_or_absent(view.zoom_scale_page_layout_view));
  emit_u32(payload, workbook_view);  // iWbkView
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginWsView), payload);
  emit_pane_and_selections(dst, view);
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtEndWsView), ByteSpan{});
}

void emit_extra_ws_views(std::vector<std::uint8_t>& dst, const SheetView& view, std::uint32_t workbook_views) {
  (void)for_each_extra_view(view.extra_views_xml, workbook_views,
                            [&](const SheetView& extra, std::uint32_t workbook_view, bool written) {
                              if (written) {
                                emit_ws_view(dst, extra, workbook_view);
                              }
                            });
}

void append_extra_view(SheetView& main, const SheetView& extra, std::uint32_t workbook_view) {
  main.extra_views_xml.append(build_sheet_view_element_xml(extra, workbook_view));
}

std::size_t unwritten_view_part_count(const SheetView& view, std::uint32_t workbook_views) {
  const auto own_parts = [](const SheetView& v) -> std::size_t {
    return (v.tail_xml.empty() ? 0U : 1U) + (view_mode_index(v.view_mode) < kViewModes.size() ? 0U : 1U);
  };
  std::size_t count = own_parts(view);
  const bool parsed = for_each_extra_view(
      view.extra_views_xml, workbook_views,
      [&](const SheetView& extra, std::uint32_t, bool written) { count += written ? own_parts(extra) : 1U; });
  return parsed ? count : count + 1U;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
