#include "io/xlsb/view_records.h"

#include <array>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <vector>

#include "io/sqref_text.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
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
  // Excel writes 0 here whichever range holds the active cell.
  emit_u32(payload, 0U);
  emit_u32(payload, static_cast<std::uint32_t>(selection.ranges.size()));
  for (const MergeRange& range : selection.ranges) {
    emit_u32(payload, range.first_row);
    emit_u32(payload, range.last_row);
    emit_u32(payload, range.first_col);
    emit_u32(payload, range.last_col);
  }
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtSel), payload);
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
  auto skip_active_range = read_u32(p);
  auto count_or = read_u32(p);
  // Sized by division so a huge count cannot wrap a 32-bit size_t.
  if (!pane_or || !row_or || !col_or || !skip_active_range || !count_or || count_or.value() == 0U ||
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
                                           a1_ranges ? std::string() : format_sqref_text(ranges)});
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

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
