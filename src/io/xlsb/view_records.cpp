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

// BrtPane flags: fFrozen, and fFrozenNoSplit for Excel's ordinary freeze.
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

void emit_selection(std::vector<std::uint8_t>& dst, const SheetSelection& selection) {
  std::vector<MergeRange> ranges;
  parse_selection_sqref_text(selection.sqref, ranges);
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  if (!a1::parse_a1_ref(selection.active_cell, &row, &col) && !ranges.empty()) {
    row = ranges.front().first_row;
    col = ranges.front().first_col;
  }
  if (ranges.empty()) {
    ranges.push_back({row, col, row, col});
  }
  std::vector<std::uint8_t> payload;
  emit_u32(payload, pane_index(selection.pane));
  emit_u32(payload, row);
  emit_u32(payload, col);
  // Excel writes 0 here whichever range holds the active cell.
  emit_u32(payload, 0U);
  emit_u32(payload, static_cast<std::uint32_t>(ranges.size()));
  for (const MergeRange& range : ranges) {
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
  // cell, the active pane and the flags. A frozen pane's splits are counts.
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
  if ((flags_or.value() & kPaneFrozen) == 0U) {
    return false;
  }
  if (!std::isfinite(rows) || !std::isfinite(cols) || rows < 0.0 || cols < 0.0 || std::trunc(rows) != rows ||
      std::trunc(cols) != cols || rows >= Sheet::kMaxRows || cols >= Sheet::kMaxCols) {
    return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtPane frozen counts invalid",
                      "context=xlsb_reader");
  }
  view.freeze_rows = static_cast<std::uint32_t>(rows);
  view.freeze_cols = static_cast<std::uint32_t>(cols);
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
  SheetSelection selection{std::string(kPaneNames[pane_or.value()]), std::string(), std::string()};
  // A1 with the single range A1 is what Excel's XLSX spells as an empty
  // `<selection pane="..."/>`, or omits altogether for the top-left pane.
  const bool default_a1 = row_or.value() == 0U && col_or.value() == 0U && ranges.size() == 1U &&
                          ranges.front().first_row == 0U && ranges.front().first_col == 0U &&
                          ranges.front().last_row == 0U && ranges.front().last_col == 0U;
  if (default_a1 && pane_or.value() == kTopLeftPane) {
    return true;
  }
  if (!default_a1) {
    selection.active_cell = a1::encode_a1(row_or.value(), col_or.value());
    selection.sqref = format_sqref_text(ranges);
  }
  view.selections.push_back(std::move(selection));
  return true;
}

void emit_pane_and_selections(std::vector<std::uint8_t>& dst, const SheetView& view) {
  const bool frozen = view.freeze_rows != 0U || view.freeze_cols != 0U;
  if (frozen) {
    std::vector<std::uint8_t> pane;
    emit_double(pane, static_cast<double>(view.freeze_cols));
    emit_double(pane, static_cast<double>(view.freeze_rows));
    std::uint32_t pane_row = view.freeze_rows;
    std::uint32_t pane_col = view.freeze_cols;
    (void)a1::parse_a1_ref(view.pane_top_left_cell, &pane_row, &pane_col);
    emit_u32(pane, pane_row);
    emit_u32(pane, pane_col);
    const std::uint32_t active = pane_index(view.effective_active_pane());
    emit_u32(pane, active < kPaneNames.size() ? active : 0U);
    emit_u8(pane, kPaneFrozen | kPaneFrozenNoSplit);
    emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtPane), pane);
  }
  // Excel writes the top-left pane's selection first, A1 when none is stored.
  const SheetSelection* top_left = nullptr;
  for (const SheetSelection& selection : view.selections) {
    if (SheetView::pane_key(selection.pane) == "topLeft") {
      top_left = &selection;
      break;
    }
  }
  emit_selection(dst, top_left != nullptr ? *top_left : SheetSelection{"topLeft", "A1", "A1"});
  // Panes only exist in a frozen view; an unfrozen one keeps the top-left
  // selection only. One record per pane, the first stored one winning.
  if (!frozen) {
    return;
  }
  std::array<bool, kPaneNames.size()> emitted{};
  emitted[kTopLeftPane] = true;
  for (const SheetSelection& selection : view.selections) {
    const std::uint32_t pane = pane_index(selection.pane);
    if (pane < kPaneNames.size() && !emitted[pane]) {
      emitted[pane] = true;
      emit_selection(dst, selection);
    }
  }
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
