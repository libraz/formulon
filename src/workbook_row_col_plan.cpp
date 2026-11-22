// Preflight and band planning of row/column insert/delete: the checks run before an edit applies and the rect and band
// edits they map through.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "pivot/pivot_table.h"
#include "sheet.h"
#include "sheet_metadata.h"
#include "styles.h"
#include "table.h"
#include "utils/a1_ref.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "workbook.h"
#include "workbook_internal.h"
#include "workbook_sheet_mutation.h"

namespace formulon {

namespace {

// Whether cell format `xf_index` draws something: a fill or a border. Fill
// records 0 and 1 are the reserved none / gray125 entries no cell shows.
bool xf_is_visible(const StylesTable& styles, std::uint32_t xf_index) {
  if (xf_index == 0U || xf_index >= styles.cell_xfs.size()) {
    return false;
  }
  const CellXf& xf = styles.cell_xfs[xf_index];
  if (xf.fill_index > 1U && xf.fill_index < styles.fills.size()) {
    const FillRecord& fill = styles.fills[xf.fill_index];
    if (fill.has_gradient || fill.pattern != 0U) {
      return true;
    }
  }
  if (xf.border_index < styles.borders.size()) {
    const BorderRecord& border = styles.borders[xf.border_index];
    const bool diagonal = (border.diagonal_up || border.diagonal_down) && border.diagonal.style != 0U;
    if (border.left.style != 0U || border.right.style != 0U || border.top.style != 0U || border.bottom.style != 0U ||
        diagonal) {
      return true;
    }
  }
  return false;
}

// What an insert would push past the grid's edge that blocks it, or null: a
// value or formula, a note, a fill or border, or a merge, as Excel refuses to
// shift non-empty cells off the sheet. A cell format that draws nothing (a
// font or number format), row and column formats, hyperlinks, validations and
// conditional formats are dropped there instead.
const char* insert_pushes_content_off(const Sheet& sheet, const StylesTable& styles, const StructuralEdit& edit) {
  if (edit.is_delete) {
    return nullptr;
  }
  const std::uint32_t bound = a1::axis_extent(edit.row_axis);
  const std::uint32_t first_lost = std::max(edit.index, edit.count >= bound ? 0U : bound - edit.count);
  const std::uint32_t across_first = edit.band.has_value() ? edit.band->first : 0U;
  const std::uint32_t across_last = edit.band.has_value() ? edit.band->last : a1::axis_extent(!edit.row_axis) - 1U;
  // The strip that falls off the grid, as an inclusive rectangle.
  const MergeRange lost = edit.row_axis ? MergeRange{first_lost, across_first, bound - 1U, across_last}
                                        : MergeRange{across_first, first_lost, across_last, bound - 1U};
  const auto in_lost = [&lost](std::uint32_t row, std::uint32_t col) {
    return lost.first_row <= row && row <= lost.last_row && lost.first_col <= col && col <= lost.last_col;
  };
  bool formatted = false;
  for (const auto& [row, cells] : sheet.rows()) {
    if (row < lost.first_row || row > lost.last_row) {
      continue;
    }
    for (const RowCells::Entry& entry : cells.entries_in(lost.first_col, lost.last_col + 1ULL)) {
      if (!entry.cell.formula_text.empty() || !entry.cell.cached_value.is_blank()) {
        return "a value or formula";
      }
      formatted = formatted || xf_is_visible(styles, entry.cell.xf_index);
    }
    const auto [run_lo, run_hi] = cells.runs_in(lost.first_col, lost.last_col + 1ULL);
    for (const StyleRun* run = run_lo; run != run_hi; ++run) {
      formatted = formatted || xf_is_visible(styles, run->xf);
    }
  }
  for (const CellComment& comment : sheet.comments()) {
    if (in_lost(comment.row, comment.col)) {
      return "a note";
    }
  }
  for (const ThreadedComment& thread : sheet.threaded_comments()) {
    if (in_lost(thread.row, thread.col)) {
      return "a note";
    }
  }
  if (formatted) {
    return "a fill or border";
  }
  for (const MergeRange& merge : sheet.merges()) {
    if (merge.first_row <= lost.last_row && lost.first_row <= merge.last_row && merge.first_col <= lost.last_col &&
        lost.first_col <= merge.last_col) {
      return "a merge";
    }
  }
  return nullptr;
}

// Where `rect` lies after `edit`, or nullopt once the edit deletes it or
// pushes it off the grid. A rectangle straddling the band edge stays put, or
// is gone with `straddle_drops`.
std::optional<MergeRange> rect_after_edit(const MergeRange& rect, const StructuralEdit& edit, bool straddle_drops) {
  if (edit.band.has_value()) {
    const BandFit fit = band_fit(rect, edit);
    if (fit == BandFit::kOutside) {
      return rect;
    }
    if (fit == BandFit::kStraddles) {
      return straddle_drops ? std::nullopt : std::optional<MergeRange>(rect);
    }
  }
  MergeRange moved = rect;
  std::uint32_t& first = edit.row_axis ? moved.first_row : moved.first_col;
  std::uint32_t& last = edit.row_axis ? moved.last_row : moved.last_col;
  const std::uint64_t bound = a1::axis_extent(edit.row_axis);
  const std::uint64_t end = edit.end();
  if (!edit.is_delete) {
    const std::uint64_t new_first = first >= edit.index ? first + static_cast<std::uint64_t>(edit.count) : first;
    const std::uint64_t new_last = last >= edit.index ? last + static_cast<std::uint64_t>(edit.count) : last;
    if (new_first >= bound) {
      return std::nullopt;
    }
    first = static_cast<std::uint32_t>(new_first);
    last = static_cast<std::uint32_t>(std::min(new_last, bound - 1U));
    return moved;
  }
  const std::int64_t new_first = first < edit.index ? first : (first >= end ? first - edit.count : edit.index);
  const std::int64_t new_last =
      last < edit.index ? last : (last >= end ? last - edit.count : static_cast<std::int64_t>(edit.index) - 1);
  if (new_last < new_first) {
    return std::nullopt;
  }
  first = static_cast<std::uint32_t>(new_first);
  last = static_cast<std::uint32_t>(new_last);
  return moved;
}

}  // namespace

namespace detail {

// A data table, a legacy array block and a PivotTable each move or go whole,
// so an edit cutting into one is refused, as a write into it is. A table grows
// and shrinks with an edit inside it instead, but a row delete may take its
// header row only with the whole table. An insert that would push any of
// them, or a value or formula, past the grid edge is refused, as Excel
// refuses to shift non-empty cells off the sheet. A band edit is also refused
// when its moving region meets one of these, or a table, across a band edge,
// and when it lies inside a PivotTable. A band insert is refused when it
// would move a PivotTable at all, and when it cuts into a table without
// spanning exactly that table's columns (rows).
Expected<void, Error> check_structure_edit(const Workbook& wb, std::size_t sheet_index, const StructuralEdit& edit) {
  const auto refuse = [&](StructureEditCut cut, const char* what) {
    std::string message(cut == StructureEditCut::kPushesOff ? "Cannot shift " : "Cannot change part of ");
    message.append(what);
    if (cut == StructureEditCut::kPushesOff) {
      message.append(" off the sheet");
    }
    return make_error(FormulonErrorCode::kPreconditionFailed, std::move(message),
                      "origin=" + std::to_string(edit.index) + " count=" + std::to_string(edit.count));
  };
  const auto refused = [](StructureEditCut cut) {
    return cut == StructureEditCut::kSplits || cut == StructureEditCut::kPushesOff;
  };
  // A whole row/column edit meets every structure along its axis alone.
  const auto fit_of = [&edit](const MergeRange& rect) {
    return edit.band.has_value() ? band_fit(rect, edit) : BandFit::kInside;
  };
  const Sheet& sheet = wb.sheet(sheet_index);
  for (const DataTable& table : sheet.data_tables()) {
    const BandFit fit = fit_of(data_table_block(table));
    if (fit == BandFit::kStraddles) {
      return refuse(StructureEditCut::kSplits, "a data table");
    }
    if (const StructureEditCut cut = data_table_structure_cut(table, edit); fit == BandFit::kInside && refused(cut)) {
      return refuse(cut, "a data table");
    }
  }
  for (const CellAddress address : sheet.all_formula_cells()) {
    const std::optional<MergeRange> found = detail::array_block_rect(sheet, address);
    if (!found.has_value()) {
      continue;
    }
    const MergeRange& block = *found;
    const BandFit fit = fit_of(block);
    if (fit == BandFit::kStraddles) {
      return refuse(StructureEditCut::kSplits, "an array");
    }
    const StructureEditCut cut = edit.row_axis ? structure_edit_cut(block.first_row, block.last_row, edit)
                                               : structure_edit_cut(block.first_col, block.last_col, edit);
    if (fit == BandFit::kInside && refused(cut)) {
      return refuse(cut, "an array");
    }
  }
  for (const std::unique_ptr<pivot::PivotTable>& pivot : sheet.pivot_tables()) {
    if (pivot == nullptr) {
      continue;
    }
    const MergeRange location = pivot_location(*pivot);
    const BandFit fit = fit_of(location);
    if (fit == BandFit::kStraddles) {
      return refuse(StructureEditCut::kSplits, "a PivotTable");
    }
    // Insert Cells never shifts a PivotTable along.
    if (edit.band.has_value() && !edit.is_delete && fit == BandFit::kInside) {
      return refuse(StructureEditCut::kSplits, "a PivotTable");
    }
    // Cells selected inside a report are the report's own to change.
    if (edit.band.has_value()) {
      const std::uint32_t last = edit.index + edit.count - 1U;
      const MergeRange selected = edit.row_axis ? MergeRange{edit.index, edit.band->first, last, edit.band->last}
                                                : MergeRange{edit.band->first, edit.index, edit.band->last, last};
      const bool within = location.first_row <= selected.first_row && selected.last_row <= location.last_row &&
                          location.first_col <= selected.first_col && selected.last_col <= location.last_col;
      if (within && !(edit.is_delete && selected == location)) {
        return refuse(StructureEditCut::kSplits, "a PivotTable");
      }
    }
    if (const StructureEditCut cut = pivot_structure_cut(*pivot, edit); fit == BandFit::kInside && refused(cut)) {
      return refuse(cut, "a PivotTable");
    }
  }
  for (const TableMetadata& table : wb.tables()) {
    const std::optional<MergeRange> ref = detail::table_rect_on_sheet(table, sheet_index);
    if (!ref.has_value()) {
      continue;
    }
    const BandFit fit = fit_of(*ref);
    if (fit == BandFit::kStraddles) {
      return refuse(StructureEditCut::kSplits, "a table");
    }
    const StructureEditCut cut = edit.row_axis ? structure_edit_cut(ref->first_row, ref->last_row, edit)
                                               : structure_edit_cut(ref->first_col, ref->last_col, edit);
    if (fit == BandFit::kInside && cut == StructureEditCut::kPushesOff) {
      return refuse(cut, "a table");
    }
    // Insert Cells growing a table spans exactly its columns (rows), totals row included.
    if (edit.band.has_value() && !edit.is_delete && fit == BandFit::kInside && cut == StructureEditCut::kSplits &&
        (edit.band->first != (edit.row_axis ? ref->first_col : ref->first_row) ||
         edit.band->last != (edit.row_axis ? ref->last_col : ref->last_row))) {
      return refuse(cut, "a table");
    }
    const std::uint64_t end = edit.end();
    if (edit.is_delete && edit.row_axis && table.header_row && fit == BandFit::kInside &&
        edit.index <= ref->first_row && ref->first_row < end && end <= ref->last_row) {
      return make_error(FormulonErrorCode::kPreconditionFailed,
                        "Cannot delete a table's header row without the whole table",
                        "origin=" + std::to_string(edit.index) + " count=" + std::to_string(edit.count));
    }
  }
  if (const char* blocker = insert_pushes_content_off(sheet, wb.styles(), edit); blocker != nullptr) {
    return refuse(StructureEditCut::kPushesOff, blocker);
  }
  return Expected<void, Error>::Ok();
}

// The argument checks of a row/column edit on `sheet_index`.
Expected<void, Error> check_edit_arguments(const Workbook& wb, std::size_t sheet_index, const StructuralEdit& edit) {
  RETURN_IF_ERROR(check_sheet_index(sheet_index, wb.sheet_count()));
  if (edit.count == 0U) {
    return make_error(FormulonErrorCode::kInvalidArgument, "count must be >= 1");
  }
  const std::uint32_t bound = a1::axis_extent(edit.row_axis);
  if (edit.index >= bound) {
    return make_error(FormulonErrorCode::kInvalidArgument, "first is out of bounds",
                      "origin=" + std::to_string(edit.index));
  }
  // Deletion helpers use `origin + count` as their exclusive endpoint.
  // Validate against the remaining grid with subtraction so an unsigned
  // count received through the C/WASM ABI cannot wrap and turn a delete
  // into a corrupting row/column move.
  if (edit.is_delete && edit.count > bound - edit.index) {
    return make_error(FormulonErrorCode::kInvalidArgument, "count exceeds the sheet bounds",
                      "origin=" + std::to_string(edit.index) + " count=" + std::to_string(edit.count));
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> apply_row_col_edit(const Workbook& wb, std::size_t sheet_index, const StructuralEdit& edit) {
  RETURN_IF_ERROR(check_edit_arguments(wb, sheet_index, edit));
  return detail::check_structure_edit(wb, sheet_index, edit);
}

// The ids of the tables with a header row whose every data row a row-axis
// delete `edit` takes without taking the whole table: the delete keeps their
// first data row, the edit's first row. Refuses the delete when it also covers
// another table's rows that keeping that row would change, so only a table
// lying wholly below the kept row may share the delete.
Expected<std::vector<std::uint32_t>, Error> tables_emptied_by(const Workbook& wb, std::size_t sheet_index,
                                                              const StructuralEdit& edit) {
  std::vector<std::uint32_t> emptied;
  if (!edit.is_delete || !edit.row_axis) {
    return emptied;
  }
  const std::uint64_t end = edit.end();
  const auto reached = [&edit](const MergeRange& rect) {
    return !edit.band.has_value() || band_fit(rect, edit) != BandFit::kOutside;
  };
  for (const TableMetadata& table : wb.tables()) {
    const std::optional<MergeRange> rect = detail::table_rect_on_sheet(table, sheet_index);
    if (!rect.has_value() || !table.header_row || !reached(*rect)) {
      continue;
    }
    const std::uint32_t last_data = rect->last_row - (table.totals_row ? 1U : 0U);
    if (rect->first_row < last_data && edit.index == rect->first_row + 1U && end > last_data) {
      emptied.push_back(table.id);
    }
  }
  if (emptied.empty()) {
    return emptied;
  }
  for (const TableMetadata& table : wb.tables()) {
    const std::optional<MergeRange> rect = detail::table_rect_on_sheet(table, sheet_index);
    if (!rect.has_value() || std::find(emptied.begin(), emptied.end(), table.id) != emptied.end() ||
        rect->last_row < edit.index || rect->first_row >= end || !reached(*rect)) {
      continue;
    }
    if (rect->first_row <= edit.index || rect->last_row >= end) {
      return make_error(FormulonErrorCode::kPreconditionFailed,
                        "Cannot delete every data row of a table together with part of another table",
                        "origin=" + std::to_string(edit.index) + " count=" + std::to_string(edit.count));
    }
  }
  return emptied;
}

Expected<StructuralEdit, Error> band_edit_of(const MergeRange& range, CellShift shift) {
  const bool known =
      shift == CellShift::kDown || shift == CellShift::kRight || shift == CellShift::kUp || shift == CellShift::kLeft;
  if (!known) {
    return make_error(FormulonErrorCode::kInvalidArgument, "shift must be down, right, up or left");
  }
  if (range.first_row > range.last_row || range.first_col > range.last_col || range.last_row >= Sheet::kMaxRows ||
      range.last_col >= Sheet::kMaxCols) {
    return make_error(FormulonErrorCode::kInvalidArgument, "range is outside the grid");
  }
  const bool vertical = shift == CellShift::kDown || shift == CellShift::kUp;
  const bool whole_cols = range.first_row == 0U && range.last_row == Sheet::kMaxRows - 1U;
  const bool whole_rows = range.first_col == 0U && range.last_col == Sheet::kMaxCols - 1U;
  if (vertical ? whole_cols : whole_rows) {
    return make_error(FormulonErrorCode::kInvalidArgument,
                      vertical ? "range covers whole columns, which cannot shift vertically"
                               : "range covers whole rows, which cannot shift horizontally");
  }
  StructuralEdit edit{vertical ? range.first_row : range.first_col,
                      vertical ? range.last_row - range.first_row + 1U : range.last_col - range.first_col + 1U,
                      shift == CellShift::kUp || shift == CellShift::kLeft, vertical};
  // A band over the whole perpendicular extent is the whole row/column edit.
  if (!(vertical ? whole_rows : whole_cols)) {
    edit.band = vertical ? Band{range.first_col, range.last_col} : Band{range.first_row, range.last_row};
  }
  return edit;
}

std::optional<MergeRange> rect_after_shifts(const MergeRange& rect, std::size_t sheet_index,
                                            const std::vector<PendingShift>& pending, bool straddle_drops) {
  std::optional<MergeRange> moved = rect;
  for (const PendingShift& shift : pending) {
    if (!moved.has_value()) {
      break;
    }
    const Expected<StructuralEdit, Error> edit = band_edit_of(shift.range, shift.shift);
    if (shift.sheet_index == sheet_index && edit.has_value()) {
      moved = rect_after_edit(*moved, edit.value(), straddle_drops);
    }
  }
  return moved;
}

Expected<void, Error> check_band_edit_after(const Workbook& wb, std::size_t sheet_index, const MergeRange& range,
                                            CellShift shift, const std::vector<PendingShift>& pending) {
  if (pending.empty()) {
    return wb.check_band_edit(sheet_index, range, shift);
  }
  RETURN_IF_ERROR(check_sheet_index(sheet_index, wb.sheet_count()));
  ASSIGN_OR_RETURN(StructuralEdit edit, band_edit_of(range, shift));
  // Maps the edit back onto the current cells, latest shift first.
  const std::uint32_t across = edit.band.has_value() ? edit.band->first : 0U;
  std::uint64_t pushed = edit.count;
  for (auto it = pending.rbegin(); it != pending.rend(); ++it) {
    const Expected<StructuralEdit, Error> prior = band_edit_of(it->range, it->shift);
    if (it->sheet_index != sheet_index || !prior.has_value() || prior.value().row_axis != edit.row_axis ||
        (prior.value().band.has_value() && (across < prior.value().band->first || across > prior.value().band->last))) {
      continue;
    }
    const StructuralEdit& before = prior.value();
    if (before.is_delete) {
      edit.index = edit.index >= before.index ? edit.index + before.count : edit.index;
    } else {
      edit.index =
          edit.index >= before.index + before.count ? edit.index - before.count : std::min(edit.index, before.index);
      // Whatever the earlier insert pushes along, this one pushes further.
      pushed += edit.is_delete ? 0U : before.count;
    }
  }
  const std::uint32_t bound = a1::axis_extent(edit.row_axis);
  edit.index = std::min(edit.index, bound - 1U);
  edit.count = static_cast<std::uint32_t>(std::min<std::uint64_t>(pushed, bound));
  return apply_row_col_edit(wb, sheet_index, edit);
}

}  // namespace detail

Expected<void, Error> Workbook::check_band_edit(std::size_t sheet_index, const MergeRange& range,
                                                CellShift shift) const {
  RETURN_IF_ERROR(check_sheet_index(sheet_index, sheets_.size()));
  ASSIGN_OR_RETURN(const StructuralEdit edit, detail::band_edit_of(range, shift));
  RETURN_IF_ERROR(detail::check_edit_arguments(*this, sheet_index, edit));
  RETURN_IF_ERROR(detail::tables_emptied_by(*this, sheet_index, edit));
  return detail::check_structure_edit(*this, sheet_index, edit);
}

}  // namespace formulon
