//
// `Workbook::project_pivot`: writes a pivot's report range into the sheet
// cells, as Excel keeps a PivotTable's output as ordinary cells that formulas
// read.
//
// Each projected cell is a literal: shared-string text, a number (a numeric
// or date axis item included) or an error. Its format follows the cell's
// layout role: field buttons carry `pivotButton`, compact row labels are
// left-aligned and indented by depth, value cells apply the data field's
// number format, and under `mergeItem` the label area is centered. Structural
// blanks get no cell; empty value cells keep their format. A report with no
// field is a bordered placeholder frame of formatted empty cells.
//
// A projection is planned (evaluated and laid out) before anything is
// written. A planned range that is occupied or overflows the sheet either
// refuses the projection or withholds the report behind `#SPILL!` in its
// anchor, per `ReportPlacement`.
//
// @size-budget: 17 KB

#include <cassert>
#include <cstddef>
#include <cstdint>
#include <map>
#include <memory>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "eval/eval_profile_scope.h"
#include "eval/pivot_formula.h"
#include "io/styles_writer.h"
#include "merge_range.h"
#include "parser/ast_format.h"
#include "pivot/field_lookup.h"
#include "pivot/filter_engine.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_locale.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "styles.h"
#include "table.h"
#include "utils/a1_ref.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "value.h"
#include "workbook.h"
#include "workbook_internal.h"

namespace formulon {
namespace {

using pivot::PivotCellRole;
using pivot::PivotRect;

constexpr std::uint8_t kAlignLeft = 1;
constexpr std::uint8_t kAlignCenter = 2;

MergeRange to_merge(const PivotRect& r) {
  return MergeRange{r.row0, r.col0, r.row1, r.col1};
}

PivotRect to_rect(const MergeRange& m) {
  return PivotRect{m.first_row, m.first_col, m.last_row, m.last_col};
}

// The range a pivot holds before its first projection: its location, and the
// page fields plus their gap row above it.
MergeRange location_extent(const pivot::PivotTable& table) {
  const std::uint32_t rows = table.span_rows() == 0U ? 1U : table.span_rows();
  const std::uint32_t cols = table.span_cols() == 0U ? 1U : table.span_cols();
  MergeRange r{table.anchor_row(), table.anchor_col(), table.anchor_row() + rows - 1U, table.anchor_col() + cols - 1U};
  const std::uint32_t page_rows = table.location_row_page_count().value_or(0U);
  if (page_rows > 0U) {
    r.first_row = table.anchor_row() > page_rows ? table.anchor_row() - page_rows - 1U : 0U;
    const std::uint32_t page_cols = table.location_col_page_count().value_or(1U);
    const std::uint32_t width = page_cols == 0U ? 2U : page_cols * 3U - 1U;
    if (table.anchor_col() + width - 1U > r.last_col) {
      r.last_col = table.anchor_col() + width - 1U;
    }
  }
  return r;
}

}  // namespace

MergeRange detail::pivot_occupied_extent(const pivot::PivotTable& table) {
  return table.projected_extent().has_value() ? to_merge(*table.projected_extent()) : location_extent(table);
}

namespace {

bool is_value_role(PivotCellRole role) {
  return role == PivotCellRole::DataValue || role == PivotCellRole::SubtotalValue ||
         role == PivotCellRole::GrandTotalValue;
}

bool is_label_role(PivotCellRole role) {
  return role == PivotCellRole::RowItem || role == PivotCellRole::DataRowLabel ||
         role == PivotCellRole::SubtotalLabel || role == PivotCellRole::GrandTotalLabel ||
         role == PivotCellRole::BlankRow;
}

// A `numFmtId` held as decimal text; nothing for any other text.
std::optional<std::uint16_t> parse_num_fmt_id(const std::string& text) {
  if (text.empty() || text.size() > 5U) {
    return std::nullopt;
  }
  std::uint32_t id = 0;
  for (const char c : text) {
    if (c < '0' || c > '9') {
      return std::nullopt;
    }
    id = id * 10U + static_cast<std::uint32_t>(c - '0');
  }
  return id > 0xFFFFU ? std::nullopt : std::optional<std::uint16_t>(static_cast<std::uint16_t>(id));
}

// The number format a numeric item of the field named `field_name` is shown
// in: the pivot field's, else its cache field's. Nothing for General.
std::optional<std::uint16_t> item_num_fmt_id(const pivot::PivotTable& table, const pivot::PivotCache& cache,
                                             const std::string& field_name) {
  for (std::size_t i = 0; i < table.fields().size(); ++i) {
    if (pivot::pivot_field_display_name(table.fields()[i]) != field_name) {
      continue;
    }
    std::optional<std::uint16_t> id = parse_num_fmt_id(table.fields()[i].number_format);
    if (!id.has_value() && table.fields()[i].number_format.empty() && i < cache.fields().size() &&
        cache.fields()[i].num_fmt_id.has_value() && *cache.fields()[i].num_fmt_id <= 0xFFFFU) {
      id = static_cast<std::uint16_t>(*cache.fields()[i].num_fmt_id);
    }
    return id.has_value() && *id != 0U ? id : std::nullopt;
  }
  return std::nullopt;
}

struct CellFormat {
  bool button = false;
  std::optional<std::uint8_t> horizontal;
  std::uint32_t indent = 0;
  std::optional<std::uint16_t> num_fmt;

  bool empty() const { return !button && !horizontal.has_value() && !num_fmt.has_value(); }
};

CellXf to_xf(const CellFormat& format) {
  CellXf xf;
  if (format.button) {
    xf.extra_attrs = " pivotButton=\"1\"";
  }
  if (format.horizontal.has_value()) {
    xf.apply_alignment = true;
    xf.has_alignment = true;
    xf.has_horizontal_align = true;
    xf.horizontal_align = *format.horizontal;
    if (format.indent > 0U) {
      xf.has_indent = true;
      xf.indent = format.indent;
    }
  }
  if (format.num_fmt.has_value()) {
    xf.apply_number_format = true;
    xf.num_fmt_id = *format.num_fmt;
  }
  return xf;
}

// One sheet cell of the report: what to store and how to format it.
struct ProjectedCell {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  Value value = Value::blank();
  CellFormat format;
};

// The report cells of `cells` as Excel writes them for `table`.
std::vector<ProjectedCell> project_cells(const pivot::PivotTable& table, const pivot::PivotCache& cache,
                                         const pivot::PivotCells& cells) {
  // An empty value cell with no data field of its own takes its column's format.
  std::map<std::uint32_t, std::string> column_format;
  std::optional<std::uint32_t> first_value_col;
  std::optional<std::uint32_t> first_item_row;
  std::optional<std::uint32_t> body_top;
  std::optional<std::uint16_t> outer_label_format;
  for (const pivot::PivotCell& cell : cells.cells) {
    if (is_value_role(cell.role)) {
      if (!cell.field_name.empty()) {
        column_format.emplace(cell.col, cell.number_format);
      }
      first_value_col = first_value_col.has_value() && *first_value_col < cell.col ? *first_value_col : cell.col;
    }
    if (cell.role == PivotCellRole::RowItem || cell.role == PivotCellRole::DataRowLabel) {
      first_item_row = first_item_row.has_value() && *first_item_row < cell.row ? *first_item_row : cell.row;
    }
    if (cell.role != PivotCellRole::None && cell.role != PivotCellRole::PageCaption &&
        cell.role != PivotCellRole::PageValue) {
      body_top = body_top.has_value() && *body_top < cell.row ? *body_top : cell.row;
    }
  }

  std::vector<ProjectedCell> out;
  out.reserve(cells.cells.size());
  std::vector<std::size_t> compact_grand_labels;
  for (const pivot::PivotCell& cell : cells.cells) {
    ProjectedCell p;
    p.row = cell.row;
    p.col = cell.col;
    p.value = cell.value;
    CellFormat& f = p.format;
    switch (cell.role) {
      case PivotCellRole::PageCaption:
      case PivotCellRole::FieldHeader:
        f.button = true;
        break;
      case PivotCellRole::DataValue:
      case PivotCellRole::SubtotalValue:
      case PivotCellRole::GrandTotalValue: {
        const std::string* code = &cell.number_format;
        if (cell.field_name.empty()) {
          const auto it = column_format.find(cell.col);
          code = it != column_format.end() ? &it->second : code;
        }
        f.num_fmt = parse_num_fmt_id(*code).value_or(0U);
        break;
      }
      default:
        break;
    }
    if (is_label_role(cell.role) && cell.label_indent.has_value()) {
      f.horizontal = kAlignLeft;
      f.indent = *cell.label_indent;
    }
    // A number or date item is stored as its value, in its field's format.
    if ((cell.role == PivotCellRole::RowItem || cell.role == PivotCellRole::ColItem) && cell.item_value.has_value()) {
      p.value = *cell.item_value;
      if (!cell.number_format.empty()) {
        f.num_fmt = item_num_fmt_id(table, cache, cell.field_name);
        if (f.num_fmt.has_value()) {
          f.horizontal = kAlignLeft;
          if (cell.role == PivotCellRole::RowItem && cell.depth == 0U && cell.label_indent.has_value()) {
            outer_label_format = f.num_fmt;
          }
        }
      }
    }
    if (table.merge_item() && body_top.has_value() && cell.row >= *body_top && !is_value_role(cell.role)) {
      // Under `mergeItem` every label of the report is centered, and so is
      // each blank of the row-label area, which a merge may span.
      const bool label_area_blank = cell.role == PivotCellRole::None && first_item_row.has_value() &&
                                    cell.row >= *first_item_row &&
                                    (!first_value_col.has_value() || cell.col < *first_value_col);
      if (cell.role != PivotCellRole::None || label_area_blank) {
        f.horizontal = kAlignCenter;
        f.indent = 0;
      }
    }
    if (cell.role == PivotCellRole::None && f.empty()) {
      continue;
    }
    if (cell.role == PivotCellRole::GrandTotalLabel && cell.label_indent.has_value()) {
      compact_grand_labels.push_back(out.size());
    }
    out.push_back(std::move(p));
  }
  // The compact grand-total label shares the outermost row field's item format.
  for (const std::size_t i : compact_grand_labels) {
    if (outer_label_format.has_value() && !out[i].format.num_fmt.has_value()) {
      out[i].format.num_fmt = outer_label_format;
    }
  }
  return out;
}

// Refuses a (sheet, pivot) pair that names no pivot table.
Expected<void, Error> check_pivot_slot(const Workbook& wb, std::size_t sheet_index, std::size_t pivot_index) {
  if (sheet_index >= wb.sheet_count() || pivot_index >= wb.sheet(sheet_index).pivot_tables().size() ||
      wb.sheet(sheet_index).pivot_tables()[pivot_index] == nullptr) {
    return make_error(FormulonErrorCode::kInvalidArgument, "index out of range",
                      "sheet=" + std::to_string(sheet_index) + " pivot=" + std::to_string(pivot_index));
  }
  return Expected<void, Error>::Ok();
}

// `Sheet!A1` for a message.
std::string cell_name(const Sheet& sheet, std::uint32_t row, std::uint32_t col) {
  std::string out;
  parser::append_sheet_name(sheet.name(), /*force_quote=*/false, out);
  out.push_back('!');
  a1::append_a1(out, row, col);
  return out;
}

// The creation refusal for a report range that `occupant` covers, or that overflows the sheet without one.
Error placement_refusal(const Sheet& sheet, const std::optional<detail::ReportOccupant>& occupant) {
  using Kind = detail::ReportOccupant::Kind;
  std::string message = "The report would extend past the edge of the sheet";
  if (occupant.has_value()) {
    switch (occupant->kind) {
      case Kind::kCell:
      case Kind::kSpill:
        message = "The report would overwrite non-empty cell " + cell_name(sheet, occupant->row, occupant->col);
        break;
      case Kind::kMerge:
        message = "The report would overlap merged cells";
        break;
      case Kind::kArray:
        message = "The report would overlap an array formula";
        break;
      case Kind::kDataTable:
        message = "The report would overlap a data table";
        break;
      case Kind::kTable:
        message = "The report would overlap a table";
        break;
      case Kind::kPivot:
        message = "The report would overlap another PivotTable";
        break;
    }
  }
  return make_error(FormulonErrorCode::kPreconditionFailed, std::move(message), "sheet=" + std::string(sheet.name()));
}

// One thin border side in an explicit RGB colour or a legacy palette index.
BorderSide thin_side(bool outer) {
  constexpr std::uint32_t kOuterArgb = 0xFF999999U;
  constexpr std::uint32_t kInnerIndexed = 65U;
  BorderSide side;
  side.style = 1;
  if (outer) {
    side.color.kind = ColorSpec::Kind::kRgb;
    side.color.rgb = kOuterArgb;
    side.color_argb = kOuterArgb;
  } else {
    side.color.kind = ColorSpec::Kind::kIndexed;
    side.color.indexed = kInnerIndexed;
  }
  return side;
}

// The border of the placeholder frame's cell at (`row`, `col`) of its grid: the outer edges in grey, the inner
// left and top edges in the system window colour.
BorderRecord placeholder_border(std::uint32_t row, std::uint32_t col) {
  BorderRecord border;
  border.left = thin_side(col == 0U);
  border.top = thin_side(row == 0U);
  if (col + 1U == pivot::kPlaceholderCols) {
    border.right = thin_side(true);
  }
  if (row + 1U == pivot::kPlaceholderRows) {
    border.bottom = thin_side(true);
  }
  return border;
}

}  // namespace

Expected<detail::PivotReportPlan, Error> detail::plan_pivot_report(const Workbook& wb, const pivot::PivotTable& table) {
  const pivot::PivotCache* cache = wb.find_pivot_cache(table.pivot_cache_id());
  if (cache == nullptr) {
    return make_error(FormulonErrorCode::kEvalPivotMissing, "pivot cache not found",
                      "cache_id=" + std::to_string(table.pivot_cache_id()));
  }
  PivotReportPlan plan;
  // A cache saved without its records cannot be re-evaluated; the cells read
  // with the file are the report.
  if (!cache->save_data() && cache->records().empty()) {
    return plan;
  }
  const eval::ScopedEvalProfile profile_scope(wb.excel_profile(), wb.text_index_unit());
  const pivot::PivotLayoutOptions options = pivot::pivot_layout_options_for(wb.excel_profile());
  eval::PivotFormulaEvaluator formulas;
  auto result_or = pivot::evaluate(table, *cache, options,
                                   pivot::PivotFilterEnv{wb.pinned_now(), wb.date1904(), &wb.styles(), &formulas});
  if (!result_or) {
    return std::move(result_or.error());
  }
  table.set_last_result(std::move(result_or.value()));
  const std::shared_ptr<const pivot::PivotResult> result = table.last_result();
  auto cells_or = pivot::layout(table, *result, options);
  if (!cells_or) {
    return std::move(cells_or.error());
  }
  plan.evaluated = true;
  plan.cells = std::move(cells_or.value());
  const pivot::PivotCells& cells = plan.cells;
  if (cells.rows > 0U && cells.cols > 0U) {
    const std::uint64_t last_row = static_cast<std::uint64_t>(cells.top) + cells.rows - 1U;
    const std::uint64_t last_col = static_cast<std::uint64_t>(cells.left) + cells.cols - 1U;
    plan.past_edge = last_row >= Sheet::kMaxRows || last_col >= Sheet::kMaxCols;
    plan.extent = MergeRange{cells.top, cells.left,
                             static_cast<std::uint32_t>(last_row < Sheet::kMaxRows ? last_row : Sheet::kMaxRows - 1U),
                             static_cast<std::uint32_t>(last_col < Sheet::kMaxCols ? last_col : Sheet::kMaxCols - 1U)};
    plan.rows_above = table.anchor_row() - cells.top;
  }
  return plan;
}

Expected<detail::PivotReportPlan, Error> detail::plan_pivot_report(Workbook& wb, std::size_t sheet_index,
                                                                   std::size_t pivot_index) {
  RETURN_IF_ERROR(check_pivot_slot(wb, sheet_index, pivot_index));
  return plan_pivot_report(wb, *wb.sheet(sheet_index).pivot_tables()[pivot_index]);
}

Expected<void, Error> Workbook::project_pivot(std::size_t sheet_index, std::size_t pivot_index,
                                              ReportPlacement placement) {
  ASSIGN_OR_RETURN(const detail::PivotReportPlan plan, detail::plan_pivot_report(*this, sheet_index, pivot_index));
  if (!plan.evaluated) {
    return Expected<void, Error>::Ok();
  }
  pivot::PivotTable& table = *sheets_[sheet_index].mutable_pivot_tables()[pivot_index];
  const pivot::PivotCache* cache = find_pivot_cache(table.pivot_cache_id());
  const pivot::PivotCells& cells = plan.cells;
  const MergeRange old_extent = detail::pivot_occupied_extent(table);
  const std::optional<MergeRange>& new_extent = plan.extent;
  std::optional<detail::ReportOccupant> occupant;
  if (new_extent.has_value() && !plan.past_edge) {
    occupant = detail::find_report_occupant(*this, sheet_index, pivot_index, *new_extent, old_extent);
  }
  if (plan.past_edge || occupant.has_value()) {
    if (placement == ReportPlacement::kRefuse) {
      return placement_refusal(sheets_[sheet_index], occupant);
    }
    return withhold_pivot_report(sheet_index, pivot_index, plan);
  }
  std::vector<ProjectedCell> projected = project_cells(table, *cache, cells);

  // The previous report goes first: its merges, then its cells and the new range's.
  std::vector<MergeRange> stale_merges;
  for (const MergeRange& merge : sheets_[sheet_index].merges()) {
    const MergeRange m = normalized(merge);
    if (contains(old_extent, m)) {
      stale_merges.push_back(m);
    }
  }
  for (const MergeRange& merge : stale_merges) {
    RETURN_IF_ERROR(remove_merges_intersecting(sheet_index, merge));
  }
  erase_cells_in(sheet_index, old_extent);
  if (new_extent.has_value()) {
    erase_cells_in(sheet_index, *new_extent);
  }

  std::map<std::string, std::uint32_t> xf_by_key;
  for (const ProjectedCell& cell : projected) {
    if (cell.value.is_text()) {
      put_cell_text(sheet_index, cell.row, cell.col, cell.value.as_text());
    } else if (!cell.value.is_blank()) {
      put_cell_value(sheet_index, cell.row, cell.col, cell.value);
    }
    if (!cell.format.empty()) {
      const CellXf xf = to_xf(cell.format);
      std::string key = cell_xf_key(xf);
      auto it = xf_by_key.find(key);
      if (it == xf_by_key.end()) {
        it = xf_by_key.emplace(std::move(key), intern_cell_xf(styles_, xf)).first;
      }
      RETURN_IF_ERROR(set_cell_xf_index(sheet_index, cell.row, cell.col, it->second));
    }
  }
  if (cells.placeholder) {
    if (styles_.borders.empty()) {
      styles_.borders.emplace_back();
    }
    for (std::uint32_t r = 0; r < pivot::kPlaceholderRows; ++r) {
      for (std::uint32_t c = 0; c < pivot::kPlaceholderCols; ++c) {
        CellXf xf;
        xf.border_index = io::intern_border(styles_, placeholder_border(r, c));
        xf.apply_border = true;
        RETURN_IF_ERROR(set_cell_xf_index(sheet_index, cells.top + r, cells.left + c, intern_cell_xf(styles_, xf)));
      }
    }
  }

  std::vector<PivotRect>& merges = table.mutable_projected_merges();
  merges.clear();
  for (const PivotRect& merge : cells.merges) {
    // The report's own merges sit inside the pivot, which `add_merge` refuses.
    put_merge(sheet_index, normalized(to_merge(merge)));
    merges.push_back(merge);
  }
  table.set_projected_extent(new_extent.has_value() ? std::optional<PivotRect>(to_rect(*new_extent)) : std::nullopt);
  table.set_projection_pending(false);
  if (table.blocked().has_value()) {
    // A recovered report takes its rendered body as its location again.
    table.set_blocked(std::nullopt);
    table.set_anchor(table.anchor_row(), table.anchor_col(), cells.body_rows, cells.body_cols);
    table.clear_span_authored();
    table.set_location_attributes(std::nullopt, std::nullopt, std::nullopt, table.location_row_page_count(),
                                  table.location_col_page_count());
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::withhold_pivot_report(std::size_t sheet_index, std::size_t pivot_index,
                                                      const detail::PivotReportPlan& plan) {
  pivot::PivotTable& table = *sheets_[sheet_index].mutable_pivot_tables()[pivot_index];
  const std::uint32_t anchor_row = table.anchor_row();
  const std::uint32_t anchor_col = table.anchor_col();
  const MergeRange own = detail::pivot_occupied_extent(table);
  // Decided before the report's own cells go, against the occupants the block decision saw.
  const SpillCause cause = plan.past_edge
                               ? SpillCause::kBeyondEdge
                               : detail::report_blocker_cause(*this, sheet_index, pivot_index, *plan.extent, own);
  const std::optional<StyleSlot> anchor_style = sheets_[sheet_index].cell_style_at(anchor_row, anchor_col);
  const std::uint32_t anchor_xf = anchor_style.has_value() ? anchor_style->xf : 0U;
  std::vector<MergeRange> own_merges;
  for (const MergeRange& merge : sheets_[sheet_index].merges()) {
    const MergeRange m = normalized(merge);
    if (contains(own, m)) {
      own_merges.push_back(m);
    }
  }
  for (const MergeRange& merge : own_merges) {
    RETURN_IF_ERROR(remove_merges_intersecting(sheet_index, merge));
  }
  // Nothing below fails: the old report's cells go, the anchor keeps its format and shows the error.
  erase_cells_in(sheet_index, own);
  put_cell_value(sheet_index, anchor_row, anchor_col, Value::error(ErrorCode::Spill));
  if (anchor_xf != 0U) {
    const Expected<void, Error> styled = set_cell_xf_index(sheet_index, anchor_row, anchor_col, anchor_xf);
    assert(styled);
    (void)styled;
  }
  const MergeRange& extent = *plan.extent;
  pivot::BlockedReport state;
  state.rows = extent.last_row - extent.first_row + 1U;
  state.cols = extent.last_col - extent.first_col + 1U;
  state.rows_above = plan.rows_above;
  state.past_edge = plan.past_edge;
  state.size_known = true;
  state.cause = cause;
  table.set_projected_extent(PivotRect{anchor_row, anchor_col, anchor_row, anchor_col});
  table.mutable_projected_merges().clear();
  table.set_blocked(state);
  table.set_auto_formatting(false);
  table.set_projection_pending(false);
  table.set_anchor(anchor_row, anchor_col, 1U, 1U);
  table.clear_span_authored();
  table.set_location_attributes(0U, 0U, 0U, table.location_row_page_count(), table.location_col_page_count());
  return Expected<void, Error>::Ok();
}

void Workbook::project_pending_pivots() {
  for (std::size_t s = 0; s < sheets_.size(); ++s) {
    for (std::size_t p = 0; p < sheets_[s].pivot_tables().size(); ++p) {
      const std::unique_ptr<pivot::PivotTable>& table = sheets_[s].pivot_tables()[p];
      if (table != nullptr && table->projection_pending()) {
        // A pivot that cannot be projected keeps its flag; the save goes on.
        (void)project_pivot(s, p, ReportPlacement::kBlock);
      }
    }
  }
}

Expected<void, Error> Workbook::clear_pivot_projection(std::size_t sheet_index, std::size_t pivot_index) {
  RETURN_IF_ERROR(check_pivot_slot(*this, sheet_index, pivot_index));
  pivot::PivotTable& table = *sheets_[sheet_index].mutable_pivot_tables()[pivot_index];
  if (!table.projected_extent().has_value()) {
    return Expected<void, Error>::Ok();
  }
  for (const PivotRect& merge : table.projected_merges()) {
    RETURN_IF_ERROR(remove_merges_intersecting(sheet_index, to_merge(merge)));
  }
  erase_cells_in(sheet_index, to_merge(*table.projected_extent()));
  table.mutable_projected_merges().clear();
  table.set_projected_extent(std::nullopt);
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
