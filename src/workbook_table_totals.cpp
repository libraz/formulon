// Totals row of a worksheet table: showing and hiding it, its default cells, and the checks guarding the toggle.

#include <algorithm>
#include <cassert>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "cell.h"
#include "data_table.h"
#include "eval/recalc_engine.h"
#include "merge_range.h"
#include "new_workbook_defaults.h"
#include "parser/ast_shift.h"
#include "parser/ref_transforms.h"
#include "parser/reference.h"
#include "parser/structured_ref_payload.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "sheet_metadata.h"
#include "sheet_name.h"
#include "style_resolve.h"
#include "styles.h"
#include "table.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "value.h"
#include "workbook.h"
#include "workbook_internal.h"
#include "workbook_ref_rewrite.h"

namespace formulon {

namespace {

int subtotal_code_of(std::string_view token) {
  for (const detail::TotalsFunction& entry : detail::kTotalsFunctions) {
    if (entry.token == token) {
      return entry.subtotal_code;
    }
  }
  return 0;
}

}  // namespace

namespace detail {

Expected<void, Error> check_totals_row_free(const Workbook& wb, std::size_t index, const TableShape& assumed,
                                            const std::vector<PendingShift>& pending) {
  const std::size_t sheet_index = wb.tables()[index].sheet_index;
  const MergeRange& rect = assumed.rect;
  const MergeRange below{rect.last_row + 1U, rect.first_col, rect.last_row + 1U, rect.last_col};
  const auto refuse = [&](const char* what) {
    std::string ref;
    append_a1_rectangle(ref, below);
    return make_error(FormulonErrorCode::kPreconditionFailed, what, "row=" + ref);
  };
  const Sheet& sheet = wb.sheet(sheet_index);
  for (std::uint32_t col = below.first_col; col <= below.last_col; ++col) {
    // A cell a pending shift inserts is empty.
    const std::optional<CellAddress> at = detail::cell_before_shifts(sheet_index, below.first_row, col, pending);
    if (!at.has_value()) {
      continue;
    }
    const Cell* cell = sheet.cell_at(at->row, at->col);
    if (cell != nullptr && (!cell->formula_text.empty() || !cell->cached_value.is_blank())) {
      return refuse("patch.totals_row needs the row below the table to be empty");
    }
    if (sheet.committed_spill_footprint_covering(at->row, at->col).has_value()) {
      return refuse("patch.totals_row needs the row below the table free of a spilled result");
    }
  }
  MergeRange grown = rect;
  grown.last_row = below.last_row;
  return check_table_range_free(wb, sheet_index, grown, index, /*unmerge_inside=*/false, pending);
}

bool band_meets_straddler(const Workbook& wb, std::size_t sheet_index, const StructuralEdit& edit,
                          const std::vector<PendingShift>& pending) {
  const Sheet& sheet = wb.sheet(sheet_index);
  const auto straddles = [&](const MergeRange& rect, bool straddle_drops) {
    const std::optional<MergeRange> moved = rect_after_shifts(rect, sheet_index, pending, straddle_drops);
    return moved.has_value() && band_fit(*moved, edit) == BandFit::kStraddles;
  };
  for (const MergeRange& merge : sheet.merges()) {
    if (straddles(merge, /*straddle_drops=*/true)) {
      return true;
    }
  }
  for (const DataTable& table : sheet.data_tables()) {
    if (straddles(data_table_block(table), false)) {
      return true;
    }
  }
  for (const std::unique_ptr<pivot::PivotTable>& pivot : sheet.pivot_tables()) {
    if (pivot != nullptr && straddles(pivot_location(*pivot), false)) {
      return true;
    }
  }
  for (const TableMetadata& table : wb.tables()) {
    const std::optional<MergeRange> ref = table_rect_on_sheet(table, sheet_index);
    if (ref.has_value() && straddles(*ref, false)) {
      return true;
    }
  }
  for (const CellAddress address : sheet.all_formula_cells()) {
    const std::optional<MergeRange> block = array_block_rect(sheet, address);
    if (block.has_value() && straddles(*block, false)) {
      return true;
    }
  }
  return false;
}

}  // namespace detail

namespace {

// Records whether the formulas it walks reach a table's totals row: a
// reference meeting the row's cells, or the table's `[#Totals]` / `[#All]`.
class TotalsRowProbe final : public parser::RefTransform {
 public:
  TotalsRowProbe(const MergeRange& row, std::string_view sheet, const sheet_names::NameList* sheet_list,
                 const std::vector<TableMetadata>& tables, std::size_t table) noexcept
      : row_(row), sheet_(sheet), sheet_list_(sheet_list), tables_(tables), table_(table) {}

  /// Whether an unqualified reference names the table's sheet.
  void set_local_means_sheet(bool local) noexcept { local_means_sheet_ = local; }
  bool hit() const noexcept { return hit_; }

  std::optional<parser::Reference> apply(const parser::Reference& ref) const override {
    note(ref, ref);
    return ref;
  }
  std::optional<std::pair<parser::Reference, parser::Reference>> apply_range(
      const parser::Reference& lhs, const parser::Reference& rhs) const override {
    note(lhs, rhs);
    return std::make_pair(lhs, rhs);
  }
  std::optional<parser::StructuredRefTarget> apply_structured_ref(std::string_view table,
                                                                  std::string_view column) const override {
    if (sheet_names::resolve_table(tables_, table) == table_) {
      const std::optional<parser::StructuredRefPayload> payload = parser::parse_structured_ref_payload_text(column);
      constexpr std::uint8_t kReachesTotals =
          parser::StructuredRefSpecifiers::kTotals | parser::StructuredRefSpecifiers::kAll;
      hit_ = hit_ || (payload.has_value() && (payload->specifiers & kReachesTotals) != 0U);
    }
    return parser::StructuredRefTarget{table, {}};
  }
  bool preserves_ref3d_coordinates() const noexcept override { return true; }

 private:
  void note(const parser::Reference& lhs, const parser::Reference& rhs) const {
    const bool on_sheet =
        lhs.sheet.empty() ? local_means_sheet_ : sheet_names::NameList::denotes(sheet_list_, lhs.sheet, sheet_);
    if (!on_sheet) {
      return;
    }
    const bool full_col = lhs.is_full_col || rhs.is_full_col;
    const bool full_row = lhs.is_full_row || rhs.is_full_row;
    const std::uint32_t first_row = full_col ? 0U : std::min(lhs.row, rhs.row);
    const std::uint32_t last_row = full_col ? Sheet::kMaxRows - 1U : std::max(lhs.row, rhs.row);
    const std::uint32_t first_col = full_row ? 0U : std::min(lhs.col, rhs.col);
    const std::uint32_t last_col = full_row ? Sheet::kMaxCols - 1U : std::max(lhs.col, rhs.col);
    hit_ = hit_ || (first_row <= row_.last_row && row_.first_row <= last_row && first_col <= row_.last_col &&
                    row_.first_col <= last_col);
  }

  MergeRange row_;
  std::string_view sheet_;
  const sheet_names::NameList* sheet_list_;
  const std::vector<TableMetadata>& tables_;
  std::size_t table_;
  bool local_means_sheet_ = false;
  mutable bool hit_ = false;
};

// Whether a cell formula or a defined name references `row`, table `index`'s
// totals row. CF and DV formulas are not "formulas or names" here, and an
// INDIRECT / OFFSET target is not followed.
bool totals_row_referenced(const std::vector<Sheet>& sheets, const std::vector<TableMetadata>& tables,
                           const std::vector<DefinedName>& names, std::size_t index, const MergeRange& row) {
  const std::size_t sheet_index = tables[index].sheet_index;
  const sheet_names::NameList sheet_list = detail::sheet_name_list(sheets);
  TotalsRowProbe probe(row, sheets[sheet_index].name(), &sheet_list, tables, index);
  for (std::size_t s = 0; s < sheets.size(); ++s) {
    probe.set_local_means_sheet(s == sheet_index);
    for (const CellAddress address : sheets[s].all_formula_cells()) {
      // The totals cells themselves go with the row.
      if (s == sheet_index && address.row == row.first_row && address.col >= row.first_col &&
          address.col <= row.last_col) {
        continue;
      }
      const Cell* cell = sheets[s].cell_at(address.row, address.col);
      if (cell != nullptr && !cell->formula_text.empty()) {
        static_cast<void>(rewrite_formula(cell->formula_text, probe));
      }
      if (probe.hit()) {
        return true;
      }
    }
  }
  for (const DefinedName& name : names) {
    probe.set_local_means_sheet(name.local_sheet_id >= 0 &&
                                static_cast<std::size_t>(name.local_sheet_id) == sheet_index);
    static_cast<void>(rewrite_name_body(name.formula, probe));
    if (probe.hit()) {
      return true;
    }
  }
  return false;
}

}  // namespace

Expected<std::optional<PendingShift>, Error> Workbook::check_table_totals_row(
    std::size_t index, bool shown, const TableShape& assumed, const std::vector<PendingShift>& pending) const {
  RETURN_IF_ERROR(detail::check_table_index(index, tables_.size()));
  if (shown == assumed.totals_row) {
    return std::optional<PendingShift>();
  }
  const TableMetadata& table = tables_[index];
  const std::size_t sheet_index = table.sheet_index;
  const MergeRange& rect = assumed.rect;
  const Band columns{rect.first_col, rect.last_col};
  if (shown) {
    if (rect.last_row + 1U >= Sheet::kMaxRows) {
      return make_error(FormulonErrorCode::kPreconditionFailed,
                        "patch.totals_row cannot be shown on a table ending on the last row of the grid",
                        "ref=" + table.ref);
    }
    const MergeRange below{rect.last_row + 1U, rect.first_col, rect.last_row + 1U, rect.last_col};
    // Something straddling the columns below keeps the cells from shifting;
    // the totals row then takes the row below only while it is empty.
    if (detail::band_meets_straddler(*this, sheet_index, StructuralEdit{below.first_row, 1U, false, true, columns},
                                     pending)) {
      RETURN_IF_ERROR(detail::check_totals_row_free(*this, index, assumed, pending));
      return std::optional<PendingShift>();
    }
    RETURN_IF_ERROR(detail::check_band_edit_after(*this, sheet_index, below, CellShift::kDown, pending));
    return std::optional<PendingShift>(PendingShift{sheet_index, below, CellShift::kDown});
  }
  const std::uint32_t data_rows = rect.last_row - rect.first_row - (assumed.header_row ? 1U : 0U);
  if (data_rows == 0U) {
    return make_error(FormulonErrorCode::kPreconditionFailed,
                      "patch.totals_row cannot be hidden when the table would keep no data row", "ref=" + table.ref);
  }
  for (std::uint32_t col = rect.first_col; col <= rect.last_col; ++col) {
    const std::optional<CellAddress> at = detail::cell_before_shifts(sheet_index, rect.last_row, col, pending);
    if (!at.has_value()) {
      continue;
    }
    const std::optional<ArrayFormulaInfo> block = array_formula_at(sheet_index, at->row, at->col);
    if ((block && (block->rows != 1U || block->cols != 1U)) ||
        data_table_at(sheet_index, at->row, at->col) != nullptr) {
      return make_error(FormulonErrorCode::kPreconditionFailed,
                        "patch.totals_row cannot be hidden while the totals row holds part of a multi-cell array "
                        "formula or data table",
                        "ref=" + table.ref);
    }
  }
  // References are read where the totals row sits now; a row a shift inserted has none.
  if (const std::optional<CellAddress> at =
          detail::cell_before_shifts(sheet_index, rect.last_row, rect.first_col, pending);
      at.has_value() && totals_row_referenced(sheets_, tables_, defined_names_, index,
                                              MergeRange{at->row, rect.first_col, at->row, rect.last_col})) {
    return make_error(FormulonErrorCode::kPreconditionFailed, "a formula or name references the totals row",
                      "ref=" + table.ref);
  }
  const MergeRange totals{rect.last_row, rect.first_col, rect.last_row, rect.last_col};
  if (detail::band_meets_straddler(*this, sheet_index, StructuralEdit{rect.last_row, 1U, true, true, columns},
                                   pending)) {
    return std::optional<PendingShift>();
  }
  RETURN_IF_ERROR(detail::check_band_edit_after(*this, sheet_index, totals, CellShift::kUp, pending));
  return std::optional<PendingShift>(PendingShift{sheet_index, totals, CellShift::kUp});
}

Expected<void, Error> Workbook::set_table_totals_row(std::size_t index, bool shown) {
  RETURN_IF_ERROR(detail::check_table_index(index, tables_.size()));
  const std::optional<MergeRange> rect = parse_a1_rectangle(tables_[index].ref);
  if (!rect) {
    return detail::invalid_argument("table ref is not a range", "ref=" + tables_[index].ref);
  }
  if (shown == tables_[index].totals_row) {
    return Expected<void, Error>::Ok();
  }
  const std::size_t sheet_index = tables_[index].sheet_index;
  // The toggle inserts or deletes cells, rewrites references and writes the
  // totals cells as one edit under a single hold of the engine mutex.
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const TableShape current{*rect, tables_[index].header_row, tables_[index].totals_row};
  ASSIGN_OR_RETURN(const std::optional<PendingShift> shift, check_table_totals_row(index, shown, current, {}));
  MergeRange next = *rect;
  if (shown) {
    if (shift.has_value()) {
      const Expected<void, Error> inserted = apply_band_edit_locked(sheet_index, shift->range, shift->shift, false);
      assert(inserted.has_value());
      RETURN_IF_ERROR(inserted);
    }
    if (!tables_[index].totals_row_shown.value_or(true)) {
      apply_totals_defaults(index);
    }
    next.last_row = rect->last_row + 1U;
  } else {
    next.last_row = rect->last_row - 1U;
  }

  const std::string ref_before = tables_[index].ref;
  // The range shrinks and the flag clears before the old totals row goes, so
  // the edit below never reads the last data row as the table's end, nor
  // itself as a structural delete of the totals row.
  tables_[index].ref.clear();
  append_a1_rectangle(tables_[index].ref, next);
  tables_[index].totals_row = shown;
  if (!shown) {
    const MergeRange totals{rect->last_row, rect->first_col, rect->last_row, rect->last_col};
    if (shift.has_value()) {
      const Expected<void, Error> deleted = apply_band_edit_locked(sheet_index, totals, CellShift::kUp, true);
      assert(deleted.has_value());
      if (!deleted) {
        tables_[index].ref = ref_before;
        tables_[index].totals_row = true;
        return deleted;
      }
    } else {
      for (std::uint32_t col = rect->first_col; col <= rect->last_col; ++col) {
        if (sheets_[sheet_index].cell_at(rect->last_row, col) != nullptr) {
          put_cell_value_locked(sheet_index, rect->last_row, col, Value::blank());
        }
      }
    }
  }
  TableMetadata& table = tables_[index];
  if (shown) {
    // Shown now, so the "never shown" marker no longer applies.
    table.totals_row_shown.reset();
  }
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  detail::finish_table_extent_change_locked(*this, sheets_, mutator, defined_names_, index);
  if (!shown) {
    return Expected<void, Error>::Ok();
  }
  for (std::size_t column = 0; column < table.columns.size(); ++column) {
    const TableColumn& col = table.columns[column];
    if (col.totals_function.empty() && col.totals_label.empty()) {
      continue;
    }
    RETURN_IF_ERROR(write_table_totals_cell_locked(index, column));
  }
  return Expected<void, Error>::Ok();
}

void Workbook::apply_totals_defaults(std::size_t index) {
  TableMetadata& table = tables_[index];
  for (TableColumn& column : table.columns) {
    column.totals_label.clear();
    column.totals_function.clear();
    column.totals_row_formula.clear();
    column.totals_row_formula_array = false;
  }
  const std::optional<MergeRange> rect = parse_a1_rectangle(table.ref);
  if (table.columns.empty() || !rect) {
    return;
  }
  // The last column sums when its non-blank data cells are all numbers not shown as dates, and there is one.
  const Sheet& sheet = sheets_[table.sheet_index];
  const std::uint32_t first = rect->first_row + (table.header_row ? 1U : 0U);
  const std::uint32_t end = rect->last_row + 1U - (table.totals_row ? 1U : 0U);  // exclusive
  bool any = false;
  bool numbers = true;
  for (std::uint32_t row = first; row < end && numbers; ++row) {
    const Value value = sheet.resolve_cell_value(row, rect->last_col);
    if (value.is_blank()) {
      continue;
    }
    any = true;
    const std::uint32_t xf = select_effective_xf(sheet, row, rect->last_col).xf_index;
    const CellXf* cell_xf =
        styles_.cell_xfs.empty() ? nullptr : &styles_.cell_xfs[xf < styles_.cell_xfs.size() ? xf : 0U];
    numbers = value.is_number() && !detail::cell_is_date_value(styles_, cell_xf, excel_profile_, value);
  }
  table.columns.back().totals_function = any && numbers ? "sum" : "count";
  if (table.columns.size() >= 2U) {
    table.columns.front().totals_label = new_workbook_defaults(excel_profile_.locale).totals_label;
  }
}

Expected<void, Error> Workbook::write_table_totals_cell(std::size_t index, std::size_t column) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  return write_table_totals_cell_locked(index, column);
}

Expected<void, Error> Workbook::write_table_totals_cell_locked(std::size_t index, std::size_t column) {
  if (index >= tables_.size() || column >= tables_[index].columns.size()) {
    return detail::invalid_argument("index out of range",
                                    "index=" + std::to_string(index) + " column=" + std::to_string(column));
  }
  const TableMetadata& table = tables_[index];
  const std::optional<MergeRange> rect = parse_a1_rectangle(table.ref);
  if (!rect) {
    return detail::invalid_argument("table ref is not a range", "ref=" + table.ref);
  }
  const TableColumn& col = table.columns[column];
  const std::uint32_t row = rect->last_row;
  const std::uint32_t cell_col = rect->first_col + static_cast<std::uint32_t>(column);
  if (const int code = subtotal_code_of(col.totals_function); code != 0) {
    put_cell_formula_locked(table.sheet_index, row, cell_col, detail::totals_subtotal_formula(table, column, code),
                            FormulaSource::kEntered, 0U, 0U);
  } else if (col.totals_function == "custom" && !col.totals_row_formula.empty()) {
    put_cell_formula_locked(table.sheet_index, row, cell_col, "=" + col.totals_row_formula, FormulaSource::kEntered, 0U,
                            0U);
  } else if (!col.totals_label.empty()) {
    put_cell_text_locked(table.sheet_index, row, cell_col, col.totals_label);
  } else {
    put_cell_value_locked(table.sheet_index, row, cell_col, Value::blank());
  }
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
