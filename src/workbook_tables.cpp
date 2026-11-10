#include <algorithm>
#include <cassert>
#include <cstddef>
#include <cstdint>
#include <iterator>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "cell.h"
#include "chart/chart_ref.h"
#include "chart/chart_template.h"
#include "data_table.h"
#include "defined_name.h"
#include "eval/dep_graph.h"
#include "eval/recalc_engine.h"
#include "eval/text_format/display_text.h"
#include "io/auto_filter_xml.h"
#include "io/xlsb/table_records.h"
#include "io/xml_escape.h"
#include "merge_range.h"
#include "new_workbook_defaults.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/ast_shift.h"
#include "parser/ref_transforms.h"
#include "parser/reference.h"
#include "parser/structured_ref_payload.h"
#include "parser/structured_ref_spelling.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "sheet_metadata.h"
#include "sheet_name.h"
#include "style_resolve.h"
#include "styles.h"
#include "table.h"
#include "table_calc_column.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/name_syntax.h"
#include "utils/status_macros.h"
#include "value.h"
#include "workbook.h"
#include "workbook_formula_index.h"
#include "workbook_internal.h"
#include "workbook_ref_rewrite.h"

namespace formulon {

namespace {

Error invalid_table(std::string message, std::string context = {}) {
  Error err = make_error(FormulonErrorCode::kInvalidArgument, std::move(message));
  err.context = std::move(context);
  return err;
}

/// True when writing `name` or `display_name` collides with either of `existing`'s names.
template <typename Existing>
bool names_collide(const Existing& existing, std::string_view name, std::string_view display_name) {
  return sheet_names::write_collides(existing.name, name) || sheet_names::write_collides(existing.display_name, name) ||
         sheet_names::write_collides(existing.name, display_name) ||
         sheet_names::write_collides(existing.display_name, display_name);
}

bool name_in_use(const std::vector<TableMetadata>& tables, const std::vector<io::xlsb::XlsbRetainedTable>& retained,
                 const std::vector<DefinedName>& defined_names, std::string_view name, std::string_view display_name,
                 std::optional<std::size_t> self) {
  for (std::size_t i = 0; i < tables.size(); ++i) {
    if (i != self && names_collide(tables[i], name, display_name)) {
      return true;
    }
  }
  for (const io::xlsb::XlsbRetainedTable& existing : retained) {
    if (names_collide(existing, name, display_name)) {
      return true;
    }
  }
  // A defined name of any scope owns its name against tables.
  for (const DefinedName& defined : defined_names) {
    if (sheet_names::write_collides(defined.name, name) || sheet_names::write_collides(defined.name, display_name)) {
      return true;
    }
  }
  return false;
}

Expected<void, Error> check_table_identity(const Workbook& wb, std::string_view name, std::string_view display_name,
                                           std::optional<std::size_t> self) {
  if (name.empty() || !is_valid_table_name(display_name)) {
    return invalid_table("name is not a valid table name",
                         "name=" + std::string(name) + " display_name=" + std::string(display_name));
  }
  if (name_in_use(wb.tables(), io::xlsb::xlsb_retained_tables(wb), wb.defined_names(), name, display_name, self)) {
    return make_error(FormulonErrorCode::kAlreadyExists, "table name already exists", "name=" + std::string(name));
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

namespace detail {

std::optional<CellAddress> cell_before_shifts(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                              const std::vector<PendingShift>& pending) {
  CellAddress at{row, col};
  for (auto it = pending.rbegin(); it != pending.rend(); ++it) {
    const Expected<StructuralEdit, Error> edit = detail::band_edit_of(it->range, it->shift);
    if (it->sheet_index != sheet_index || !edit.has_value()) {
      continue;
    }
    const StructuralEdit& e = edit.value();
    std::uint32_t& coord = e.row_axis ? at.row : at.col;
    const std::uint32_t across = e.row_axis ? at.col : at.row;
    if (e.band.has_value() && (across < e.band->first || across > e.band->last)) {
      continue;
    }
    if (!e.is_delete) {
      if (coord >= e.index && coord - e.index < e.count) {
        return std::nullopt;
      }
      coord = coord >= e.index ? coord - e.count : coord;
    } else if (coord >= e.index) {
      if (static_cast<std::uint64_t>(coord) + e.count >= a1::axis_extent(e.row_axis)) {
        return std::nullopt;
      }
      coord += e.count;
    }
  }
  return at;
}

Expected<void, Error> check_table_range_free(const Workbook& wb, std::size_t sheet_index, const MergeRange& rect,
                                             std::optional<std::size_t> self, bool unmerge_inside,
                                             const std::vector<PendingShift>& pending) {
  const auto refuse = [&](const char* what) {
    std::string ref;
    append_a1_rectangle(ref, rect);
    return make_error(FormulonErrorCode::kPreconditionFailed, what, "ref=" + ref);
  };
  // Whether a structure now at `at` lies over `rect` once the pending shifts have moved it.
  const auto meets = [&](const MergeRange& at) {
    const std::optional<MergeRange> moved = rect_after_shifts(at, sheet_index, pending, /*straddle_drops=*/false);
    return moved.has_value() && intersects(rect, *moved);
  };
  for (std::size_t i = 0; i < wb.tables().size(); ++i) {
    const std::optional<MergeRange> other_rect = table_rect_on_sheet(wb.tables()[i], sheet_index);
    if (i != self && other_rect && meets(*other_rect)) {
      return refuse("table range overlaps another table");
    }
  }
  for (const io::xlsb::XlsbRetainedTable& kept : io::xlsb::xlsb_retained_tables(wb)) {
    const std::optional<MergeRange> kept_rect = parse_a1_rectangle(kept.ref);
    if (kept.sheet_index == sheet_index && kept_rect && meets(*kept_rect)) {
      return refuse("table range overlaps another table");
    }
  }
  const Sheet& sheet = wb.sheet(sheet_index);
  for (const std::unique_ptr<pivot::PivotTable>& pivot : sheet.pivot_tables()) {
    if (pivot != nullptr && meets(pivot_occupied_extent(*pivot))) {
      return refuse("table range overlaps a PivotTable");
    }
  }
  for (const DataTable& data_table : sheet.data_tables()) {
    if (meets(normalized(data_table.ref))) {
      return refuse("table range overlaps a data table");
    }
  }
  // A shift can bring a block from below the range into it.
  const std::vector<CellAddress> formulas =
      pending.empty() ? sheet.formula_cells_in(0U, 0U, rect.last_row, rect.last_col) : sheet.all_formula_cells();
  for (const CellAddress address : formulas) {
    const Cell* cell = sheet.cell_at(address.row, address.col);
    if (cell == nullptr || cell->array_block_rows == 0U ||
        (cell->array_block_rows == 1U && cell->array_block_cols == 1U)) {
      continue;
    }
    const MergeRange block{address.row, address.col, address.row + cell->array_block_rows - 1U,
                           address.col + cell->array_block_cols - 1U};
    if (meets(block)) {
      return refuse("range must not hold a multi-cell array formula");
    }
  }
  for (const MergeRange& merge : sheet.merges()) {
    // A shift unmerges a merge straddling its band.
    const std::optional<MergeRange> m =
        rect_after_shifts(normalized(merge), sheet_index, pending, /*straddle_drops=*/true);
    if (m && intersects(*m, rect) && !(unmerge_inside && contains(rect, *m))) {
      return refuse(unmerge_inside ? "table range cuts through merged cells" : "table range overlaps merged cells");
    }
  }
  return Expected<void, Error>::Ok();
}

}  // namespace detail

Expected<std::size_t, Error> Workbook::add_table(TableMetadata table) {
  const auto rect = parse_a1_rectangle(table.ref);
  if (table.sheet_index >= sheet_count() || table.ref.empty() || table.name.empty() || table.display_name.empty() ||
      table.columns.empty()) {
    return invalid_table("sheet_index is out of range, or name, range, or column_names is empty");
  }
  const std::size_t width = rect ? static_cast<std::size_t>(rect->last_col - rect->first_col) + 1U : 0U;
  if (width != table.columns.size()) {
    return invalid_table("column_count does not match the range width",
                         "ref=" + table.ref + " column_count=" + std::to_string(table.columns.size()));
  }
  // A header-only range takes a blank data row below it; one row cannot hold both a header and a totals row.
  const bool header_only = table.header_row && rect->first_row == rect->last_row;
  if (header_only && table.totals_row) {
    return invalid_table("header_row and totals_row need a range of at least two rows", "ref=" + table.ref);
  }
  RETURN_IF_ERROR(check_table_identity(*this, table.name, table.display_name, std::nullopt));
  MergeRange extent = *rect;
  std::vector<PendingShift> shifts;
  if (header_only) {
    if (rect->last_row + 1U >= Sheet::kMaxRows) {
      return make_error(FormulonErrorCode::kPreconditionFailed,
                        "range cannot be a header row on the last row of the grid", "ref=" + table.ref);
    }
    const MergeRange placeholder{rect->last_row + 1U, rect->first_col, rect->last_row + 1U, rect->last_col};
    for (std::uint32_t col = placeholder.first_col; col <= placeholder.last_col && shifts.empty(); ++col) {
      if (detail::cell_is_occupied(sheets_[table.sheet_index], placeholder.first_row, col)) {
        RETURN_IF_ERROR(check_band_edit(table.sheet_index, placeholder, CellShift::kDown));
        shifts.push_back(PendingShift{table.sheet_index, placeholder, CellShift::kDown});
      }
    }
    extent.last_row = placeholder.last_row;
  }
  RETURN_IF_ERROR(
      detail::check_table_range_free(*this, table.sheet_index, extent, std::nullopt, /*unmerge_inside=*/true, shifts));
  // An occupied row under a header-only range moves down within the table's columns.
  if (!shifts.empty()) {
    const Expected<void, Error> shifted = insert_cells(table.sheet_index, shifts.front().range, CellShift::kDown);
    assert(shifted.has_value());
    RETURN_IF_ERROR(shifted);
  }
  if (header_only) {
    table.ref.clear();
    append_a1_rectangle(table.ref, extent);
    table.insert_row = true;
    table.insert_row_shift = !shifts.empty();
  }
  table.id = next_free_table_id(0);
  std::vector<std::string> names;
  names.reserve(table.columns.size());
  for (const TableColumn& column : table.columns) {
    names.push_back(column.name);
  }
  detail::name_new_table_columns(names, new_workbook_defaults(excel_profile_.locale).column_word);
  std::vector<std::uint32_t> renamed_columns;
  for (std::size_t i = 0; i < names.size(); ++i) {
    TableColumn& column = table.columns[i];
    if (column.name != names[i]) {
      renamed_columns.push_back(static_cast<std::uint32_t>(i));
      column.name = std::move(names[i]);
    }
    column.id = static_cast<std::uint32_t>(i) + 1U;
  }
  // Excel unmerges the merges a new table covers.
  RETURN_IF_ERROR(remove_merges_intersecting(table.sheet_index, extent));
  std::string escaped_ref;
  io::AppendXmlAttrEscaped(escaped_ref, table.ref);
  table.auto_filter_xml.set(io::auto_filter_from_xml("<autoFilter ref=\"" + escaped_ref + "\"/>"));
  // A column already holding one formula becomes a calculated column; its cells stay as they are.
  for (std::size_t i = 0; i < table.columns.size(); ++i) {
    TableColumn& column = table.columns[i];
    if (!column.calculated_column_formula.empty()) {
      continue;
    }
    if (std::optional<std::string> inferred = infer_calculated_column(sheets_[table.sheet_index], table, i)) {
      column.calculated_column_formula = std::move(*inferred);
      column.calculated_column_formula_array = false;
    }
  }
  // A new table has never shown a totals row; one created with it is shown from the start.
  table.totals_row_shown = false;
  tables_.push_back(std::move(table));
  const std::size_t index = tables_.size() - 1U;
  const TableMetadata& created = tables_.back();
  if (created.header_row) {
    for (const std::uint32_t i : renamed_columns) {
      put_cell_text(created.sheet_index, rect->first_row, rect->first_col + i, created.columns[i].name);
    }
  }
  if (created.totals_row) {
    tables_[index].totals_row_shown.reset();
    bool blank = true;
    for (std::uint32_t col = rect->first_col; col <= rect->last_col && blank; ++col) {
      blank = !detail::cell_is_occupied(sheets_[created.sheet_index], extent.last_row, col);
    }
    bool configured = false;
    for (const TableColumn& col : created.columns) {
      configured =
          configured || !col.totals_function.empty() || !col.totals_label.empty() || !col.totals_row_formula.empty();
    }
    // Content in the totals row wins, each column read from its own cell; columns given settings keep them; only a
    // blank row over bare columns takes the defaults.
    if (!blank) {
      for (std::size_t column = 0; column < created.columns.size(); ++column) {
        sync_table_totals_cell(index, column);
      }
    } else if (!configured) {
      apply_totals_defaults(index);
      for (std::size_t column = 0; column < created.columns.size(); ++column) {
        const TableColumn& col = created.columns[column];
        if (!col.totals_function.empty() || !col.totals_label.empty()) {
          RETURN_IF_ERROR(write_table_totals_cell(index, column));
        }
      }
    }
  }
  // A structured reference can already exist in a formula authored ahead of
  // its table, so both names need the scoped reindex a defined-name edit gets.
  reindex_formulas_for_table_change({created.name, created.display_name});
  return index;
}

namespace {

// `Workbook::check_table_resize` for table `index` whose range is `current`
// (none when unparsable) and which shows a header row when `header_row`.
Expected<void, Error> check_resize(const Workbook& wb, std::size_t index, const MergeRange& rect,
                                   const MergeRange* current, bool header_row,
                                   const std::vector<PendingShift>& pending) {
  const TableMetadata& table = wb.tables()[index];
  if (current != nullptr && *current == rect) {
    return Expected<void, Error>::Ok();
  }
  if (current != nullptr && current->first_row != rect.first_row) {
    return invalid_table("table range must keep its first row", "ref=" + table.ref);
  }
  if (header_row && rect.first_row == rect.last_row) {
    return invalid_table("table range has no data row", "ref=" + table.ref);
  }
  // Excel answers such a resize with a silent no-op; the engine refuses it instead.
  if (table.insert_row_shift) {
    return make_error(FormulonErrorCode::kPreconditionFailed,
                      "table range cannot change while the cells its creation shifted down are marked",
                      "ref=" + table.ref);
  }
  return detail::check_table_range_free(wb, table.sheet_index, rect, index, /*unmerge_inside=*/false, pending);
}

}  // namespace

Expected<void, Error> Workbook::check_table_resize(std::size_t index, const MergeRange& rect) const {
  if (index >= tables_.size()) {
    return invalid_table("index out of range", "index=" + std::to_string(index));
  }
  const std::optional<MergeRange> current = parse_a1_rectangle(tables_[index].ref);
  return check_resize(*this, index, rect, current ? &*current : nullptr, tables_[index].header_row, {});
}

Expected<void, Error> Workbook::check_table_resize(std::size_t index, const MergeRange& rect, const TableShape& assumed,
                                                   const std::vector<PendingShift>& pending) const {
  if (index >= tables_.size()) {
    return invalid_table("index out of range", "index=" + std::to_string(index));
  }
  return check_resize(*this, index, rect, &assumed.rect, assumed.header_row, pending);
}

Expected<void, Error> Workbook::resize_table(std::size_t index, const MergeRange& rect) {
  RETURN_IF_ERROR(check_table_resize(index, rect));
  const std::optional<MergeRange> current = parse_a1_rectangle(tables_[index].ref);
  if (current && *current == rect) {
    return Expected<void, Error>::Ok();
  }
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  tables_[index].ref.clear();
  append_a1_rectangle(tables_[index].ref, rect);
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  detail::finish_table_extent_change_locked(*this, sheets_, mutator, defined_names_, index);
  return Expected<void, Error>::Ok();
}

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
    const Cell* cell = sheet.cell_at(address.row, address.col);
    if (cell != nullptr && cell->array_block_rows != 0U &&
        straddles(MergeRange{address.row, address.col, address.row + cell->array_block_rows - 1U,
                             address.col + cell->array_block_cols - 1U},
                  false)) {
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
    static_cast<void>(rewrite_formula(name.formula, probe));
    if (probe.hit()) {
      return true;
    }
  }
  return false;
}

}  // namespace

Expected<std::optional<PendingShift>, Error> Workbook::check_table_totals_row(
    std::size_t index, bool shown, const TableShape& assumed, const std::vector<PendingShift>& pending) const {
  if (index >= tables_.size()) {
    return invalid_table("index out of range", "index=" + std::to_string(index));
  }
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
  if (index >= tables_.size()) {
    return invalid_table("index out of range", "index=" + std::to_string(index));
  }
  const std::optional<MergeRange> rect = parse_a1_rectangle(tables_[index].ref);
  if (!rect) {
    return invalid_table("table ref is not a range", "ref=" + tables_[index].ref);
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
    return invalid_table("index out of range", "index=" + std::to_string(index) + " column=" + std::to_string(column));
  }
  const TableMetadata& table = tables_[index];
  const std::optional<MergeRange> rect = parse_a1_rectangle(table.ref);
  if (!rect) {
    return invalid_table("table ref is not a range", "ref=" + table.ref);
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

void Workbook::set_tables(std::vector<TableMetadata> tables) {
  std::vector<std::string> names;
  for (const std::vector<TableMetadata>* list : {&tables_, &tables}) {
    for (const TableMetadata& table : *list) {
      names.push_back(table.name);
      names.push_back(table.display_name);
    }
  }
  tables_ = std::move(tables);
  reindex_formulas_for_table_change(names);
}

std::uint32_t Workbook::next_free_table_id(std::uint32_t after) const {
  for (const TableMetadata& existing : tables_) {
    after = after > existing.id ? after : existing.id;
  }
  const std::vector<io::xlsb::XlsbRetainedTable> retained = io::xlsb::xlsb_retained_tables(*this);
  std::uint32_t next_id = after + 1U;
  // A table part kept verbatim still owns its id.
  for (bool taken = true; taken;) {
    taken = false;
    for (const io::xlsb::XlsbRetainedTable& kept : retained) {
      if (kept.id == next_id) {
        ++next_id;
        taken = true;
      }
    }
  }
  return next_id;
}

std::string Workbook::next_default_table_name(std::string_view word) const {
  const std::vector<io::xlsb::XlsbRetainedTable> retained = io::xlsb::xlsb_retained_tables(*this);
  for (std::size_t n = 1;; ++n) {
    const std::string candidate = std::string(word) + std::to_string(n);
    if (!name_in_use(tables_, retained, defined_names_, candidate, candidate, std::nullopt)) {
      return candidate;
    }
  }
}

std::optional<std::size_t> Workbook::table_containing(std::size_t sheet_index, std::uint32_t row,
                                                      std::uint32_t col) const noexcept {
  for (std::size_t i = 0; i < tables_.size(); ++i) {
    if (tables_[i].sheet_index != sheet_index) {
      continue;
    }
    const std::optional<MergeRange> rect = parse_a1_rectangle(tables_[i].ref);
    if (rect && row >= rect->first_row && row <= rect->last_row && col >= rect->first_col && col <= rect->last_col) {
      return i;
    }
  }
  return std::nullopt;
}

bool Workbook::is_table_name(std::string_view name) const noexcept {
  return sheet_names::resolve_table(tables_, name) != sheet_names::kNotFound;
}

void Workbook::reindex_formulas_for_table_change(const std::vector<std::string>& table_names) {
  if (table_names.empty()) {
    return;
  }
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  reindex_formulas_referencing_table(sheets_, mutator, *this, table_names);
  sync_table_spill_barriers_locked(sheets_, mutator, tables_);
}

namespace {

/// Maps a structured reference's table name that resolves to `tables[index]`
/// (still under its old names) to the table's new display name.
class TableRenameTransform final : public parser::RefTransform {
 public:
  TableRenameTransform(const std::vector<TableMetadata>& tables, std::size_t index, std::string_view new_name)
      : tables_(tables), index_(index), new_name_(new_name) {}

  std::optional<parser::Reference> apply(const parser::Reference& ref) const override { return ref; }

  std::optional<parser::StructuredRefTarget> apply_structured_ref(std::string_view table,
                                                                  std::string_view /*column*/) const override {
    if (sheet_names::resolve_table(tables_, table) == index_) {
      return parser::StructuredRefTarget{new_name_, {}};
    }
    return parser::StructuredRefTarget{table, {}};
  }

 private:
  const std::vector<TableMetadata>& tables_;
  std::size_t index_;
  std::string_view new_name_;
};

}  // namespace

Expected<void, Error> Workbook::rename_table(std::size_t index, std::string name, std::string display_name) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (index >= tables_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "index out of range",
                      "index=" + std::to_string(index) + " table_count=" + std::to_string(tables_.size()));
  }
  RETURN_IF_ERROR(check_table_identity(*this, name, display_name, index));
  const std::string old_name = tables_[index].name;
  const std::string old_display_name = tables_[index].display_name;
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  const TableRenameTransform transform(tables_, index, display_name);
  std::vector<std::uint32_t> ignored_cache_ids;
  bool defined_names_changed = false;
  rewrite_workbook_references(sheets_, defined_names_, tables_, pivot_caches_, transform, mutator, {}, {}, {},
                              ignored_cache_ids, defined_names_changed);
  RETURN_IF_ERROR(chart_ref::rewrite_all(*this, transform));
  for (std::unique_ptr<pivot::PivotCache>& cache : pivot_caches_) {
    if (cache == nullptr) {
      continue;
    }
    pivot::WorksheetSource& source = cache->mutable_worksheet_source();
    if (!source.name.empty() && sheet_names::resolve_table(tables_, source.name) == index) {
      source.name = display_name;
    }
  }
  tables_[index].name = std::move(name);
  tables_[index].display_name = std::move(display_name);
  if (defined_names_changed) {
    mark_all_formulas_dirty_locked(sheets_, mutator);
  }
  reindex_formulas_referencing_table(sheets_, mutator, *this,
                                     {old_name, old_display_name, tables_[index].name, tables_[index].display_name});
  return Expected<void, Error>::Ok();
}

namespace {

/// Maps a structured reference through `tables[table]` naming its column
/// `column` (still under its old name) to the same reference naming
/// `new_column`, re-spelled in the stored form.
class TableColumnRenameTransform final : public parser::RefTransform {
 public:
  TableColumnRenameTransform(const std::vector<TableMetadata>& tables, std::size_t table, std::size_t column,
                             std::string_view new_column)
      : tables_(tables), table_(table), column_(column), new_column_(new_column) {}

  std::optional<parser::Reference> apply(const parser::Reference& ref) const override { return ref; }

  std::optional<parser::StructuredRefTarget> apply_structured_ref(std::string_view table,
                                                                  std::string_view column) const override {
    if (sheet_names::resolve_table(tables_, table) != table_) {
      return parser::StructuredRefTarget{table, {}};
    }
    std::optional<parser::StructuredRefPayload> payload = parser::parse_structured_ref_payload_text(column);
    if (!payload) {
      return parser::StructuredRefTarget{table, {}};
    }
    bool renamed = false;
    for (std::string& name : payload->columns) {
      if (sheet_names::resolve_named(tables_[table_].columns, name) == column_) {
        name.assign(new_column_);
        renamed = true;
      }
    }
    if (!renamed) {
      return parser::StructuredRefTarget{table, {}};
    }
    return parser::StructuredRefTarget{table, parser::stored_payload(*payload)};
  }

 private:
  const std::vector<TableMetadata>& tables_;
  std::size_t table_;
  std::size_t column_;
  std::string_view new_column_;
};

}  // namespace

Expected<std::string, Error> Workbook::rename_table_column(std::size_t table, std::size_t column,
                                                           std::string_view name) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (table >= tables_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "table index out of range",
                      "table=" + std::to_string(table) + " table_count=" + std::to_string(tables_.size()));
  }
  std::vector<TableColumn>& columns = tables_[table].columns;
  if (column >= columns.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "column index out of range",
                      "column=" + std::to_string(column) + " column_count=" + std::to_string(columns.size()));
  }
  std::vector<std::string> others;
  others.reserve(columns.size());
  for (std::size_t i = 0; i < columns.size(); ++i) {
    if (i != column) {
      others.push_back(columns[i].name);
    }
  }
  // An empty header takes `<word>N` from N = 1; a duplicate takes `<name>N` from N = 2.
  std::string final_name(name);
  if (name.empty()) {
    final_name = detail::first_free_numbered_name(new_workbook_defaults(excel_profile_.locale).column_word, 1U, others);
  } else if (sheet_names::write_collides(others, name, sheet_names::kNotFound)) {
    final_name = detail::first_free_numbered_name(name, 2U, others);
  }

  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  const std::string old_name = columns[column].name;
  const bool renamed = final_name != old_name;
  if (renamed) {
    const TableColumnRenameTransform transform(tables_, table, column, final_name);
    std::vector<std::uint32_t> ignored_cache_ids;
    bool defined_names_changed = false;
    rewrite_workbook_references(sheets_, defined_names_, tables_, pivot_caches_, transform, mutator, {}, {}, {},
                                ignored_cache_ids, defined_names_changed);
    RETURN_IF_ERROR(chart_ref::rewrite_all(*this, transform));
    tables_[table].columns[column].name = final_name;
    if (defined_names_changed) {
      mark_all_formulas_dirty_locked(sheets_, mutator);
    }
  }
  const std::optional<MergeRange> rect = parse_a1_rectangle(tables_[table].ref);
  if (tables_[table].header_row && rect.has_value() && tables_[table].sheet_index < sheets_.size()) {
    const std::uint32_t header_col = rect->first_col + static_cast<std::uint32_t>(column);
    sheets_[tables_[table].sheet_index].set_cell_text(rect->first_row, header_col, final_name);
    const eval::CellNodeId node = make_node(tables_[table].sheet_index, rect->first_row, header_col);
    detail::mark_dependents_dirty(mutator, node);
  }
  if (renamed) {
    reindex_formulas_referencing_table(sheets_, mutator, *this, {tables_[table].name, tables_[table].display_name});
  }
  return final_name;
}

namespace detail {

std::string totals_subtotal_formula(const TableMetadata& table, std::size_t column, int code) {
  parser::StructuredRefPayload payload;
  payload.columns.push_back(table.columns[column].name);
  return "=SUBTOTAL(" + std::to_string(code) + "," + table.display_name + "[" + parser::stored_payload(payload) + "])";
}

std::string first_free_numbered_name(std::string_view stem, std::size_t first, const std::vector<std::string>& names) {
  for (std::size_t n = first;; ++n) {
    std::string candidate(stem);
    candidate += std::to_string(n);
    if (!sheet_names::write_collides(names, candidate, sheet_names::kNotFound)) {
      return candidate;
    }
  }
}

void name_new_table_columns(std::vector<std::string>& names, std::string_view column_word) {
  std::vector<std::string> reserved = names;
  for (std::size_t i = 0; i < names.size(); ++i) {
    const bool blank = names[i].empty();
    bool repeats = false;
    for (std::size_t j = 0; j < i && !blank && !repeats; ++j) {
      repeats = sheet_names::write_collides(names[j], names[i]);
    }
    if (blank || repeats) {
      names[i] = blank ? first_free_numbered_name(column_word, 1U, reserved)
                       : first_free_numbered_name(names[i], 2U, reserved);
      reserved.push_back(names[i]);
    }
  }
}

std::string next_table_column_name(std::string_view left_name, std::string_view column_word,
                                   const std::vector<std::string>& taken) {
  std::size_t stem = left_name.size();
  while (stem > 0U && left_name[stem - 1U] >= '0' && left_name[stem - 1U] <= '9') {
    --stem;
  }
  if (stem == left_name.size()) {
    return first_free_numbered_name(column_word, 1U, taken);
  }
  // Counts the trailing digits on in place, so leading zeros keep the width.
  std::string next(left_name);
  std::size_t at = next.size();
  while (at > stem && next[at - 1U] == '9') {
    next[--at] = '0';
  }
  if (at > stem) {
    ++next[at - 1U];
  } else {
    next.insert(stem, 1U, '1');
  }
  if (!sheet_names::write_collides(taken, next, sheet_names::kNotFound)) {
    return next;
  }
  return first_free_numbered_name(left_name, 2U, taken);
}

void add_table_columns_locked(TableMetadata& table, std::size_t offset, std::uint32_t count, const MergeRange& rect,
                              Sheet& sheet, std::size_t sheet_index, const eval::RecalcEngine::LockedMutator& mutator,
                              std::string_view column_word) {
  offset = std::min(offset, table.columns.size());
  std::uint32_t next_id = 1;
  std::vector<std::string> taken;
  taken.reserve(table.columns.size() + count);
  for (const TableColumn& column : table.columns) {
    next_id = std::max(next_id, column.id + 1U);
    taken.push_back(column.name);
  }
  std::string left = offset == 0U ? std::string() : table.columns[offset - 1U].name;
  std::vector<TableColumn> added;
  added.reserve(count);
  for (std::uint32_t i = 0; i < count; ++i) {
    TableColumn column;
    column.id = next_id++;
    column.name = next_table_column_name(left, column_word, taken);
    taken.push_back(column.name);
    left = column.name;
    if (table.header_row) {
      const std::uint32_t col = rect.first_col + static_cast<std::uint32_t>(offset) + i;
      sheet.set_cell_text(rect.first_row, col, column.name);
      mark_dependents_dirty(mutator, make_node(sheet_index, rect.first_row, col));
    }
    added.push_back(std::move(column));
  }
  table.columns.insert(table.columns.begin() + static_cast<std::ptrdiff_t>(offset),
                       std::make_move_iterator(added.begin()), std::make_move_iterator(added.end()));
}

void fill_calculated_columns_locked(const Workbook& wb, Sheet& sheet, const eval::RecalcEngine::LockedMutator& mutator,
                                    const TableMetadata& table, std::uint32_t first_row, std::uint32_t last_row,
                                    std::size_t first_column, std::size_t last_column) {
  const std::optional<MergeRange> rect = parse_a1_rectangle(table.ref);
  if (!rect.has_value() || first_row > last_row) {
    return;
  }
  // The stored formula reads as entered in the first data row.
  const std::uint32_t anchor_row = rect->first_row + (table.header_row ? 1U : 0U);
  Arena formula_arena;
  Arena shift_arena;
  for (std::size_t i = first_column; i <= last_column && i < table.columns.size(); ++i) {
    const TableColumn& column = table.columns[i];
    if (column.calculated_column_formula.empty()) {
      continue;
    }
    const std::string text =
        Workbook::normalize_formula_text("=" + column.calculated_column_formula, Workbook::FormulaSource::kRendered);
    const parser::AstNode* root = parse_stored_formula(text, formula_arena);
    if (root == nullptr) {
      continue;
    }
    const std::uint32_t col = rect->first_col + static_cast<std::uint32_t>(i);
    const std::uint32_t block = column.calculated_column_formula_array ? 1U : 0U;
    for (std::uint32_t row = first_row; row <= last_row; ++row) {
      shift_arena.reset();
      const parser::AstNode* shifted = parser::shift_relative_refs(
          *root, shift_arena, static_cast<std::int32_t>(row) - static_cast<std::int32_t>(anchor_row), 0);
      if (shifted == nullptr) {
        continue;
      }
      sheet.set_cell_formula(row, col, "=" + parser::format_formula(*shifted));
      sheet.set_cell_array_block(row, col, block, block);
      const eval::CellNodeId node = make_node(table.sheet_index, row, col);
      mutator.register_formula(node, *shifted, wb);
      mutator.mark_dirty(node);
      mark_dependents_dirty(mutator, node);
    }
  }
}

void shift_calculated_column_anchor(TableMetadata& table, std::int32_t row_delta) {
  Arena formula_arena;
  Arena shift_arena;
  for (TableColumn& column : table.columns) {
    if (column.calculated_column_formula.empty()) {
      continue;
    }
    const std::string text =
        Workbook::normalize_formula_text("=" + column.calculated_column_formula, Workbook::FormulaSource::kRendered);
    const parser::AstNode* root = parse_stored_formula(text, formula_arena);
    if (root == nullptr) {
      continue;
    }
    shift_arena.reset();
    const parser::AstNode* shifted = parser::shift_relative_refs(*root, shift_arena, row_delta, 0);
    if (shifted != nullptr) {
      column.calculated_column_formula = parser::format_formula(*shifted);
    }
  }
}

void finish_table_extent_change_locked(Workbook& wb, std::vector<Sheet>& sheets,
                                       const eval::RecalcEngine::LockedMutator& mutator,
                                       const std::vector<DefinedName>& defined_names, std::size_t index) {
  TableMetadata& table = wb.mutable_tables()[index];
  table.insert_row = false;
  table.insert_row_shift = false;
  const std::optional<MergeRange> rect = parse_a1_rectangle(table.ref);
  if (AutoFilter* filter = table.auto_filter_xml.get(); filter != nullptr && !filter->is_opaque() && rect) {
    filter->range = *rect;
    if (table.totals_row && rect->last_row > rect->first_row) {
      filter->range.last_row = rect->last_row - 1U;
    }
  }
  // Every structured reference into the table may resolve differently now.
  reindex_formulas_referencing_table(sheets, mutator, wb, {table.name, table.display_name});
  sync_table_spill_barriers_locked(sheets, mutator, wb.tables());
  mark_row_visibility_dependents_dirty_locked(sheets, mutator, defined_names);
}

bool cell_is_date_value(const StylesTable& styles, const CellXf* xf, ExcelProfile profile, const Value& value) {
  return value.is_number() && xf != nullptr &&
         chart_is_date_format(text_format::number_format_code_for_xf(styles, xf, profile));
}

bool cell_is_occupied(const Sheet& sheet, std::uint32_t row, std::uint32_t col) {
  const Cell* cell = sheet.cell_at(row, col);
  return (cell != nullptr && (!cell->formula_text.empty() || !cell->cached_value.is_blank())) ||
         sheet.committed_spill_footprint_covering(row, col).has_value();
}

}  // namespace detail

}  // namespace formulon
