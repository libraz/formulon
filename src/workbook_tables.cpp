#include <algorithm>
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
#include "chart/chart_ref.h"
#include "data_table.h"
#include "defined_name.h"
#include "eval/dep_graph.h"
#include "eval/recalc_engine.h"
#include "io/auto_filter_xml.h"
#include "io/xlsb/table_records.h"
#include "io/xml_escape.h"
#include "merge_range.h"
#include "new_workbook_defaults.h"
#include "parser/ref_transforms.h"
#include "parser/reference.h"
#include "parser/structured_ref_payload.h"
#include "parser/structured_ref_spelling.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "sheet_metadata.h"
#include "sheet_name.h"
#include "table.h"
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

// Refusal of `rect` on `sheet_index` as the range of a table other than
// `self`, on the first structure Excel keeps a table from overlapping. With
// `unmerge_inside`, as on creation, a merge wholly inside is not one.
Expected<void, Error> check_table_range_free(const Workbook& wb, std::size_t sheet_index, const MergeRange& rect,
                                             std::optional<std::size_t> self, bool unmerge_inside) {
  const auto refuse = [&](const char* what) {
    std::string ref;
    append_a1_rectangle(ref, rect);
    return make_error(FormulonErrorCode::kPreconditionFailed, what, "ref=" + ref);
  };
  for (std::size_t i = 0; i < wb.tables().size(); ++i) {
    const std::optional<MergeRange> other_rect = detail::table_rect_on_sheet(wb.tables()[i], sheet_index);
    if (i != self && other_rect && intersects(rect, *other_rect)) {
      return refuse("table range overlaps another table");
    }
  }
  for (const io::xlsb::XlsbRetainedTable& kept : io::xlsb::xlsb_retained_tables(wb)) {
    const std::optional<MergeRange> kept_rect = parse_a1_rectangle(kept.ref);
    if (kept.sheet_index == sheet_index && kept_rect && intersects(rect, *kept_rect)) {
      return refuse("table range overlaps another table");
    }
  }
  const Sheet& sheet = wb.sheet(sheet_index);
  for (const std::unique_ptr<pivot::PivotTable>& pivot : sheet.pivot_tables()) {
    if (pivot != nullptr && intersects(detail::pivot_occupied_extent(*pivot), rect)) {
      return refuse("table range overlaps a PivotTable");
    }
  }
  for (const DataTable& data_table : sheet.data_tables()) {
    if (intersects(normalized(data_table.ref), rect)) {
      return refuse("table range overlaps a data table");
    }
  }
  for (const CellAddress address : sheet.formula_cells_in(0U, 0U, rect.last_row, rect.last_col)) {
    const Cell* cell = sheet.cell_at(address.row, address.col);
    if (cell == nullptr || cell->array_block_rows == 0U ||
        (cell->array_block_rows == 1U && cell->array_block_cols == 1U)) {
      continue;
    }
    const MergeRange block{address.row, address.col, address.row + cell->array_block_rows - 1U,
                           address.col + cell->array_block_cols - 1U};
    if (intersects(block, rect)) {
      return refuse("range must not hold a multi-cell array formula");
    }
  }
  for (const MergeRange& merge : sheet.merges()) {
    const MergeRange m = normalized(merge);
    if (intersects(m, rect) && !(unmerge_inside && contains(rect, m))) {
      return refuse(unmerge_inside ? "table range cuts through merged cells" : "table range overlaps merged cells");
    }
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

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
  RETURN_IF_ERROR(check_table_identity(*this, table.name, table.display_name, std::nullopt));
  RETURN_IF_ERROR(check_table_range_free(*this, table.sheet_index, *rect, std::nullopt, /*unmerge_inside=*/true));
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
  RETURN_IF_ERROR(remove_merges_intersecting(table.sheet_index, *rect));
  std::string escaped_ref;
  io::AppendXmlAttrEscaped(escaped_ref, table.ref);
  table.auto_filter_xml.set(io::auto_filter_from_xml("<autoFilter ref=\"" + escaped_ref + "\"/>"));
  tables_.push_back(std::move(table));
  const TableMetadata& created = tables_.back();
  if (created.header_row) {
    for (const std::uint32_t i : renamed_columns) {
      put_cell_text(created.sheet_index, rect->first_row, rect->first_col + i, created.columns[i].name);
    }
  }
  // A structured reference can already exist in a formula authored ahead of
  // its table, so both names need the scoped reindex a defined-name edit gets.
  reindex_formulas_for_table_change({created.name, created.display_name});
  return tables_.size() - 1U;
}

Expected<void, Error> Workbook::check_table_resize(std::size_t index, const MergeRange& rect) const {
  if (index >= tables_.size()) {
    return invalid_table("index out of range", "index=" + std::to_string(index));
  }
  const TableMetadata& table = tables_[index];
  const std::optional<MergeRange> current = parse_a1_rectangle(table.ref);
  if (current && *current == rect) {
    return Expected<void, Error>::Ok();
  }
  if (current && current->first_row != rect.first_row) {
    return invalid_table("table range must keep its first row", "ref=" + table.ref);
  }
  if (table.header_row && rect.first_row == rect.last_row) {
    return invalid_table("table range has no data row", "ref=" + table.ref);
  }
  return check_table_range_free(*this, table.sheet_index, rect, index, /*unmerge_inside=*/false);
}

namespace {

/// `SUBTOTAL` function number of each built-in `ST_TotalsRowFunction` token;
/// the 1xx forms skip hidden rows, as the desktop application writes them.
struct TotalsFunction {
  std::string_view token;
  int subtotal_code;
};
constexpr TotalsFunction kTotalsFunctions[] = {
    {"sum", 109},   {"min", 105},       {"max", 104},    {"average", 101},
    {"count", 103}, {"countNums", 102}, {"stdDev", 107}, {"var", 110},
};

int subtotal_code_of(std::string_view token) {
  for (const TotalsFunction& entry : kTotalsFunctions) {
    if (entry.token == token) {
      return entry.subtotal_code;
    }
  }
  return 0;
}

// Refusal of the row below `rect` (within its columns) as a new totals row,
// on the first thing occupying it. `rect` must end above the grid's last row.
Expected<void, Error> check_totals_row_free(const Workbook& wb, std::size_t index, const MergeRange& rect) {
  const std::size_t sheet_index = wb.tables()[index].sheet_index;
  const MergeRange below{rect.last_row + 1U, rect.first_col, rect.last_row + 1U, rect.last_col};
  const auto refuse = [&](const char* what) {
    std::string ref;
    append_a1_rectangle(ref, below);
    return make_error(FormulonErrorCode::kPreconditionFailed, what, "row=" + ref);
  };
  const Sheet& sheet = wb.sheet(sheet_index);
  for (std::uint32_t col = below.first_col; col <= below.last_col; ++col) {
    const Cell* cell = sheet.cell_at(below.first_row, col);
    if (cell != nullptr && (!cell->formula_text.empty() || !cell->cached_value.is_blank())) {
      return refuse("patch.totals_row needs the row below the table to be empty");
    }
    if (sheet.committed_spill_footprint_covering(below.first_row, col).has_value()) {
      return refuse("patch.totals_row needs the row below the table free of a spilled result");
    }
  }
  MergeRange grown = rect;
  grown.last_row = below.last_row;
  return wb.check_table_resize(index, grown);
}

// Whether a merge, table, PivotTable, multi-cell array formula or data table
// reaches `edit`'s moving region across its band edge, which keeps a totals
// row toggle from shifting cells.
bool band_meets_straddler(const Workbook& wb, std::size_t sheet_index, const StructuralEdit& edit) {
  const Sheet& sheet = wb.sheet(sheet_index);
  const auto straddles = [&edit](const MergeRange& rect) { return band_fit(rect, edit) == BandFit::kStraddles; };
  if (std::any_of(sheet.merges().begin(), sheet.merges().end(), straddles)) {
    return true;
  }
  for (const DataTable& table : sheet.data_tables()) {
    if (straddles(data_table_block(table))) {
      return true;
    }
  }
  for (const std::unique_ptr<pivot::PivotTable>& pivot : sheet.pivot_tables()) {
    if (pivot != nullptr && straddles(pivot_location(*pivot))) {
      return true;
    }
  }
  for (const TableMetadata& table : wb.tables()) {
    const std::optional<MergeRange> ref = detail::table_rect_on_sheet(table, sheet_index);
    if (ref.has_value() && straddles(*ref)) {
      return true;
    }
  }
  for (const CellAddress address : sheet.all_formula_cells()) {
    const Cell* cell = sheet.cell_at(address.row, address.col);
    if (cell != nullptr && cell->array_block_rows != 0U &&
        straddles(MergeRange{address.row, address.col, address.row + cell->array_block_rows - 1U,
                             address.col + cell->array_block_cols - 1U})) {
      return true;
    }
  }
  return false;
}

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
  const Band columns{rect->first_col, rect->last_col};
  // The toggle inserts or deletes cells, rewrites references and writes the
  // totals cells as one edit under a single hold of the engine mutex.
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  MergeRange next = *rect;
  if (shown) {
    if (rect->last_row + 1U >= Sheet::kMaxRows) {
      return make_error(FormulonErrorCode::kPreconditionFailed,
                        "patch.totals_row cannot be shown on a table ending on the last row of the grid",
                        "ref=" + tables_[index].ref);
    }
    const MergeRange below{rect->last_row + 1U, rect->first_col, rect->last_row + 1U, rect->last_col};
    // Something straddling the columns below keeps the cells from shifting;
    // the totals row then takes the row below only while it is empty.
    if (band_meets_straddler(*this, sheet_index, StructuralEdit{below.first_row, 1U, false, true, columns})) {
      RETURN_IF_ERROR(check_totals_row_free(*this, index, *rect));
    } else {
      RETURN_IF_ERROR(apply_band_edit_locked(sheet_index, below, CellShift::kDown, /*is_delete=*/false));
    }
    next.last_row = below.last_row;
  } else {
    const std::uint32_t data_rows = rect->last_row - rect->first_row - (tables_[index].header_row ? 1U : 0U);
    if (data_rows == 0U) {
      return make_error(FormulonErrorCode::kPreconditionFailed,
                        "patch.totals_row cannot be hidden when the table would keep no data row",
                        "ref=" + tables_[index].ref);
    }
    for (std::uint32_t col = rect->first_col; col <= rect->last_col; ++col) {
      const std::optional<ArrayFormulaInfo> block = array_formula_at(sheet_index, rect->last_row, col);
      if ((block && (block->rows != 1U || block->cols != 1U)) ||
          data_table_at(sheet_index, rect->last_row, col) != nullptr) {
        return make_error(FormulonErrorCode::kPreconditionFailed,
                          "patch.totals_row cannot be hidden while the totals row holds part of a multi-cell array "
                          "formula or data table",
                          "ref=" + tables_[index].ref);
      }
    }
    const MergeRange totals{rect->last_row, rect->first_col, rect->last_row, rect->last_col};
    if (totals_row_referenced(sheets_, tables_, defined_names_, index, totals)) {
      return make_error(FormulonErrorCode::kPreconditionFailed, "a formula or name references the totals row",
                        "ref=" + tables_[index].ref);
    }
    next.last_row = rect->last_row - 1U;
  }

  const bool shifts_back =
      !shown && !band_meets_straddler(*this, sheet_index, StructuralEdit{rect->last_row, 1U, true, true, columns});
  const std::string ref_before = tables_[index].ref;
  // The range shrinks before the old totals row goes, so the edit below never
  // reads the last data row as the table's end.
  tables_[index].ref.clear();
  append_a1_rectangle(tables_[index].ref, next);
  tables_[index].totals_row = shown;
  if (!shown) {
    const MergeRange totals{rect->last_row, rect->first_col, rect->last_row, rect->last_col};
    if (shifts_back) {
      if (Expected<void, Error> deleted = apply_band_edit_locked(sheet_index, totals, CellShift::kUp, true); !deleted) {
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
  // The extent changed, so every structured reference into the table may
  // resolve differently; reindex the formulas that name it.
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  reindex_formulas_referencing_table(sheets_, mutator, *this, {table.name, table.display_name});
  sync_table_spill_barriers_locked(sheets_, mutator, tables_);
  mark_row_visibility_dependents_dirty_locked(sheets_, mutator, defined_names_);
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
    parser::StructuredRefPayload payload;
    payload.columns.push_back(col.name);
    put_cell_formula_locked(
        table.sheet_index, row, cell_col,
        "=SUBTOTAL(" + std::to_string(code) + "," + table.display_name + "[" + parser::stored_payload(payload) + "])",
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

}  // namespace detail

}  // namespace formulon
