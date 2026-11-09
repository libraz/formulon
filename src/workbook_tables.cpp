#include <cstddef>
#include <cstdint>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_set>
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

/// True when `existing` owns `name` or `display_name` as either of its names.
template <typename Existing>
bool names_collide(const Existing& existing, std::string_view name, std::string_view display_name) {
  return collection_names::equal(existing.name, name) || collection_names::equal(existing.display_name, name) ||
         collection_names::equal(existing.name, display_name) ||
         collection_names::equal(existing.display_name, display_name);
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
    if (collection_names::equal(defined.name, name) || collection_names::equal(defined.name, display_name)) {
      return true;
    }
  }
  return false;
}

Expected<void, Error> check_table_identity(const Workbook& wb, std::string_view name, std::string_view display_name,
                                           std::optional<std::size_t> self) {
  if (name.empty() || !is_valid_table_name(display_name)) {
    return invalid_table("invalid table name",
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
      return refuse("multi-cell array formulas are not valid in a table");
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
    return invalid_table("invalid sheet, table identity, range, or columns");
  }
  const std::size_t width = rect ? static_cast<std::size_t>(rect->last_col - rect->first_col) + 1U : 0U;
  if (width != table.columns.size()) {
    return invalid_table("column_count does not match the range width",
                         "ref=" + table.ref + " column_count=" + std::to_string(table.columns.size()));
  }
  RETURN_IF_ERROR(check_table_identity(*this, table.name, table.display_name, std::nullopt));
  RETURN_IF_ERROR(check_table_range_free(*this, table.sheet_index, *rect, std::nullopt, /*unmerge_inside=*/true));
  table.id = next_free_table_id(0);
  std::unordered_set<std::string> column_keys;
  column_keys.reserve(table.columns.size());
  std::uint32_t column_id = 1;
  for (TableColumn& column : table.columns) {
    if (column.name.empty()) {
      return invalid_table("column name is empty");
    }
    if (!column_keys.insert(collection_names::fold_key(column.name)).second) {
      return invalid_table("column name is duplicated");
    }
    column.id = column_id++;
  }
  // Excel unmerges the merges a new table covers.
  RETURN_IF_ERROR(remove_merges_intersecting(table.sheet_index, *rect));
  std::string escaped_ref;
  io::AppendXmlAttrEscaped(escaped_ref, table.ref);
  table.auto_filter_xml.set(io::auto_filter_from_xml("<autoFilter ref=\"" + escaped_ref + "\"/>"));
  tables_.push_back(std::move(table));
  // A structured reference can already exist in a formula authored ahead of
  // its table, so both names need the scoped reindex a defined-name edit gets.
  const TableMetadata& created = tables_.back();
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
      return refuse("the row below the table is not empty");
    }
    if (sheet.committed_spill_footprint_covering(below.first_row, col).has_value()) {
      return refuse("the row below the table holds a spilled result");
    }
  }
  MergeRange grown = rect;
  grown.last_row = below.last_row;
  return wb.check_table_resize(index, grown);
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
  MergeRange next = *rect;
  if (shown) {
    if (rect->last_row + 1U >= Sheet::kMaxRows) {
      return make_error(FormulonErrorCode::kPreconditionFailed, "the table ends on the last row of the grid",
                        "ref=" + tables_[index].ref);
    }
    RETURN_IF_ERROR(check_totals_row_free(*this, index, *rect));
    next.last_row = rect->last_row + 1U;
  } else {
    const std::uint32_t data_rows = rect->last_row - rect->first_row - (tables_[index].header_row ? 1U : 0U);
    if (data_rows == 0U) {
      return make_error(FormulonErrorCode::kPreconditionFailed, "the table would keep no data row",
                        "ref=" + tables_[index].ref);
    }
    for (std::uint32_t col = rect->first_col; col <= rect->last_col; ++col) {
      const std::optional<ArrayFormulaInfo> block = array_formula_at(tables_[index].sheet_index, rect->last_row, col);
      if ((block && (block->rows != 1U || block->cols != 1U)) ||
          data_table_at(tables_[index].sheet_index, rect->last_row, col) != nullptr) {
        return make_error(FormulonErrorCode::kPreconditionFailed,
                          "the totals row holds part of a multi-cell array formula or data table",
                          "ref=" + tables_[index].ref);
      }
    }
    next.last_row = rect->last_row - 1U;
  }

  TableMetadata& table = tables_[index];
  table.ref.clear();
  append_a1_rectangle(table.ref, next);
  table.totals_row = shown;
  if (shown) {
    // Shown now, so the "never shown" marker no longer applies.
    table.totals_row_shown.reset();
  }
  // The extent changed, so every structured reference into the table may
  // resolve differently; reindex the formulas that name it.
  reindex_formulas_for_table_change({table.name, table.display_name});
  mark_row_visibility_dependents_dirty();
  if (!shown) {
    for (std::uint32_t col = rect->first_col; col <= rect->last_col; ++col) {
      if (sheets_[table.sheet_index].cell_at(rect->last_row, col) != nullptr) {
        RETURN_IF_ERROR(set_cell_value(table.sheet_index, rect->last_row, col, Value::blank()));
      }
    }
    return Expected<void, Error>::Ok();
  }
  for (std::size_t column = 0; column < table.columns.size(); ++column) {
    const TableColumn& col = table.columns[column];
    if (col.totals_function.empty() && col.totals_label.empty()) {
      continue;
    }
    RETURN_IF_ERROR(write_table_totals_cell(index, column));
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::write_table_totals_cell(std::size_t index, std::size_t column) {
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
    return set_cell_formula(
        table.sheet_index, row, cell_col,
        "=SUBTOTAL(" + std::to_string(code) + "," + table.display_name + "[" + parser::stored_payload(payload) + "])");
  }
  if (col.totals_function == "custom" && !col.totals_row_formula.empty()) {
    return set_cell_formula(table.sheet_index, row, cell_col, "=" + col.totals_row_formula);
  }
  if (!col.totals_label.empty()) {
    return set_cell_text(table.sheet_index, row, cell_col, col.totals_label);
  }
  return set_cell_value(table.sheet_index, row, cell_col, Value::blank());
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
  for (const TableMetadata& table : tables_) {
    if (collection_names::equal(table.name, name) || collection_names::equal(table.display_name, name)) {
      return true;
    }
  }
  return false;
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

/// Maps a structured reference's table name from either of a table's old
/// names to its new display name.
class TableRenameTransform final : public parser::RefTransform {
 public:
  TableRenameTransform(std::string_view old_name, std::string_view old_display_name, std::string_view new_name)
      : old_name_(old_name), old_display_name_(old_display_name), new_name_(new_name) {}

  std::optional<parser::Reference> apply(const parser::Reference& ref) const override { return ref; }

  std::optional<parser::StructuredRefTarget> apply_structured_ref(std::string_view table,
                                                                  std::string_view /*column*/) const override {
    if (collection_names::equal(table, old_name_) || collection_names::equal(table, old_display_name_)) {
      return parser::StructuredRefTarget{new_name_, {}};
    }
    return parser::StructuredRefTarget{table, {}};
  }

 private:
  std::string_view old_name_;
  std::string_view old_display_name_;
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
  const TableRenameTransform transform(old_name, old_display_name, display_name);
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
    if (!source.name.empty() &&
        (collection_names::equal(source.name, old_name) || collection_names::equal(source.name, old_display_name))) {
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

/// Maps a structured reference through one table naming `old_column` to the
/// same reference naming `new_column`, re-spelled in the stored form.
class TableColumnRenameTransform final : public parser::RefTransform {
 public:
  TableColumnRenameTransform(std::string_view table_name, std::string_view table_display_name,
                             std::string_view old_column, std::string_view new_column)
      : table_name_(table_name),
        table_display_name_(table_display_name),
        old_column_(old_column),
        new_column_(new_column) {}

  std::optional<parser::Reference> apply(const parser::Reference& ref) const override { return ref; }

  std::optional<parser::StructuredRefTarget> apply_structured_ref(std::string_view table,
                                                                  std::string_view column) const override {
    if (!collection_names::equal(table, table_name_) && !collection_names::equal(table, table_display_name_)) {
      return parser::StructuredRefTarget{table, {}};
    }
    std::optional<parser::StructuredRefPayload> payload = parser::parse_structured_ref_payload_text(column);
    if (!payload) {
      return parser::StructuredRefTarget{table, {}};
    }
    bool renamed = false;
    for (std::string& name : payload->columns) {
      if (collection_names::equal(name, old_column_)) {
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
  std::string_view table_name_;
  std::string_view table_display_name_;
  std::string_view old_column_;
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
  const auto taken = [&columns, column](std::string_view candidate) {
    for (std::size_t i = 0; i < columns.size(); ++i) {
      if (i != column && collection_names::equal(columns[i].name, candidate)) {
        return true;
      }
    }
    return false;
  };
  // An empty header takes `<word>N` from N = 1; a duplicate takes `<name>N` from N = 2.
  std::string final_name(name);
  if (name.empty() || taken(name)) {
    const std::string stem =
        name.empty() ? std::string(new_workbook_defaults(excel_profile_.locale).column_word) : std::string(name);
    for (std::size_t n = name.empty() ? 1U : 2U;; ++n) {
      final_name = stem + std::to_string(n);
      if (!taken(final_name)) {
        break;
      }
    }
  }

  const TableMetadata& meta = tables_[table];
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  const std::string old_name = columns[column].name;
  const bool renamed = final_name != old_name;
  if (renamed) {
    const TableColumnRenameTransform transform(meta.name, meta.display_name, old_name, final_name);
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

}  // namespace formulon
