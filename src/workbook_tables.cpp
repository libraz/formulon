#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_set>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "io/auto_filter_xml.h"
#include "io/xlsb/table_records.h"
#include "io/xml_escape.h"
#include "table.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/strings.h"
#include "workbook.h"

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
  return strings::case_insensitive_eq(existing.name, name) ||
         strings::case_insensitive_eq(existing.display_name, name) ||
         strings::case_insensitive_eq(existing.name, display_name) ||
         strings::case_insensitive_eq(existing.display_name, display_name);
}

bool name_in_use(const std::vector<TableMetadata>& tables, const std::vector<io::xlsb::XlsbRetainedTable>& retained,
                 std::string_view name, std::string_view display_name) {
  for (const TableMetadata& existing : tables) {
    if (names_collide(existing, name, display_name)) {
      return true;
    }
  }
  for (const io::xlsb::XlsbRetainedTable& existing : retained) {
    if (names_collide(existing, name, display_name)) {
      return true;
    }
  }
  return false;
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
  const std::vector<io::xlsb::XlsbRetainedTable> retained = io::xlsb::xlsb_retained_tables(*this);
  if (name_in_use(tables_, retained, table.name, table.display_name)) {
    return invalid_table("table name already exists", "reason=table name already exists name=" + table.name);
  }
  std::uint32_t next_id = 1;
  for (const TableMetadata& existing : tables_) {
    next_id = next_id > existing.id + 1U ? next_id : existing.id + 1U;
  }
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
  table.id = next_id;
  std::unordered_set<std::string> column_keys;
  column_keys.reserve(table.columns.size());
  std::uint32_t column_id = 1;
  for (TableColumn& column : table.columns) {
    if (column.name.empty()) {
      return invalid_table("column name is empty");
    }
    if (!column_keys.insert(strings::to_ascii_lower(column.name)).second) {
      return invalid_table("column name is duplicated");
    }
    column.id = column_id++;
  }
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

std::string Workbook::next_default_table_name(std::string_view word) const {
  const std::vector<io::xlsb::XlsbRetainedTable> retained = io::xlsb::xlsb_retained_tables(*this);
  for (std::size_t n = 1;; ++n) {
    const std::string candidate = std::string(word) + std::to_string(n);
    if (!name_in_use(tables_, retained, candidate, candidate)) {
      return candidate;
    }
  }
}

}  // namespace formulon
