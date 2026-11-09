#include "workbook_table_loss.h"

#include <algorithm>
#include <cstddef>
#include <optional>
#include <string>
#include <string_view>

#include "parser/structured_ref_payload.h"
#include "sheet_name.h"

namespace formulon {

std::optional<parser::StructuredRefTarget> TableLossTransform::apply_structured_ref(std::string_view table,
                                                                                    std::string_view column) const {
  for (const TableLoss& loss : losses_) {
    if (!sheet_names::NameList::denotes(table_list_, table, loss.display_name) &&
        !sheet_names::NameList::denotes(table_list_, table, loss.name)) {
      continue;
    }
    if (loss.whole) {
      return std::nullopt;
    }
    const auto payload = parser::parse_structured_ref_payload_text(column);
    if (!payload) {
      break;
    }
    // A column spelling resolves among the table's columns before the edit.
    const auto deleted = [&loss](const std::string& name) {
      const std::size_t found = sheet_names::resolve(loss.all_columns, name);
      return found != sheet_names::kNotFound &&
             std::find(loss.columns.begin(), loss.columns.end(), loss.all_columns[found]) != loss.columns.end();
    };
    if (std::any_of(payload->columns.begin(), payload->columns.end(), deleted)) {
      return std::nullopt;
    }
    break;
  }
  return inner_.apply_structured_ref(table, column);
}

}  // namespace formulon
