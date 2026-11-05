#include "workbook_table_loss.h"

#include <algorithm>
#include <optional>
#include <string>
#include <string_view>

#include "parser/structured_ref_payload.h"
#include "utils/strings.h"

namespace formulon {

std::optional<parser::StructuredRefTarget> TableLossTransform::apply_structured_ref(std::string_view table,
                                                                                    std::string_view column) const {
  for (const TableLoss& loss : losses_) {
    if (!strings::case_insensitive_eq(table, loss.name) && !strings::case_insensitive_eq(table, loss.display_name)) {
      continue;
    }
    if (loss.whole) {
      return std::nullopt;
    }
    const auto payload = parser::parse_structured_ref_payload_text(column);
    if (!payload) {
      break;
    }
    const auto deleted = [&loss](const std::string& name) {
      return std::any_of(loss.columns.begin(), loss.columns.end(),
                         [&name](const std::string& gone) { return strings::case_insensitive_eq(gone, name); });
    };
    if (std::any_of(payload->columns.begin(), payload->columns.end(), deleted)) {
      return std::nullopt;
    }
    break;
  }
  return inner_.apply_structured_ref(table, column);
}

}  // namespace formulon
