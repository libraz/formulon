#include "workbook_table_loss.h"

#include <algorithm>
#include <optional>
#include <string>
#include <string_view>

#include "eval/structured_ref.h"
#include "utils/strings.h"

namespace formulon {

std::optional<std::string_view> TableLossTransform::apply_structured_ref(std::string_view table,
                                                                         std::string_view column) const {
  for (const TableLoss& loss : losses_) {
    if (!strings::case_insensitive_eq(table, loss.name) && !strings::case_insensitive_eq(table, loss.display_name)) {
      continue;
    }
    if (loss.whole) {
      return std::nullopt;
    }
    const auto selector = eval::parse_structured_ref_payload(column);
    if (!selector) {
      break;
    }
    const auto deleted = [&loss](std::string_view written) {
      const std::string name = eval::unescape_structured_ref_name(written);
      return std::any_of(loss.columns.begin(), loss.columns.end(),
                         [&name](const std::string& gone) { return strings::case_insensitive_eq(gone, name); });
    };
    if ((selector.value().column_first_set && deleted(selector.value().column_first)) ||
        (selector.value().column_last_set && deleted(selector.value().column_last))) {
      return std::nullopt;
    }
    break;
  }
  return inner_.apply_structured_ref(table, column);
}

}  // namespace formulon
