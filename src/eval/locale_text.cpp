#include "eval/locale_text.h"

#include <cstddef>

#include "eval/eval_profile_scope.h"
#include "excel_locale.h"
#include "utils/double_format.h"

namespace formulon {
namespace eval {

std::string locale_number_text(double value) {
  std::string out;
  format_double(out, value);
  const char separator = locale_facts(current_eval_profile()).decimal_separator;
  if (separator != '.') {
    for (char& c : out) {
      if (c == '.') {
        c = separator;
      }
    }
  }
  return out;
}

std::string_view locale_bool_text(bool value) noexcept {
  const LocaleFacts& facts = locale_facts(current_eval_profile());
  return value ? facts.true_name : facts.false_name;
}

std::string_view locale_error_text(ErrorCode code) noexcept {
  return locale_facts(current_eval_profile()).error_names[static_cast<std::size_t>(code)];
}

}  // namespace eval
}  // namespace formulon
