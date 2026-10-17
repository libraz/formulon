#include "eval/locale_text.h"

#include <cstddef>
#include <cstdlib>
#include <string>

#include "eval/eval_profile_scope.h"
#include "excel_locale.h"
#include "utils/double_format.h"
#include "utils/subnormal_flush.h"

namespace formulon {
namespace eval {

char locale_decimal_separator() noexcept {
  return locale_facts(current_eval_profile()).decimal_separator;
}

std::string locale_number_text(double value) {
  // A subnormal reads as a signed zero.
  if (value != 0.0 && flush_subnormal(value) == 0.0) {
    return value < 0.0 ? "-0" : "0";
  }
  std::string out;
  format_double(out, value);
  // Excel keeps 14 significant digits once the decimal exponent reaches 99 in magnitude.
  if (const std::size_t e = out.find('E'); e != std::string::npos) {
    const int exponent = std::abs(std::atoi(out.c_str() + e + 1));
    if (exponent >= 99) {
      out.clear();
      format_double(out, value, 14);
    }
  }
  const char separator = locale_decimal_separator();
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
