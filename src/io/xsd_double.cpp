//
// Implementation of `parse_xsd_double`. See `xsd_double.h` for the contract.

#include "io/xsd_double.h"

#include <cmath>
#include <cstddef>
#include <string_view>

#include "utils/double_parse.h"

namespace formulon::io {

bool parse_xsd_double(std::string_view text, double* out) {
  const ParsedDouble parsed = parse_double_prefix(text);
  if (parsed.consumed == 0) {
    return false;
  }
  // Trailing garbage is not allowed; trailing whitespace is.
  for (std::size_t i = parsed.consumed; i < text.size(); ++i) {
    if (text[i] != ' ' && text[i] != '\t' && text[i] != '\r' && text[i] != '\n') {
      return false;
    }
  }
  // A magnitude the file states but IEEE 754 cannot hold is a rejection,
  // not a value. `1e999` saturates to +inf and `1e-999` collapses to 0;
  // storing either would silently change the number the cell carries, and
  // the writer turns a non-finite back into `#NUM!` on the next save, so
  // the loss surfaces far from its cause. A subnormal result is out of the
  // normal range but representable, so only an exact zero is rejected here.
  if (parsed.out_of_range && (parsed.value == 0.0 || !std::isfinite(parsed.value))) {
    return false;
  }
  *out = parsed.value;
  return true;
}

bool parse_xsd_nonneg_double(std::string_view text, double* out) {
  double value = 0.0;
  if (!parse_xsd_double(text, &value)) {
    return false;
  }
  if (value < 0.0) {
    return false;
  }
  *out = value;
  return true;
}

}  // namespace formulon::io
