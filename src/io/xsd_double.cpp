//
// Implementation of `parse_xsd_double`. See `xsd_double.h` for the contract.

#include "io/xsd_double.h"

#include <cerrno>
#include <cmath>
#include <cstdlib>
#include <cstring>
#include <string>
#include <string_view>

namespace formulon::io {
namespace {

/// Rejects the three lexical forms `std::strtod` accepts and the xs:double
/// lexical space does not: hexadecimal (`0x10`, which would silently read
/// as 16), and the `inf` / `infinity` / `nan` spellings in any case.
/// Excel emits none of them, and each produces a value the writer cannot
/// round-trip.
///
/// Leading whitespace is skipped first, matching what `strtod` itself
/// consumes — otherwise a padded `" 0x10"` would slip past the gate.
bool StartsOutsideXsdDoubleSpace(std::string_view text) noexcept {
  std::size_t i = 0;
  while (i < text.size() && (text[i] == ' ' || text[i] == '\t' || text[i] == '\r' || text[i] == '\n')) {
    ++i;
  }
  if (i >= text.size()) {
    return true;
  }
  if (text[i] == '+' || text[i] == '-') {
    ++i;
  }
  if (i >= text.size()) {
    return true;
  }
  const char lead = text[i];
  if (lead == 'i' || lead == 'I' || lead == 'n' || lead == 'N') {
    return true;
  }
  return lead == '0' && (i + 1) < text.size() && (text[i + 1] == 'x' || text[i + 1] == 'X');
}

}  // namespace

bool parse_xsd_double(std::string_view text, double* out) {
  if (text.empty()) {
    return false;
  }
  if (StartsOutsideXsdDoubleSpace(text)) {
    return false;
  }
  // strtod requires a NUL-terminated string. A pugixml `text().get()` already
  // is, but a caller may hand in a `string_view` substring, so copy into a
  // small stack buffer and fall back to the heap only for oversized input.
  char small_buf[64];
  const char* nstr = nullptr;
  std::string heap;
  if (text.size() < sizeof(small_buf)) {
    std::memcpy(small_buf, text.data(), text.size());
    small_buf[text.size()] = '\0';
    nstr = small_buf;
  } else {
    heap.assign(text.data(), text.size());
    nstr = heap.c_str();
  }
  char* end = nullptr;
  errno = 0;
  const double value = std::strtod(nstr, &end);
  if (end == nstr) {
    return false;
  }
  // Trailing garbage is not allowed; trailing whitespace is.
  while (end != nullptr && *end != '\0') {
    if (*end != ' ' && *end != '\t' && *end != '\r' && *end != '\n') {
      return false;
    }
    ++end;
  }
  // A magnitude the file states but IEEE 754 cannot hold is a rejection,
  // not a value. `1e999` saturates to +inf and `1e-999` collapses to 0;
  // storing either would silently change the number the cell carries, and
  // the writer turns a non-finite back into `#NUM!` on the next save, so
  // the loss surfaces far from its cause. A subnormal result keeps its
  // ERANGE but is representable, so only an exact zero is rejected here.
  if (errno == ERANGE && (value == 0.0 || !std::isfinite(value))) {
    return false;
  }
  if (!std::isfinite(value)) {
    return false;
  }
  *out = value;
  return true;
}

}  // namespace formulon::io
