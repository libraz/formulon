//
// Implementation of `parse_xsd_double`. See `xsd_double.h` for the contract.

#include "io/xsd_double.h"

#include <cstdlib>
#include <cstring>
#include <string>
#include <string_view>

namespace formulon::io {

bool parse_xsd_double(std::string_view text, double* out) {
  if (text.empty()) {
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
  *out = value;
  return true;
}

}  // namespace formulon::io
