#include "utils/strings.h"

namespace formulon {
namespace strings {

bool case_insensitive_eq_bytes(const char* a, const char* b, std::size_t n) noexcept {
  for (std::size_t i = 0; i < n; ++i) {
    if (ascii_to_lower(a[i]) != ascii_to_lower(b[i])) {
      return false;
    }
  }
  return true;
}

}  // namespace strings
}  // namespace formulon
