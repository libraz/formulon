#pragma once

#include <cstddef>

// An out-index slot for C ABI adds whose allocated index the test does not read.
inline std::size_t* ignored_index() {
  static std::size_t slot = 0;
  return &slot;
}
