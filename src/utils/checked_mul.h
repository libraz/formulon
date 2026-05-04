// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// `checked_mul_size_t(a, b)` performs `a * b` with overflow detection,
// returning `Expected<std::size_t, Error>`. On 32-bit `size_t` (the WASM
// main build) this catches `rows * cols` cases that a bare cast would
// silently truncate; on 64-bit it is identical-after-inlining to a plain
// multiply.
//
// This guard is defensive: most call sites multiply two small `uint32_t`
// dimensions, and on 64-bit `size_t` there is no observable difference from
// the existing `static_cast<size_t>(a) * static_cast<size_t>(b)` pattern. On
// 32-bit `size_t`, however, an attacker-controlled pair of dimensions can
// wrap, causing downstream code to allocate a small buffer and then index
// well past its end. Surface the overflow as a recoverable error here so
// the caller can map it to an Excel-visible `#NUM!` instead.

#ifndef FORMULON_UTILS_CHECKED_MUL_H_
#define FORMULON_UTILS_CHECKED_MUL_H_

#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>

#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

/// Multiplies two non-negative integer-like operands as `std::size_t` with
/// overflow detection. Returns `kFnOverflow` when the product would exceed
/// `std::numeric_limits<std::size_t>::max()`.
inline Expected<std::size_t, Error> checked_mul_size_t(std::size_t a, std::size_t b) {
  if (a == 0U || b == 0U) {
    return std::size_t{0};
  }
  if (a > std::numeric_limits<std::size_t>::max() / b) {
    return make_error(FormulonErrorCode::kFnOverflow, "checked_mul_size_t: size_t overflow",
                      "a=" + std::to_string(a) + " b=" + std::to_string(b));
  }
  return a * b;
}

}  // namespace formulon

#endif  // FORMULON_UTILS_CHECKED_MUL_H_
