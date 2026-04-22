// Copyright 2026 libraz. Licensed under the MIT License.
//
// Abseil-style status propagation macros built on top of `Expected<T, E>`.
// The macros bail out of the enclosing function on the first failure while
// keeping the call site readable.
//
// Usage:
//   RETURN_IF_ERROR(validate_input(foo));
//   ASSIGN_OR_RETURN(auto bytes, io::read_file(path));
//
// Both macros return `.error()` to the caller, meaning the caller's return
// type must be convertible from the error type (typically `Error`).

#ifndef FORMULON_UTILS_STATUS_MACROS_H_
#define FORMULON_UTILS_STATUS_MACROS_H_

#include <utility>

#include "utils/expected.h"

// Two-level concatenation macros. The indirection is required so that the
// token being pasted (typically `__COUNTER__`) is expanded before pasting.
#define FM_STATUS_MACROS_CONCAT_INNER(a, b) a##b
#define FM_STATUS_MACROS_CONCAT(a, b) FM_STATUS_MACROS_CONCAT_INNER(a, b)

// Per-call-site unique identifier. `__COUNTER__` is preferred because it is
// guaranteed unique within a translation unit; `__LINE__` is used as a
// fallback on the rare compilers that lack `__COUNTER__`.
#if defined(__COUNTER__)
#define FM_STATUS_MACROS_UNIQUE(prefix) FM_STATUS_MACROS_CONCAT(prefix, __COUNTER__)
#else
#define FM_STATUS_MACROS_UNIQUE(prefix) FM_STATUS_MACROS_CONCAT(prefix, __LINE__)
#endif

/// Returns from the enclosing function if `expr` evaluates to an error-state
/// `Expected<...>`. On success falls through.
#define RETURN_IF_ERROR(expr)    \
  do {                           \
    auto _fm_status = (expr);    \
    if (!_fm_status) {           \
      return _fm_status.error(); \
    }                            \
  } while (0)

// Implementation helper for ASSIGN_OR_RETURN. `tmp` is a compiler-generated
// unique identifier supplied by the public macro.
#define FM_ASSIGN_OR_RETURN_IMPL(tmp, lhs, expr) \
  auto tmp = (expr);                             \
  if (!tmp) {                                    \
    return tmp.error();                          \
  }                                              \
  lhs = std::move(tmp.value())

/// Evaluates `expr` (which must return an `Expected<T, E>`); on success binds
/// `lhs` to the unwrapped value, otherwise returns the error to the caller.
#define ASSIGN_OR_RETURN(lhs, expr) FM_ASSIGN_OR_RETURN_IMPL(FM_STATUS_MACROS_UNIQUE(_fm_tmp_), lhs, expr)

#endif  // FORMULON_UTILS_STATUS_MACROS_H_
