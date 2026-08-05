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
// token being pasted (`__LINE__`) is expanded before pasting.
#define FM_CONCAT_INNER(a, b) a##b
#define FM_CONCAT(a, b) FM_CONCAT_INNER(a, b)

// Per-call-site unique identifier. We use `__LINE__` rather than the more
// robust `__COUNTER__` because Emscripten's bundled clang rejects the
// latter under `-Werror -Wc2y-extensions` (it has not yet promoted the
// long-standing extension to a non-diagnosed builtin). The trade-off is
// that callers MUST NOT expand `ASSIGN_OR_RETURN` twice on the same
// source line: the resulting identifiers would collide. No call site in
// the codebase does so today; the convention is one macro per line.
#define FM_UNIQUE(prefix) FM_CONCAT(prefix, __LINE__)

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
#define ASSIGN_OR_RETURN(lhs, expr) FM_ASSIGN_OR_RETURN_IMPL(FM_UNIQUE(_fm_tmp_), lhs, expr)

#endif  // FORMULON_UTILS_STATUS_MACROS_H_
