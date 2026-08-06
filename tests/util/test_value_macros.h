//
// Value-shape assertion macros for unit tests.
//
// Each macro pairs an `ASSERT_TRUE(_v.is_X())` with a typed value
// check. Using ASSERT_TRUE (not EXPECT_TRUE) on the shape guard is
// load-bearing: if the value's kind is wrong, subsequent accessors
// (e.g. `as_number()` on a non-Number) would abort. Failing fast lets
// the per-test diagnostic surface the unexpected kind via
// `debug_to_string()` rather than terminating the whole binary.
//
// The macros wrap their bodies in `do { ... } while (0)` so they
// compose with `if`/`else` chains without requiring trailing braces
// and so multi-statement bodies cannot leak a stray `;`.

#ifndef FORMULON_TESTS_UTIL_TEST_VALUE_MACROS_H_
#define FORMULON_TESTS_UTIL_TEST_VALUE_MACROS_H_

#include <gtest/gtest.h>

#include "value.h"

// Asserts `v` is a Number and its payload equals `expected` under the
// `EXPECT_DOUBLE_EQ` (4-ULP) tolerance. Use the `_NEAR` variant for
// tests that need a custom tolerance.
#define EXPECT_VALUE_NUMBER(v, expected)                                            \
  do {                                                                              \
    const ::formulon::Value& _v = (v);                                              \
    ASSERT_TRUE(_v.is_number()) << "expected Number, got " << _v.debug_to_string(); \
    EXPECT_DOUBLE_EQ(_v.as_number(), (expected));                                   \
  } while (0)

// Like `EXPECT_VALUE_NUMBER` but with caller-supplied absolute
// tolerance. Useful for stat / financial / engineering functions whose
// reference values come from hand-computed identities rather than
// IEEE-754 bit-exact reproduction.
#define EXPECT_VALUE_NUMBER_NEAR(v, expected, tol)                                  \
  do {                                                                              \
    const ::formulon::Value& _v = (v);                                              \
    ASSERT_TRUE(_v.is_number()) << "expected Number, got " << _v.debug_to_string(); \
    EXPECT_NEAR(_v.as_number(), (expected), (tol));                                 \
  } while (0)

// Asserts `v` is a Text and its payload (as `std::string_view`) equals
// `expected`. `expected` may be any type implicitly convertible to
// `std::string_view`.
#define EXPECT_VALUE_TEXT(v, expected)                                          \
  do {                                                                          \
    const ::formulon::Value& _v = (v);                                          \
    ASSERT_TRUE(_v.is_text()) << "expected Text, got " << _v.debug_to_string(); \
    EXPECT_EQ(_v.as_text(), ::std::string_view(expected));                      \
  } while (0)

// Asserts `v` is a Bool and its payload equals `expected`.
#define EXPECT_VALUE_BOOL(v, expected)                                             \
  do {                                                                             \
    const ::formulon::Value& _v = (v);                                             \
    ASSERT_TRUE(_v.is_boolean()) << "expected Bool, got " << _v.debug_to_string(); \
    EXPECT_EQ(_v.as_boolean(), (expected));                                        \
  } while (0)

// Asserts `v` is an Error and its `ErrorCode` payload equals `code`.
// Eliminates the
// `ASSERT_TRUE(v.is_error()); EXPECT_EQ(v.as_error(), ErrorCode::...);`
// idiom duplicated across the test tree.
#define EXPECT_VALUE_ERROR(v, code)                                               \
  do {                                                                            \
    const ::formulon::Value& _v = (v);                                            \
    ASSERT_TRUE(_v.is_error()) << "expected Error, got " << _v.debug_to_string(); \
    EXPECT_EQ(_v.as_error(), (code));                                             \
  } while (0)

// Asserts `v` is the Blank variant. No payload to check.
#define EXPECT_VALUE_BLANK(v)                                                     \
  do {                                                                            \
    const ::formulon::Value& _v = (v);                                            \
    EXPECT_TRUE(_v.is_blank()) << "expected Blank, got " << _v.debug_to_string(); \
  } while (0)

#endif  // FORMULON_TESTS_UTIL_TEST_VALUE_MACROS_H_
