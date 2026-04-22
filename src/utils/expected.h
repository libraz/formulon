// Copyright 2026 libraz. Licensed under the MIT License.
//
// `Expected<T, E>` is Formulon's exclusive error-propagation primitive. The
// engine is built with `-fno-exceptions -fno-rtti`; every fallible API signals
// failure by returning `Expected<T, Error>` rather than throwing. See
// backup/plans/23-error-codes.md for the full rationale.
//
// Key design points:
//   * The internal state is a `std::variant<T, E>` so the same storage can
//     carry either the success payload or the error value.
//   * `value()` and `error()` are contract-checked; a violation aborts the
//     process via `FM_CHECK`. The engine must not attempt to read a value
//     that has not been produced.
//   * `take()` is a one-shot move accessor. In debug builds we track a
//     "consumed" flag so accidental re-reads abort loudly.
//   * A `void` specialization supports APIs that succeed without a payload
//     (e.g. validators).

#ifndef FORMULON_UTILS_EXPECTED_H_
#define FORMULON_UTILS_EXPECTED_H_

#include <cstdio>
#include <cstdlib>
#include <type_traits>
#include <utility>
#include <variant>

#include "utils/error.h"

// Formulon's abort-on-contract-violation macro. Callers use this in place of
// `throw` / `std::terminate`. On Debug builds we print the location before
// calling `std::abort`; on Release builds we keep the abort but skip the
// stderr message to stay small.
#if defined(NDEBUG)
#define FM_CHECK(cond, msg) \
  do {                      \
    if (!(cond)) {          \
      std::abort();         \
    }                       \
  } while (0)
#else
#define FM_CHECK(cond, msg)                                                              \
  do {                                                                                   \
    if (!(cond)) {                                                                       \
      std::fprintf(stderr, "FM_CHECK failed at %s:%d: %s\n", __FILE__, __LINE__, (msg)); \
      std::abort();                                                                      \
    }                                                                                    \
  } while (0)
#endif

namespace formulon {

/// Result type for every fallible Formulon API.
///
/// `Expected<T, E>` holds either a value of type `T` (success) or a value of
/// type `E` (failure, typically `Error`). Instances are constructed through
/// the implicit converting constructors from `T` / `E`, or the named factory
/// methods `Ok` / `Err` when the types overlap.
///
/// The monadic helpers `map`, `and_then` and `or_else` make it easy to chain
/// transformations without manual error plumbing. For the common early-return
/// pattern, prefer the `RETURN_IF_ERROR` / `ASSIGN_OR_RETURN` macros defined
/// in `status_macros.h`.
template <class T, class E = Error>
class Expected {
 public:
  using value_type = T;
  using error_type = E;

  /// Builds a success-state expected from a value.
  Expected(T v) noexcept(std::is_nothrow_move_constructible_v<T>)  // NOLINT(google-explicit-constructor)
      : payload_(std::in_place_index<0>, std::move(v)) {}

  /// Builds an error-state expected from an error value.
  Expected(E e) noexcept(std::is_nothrow_move_constructible_v<E>)  // NOLINT(google-explicit-constructor)
      : payload_(std::in_place_index<1>, std::move(e)) {}

  /// Named factory for a success-state expected.
  static Expected<T, E> Ok(T v) { return Expected<T, E>(std::move(v)); }

  /// Named factory for an error-state expected.
  static Expected<T, E> Err(E e) { return Expected<T, E>(std::move(e)); }

  /// Returns true when the expected carries a value (success state).
  bool has_value() const noexcept { return payload_.index() == 0; }

  /// Truthy iff `has_value()`.
  explicit operator bool() const noexcept { return has_value(); }

  /// Accesses the held value. Aborts if the expected is in the error state.
  const T& value() const& {
    FM_CHECK(has_value(), "Expected::value() on error-state");
    return std::get<0>(payload_);
  }

  /// Accesses the held value on an rvalue. Aborts if in the error state.
  T& value() & {
    FM_CHECK(has_value(), "Expected::value() on error-state");
    return std::get<0>(payload_);
  }

  /// Moves the held value out of the expected.
  ///
  /// The caller assumes ownership. In debug builds the "taken" state is
  /// tracked so a second invocation aborts rather than silently returning a
  /// moved-from value.
  T take() {
    FM_CHECK(has_value(), "Expected::take() on error-state");
#if !defined(NDEBUG)
    FM_CHECK(!consumed_, "Expected::take() called twice");
    consumed_ = true;
#endif
    return std::move(std::get<0>(payload_));
  }

  /// Accesses the held error. Aborts if the expected holds a value.
  const E& error() const& {
    FM_CHECK(!has_value(), "Expected::error() on value-state");
    return std::get<1>(payload_);
  }

  /// Accesses the held error on an rvalue. Aborts if in the value state.
  E& error() & {
    FM_CHECK(!has_value(), "Expected::error() on value-state");
    return std::get<1>(payload_);
  }

  /// Returns `f(value())` wrapped in an `Expected<U, E>` on success, or
  /// forwards the error untouched on failure.
  template <class F>
  auto map(F&& f) const& -> Expected<std::invoke_result_t<F, const T&>, E> {
    using U = std::invoke_result_t<F, const T&>;
    if (has_value()) {
      return Expected<U, E>(std::forward<F>(f)(std::get<0>(payload_)));
    }
    return Expected<U, E>(std::get<1>(payload_));
  }

  /// Invokes `f(value())`, which must itself return an `Expected<U, E>`, and
  /// forwards its result. Forwards the error untouched on failure.
  template <class F>
  auto and_then(F&& f) const& -> std::invoke_result_t<F, const T&> {
    using ResultT = std::invoke_result_t<F, const T&>;
    if (has_value()) {
      return std::forward<F>(f)(std::get<0>(payload_));
    }
    return ResultT(std::get<1>(payload_));
  }

  /// Recovers from an error by invoking `f(error())`, which must return an
  /// `Expected<T, E>`. On success the original value is forwarded through.
  template <class F>
  Expected<T, E> or_else(F&& f) const& {
    if (has_value()) {
      return Expected<T, E>(std::get<0>(payload_));
    }
    return std::forward<F>(f)(std::get<1>(payload_));
  }

 private:
  std::variant<T, E> payload_;
#if !defined(NDEBUG)
  mutable bool consumed_ = false;
#endif
};

/// Specialization for APIs that signal success without a payload.
///
/// Construct success via the nullary `Ok()` factory, failure via `Err(E)`.
template <class E>
class Expected<void, E> {
 public:
  using value_type = void;
  using error_type = E;

  /// Builds a success-state expected.
  Expected() noexcept : has_value_(true) {}

  /// Builds an error-state expected from an error value.
  Expected(E e) noexcept(std::is_nothrow_move_constructible_v<E>)  // NOLINT(google-explicit-constructor)
      : error_(std::move(e)), has_value_(false) {}

  /// Named factory for a success-state expected.
  static Expected<void, E> Ok() { return Expected<void, E>(); }

  /// Named factory for an error-state expected.
  static Expected<void, E> Err(E e) { return Expected<void, E>(std::move(e)); }

  /// Returns true when the expected is in the success state.
  bool has_value() const noexcept { return has_value_; }

  /// Truthy iff `has_value()`.
  explicit operator bool() const noexcept { return has_value_; }

  /// Accesses the held error. Aborts if the expected is in the success state.
  const E& error() const& {
    FM_CHECK(!has_value_, "Expected<void>::error() on value-state");
    return error_;
  }

  /// Accesses the held error on an rvalue. Aborts if in the success state.
  E& error() & {
    FM_CHECK(!has_value_, "Expected<void>::error() on value-state");
    return error_;
  }

 private:
  E error_{};
  bool has_value_ = false;
};

}  // namespace formulon

#endif  // FORMULON_UTILS_EXPECTED_H_
