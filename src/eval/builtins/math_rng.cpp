// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Formulon's random-number scalar built-ins: RAND and
// RANDBETWEEN. Both are volatile in Excel's sense -- they re-evaluate on
// every call. The engine achieves this for free because `FunctionDef` has
// no cached-result slot; the dispatcher re-runs `impl` on each reference.
//
// RNG storage: a single `thread_local std::mt19937_64` lazily seeded from
// `std::random_device` the first time it is touched on each thread. The
// engine is reachable from multiple host threads via the WASM binding;
// per-thread state avoids contention and guarantees that independent
// sheets/threads pull from independent sequences without a global mutex.

#include "eval/builtins/math_rng.h"

#include <cmath>
#include <cstdint>
#include <random>

#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Per-thread 64-bit Mersenne Twister. Lazily seeded from `std::random_device`
// on first touch. Encapsulated behind a function to guarantee correct
// thread-local initialisation order across translation units and to avoid
// leaking RNG state into the public header.
std::mt19937_64& rng_for_current_thread() {
  thread_local std::mt19937_64 rng{std::random_device{}()};
  return rng;
}

/// RAND(). Returns a uniform random double in the half-open interval
/// [0.0, 1.0). Zero arguments. Volatile: each call draws a fresh sample
/// from the thread-local Mersenne Twister.
Value Rand_(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  std::uniform_real_distribution<double> dist(0.0, 1.0);
  return Value::number(dist(rng_for_current_thread()));
}

/// RANDBETWEEN(bottom, top). Returns a uniform random integer in the closed
/// interval `[ceil(bottom), floor(top)]`. Both arguments are coerced to
/// number (failure surfaces `#VALUE!`). Excel rounds `bottom` up and `top`
/// down AFTER coercion, so e.g. `RANDBETWEEN(3.2, 7.9)` draws from
/// `[4, 7]`. If `ceil(bottom) > floor(top)` the call surfaces `#NUM!`.
/// Volatile: each call draws a fresh sample from the thread-local
/// Mersenne Twister.
Value RandBetween_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto bottom = coerce_to_number(args[0]);
  if (!bottom) {
    return Value::error(bottom.error());
  }
  auto top = coerce_to_number(args[1]);
  if (!top) {
    return Value::error(top.error());
  }
  const double lo_d = std::ceil(bottom.value());
  const double hi_d = std::floor(top.value());
  if (std::isnan(lo_d) || std::isinf(lo_d) || std::isnan(hi_d) || std::isinf(hi_d)) {
    return Value::error(ErrorCode::Num);
  }
  if (lo_d > hi_d) {
    return Value::error(ErrorCode::Num);
  }
  const auto lo = static_cast<std::int64_t>(lo_d);
  const auto hi = static_cast<std::int64_t>(hi_d);
  std::uniform_int_distribution<std::int64_t> dist(lo, hi);
  return Value::number(static_cast<double>(dist(rng_for_current_thread())));
}

}  // namespace

void register_math_rng_builtins(FunctionRegistry& registry) {
  registry.register_function(FunctionDef{"RAND", 0u, 0u, &Rand_});
  registry.register_function(FunctionDef{"RANDBETWEEN", 2u, 2u, &RandBetween_});
}

}  // namespace eval
}  // namespace formulon
