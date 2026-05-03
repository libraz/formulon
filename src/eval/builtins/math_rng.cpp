// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of Formulon's random-number scalar built-ins: RAND and
// RANDBETWEEN. Both are volatile in Excel's sense -- they re-evaluate on
// every call. The engine achieves this for free because `FunctionDef` has
// no cached-result slot; the dispatcher re-runs `impl` on each reference.
//
// The thread-local Mersenne Twister itself lives in `eval/rng.{h,cpp}` so
// that the dynamic-array RANDARRAY impl in `builtins/dynamic_array.cpp`
// can pull from the same per-thread sequence as the scalar variants.

#include "eval/builtins/math_rng.h"

#include <cmath>
#include <cstdint>
#include <random>

#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "eval/rng.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

/// RAND(). Returns a uniform random double in the half-open interval
/// [0.0, 1.0). Zero arguments. Volatile: each call draws a fresh sample
/// from the thread-local Mersenne Twister.
Value Rand_(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  std::uniform_real_distribution<double> dist(0.0, 1.0);
  return Value::number(dist(thread_local_rng()));
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
  return Value::number(static_cast<double>(dist(thread_local_rng())));
}

}  // namespace

void register_math_rng_builtins(FunctionRegistry& registry) {
  registry.register_function(FunctionDef{"RAND", 0u, 0u, &Rand_});
  registry.register_function(FunctionDef{"RANDBETWEEN", 2u, 2u, &RandBetween_});
}

}  // namespace eval
}  // namespace formulon
