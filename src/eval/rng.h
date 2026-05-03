// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Shared thread-local random-number engine used by Excel's volatile RNG
// builtins (RAND, RANDBETWEEN, RANDARRAY). Splitting the RNG out of the
// scalar `math_rng` TU lets the dynamic-array spilling RANDARRAY in
// `builtins/dynamic_array.cpp` draw from the same per-thread sequence as
// the scalar variants — matching Mac Excel's behaviour where consecutive
// RAND / RANDARRAY calls share a single deterministic stream per recalc.

#ifndef FORMULON_EVAL_RNG_H_
#define FORMULON_EVAL_RNG_H_

#include <random>

namespace formulon {
namespace eval {

/// Returns a reference to the current thread's 64-bit Mersenne Twister.
/// Lazily seeded from `std::random_device` on first touch. Per-thread
/// storage avoids contention and gives independent host threads
/// independent sequences without a global mutex; see
/// `builtins/math_rng.cpp` for the original rationale.
std::mt19937_64& thread_local_rng();

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_RNG_H_
