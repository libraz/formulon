
#ifndef FORMULON_EVAL_DYNAMIC_ARRAY_LIMITS_H_
#define FORMULON_EVAL_DYNAMIC_ARRAY_LIMITS_H_

#include <cstddef>

namespace formulon::eval {

// Shared dynamic-array allocation ceiling. This bounds producers before they
// materialise their row-major Value buffers and keeps Excel-grid-shaped calls
// such as MAKEARRAY(1048576, 2, ...) from becoming unbounded allocations.
inline constexpr std::size_t kMaxSequenceCells = 1U << 20;

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_DYNAMIC_ARRAY_LIMITS_H_
