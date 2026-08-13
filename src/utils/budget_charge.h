//
// One shared "charge, then allocate" seam for the resource ceilings declared
// in `utils/resource_budget.h`.
//
// Every call site that sizes a container from attacker-controlled input owes
// the same three steps in the same order: compute the count in 64-bit
// arithmetic, charge it to a `ResourceBudget`, and only then narrow it to
// `size_t` and reserve. Re-deriving that order per seam is how a step gets
// skipped, so the order lives here once and the call sites spell out only
// their own context string.

#ifndef FORMULON_UTILS_BUDGET_CHARGE_H_
#define FORMULON_UTILS_BUDGET_CHARGE_H_

#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <string_view>
#include <utility>

#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"

namespace formulon {

/// Charges `count` work units to `budget`, prepending `context` to the
/// budget's own `used/requested/ceiling` diagnostics when the ceiling is hit.
/// The caller's context is what makes an exceeded ceiling attributable to a
/// specific record, anchor or formula rather than to the engine at large.
inline Expected<void, Error> charge(ResourceBudget& budget, std::uint64_t count, std::string context) {
  auto charged = budget.consume(count);
  if (charged) {
    return Expected<void, Error>::Ok();
  }
  const std::string_view budget_context = charged.error().context;
  if (!context.empty() && !budget_context.empty()) {
    context.push_back(' ');
  }
  context.append(budget_context);
  Error error = charged.error();
  error.context = std::move(context);
  return error;
}

/// Charges `count` to `budget` and, only once the charge succeeded, reserves
/// that many elements in `container`.
///
/// Reserving before charging is the bug this exists to prevent: a rejected
/// count has already committed its allocation by then, so the ceiling bounds
/// only the work that follows the reservation and not the reservation itself.
/// Charging first also makes the narrowing to `size_t` safe, because the
/// accepted count is bounded by a ceiling the caller sized to fit `size_t`.
template <class Container>
Expected<void, Error> charge_then_reserve(ResourceBudget& budget, Container& container, std::uint64_t count,
                                          std::string context) {
  if constexpr (sizeof(std::size_t) < sizeof(std::uint64_t)) {
    // A ceiling above `SIZE_MAX` cannot be honoured by a 32-bit `size_t`
    // container (the WASM main build), so reject before either step runs.
    if (count > static_cast<std::uint64_t>(std::numeric_limits<std::size_t>::max())) {
      return make_error(FormulonErrorCode::kSecResourceLimit, "reserve count exceeds addressable size",
                        context + " requested=" + std::to_string(count));
    }
  }
  auto charged = charge(budget, count, std::move(context));
  if (!charged) {
    return charged.error();
  }
  container.reserve(static_cast<std::size_t>(count));
  return Expected<void, Error>::Ok();
}

}  // namespace formulon

#endif  // FORMULON_UTILS_BUDGET_CHARGE_H_
