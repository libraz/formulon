// Conservative result-shape analysis for partial-recalc candidate discovery.
//
// The analysis deliberately answers only "may this formula produce an array
// result?". It never claims that a formula will spill at runtime; committed
// spill geometry remains the source of truth for derived dependency edges.
// Keeping this predicate separate from evaluation lets the partial engine
// admit clean potential producers without enumerating every stored cell.

#ifndef FORMULON_EVAL_SPILL_POTENTIAL_H_
#define FORMULON_EVAL_SPILL_POTENTIAL_H_

#include <cstdint>

#include "parser/ast.h"

namespace formulon::eval {

class FunctionRegistry;

/// Static result-shape knowledge used by the partial-recalc potential index.
/// `kNeedsRegistry` means the formula is sensitive to the FunctionRegistry
/// supplied to the eventual recalc call (for example a host UDF that reuses
/// an eager built-in name). Such formulas stay in the index and are checked
/// again against that runtime registry before entering a closure.
enum class SpillPotential : std::uint8_t {
  kNever,
  kNeedsRegistry,
  kMaySpill,
};

/// Classifies a formula without a caller registry. Known lazy/intrinsic
/// shapes are resolved statically; eager calls retain `kNeedsRegistry` so a
/// custom runtime registry cannot turn a provably scalar default call into a
/// missed array producer.
SpillPotential spill_potential(const parser::AstNode& root) noexcept;

/// Classifies a formula using the registry that will perform evaluation.
/// Unknown calls remain conservative (`kMaySpill`) so custom/UDF producers
/// cannot disappear from a partial closure.
SpillPotential spill_potential(const parser::AstNode& root, const FunctionRegistry& registry) noexcept;

/// Returns true when `root` may evaluate to a two-dimensional array. The
/// result is intentionally conservative for names, lambdas, structured
/// references, and unknown function calls so custom/UDF array producers are
/// not omitted from a partial closure.
bool may_produce_spill(const parser::AstNode& root) noexcept;

/// Registry-aware variant. Registered built-ins use their declared shape;
/// unknown calls remain conservatively array-capable. This lets the engine
/// distinguish known scalar calls such as NOW() from custom/UDF calls while
/// preserving the no-registry convenience predicate above.
bool may_produce_spill(const parser::AstNode& root, const FunctionRegistry& registry) noexcept;

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_SPILL_POTENTIAL_H_
