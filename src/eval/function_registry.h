// Copyright 2026 libraz. Licensed under the MIT License.
//
// Function dispatch table for the Formulon calc engine.
//
// `FunctionRegistry` is a case-insensitive lookup keyed by canonical
// UPPERCASE name. Every Formulon built-in (and every host-registered
// extension) is described by a `FunctionDef`. The evaluator pre-evaluates
// each argument and short-circuits on the left-most error before invoking
// the entry's `impl`; the callee remains responsible for any per-argument
// coercion via the helpers in `eval/coerce.h`.
//
// See `backup/plans/02-calc-engine.md` §2.5.1 for the design rationale.

#ifndef FORMULON_EVAL_FUNCTION_REGISTRY_H_
#define FORMULON_EVAL_FUNCTION_REGISTRY_H_

#include <cstddef>
#include <cstdint>
#include <limits>
#include <memory>
#include <string_view>

#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {

/// Sentinel for `FunctionDef::max_arity` denoting an unbounded variadic.
inline constexpr std::uint32_t kVariadic = std::numeric_limits<std::uint32_t>::max();

/// Built-in function entry. Every Formulon function (and every host-registered
/// extension) is described by a `FunctionDef` and looked up by its canonical
/// UPPERCASE name.
struct FunctionDef {
  /// Canonical UPPERCASE name (e.g. `"SUM"`). Must be ASCII.
  std::string_view canonical_name;
  /// Inclusive minimum number of arguments accepted.
  std::uint32_t min_arity;
  /// Inclusive maximum number of arguments. Use `kVariadic` for unlimited.
  std::uint32_t max_arity;
  /// Implementation. The dispatcher pre-evaluates every argument and
  /// short-circuits on the left-most error before invoking `impl`. The
  /// callee is responsible for any further per-argument coercion.
  Value (*impl)(const Value* args, std::uint32_t arity, Arena& arena);
};

/// Case-insensitive function lookup table. Names are stored UPPERCASE
/// internally; lookup uppercases ASCII input bytes before comparing. Insertion
/// order is not preserved.
class FunctionRegistry {
 public:
  FunctionRegistry();
  ~FunctionRegistry();

  FunctionRegistry(const FunctionRegistry&) = delete;
  FunctionRegistry& operator=(const FunctionRegistry&) = delete;
  FunctionRegistry(FunctionRegistry&&) noexcept;
  FunctionRegistry& operator=(FunctionRegistry&&) noexcept;

  /// Registers `def` keyed by its canonical name. Returns false if a function
  /// with the same name was already registered (no overwrite).
  bool register_function(const FunctionDef& def);

  /// Returns a pointer to the registered definition for `name`, or nullptr if
  /// no such function exists. `name` is normalized to uppercase ASCII before
  /// lookup; non-ASCII bytes are compared verbatim.
  const FunctionDef* lookup(std::string_view name) const noexcept;

  /// Returns the number of registered functions.
  std::size_t size() const noexcept;

 private:
  // PIMPL via unique_ptr<unordered_map<...>> to keep the header dep-light.
  struct Impl;
  std::unique_ptr<Impl> impl_;
};

/// Returns the process-wide default registry, populated with all Formulon
/// built-ins. Thread-safe lazy initialization (Meyers singleton). The returned
/// reference has program lifetime; callers must not free or mutate it.
const FunctionRegistry& default_registry();

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_FUNCTION_REGISTRY_H_
