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
  /// When true (the default, matching Excel for the vast majority of
  /// functions), the dispatcher in `tree_walker` short-circuits on the
  /// left-most argument that evaluates to an error and never invokes the
  /// `impl` callback. When false, errors are passed through as raw `Value`
  /// arguments - required by the IS* type-predicate family
  /// (`ISERROR`, `ISERR`, `ISNA`, `ISNUMBER`, `ISTEXT`, `ISBLANK`,
  /// `ISLOGICAL`), which must be able to inspect error-typed inputs.
  bool propagate_errors = true;

  /// When true, the dispatcher in `tree_walker` expands any argument whose
  /// AST is `NodeKind::RangeOp` by walking the referenced rectangle on the
  /// bound sheet and flattening the cell values (row-major) into the args
  /// vector. When false, a RangeOp argument evaluates to `#VALUE!` as
  /// before. Only simple literal ranges (Ref:Ref) are expanded; any other
  /// RangeOp shape (e.g. INDIRECT-produced) surfaces as `#REF!`.
  bool accepts_ranges = false;

  /// Provenance-aware filter: when set, only Number-typed values flowing in
  /// through a RangeOp (or range-like) argument survive flattening; Bool,
  /// Text, and Blank cells inside the range are dropped before reaching the
  /// impl. Direct scalar arguments are unaffected and continue to coerce
  /// through the normal rules. Mirrors Excel's behaviour for SUM / AVERAGE /
  /// MIN / MAX / PRODUCT, where a bool cell sourced from a range is skipped
  /// silently while a bool passed directly coerces to 1 / 0.
  bool range_filter_numeric_only = false;

  /// Provenance-aware filter for AND / OR: when set, Number and Bool values
  /// flowing in through a RangeOp argument survive (numbers coerce to bool
  /// via 0 / non-zero), and Text / Blank cells inside the range are
  /// dropped silently. Direct scalar arguments still flow through the
  /// dispatcher's normal coercion, so a text literal passed directly
  /// surfaces as #VALUE!.
  bool range_filter_bool_coercible = false;

  /// Provenance-aware transform for the "A" family
  /// (AVERAGEA / MAXA / MINA / VARA / VARPA / STDEVA / STDEVPA): when set,
  /// range-sourced cells are coerced to numbers before reaching the impl --
  /// Bool becomes 0 / 1, Text (including the empty string) becomes 0, Blank
  /// cells are dropped, and Number cells are kept unchanged. Direct scalar
  /// arguments pass through unchanged (so a direct TRUE still coerces to 1
  /// via the impl's own `coerce_to_number`, and a direct non-numeric text
  /// still surfaces `#VALUE!`).
  bool range_filter_a_coerce = false;

  /// Mac-aligned policy for arguments that resolve to `Value::blank()`.
  /// Only applies to scalar (non-range, non-array-literal, non-OFFSET-call)
  /// arg slots; cells inside an expanded RangeOp or `OFFSET(...)` call are
  /// always coerced to 0 regardless of this setting (matching Excel's
  /// "blank in a range = 0" rule for SUM/AVERAGE/GCD-range/etc.).
  ///
  ///   * `Allow` (default) - Blank coerces to 0 via `coerce_to_number`.
  ///   * `RejectLiteralEmpty` - Only the literal-empty slot (parser-injected
  ///     `Literal(blank)` for `FN(,X)` / `FN(X,)` / `FN(,)`) surfaces
  ///     `blank_scalar_error`. A `Ref` to a blank cell still coerces to 0.
  ///     Used by `MROUND`, where Mac returns `#N/A` for `=MROUND(,5)` but
  ///     `0` for `=MROUND(A1,B1)` with A1/B1 blank.
  ///   * `RejectAnyScalar` - Both literal-empty slots AND `Ref`-to-blank
  ///     surface `blank_scalar_error`. Used by `GCD` / `LCM`, where Mac
  ///     surfaces `#VALUE!` for `=GCD(A1,B1,C1)` (all blank refs) but
  ///     returns `0` for `=GCD(A1:C1)` (range form, same blank cells).
  enum class BlankScalarPolicy : std::uint8_t {
    Allow = 0,
    RejectLiteralEmpty = 1,
    RejectAnyScalar = 2,
  };
  BlankScalarPolicy blank_scalar_policy = BlankScalarPolicy::Allow;
  /// Error code surfaced when `blank_scalar_policy` fires. Ignored when
  /// the policy is `Allow`. Defaults to `Value` purely so the field has a
  /// well-defined value; consumers must set this together with the policy.
  ErrorCode blank_scalar_error = ErrorCode::Value;
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

  /// Invokes `cb(name, ctx)` once for each registered function, where `name`
  /// is the canonical UPPERCASE key under which the entry was stored. Order
  /// is unspecified; callbacks must not mutate the registry. Exposed
  /// primarily for drift-detection tooling (see
  /// `tests/unit/registry_catalog_test.cpp` and `tools/catalog/status.py`).
  void for_each_name(void (*cb)(std::string_view name, void* ctx), void* ctx) const;

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
