//
// Lazy-dispatch seam between `tree_walker.cpp` and the per-family lazy-impl
// translation units.
//
// Why this header exists: the dispatch table itself (a constexpr array
// mapping canonical UPPERCASE function names to `LazyImpl` pointers) used
// to live at the top of `tree_walker.cpp`. The table forced
// `tree_walker.cpp` to `#include` every per-family header (30+ entries),
// which made the evaluator's main TU a hot rebuild target whenever a
// single new lazy function was added. By moving the table into its own
// TU (`tree_walker_lazy_table.cpp`) and exposing only the two narrow
// hooks below, adding a new lazy entry only rebuilds:
//   - the family TU that owns the new impl, and
//   - `tree_walker_lazy_table.cpp`.
// `tree_walker.cpp` itself is untouched.
//
// Both functions are stateless and thread-safe. `find_lazy_impl` is
// case-insensitive on ASCII (matches `case_insensitive_eq` semantics).

#ifndef FORMULON_EVAL_TREE_WALKER_LAZY_TABLE_H_
#define FORMULON_EVAL_TREE_WALKER_LAZY_TABLE_H_

#include <cstdint>
#include <string_view>

#include "eval/function_registry.h"
#include "eval/lazy_impls.h"

namespace formulon {
namespace eval {

/// Result-shape policy used by the partial-recalculation potential-producer
/// index.  Lazy forms are kept in the same table as their dispatch function
/// so shape metadata cannot silently drift away from evaluation semantics.
enum class LazyResultShape : std::uint8_t {
  kNotLazy,
  kScalar,
  kBroadcast,
  kArray,
  kReduce,
};

/// What a lazy form's call may return when it is a reference.
enum class RefResult : std::uint8_t {
  kNone,           ///< Never a reference.
  kBuilds,         ///< Builds a reference from its arguments (OFFSET, INDIRECT).
  kArg0,           ///< Returns a part of its first argument's reference (INDEX, TAKE, ...).
  kTrimmed,        ///< Rebuilds its first argument from its populated extent (TRIMRANGE).
  kProjected,      ///< Projects its argument onto the formula cell (SINGLE).
  kIfBranch,       ///< The chosen branch of IF.
  kChooseValue,    ///< The chosen value of CHOOSE.
  kIfsValue,       ///< The chosen value of IFS.
  kSwitchResult,   ///< The matched result of SWITCH.
  kXlookupResult,  ///< The return array element of XLOOKUP.
};

/// A set of parameter positions (0-based) of a lazy form.
enum class ParamSet : std::uint8_t {
  kNone,
  kFirst,           ///< 0
  kSecond,          ///< 1
  kFirstAndSecond,  ///< 0, 1
  kFirstAndThird,   ///< 0, 2
  kAllButFirst,     ///< 1, 2, ...
  kEven,            ///< 0, 2, 4, ...
  kEvenFromThird,   ///< 2, 4, 6, ...
  kOdd,             ///< 1, 3, 5, ...
  kFirstAndOdd,     ///< 0, 1, 3, 5, ...
  kFourthOn,        ///< 3, 4, ...
  kFourthOnByForm,  ///< 3, 4, ... when the first argument is a function number 1-19, else none (resolved per call)
  kAll,             ///< every position
};

/// True when position `index` (0-based) belongs to `set`.
constexpr bool param_in_set(ParamSet set, std::uint32_t index) noexcept {
  switch (set) {
    case ParamSet::kNone:
      return false;
    case ParamSet::kFirst:
      return index == 0U;
    case ParamSet::kSecond:
      return index == 1U;
    case ParamSet::kFirstAndSecond:
      return index <= 1U;
    case ParamSet::kFirstAndThird:
      return index == 0U || index == 2U;
    case ParamSet::kAllButFirst:
      return index >= 1U;
    case ParamSet::kEven:
      return index % 2U == 0U;
    case ParamSet::kEvenFromThird:
      return index >= 2U && index % 2U == 0U;
    case ParamSet::kOdd:
      return index % 2U == 1U;
    case ParamSet::kFirstAndOdd:
      return index == 0U || index % 2U == 1U;
    case ParamSet::kFourthOn:
    case ParamSet::kFourthOnByForm:
      return index >= 3U;
    case ParamSet::kAll:
      return true;
  }
  return false;
}

/// How a lone value (neither a reference nor an array) in a scalar-policy
/// parameter is answered.
enum class ScalarArg : std::uint8_t {
  kAccept,
  kRejectValue,          ///< #VALUE!
  kRejectNA,             ///< #N/A
  kRejectNum,            ///< #NUM!
  kRejectDiv0,           ///< #DIV/0!
  kRejectNAOrTextValue,  ///< #N/A, or #VALUE! for text
  kRejectTextValue,      ///< #VALUE! for text; any other lone value is read
};

/// `LazyEntry::min_arity` / `max_arity` value of an open argument list.
inline constexpr std::uint8_t kLazyVariadic = 0xFF;

/// One row of the lazy dispatch table.
struct LazyEntry {
  const char* name;  // canonical UPPERCASE
  LazyImpl impl;
  // Unannotated lazy forms are conservatively array-capable. A new lazy
  // implementation must opt into kScalar/kReduce/kBroadcast only after its
  // return paths have been checked; this prevents a catalog addition from
  // creating a silent partial-recalc false negative.
  LazyResultShape shape = LazyResultShape::kArray;
  // Written argument counts Excel accepts at entry, as `FunctionDef`'s
  // min_arity / max_arity (`kLazyVariadic` for an open list). The impl still
  // checks its own arguments.
  std::uint8_t min_arity;
  std::uint8_t max_arity;
  RefResult ref_result = RefResult::kNone;
  ParamSet ref_params = ParamSet::kNone;   // parameters that take only a reference
  ParamSet grid_params = ParamSet::kNone;  // parameters read through `grid_sel`
  GridArg grid_sel = GridArg::kElementError;
  GridArg grid_rest = GridArg::kElementError;
  ParamSet scalar_params = ParamSet::kNone;  // parameters under the lone-value policy
  ScalarArg scalar_error = ScalarArg::kAccept;
  ParamSet lift_params = ParamSet::kNone;  // value parameters evaluated once to join the lift
};

static_assert(sizeof(LazyEntry) == (sizeof(void*) == 8 ? 32 : 20), "LazyEntry must not grow");

/// Returns the row for `name` (canonical UPPERCASE, case-insensitive), or
/// nullptr when it is not a lazy form.
const LazyEntry* find_lazy_entry(std::string_view name) noexcept;

/// Looks up `name` (canonical UPPERCASE, case-insensitive) in the lazy
/// dispatch table. Returns the bound `LazyImpl` pointer on hit, or
/// `nullptr` when the name is not registered as a lazy form. Callers
/// invoke the returned pointer directly with the `Call` AST node + the
/// usual `Arena` / `FunctionRegistry` / `EvalContext` triple.
LazyImpl find_lazy_impl(std::string_view name) noexcept;

/// Returns the result-shape policy for a lazy form, or `kNotLazy` when the
/// name is not present in the dispatch table.
LazyResultShape find_lazy_result_shape(std::string_view name) noexcept;

/// Stores the inclusive argument-count bounds of the lazy form `name` (the
/// meaning of `FunctionDef::min_arity` / `max_arity`, `kVariadic` for an
/// open list) and returns true, or returns false when `name` is not a lazy
/// form.
bool find_lazy_arity(std::string_view name, std::uint32_t* min_arity, std::uint32_t* max_arity) noexcept;

/// The canonical UPPERCASE spelling of the lazy form `name`, in static
/// storage, or an empty view when `name` is not a lazy form.
std::string_view find_lazy_canonical_name(std::string_view name) noexcept;

/// Returns a nullptr-terminated array of canonical UPPERCASE names in
/// the lazy dispatch table. Storage has static duration; callers must
/// not free the array. Used by `tree_walker.cpp::lazy_form_names`
/// (declared in `tree_walker.h`) to expose the routing decisions to
/// catalog/diagnostics consumers.
const char* const* lazy_table_names();

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_TREE_WALKER_LAZY_TABLE_H_
