//
// Function-call dispatch path of the tree-walk evaluator. Split out of
// `tree_walker.cpp` to keep the recursive walker compile unit small;
// see `tree_walker/dispatch.h` for the public contract.
//
// Lazy entries (`IF`, `IFERROR`, `IFNA`, the `*IF`/`*IFS` aggregators,
// the lookup family, ...) are routed through the central lazy dispatch
// table in `tree_walker_lazy_table.cpp`; each impl owns its own arity
// check and chooses which subtrees to evaluate.
//
// All other names are routed through `FunctionRegistry`:
//   * Unknown name -> #NAME?
//   * Arity violation -> #VALUE!
//   * Otherwise every argument is pre-evaluated in order; by default the
//     left-most error short-circuits before the impl runs, but an entry
//     whose `propagate_errors` flag is `false` (the IS* type-predicate
//     family) opts out and receives raw error values among its arguments.
//
// Range-aware `accepts_ranges` entries expand range-shaped arguments
// (RangeOp, SpillRef, Ref3D, IntersectOp, StructuredRef, ArrayLiteral,
// and range-producing calls like OFFSET / CHOOSE / IF / ROW / COLUMN)
// into a flat values vector before invoking the impl. The expansion
// applies the per-cell `range_filter_*` rules so blank / text / bool
// cells are dropped or coerced consistently across input shapes.

#include "eval/tree_walker/dispatch.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/coerce.h"
#include "eval/declared_rect.h"
#include "eval/defined_name_resolve.h"
#include "eval/dynamic_array/anchor.h"
#include "eval/error_origin.h"
#include "eval/eval_context.h"
#include "eval/external_ref.h"
#include "eval/function_registry.h"
#include "eval/lambda_value.h"
#include "eval/lazy_impls.h"
#include "eval/name_env.h"
#include "eval/name_env_resolve.h"
#include "eval/omitted_arg.h"
#include "eval/range_args.h"
#include "eval/range_resolvers.h"
#include "eval/reference/deref.h"
#include "eval/scalar_args.h"
#include "eval/tail_array.h"
#include "eval/tree_walker/broadcast.h"
#include "eval/tree_walker/depth_guard.h"
#include "eval/tree_walker/dispatch_internal.h"
#include "eval/tree_walker_lazy_table.h"
#include "parser/ast.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {

std::string_view strip_future_prefix(std::string_view name) noexcept {
  constexpr std::string_view kXlws = "_xlfn._xlws.";
  constexpr std::string_view kXlfn = "_xlfn.";
  if (name.size() > kXlws.size() && strings::case_insensitive_eq(name.substr(0, kXlws.size()), kXlws)) {
    return name.substr(kXlws.size());
  }
  if (name.size() > kXlfn.size() && strings::case_insensitive_eq(name.substr(0, kXlfn.size()), kXlfn)) {
    return name.substr(kXlfn.size());
  }
  return name;
}

bool is_reference_shape(const parser::AstNode& node) noexcept {
  switch (node.kind()) {
    case parser::NodeKind::Ref:
    case parser::NodeKind::RangeOp:
    case parser::NodeKind::UnionOp:
    case parser::NodeKind::IntersectOp:
    case parser::NodeKind::SpillRef:
    case parser::NodeKind::StructuredRef:
      return true;
    case parser::NodeKind::Ref3D:
      return !is_grouped_ref3d(node);
    default:
      return false;
  }
}

bool is_full_axis_range(const parser::Reference& lhs, const parser::Reference& rhs) {
  const Expected<DeclaredRect, ErrorCode> rect = declared_rect(lhs, rhs);
  return rect && (rect.value().rows() == Sheet::kMaxRows || rect.value().cols() == Sheet::kMaxCols);
}

namespace {

// A scalar (non-range-aware) function applied to its collected arguments:
// element-wise over an array or whole column / row argument, else once.
Value call_scalar_function(const FunctionDef& def, std::vector<Value>& values, std::vector<const TailArray*>& tails,
                           Arena& arena, const TailArray** tail_out) {
  if (!tails.empty()) {
    tails.resize(values.size(), nullptr);
    const Shaped result = broadcast_scalar_call_shaped(def, values, tails, arena);
    *tail_out = result.tail_array;
    return result.value;
  }
  std::uint32_t out_rows = 1;
  std::uint32_t out_cols = 1;
  bool any_array = false;
  for (const Value& v : values) {
    if (v.is_array()) {
      any_array = true;
      const ArrayValue* a = v.as_array();
      out_rows = std::max(out_rows, a->rows);
      out_cols = std::max(out_cols, a->cols);
    }
  }
  if (any_array) {
    return broadcast_scalar_call(def, values, out_rows, out_cols, arena);
  }
  return call_scalar_impl(def, values.data(), static_cast<std::uint32_t>(values.size()), arena);
}

// The `#VALUE!` of a text literal that failed ahead of an error argument, with the origin slot as it stood then.
Value literal_text_failure(const std::optional<ErrorOriginSnapshot>& origin) {
  if (origin.has_value()) {
    restore_error_origin(*origin);
  }
  return Value::error(ErrorCode::Value);
}

// The error a lone value (neither a reference nor an array) answers in a
// parameter under `policy`, or nullopt when it is accepted.
std::optional<ErrorCode> lone_value_error(ScalarArg policy, const Value& v) {
  if (v.is_ref() || v.is_array() || v.is_error()) {
    return std::nullopt;
  }
  switch (policy) {
    case ScalarArg::kAccept:
      return std::nullopt;
    case ScalarArg::kRejectValue:
      return ErrorCode::Value;
    case ScalarArg::kRejectNA:
      return ErrorCode::NA;
    case ScalarArg::kRejectNum:
      return ErrorCode::Num;
    case ScalarArg::kRejectDiv0:
      return ErrorCode::Div0;
    case ScalarArg::kRejectNAOrTextValue:
      return v.is_text() ? ErrorCode::Value : ErrorCode::NA;
    case ScalarArg::kRejectTextValue:
      return v.is_text() ? std::optional<ErrorCode>(ErrorCode::Value) : std::nullopt;
  }
  return std::nullopt;
}

// A node that names a reference where it is written, read by the impl itself.
bool is_written_reference(const parser::AstNode& node) {
  switch (node.kind()) {
    case parser::NodeKind::Ref3D:
    case parser::NodeKind::ExternalRef:
      return true;
    case parser::NodeKind::ImplicitIntersection:
      return is_reference_shape(node.as_implicit_intersection_operand());
    default:
      return is_reference_shape(node);
  }
}

// A node whose value is fixed where it is written: re-reading it reads no cell.
bool is_constant_node(const parser::AstNode& node) {
  return node.kind() == parser::NodeKind::Literal || node.kind() == parser::NodeKind::ArrayLiteral ||
         node.kind() == parser::NodeKind::ErrorLiteral;
}

// The class position `index` reads a grid with, `grid_params` resolved for the call.
GridArg grid_class(const LazyEntry& entry, ParamSet grid_params, std::uint32_t index) {
  return param_in_set(grid_params, index) ? entry.grid_sel : entry.grid_rest;
}

// `entry.grid_params` for one call: a form-dependent set selected when the
// call's evaluated first argument `first` is a function number 1-19.
ParamSet resolve_grid_params(const LazyEntry& entry, const Value& first) {
  if (entry.grid_params != ParamSet::kFourthOnByForm) {
    return entry.grid_params;
  }
  const double form = first.is_number() ? std::trunc(first.as_number()) : 0.0;
  return form >= 1.0 && form <= 19.0 ? ParamSet::kFourthOn : ParamSet::kNone;
}

// Whether any argument of `entry` is evaluated before its impl runs.
bool prepares_args(const LazyEntry& entry) {
  return entry.scalar_params != ParamSet::kNone || entry.grid_params != ParamSet::kNone ||
         entry.grid_rest != GridArg::kElementError || entry.lift_params != ParamSet::kNone;
}

// A value substituted for an argument; an evaluated blank is a blank cell, not
// an omitted slot (`COUNTIF(A1:A5,B2)` with B2 empty).
const parser::AstNode* value_literal(Arena& arena, const Value& v) {
  return parser::make_literal(
      arena, v.is_blank() && !v.blank_projects_to_zero() ? Value::blank(BlankGridProjection::kValueArrayZero) : v);
}

// One synthesized call of a lift: the prepared arguments, a lifted position
// taking its element.
struct LiftCall {
  const parser::AstNode& call;
  const LazyEntry& entry;
  const parser::AstNode* const* nodes;
  const std::uint8_t* lifted;
  Arena& arena;
  const FunctionRegistry& registry;
  const EvalContext& ctx;
};

Value run_lifted_call(const Value* elements, std::uint32_t arity, const void* state) {
  const LiftCall& lift = *static_cast<const LiftCall*>(state);
  std::vector<const parser::AstNode*> args(lift.nodes, lift.nodes + arity);
  for (std::uint32_t i = 0; i < arity; ++i) {
    if (lift.lifted[i] != 0U && (args[i] = value_literal(lift.arena, elements[i])) == nullptr) {
      return Value::error(ErrorCode::Num);
    }
  }
  const parser::AstNode* call = parser::make_call(lift.arena, lift.call.as_call_name(), args.data(), arity);
  if (call == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return run_lazy_call(*call, lift.entry, lift.arena, lift.registry, lift.ctx);
}

// A value that stands as one element of a grid: its first cell for an array.
Value element_of(const Value& v) {
  return v.is_array() ? v.as_array()->cells[0] : v;
}

// `results` (rows x cols) as one grid: a one-area reference is an element,
// an error an error element, any other reference #VALUE!, and any other value
// a value element.
Value collect_grid(const std::vector<Value>& results, std::uint32_t rows, std::uint32_t cols, Arena& arena) {
  std::vector<RefArea> areas(results.size());
  std::vector<std::uint16_t> books(results.size(), 0U);
  std::vector<Value> values(results.size(), Value::blank());
  bool other_book = false;
  bool has_values = false;
  for (std::size_t i = 0; i < results.size(); ++i) {
    const Value& v = results[i];
    if (v.is_ref() && v.ref_form() == RefForm::kArea) {
      const RefView view = v.as_ref();
      areas[i] = view.areas[0];
      books[i] = view.book;
      other_book = other_book || view.book != 0U;
    } else if (v.is_ref() || v.is_error()) {
      areas[i] = RefArea::error_element(v.is_error() ? v.as_error() : ErrorCode::Value);
    } else {
      areas[i] = RefArea::value_element();
      values[i] = element_of(v);
      has_values = true;
    }
  }
  return make_ref(arena, RefForm::kGrid, 0U, areas.data(), static_cast<std::uint32_t>(areas.size()), rows, cols,
                  other_book ? books.data() : nullptr, has_values ? values.data() : nullptr);
}

// The per-position results of a lift as one value: a grid when any is a
// reference (`collect_grid`), an array otherwise (an array result keeps its
// top-left element).
Value collect_lifted(const std::vector<Value>& results, std::uint32_t rows, std::uint32_t cols, Arena& arena) {
  if (std::any_of(results.begin(), results.end(), [](const Value& v) { return v.is_ref(); })) {
    return collect_grid(results, rows, cols, arena);
  }
  Value* cells = nullptr;
  ArrayValue* out = allocate_derived_array_value(rows, cols, arena, cells);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::size_t i = 0; i < results.size(); ++i) {
    cells[i] = element_of(results[i]);
  }
  return Value::array(out);
}

// The elements of the grids a reshaping call was given, in the order their
// `Value::grid_element` indices count them.
struct ReshapeSource {
  std::vector<RefArea> areas;
  std::vector<std::uint16_t> books;
  std::vector<Value> values;
};

// The grid `ref`'s elements appended to `source`, as an array of
// `Value::grid_element` indices of the grid's shape.
Value index_grid_elements(const Value& ref, ReshapeSource* source, Arena& arena) {
  const RefView view = ref.as_ref();
  Value* cells = nullptr;
  ArrayValue* out = allocate_derived_array_value(view.rows, view.cols, arena, cells);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  const std::size_t n = static_cast<std::size_t>(view.rows) * view.cols;
  for (std::size_t i = 0; i < n; ++i) {
    cells[i] = Value::grid_element(static_cast<std::uint32_t>(source->areas.size()));
    source->areas.push_back(view.areas[i]);
    source->books.push_back(view.element_books != nullptr ? view.element_books[i] : view.book);
    source->values.push_back(view.element_values != nullptr ? view.element_values[i] : Value::blank());
  }
  return Value::array(out);
}

// A reshaping call's result mapped back onto the elements of `source`: a grid
// of the result's shape, an index cell its element, any other cell a value
// element.
Value map_reshaped(const Value& result, const ReshapeSource& source, Arena& arena) {
  if (result.is_ref()) {
    return result;
  }
  const std::uint32_t rows = result.is_array() ? result.as_array()->rows : 1U;
  const std::uint32_t cols = result.is_array() ? result.as_array()->cols : 1U;
  const Value* cells = result.is_array() ? result.as_array()->cells : &result;
  const std::size_t n = static_cast<std::size_t>(rows) * cols;
  std::vector<RefArea> areas(n);
  std::vector<std::uint16_t> books(n, 0U);
  std::vector<Value> values(n, Value::blank());
  bool other_book = false;
  bool has_values = false;
  for (std::size_t i = 0; i < n; ++i) {
    const Value& v = cells[i];
    if (v.is_grid_element() && v.grid_element_index() < source.areas.size()) {
      const std::uint32_t k = v.grid_element_index();
      areas[i] = source.areas[k];
      books[i] = source.books[k];
      values[i] = source.values[k];
    } else if (v.is_error()) {
      areas[i] = RefArea::error_element(v.as_error());
    } else {
      areas[i] = RefArea::value_element();
      values[i] = v;
    }
    other_book = other_book || books[i] != 0U;
    has_values = has_values || areas[i].is_value_element();
  }
  return make_ref(arena, RefForm::kGrid, 0U, areas.data(), static_cast<std::uint32_t>(n), rows, cols,
                  other_book ? books.data() : nullptr, has_values ? values.data() : nullptr);
}

// Evaluates the arguments of a lazy call that `entry` reads before its impl:
// a lone value under the lone-value policy is replaced by the error the
// policy answers; a grid in a classed position is read by its class, and one
// in a `kPerRef` position, like an array in a lift position, lifts the call;
// one in a `kReshape` position is given as an array of element indices, its
// elements appended to `*reshape`. A constant, an omitted slot and a written
// reference stay in place. Returns the call to run, or null with `*result`
// holding the lifted call's value (or the error of an arena that cannot hold
// the new call).
[[gnu::noinline]] const parser::AstNode* prepare_lazy_args(const parser::AstNode& call, const LazyEntry& entry,
                                                           Arena& arena, const FunctionRegistry& registry,
                                                           const EvalContext& ctx, ReshapeSource* reshape,
                                                           Value* result) {
  const std::uint32_t arity = call.as_call_arity();
  std::vector<const parser::AstNode*> args(arity);
  for (std::uint32_t i = 0; i < arity; ++i) {
    args[i] = &call.as_call_arg(i);
  }
  std::vector<Value> lifting(arity, Value::blank());
  std::vector<std::uint8_t> lifted(arity, 0U);
  bool changed = false;
  bool lifts = false;
  auto replace = [&](std::uint32_t i, const Value& v) {
    args[i] = value_literal(arena, v);
    changed = true;
    return args[i] != nullptr;
  };
  ParamSet grid_params = entry.grid_params;
  if (grid_params == ParamSet::kFourthOnByForm && arity > 0U) {
    const parser::AstNode& first = call.as_call_arg(0);
    const Value v = is_constant_node(first) ? eval_node(first, arena, registry, ctx)
                                            : densify(eval_node_shaped(first, arena, registry, ctx), arena);
    grid_params = resolve_grid_params(entry, v);
    if (!is_constant_node(first) && !replace(0U, v)) {
      *result = Value::error(ErrorCode::Num);
      return nullptr;
    }
  }
  for (std::uint32_t i = 0; i < arity; ++i) {
    const parser::AstNode& arg = *args[i];
    if (is_omitted_arg(arg)) {
      continue;
    }
    if (param_in_set(entry.lift_params, i)) {
      const Value v = densify(eval_node_shaped(arg, arena, registry, ctx), arena);
      if (v.is_array()) {
        lifting[i] = v;
        lifted[i] = 1U;
        lifts = true;
      } else if (!is_constant_node(arg) && !replace(i, v)) {
        break;
      }
      continue;
    }
    const GridArg cls = grid_class(entry, grid_params, i);
    const bool scalar = param_in_set(entry.scalar_params, i);
    // Outside the lone-value policy only something that may be a reference can be a grid; anything else stays
    // in place, so a derived whole column keeps its declared size (`A:A&""`).
    if ((!scalar && (cls == GridArg::kElementError || !evaluates_to_reference(arg))) || is_written_reference(arg) ||
        arg.kind() == parser::NodeKind::ArrayLiteral || arg.kind() == parser::NodeKind::ErrorLiteral) {
      continue;
    }
    const bool constant = arg.kind() == parser::NodeKind::Literal;
    const Value v = constant ? arg.as_literal() : eval_node_ref(arg, arena, registry, ctx);
    if (v.is_ref() && v.ref_form() == RefForm::kGrid && cls != GridArg::kElementError) {
      if (cls == GridArg::kPerRef) {
        lifting[i] = v;
        lifted[i] = 1U;
        lifts = true;
        continue;
      }
      const bool kept = cls == GridArg::kKeepRef || cls == GridArg::kLambdaElement;
      const Value read =
          cls == GridArg::kReshape
              ? index_grid_elements(v, reshape, arena)
              : (kept ? v : materialise_grid(v, cls, nullptr, arena, registry, ctx, ReadCount::kCounted));
      if (!replace(i, read)) {
        break;
      }
      continue;
    }
    const std::optional<ErrorCode> rejected = scalar ? lone_value_error(entry.scalar_error, v) : std::nullopt;
    if (constant && !rejected) {
      continue;
    }
    if (!replace(i, rejected ? Value::error(*rejected) : v)) {
      break;
    }
  }
  if (std::find(args.begin(), args.end(), nullptr) != args.end()) {
    *result = Value::error(ErrorCode::Num);
    return nullptr;
  }
  if (lifts) {
    // Every other argument is read once, not once per element.
    for (std::uint32_t i = 0; i < arity; ++i) {
      const parser::AstNode& arg = *args[i];
      if (lifted[i] != 0U || is_constant_node(arg) || is_omitted_arg(arg)) {
        continue;
      }
      if (!replace(i, eval_node_ref(arg, arena, registry, ctx))) {
        *result = Value::error(ErrorCode::Num);
        return nullptr;
      }
    }
    std::uint32_t rows = 1U;
    std::uint32_t cols = 1U;
    lift_extent(lifting.data(), arity, nullptr, &rows, &cols);
    const LiftCall lift{call, entry, args.data(), lifted.data(), arena, registry, ctx};
    std::vector<Value> results(static_cast<std::size_t>(rows) * cols, Value::blank());
    lift_positions(lifting.data(), arity, nullptr, rows, cols, arena, &run_lifted_call, &lift, results.data());
    *result = collect_lifted(results, rows, cols, arena);
    return nullptr;
  }
  if (!changed) {
    return &call;
  }
  const parser::AstNode* prepared = parser::make_call(arena, call.as_call_name(), args.data(), arity);
  if (prepared == nullptr) {
    *result = Value::error(ErrorCode::Num);
  }
  return prepared;
}

// The body of `dispatch_call`. A scalar function applied over a whole column /
// row returns through `*tail_out` (the returned value is then unused).
Value dispatch_call_impl(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, const TailArray** tail_out) {
  const std::string_view name = strip_future_prefix(node.as_call_name());
  const std::uint32_t arity = node.as_call_arity();

  // Name-bound lambda dispatch: when the formula text reads `f(x)` and `f`
  // resolves to a runtime `LambdaValue` via the lexical name environment
  // (LET-bound or LAMBDA-parameter), invoke it as if the user had written
  // an explicit IIFE. The lookup runs *before* the registry / lazy table so
  // a LET binding can shadow a built-in name (matching Excel's semantics).
  // A name bound to a reference is `#REF!`, as calling a cell is; any other
  // bound non-Lambda value is `#VALUE!`; a bound error propagates verbatim.
  // Unbound names fall through to the existing registry path.
  if (const NameEnv* env = ctx.name_env(); env != nullptr) {
    if (const auto* binding = env->lookup(name); binding != nullptr) {
      if (binding->reference() != nullptr) {
        return Value::error(ErrorCode::Ref);
      }
      const auto read = [&](const Value& ref) { return deref_shaped(ref, arena, registry, ctx, ReadCount::kCounted); };
      const Value bound = NameEnv::binding_value(*binding, arena, read);
      if (bound.is_lambda()) {
        return invoke_lambda(bound.as_lambda(), arity, node.as_call_args(), arena, registry, ctx);
      }
      if (bound.is_error()) {
        return bound;
      }
      return Value::error(ErrorCode::Value);
    }
  }

  // A visible defined name has precedence over the lazy table and the
  // function registry. This is the call-shaped counterpart of NameRef
  // resolution: a definition may evaluate to a LambdaValue and is then
  // invoked with the ordinary AST-backed lambda path; an error value is
  // propagated, while a successfully evaluated non-lambda is not callable.
  // `resolve_defined_name` returns the LambdaValue after its expansion frame
  // has been removed, so a named lambda may recurse by resolving itself on
  // each nested call without being mistaken for a circular defined-name
  // reference.
  if (find_defined_name(ctx, name) != nullptr) {
    const Value defined = resolve_defined_name(name, arena, registry, ctx);
    if (defined.is_error()) {
      return defined;
    }
    if (!defined.is_lambda()) {
      return Value::error(ErrorCode::Value);
    }
    return invoke_lambda(defined.as_lambda(), arity, node.as_call_args(), arena, registry, ctx);
  }

  if (const LazyEntry* entry = find_lazy_entry(name); entry != nullptr) {
    return run_lazy_call(node, *entry, arena, registry, ctx);
  }

  const FunctionDef* def = lookup_on_current_host(registry, name);
  if (def == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  // The pre-expansion arity guards min_arity / max_arity. This happens to
  // align with Excel's behaviour for the range-aware aggregators:
  // `=SUM()` is rejected at parse time, and `=SUM(A1:A1)` passes the
  // `min_arity = 1` check even though its expansion might be empty (which
  // cannot happen with a finite valid rectangle today).
  if (arity < def->min_arity || arity > def->max_arity) {
    return Value::error(ErrorCode::Value);
  }

  // Pre-evaluate arguments left-to-right. A range-aware function stops at the
  // first error; a scalar one evaluates every argument and hands errors to
  // its impl, whose coercion reports the first failing argument
  // (`call_scalar_impl`). Functions that inspect error arguments (e.g.
  // `ISERROR`) clear `propagate_errors`. For a range-aware function, any
  // argument whose AST node is a simple RangeOp (Ref:Ref) is flattened into
  // the values vector in row-major order.
  std::vector<Value> values;
  values.reserve(arity);
  // Parallel to `values`: the whole column / row a scalar function's argument
  // evaluated to, whose `values` slot is then a placeholder.
  std::vector<const TailArray*> tails;
  // Derived whole column / row arguments of a reducer with `FoldOps`.
  std::vector<StreamedTail> streamed;
  // Tracks whether any argument slot was range-shaped (RangeOp / OFFSET-call
  // / ArrayLiteral). The deferred `RejectAllScalarsBlank` policy surfaces
  // #VALUE! for `=GCD(A1,B1,C1)` (all blank
  // scalar refs) when there is NO range-shaped arg in the call. The mixed
  // form `=GCD(A1:B1, C1)` over the same blank cells still returns 0,
  // because the range arg "rescues" the policy.
  bool had_range_shaped_arg = false;
  // Tracks whether a scalar reference resolved to a blank under the
  // `RejectAllScalarsBlank` policy. Such a reference is omitted from the
  // callee's values; the deferred error fires only when it was the entire
  // scalar argument set.
  bool saw_blank_scalar_ref = false;
  // Set once a direct text literal that does not coerce to a number has been
  // seen by a numeric or A-family aggregator.
  bool literal_text_not_numeric = false;
  // The origin slot when that literal failed: its #VALUE! is the left error and keeps that state.
  std::optional<ErrorOriginSnapshot> literal_text_origin;
  // Set when a `first_arg_scalar` / `last_arg_scalar` slot evaluated to an
  // array or grid; the call is then evaluated once per element.
  bool lift_scalar_slot = false;
  // A scalar function that propagates errors keeps an error argument in its
  // slot, as does a range-aware one that reads its values in order; the origin
  // of the one it reports is kept for the result.
  const bool in_order_ranges = reads_range_errors_in_order(*def);
  const bool passes_errors = def->propagate_errors && (!def->accepts_ranges || in_order_ranges);
  std::optional<ErrorOriginSnapshot> reported_origin;
  const auto pass_error = [&](const Value& err) {
    if (!reported_origin.has_value() || def->checks_args_last_first) {
      reported_origin = snapshot_error_origin();
    }
    values.push_back(err);
  };
  // Range cells and failed expansions reach `values` without `pass_error`; the
  // first such error keeps the origin left by the argument that produced it.
  const bool defers_cell_errors = defers_range_cell_errors(*def);
  std::size_t origin_scanned = 0;
  const auto note_collected_error = [&]() {
    if ((!in_order_ranges && !defers_cell_errors) || reported_origin.has_value()) {
      return;
    }
    for (; origin_scanned < values.size(); ++origin_scanned) {
      if (values[origin_scanned].is_error()) {
        reported_origin = snapshot_error_origin();
        return;
      }
    }
  };
  const std::uint32_t evaluated_arity =
      def->analysis_toolpak_args && !def->atp_omitted_optional_is_na
          ? atp_evaluated_arity(node, def->min_arity, def->max_arity)
          : (def->omitted_defaults != nullptr ? trim_absent_omitted_args(node, *def, arity) : arity);
  for (std::uint32_t i = 0; i < evaluated_arity; ++i) {
    note_collected_error();
    const parser::AstNode& raw_arg = node.as_call_arg(i);
    // An omitted slot with a declared default takes it; a blank reference does not.
    Value omitted_default = Value::blank();
    if (def->omitted_defaults != nullptr && is_omitted_arg(raw_arg) && omitted_arg_value(*def, i, &omitted_default)) {
      values.push_back(omitted_default);
      continue;
    }
    if ((def->last_arg_scalar && i > 0U && i + 1U == evaluated_arity) || (def->first_arg_scalar && i == 0U)) {
      const Value v = densify(eval_node_shaped(raw_arg, arena, registry, ctx), arena);
      if (v.is_error()) {
        if (def->propagate_errors) {
          return literal_text_not_numeric ? literal_text_failure(literal_text_origin) : v;
        }
        forget_error_origin();
      }
      std::uint32_t lift_rows = 0;
      std::uint32_t lift_cols = 0;
      if (lift_extent(&v, 1U, nullptr, &lift_rows, &lift_cols)) {
        lift_scalar_slot = true;
      } else if (def->analysis_toolpak_args) {
        const Value err = atp_arg_error(raw_arg, v, atp_slot_required(i, def->min_arity, def->max_arity));
        if (err.is_error()) {
          if (!absorb_arg_error(*def, err, &values)) {
            return err;
          }
          continue;
        }
      }
      values.push_back(v);
      continue;
    }
    // A name bound to a reference (`SUM(r)` with `r` bound by LET) is read as
    // that reference, so its cells take the range filters; a name bound to a
    // value flows through the scalar branch below.
    const parser::AstNode& arg_node = resolve_name_ast(raw_arg, ctx.name_env());
    Value range_arg_result = Value::blank();
    switch (collect_range_arg(*def, arg_node, arena, registry, ctx, values, had_range_shaped_arg, range_arg_result)) {
      case RangeArgOutcome::kCollected:
        continue;
      case RangeArgOutcome::kReturn:
        return range_arg_result;
      case RangeArgOutcome::kNotRange:
        break;
    }
    // A scalar function keeps a whole column / row at its declared size for
    // the element-wise broadcast below; a range-aware one takes it expanded.
    const Shaped shaped_arg = eval_node_shaped(arg_node, arena, registry, ctx);
    if (shaped_arg.tail_array != nullptr && !def->accepts_ranges) {
      tails.resize(values.size() + 1U, nullptr);
      tails[values.size()] = shaped_arg.tail_array;
      values.push_back(Value::blank());
      continue;
    }
    // A whole column / row read straight from a reference (a defined name
    // over `A:A`) passes only its populated head, as `SUM(A:A)` does; a
    // derived one (`A:A*1`) is expanded below.
    if (shaped_arg.tail_array != nullptr && shaped_arg.tail_array->from_reference) {
      had_range_shaped_arg = true;
      const TailArray& ta = *shaped_arg.tail_array;
      const std::size_t span = ta.axis == TailAxis::kRows ? ta.cols : ta.rows;
      Value range_err = Value::blank();
      if (!append_range_sourced_values(*def, ta.cells, static_cast<std::size_t>(ta.head) * span, &values, &range_err)) {
        return range_err;
      }
      continue;
    }
    // A reducer that takes values one at a time streams a derived one
    // (`-A:BB`), which can hold more cells than an array may.
    if (shaped_arg.tail_array != nullptr && def->accepts_ranges && def->fold != nullptr) {
      had_range_shaped_arg = true;
      if (def->propagate_errors) {
        const Value err = first_tail_error(*shaped_arg.tail_array);
        if (err.is_error()) {
          return err;
        }
      }
      streamed.push_back(StreamedTail{values.size(), shaped_arg.tail_array});
      continue;
    }
    Value v = densify(shaped_arg, arena);
    if (v.is_error() && passes_errors) {
      pass_error(v);
      continue;
    }
    // A cell reference in the array slot holds its error until the last value has coerced.
    if (v.is_error() && defers_cell_errors && arg_node.kind() == parser::NodeKind::Ref) {
      values.push_back(v);
      continue;
    }
    if (v.is_error()) {
      // A text literal that cannot be a number already failed ahead of this
      // error, and the earlier failure is the one reported.
      if (def->propagate_errors) {
        return literal_text_not_numeric ? literal_text_failure(literal_text_origin) : v;
      }
      // A function that inspects its error arguments consumes this one.
      forget_error_origin();
    }
    if ((def->range_filter_numeric_only || def->range_filter_a_coerce || defers_cell_errors) &&
        arg_node.kind() == parser::NodeKind::Literal && v.is_text() && !coerce_to_number(v)) {
      if (!literal_text_not_numeric) {
        literal_text_origin = snapshot_error_origin();
      }
      literal_text_not_numeric = true;
    }
    // Generic Array-result flatten for range-aware aggregators. Lazy
    // builtins (`ANCHORARRAY`, `SEQUENCE`, `TRANSPOSE`, `MUNIT`, ...)
    // return `Value::Array`; without this branch SUM / AVERAGE / MIN /
    // MAX would receive the Array as a single scalar slot and fail with
    // `#VALUE!` from `coerce_to_number`. Flattening row-major mirrors the
    // SpillRef / RangeOp branches above, applying the same provenance
    // filters so blank / text / bool cells are dropped or coerced
    // consistently. Calls returning a reference reach `continue` in
    // `collect_range_arg` before getting here, so this branch catches the
    // Array-returning calls.
    if (def->accepts_ranges && v.is_array()) {
      had_range_shaped_arg = true;
      const ArrayValue* arr = v.as_array();
      const std::size_t n = static_cast<std::size_t>(arr->rows) * static_cast<std::size_t>(arr->cols);
      Value range_err = Value::blank();
      if (!append_array_sourced_values(*def, arr->cells, n, &values, &range_err)) {
        return range_err;
      }
      continue;
    }
    // RangeOp / OFFSET-call / ArrayLiteral args were handled above and
    // `continue`'d, so reaching this point implies a scalar arg slot
    // (Literal, Ref, BinaryOp, ...). The Analysis-ToolPak rule fires
    // eagerly so the left-most offending slot wins.
    if (def->analysis_toolpak_args) {
      const bool required = def->atp_omitted_optional_is_na || atp_slot_required(i, def->min_arity, def->max_arity);
      const Value err = atp_arg_error(arg_node, v, required);
      if (err.is_error()) {
        if (passes_errors) {
          pass_error(err);
        } else if (!absorb_arg_error(*def, err, &values)) {
          return err;
        }
        continue;
      }
    }
    // `RejectAllScalarsBlank` (GCD / LCM) omits blank scalar references so
    // a blank mixed with a numeric literal or reference cannot become zero
    // inside the callee. Range-shaped blanks are handled above and remain
    // zero, preserving the range form's semantics.
    if (def->blank_scalar_policy == FunctionDef::BlankScalarPolicy::RejectAllScalarsBlank &&
        arg_node.kind() == parser::NodeKind::Ref && v.kind() == ValueKind::Blank) {
      saw_blank_scalar_ref = true;
      continue;
    }
    if (def->blank_scalar_policy == FunctionDef::BlankScalarPolicy::RejectAnyScalarBlank &&
        arg_node.kind() == parser::NodeKind::Ref && v.kind() == ValueKind::Blank) {
      return Value::error(def->blank_scalar_error);
    }
    // Provenance-aware filter also applies to a single-cell `Ref` argument
    // for range-aware aggregators. Excel treats `MIN(A1, A2, A3)` the same
    // way it treats `MIN(A1:A3)`: Text / Bool / Blank cells are silently
    // skipped for `range_filter_numeric_only`, coerced for the A-family,
    // etc. Direct scalar literals (numbers, bool literals, text literals)
    // still use strict coercion in the impl.
    if (def->accepts_ranges && arg_node.kind() == parser::NodeKind::Ref) {
      Value range_err = Value::blank();
      if (!append_range_sourced_value(*def, v, &values, &range_err)) {
        return range_err;
      }
      continue;
    }
    // An omitted slot of a range-aware aggregator is a counted 0 (COUNTA(3,,4) = 3).
    if (def->accepts_ranges && v.kind() == ValueKind::Blank && is_omitted_arg(arg_node)) {
      values.push_back(Value::blank(BlankGridProjection::kValueArrayZero));
      continue;
    }
    values.push_back(v);
  }
  // Deferred fire-point for `RejectAllScalarsBlank`: only surface the error
  // when blank scalar references were the entire scalar argument set. A
  // numeric literal or reference mixed with a blank therefore proceeds with
  // the blank omitted, while any range-shaped argument keeps its existing
  // blank-as-zero behavior.
  if (def->blank_scalar_policy == FunctionDef::BlankScalarPolicy::RejectAllScalarsBlank && saw_blank_scalar_ref &&
      values.empty() && !had_range_shaped_arg) {
    return Value::error(def->blank_scalar_error);
  }
  note_collected_error();
  if (defers_cell_errors) {
    const auto deferred = std::find_if(values.begin(), values.end(), [](const Value& v) { return v.is_error(); });
    if (deferred != values.end()) {
      const Value cell_error = *deferred;
      forget_error_origin();
      bool reported_cell_error = false;
      // The last value coerces first; its failure is reported ahead of the cell's error.
      const auto after_last_value = [&](const Value* cell_args, std::uint32_t n) {
        auto last = coerce_to_number(cell_args[n - 1U]);
        reported_cell_error = reported_cell_error || static_cast<bool>(last);
        return last ? cell_error : Value::error(last.error());
      };
      const auto size = static_cast<std::uint32_t>(values.size());
      const Value result = lift_scalar_slot ? lift_over_array_args(values.data(), size, arena, after_last_value)
                                            : after_last_value(values.data(), size);
      if (result.is_error() && reported_cell_error && reported_origin.has_value()) {
        restore_error_origin(*reported_origin);
      }
      return result;
    }
  }
  // Dynamic-array spill for scalar (non-range-aware) functions: when any
  // collected argument is an Array — a bounded range materialised above,
  // or a nested array-returning call (`=ROUND(SEQUENCE(3),0)`) — Excel 365
  // evaluates the function element-wise over the broadcast rectangle and
  // spills the result. Range-aware aggregators already flattened their
  // array arguments into `values`, so they never take this path.
  if (!def->accepts_ranges) {
    if (reported_origin.has_value()) {
      forget_error_origin();
    }
    const Value result = call_scalar_function(*def, values, tails, arena, tail_out);
    // The impl's own failure keeps what it noted; a reported error argument
    // keeps the origin it arrived with.
    restore_reported_origin(result, values.data(), static_cast<std::uint32_t>(values.size()),
                            def->checks_args_last_first, reported_origin);
    return result;
  }
  if (!streamed.empty()) {
    return fold_with_tails(*def, values, streamed);
  }
  if (lift_scalar_slot) {
    return lift_over_array_args(
        values.data(), static_cast<std::uint32_t>(values.size()), arena,
        [&](const Value* cell_args, std::uint32_t n) { return def->impl(cell_args, n, arena); });
  }
  // Hand the post-expansion size to the impl; aggregator bodies walk the
  // flattened vector directly.
  if (!in_order_ranges) {
    return def->impl(values.data(), static_cast<std::uint32_t>(values.size()), arena);
  }
  if (reported_origin.has_value()) {
    forget_error_origin();
  }
  const auto size = static_cast<std::uint32_t>(values.size());
  const Value result = report_error_arg(def->impl(values.data(), size, arena), values.data(), size, false);
  restore_reported_origin(result, values.data(), size, false, reported_origin);
  return result;
}

}  // namespace

Value run_lazy_call(const parser::AstNode& node, const LazyEntry& entry, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx) {
  const parser::AstNode* call = &node;
  ReshapeSource reshape;
  if (prepares_args(entry)) {
    Value prepared = Value::blank();
    call = prepare_lazy_args(node, entry, arena, registry, ctx, &reshape, &prepared);
    if (call == nullptr) {
      return prepared;
    }
  }
  // Lazy callees read their arguments as rectangles, so whole-axis
  // references among them are walked to one shared length and agree on
  // shape however far each is populated. Eager callees flatten ranges and
  // have no shape to agree on.
  WholeAxisScope scope;
  scope.call = call;
  const Value result = entry.impl(*call, arena, registry, ctx.with_whole_axis_scope(&scope));
  return reshape.areas.empty() ? result : map_reshaped(result, reshape, arena);
}

Shaped dispatch_call(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  // A whole-axis scope belongs to the arguments of the call that installed
  // it; a nested call starts from its own references.
  if (ctx.whole_axis_scope() != nullptr) {
    return dispatch_call(node, arena, registry, ctx.with_whole_axis_scope(nullptr));
  }
  Shaped out;
  out.value = dispatch_call_impl(node, arena, registry, ctx, &out.tail_array);
  return out;
}

}  // namespace eval
}  // namespace formulon
