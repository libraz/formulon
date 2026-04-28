// Copyright 2026 libraz. Licensed under the MIT License.

#include "eval/groupby_pivotby_lazy.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/jp_fold.h"
#include "eval/lambda_value.h"
#include "eval/lazy_impls.h"
#include "eval/name_env.h"
#include "eval/name_env_resolve.h"
#include "eval/shape_ops_lazy.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace {

// ---------------------------------------------------------------------------
// Aggregator resolution
// ---------------------------------------------------------------------------

// Discriminated reference to the aggregator selected for the call. Form C
// (bare function name) cannot be wrapped in a synthetic LambdaValue because a
// LambdaValue requires an AST body; per-group dispatch therefore branches on
// the kind tag and either invokes the lambda body or calls the registry impl.
struct AggregatorRef {
  enum class Kind { Lambda, Function };
  Kind kind = Kind::Lambda;
  const LambdaValue* lambda = nullptr;        // valid when kind == Lambda
  const FunctionDef* function_def = nullptr;  // valid when kind == Function
};

// Resolves the third argument (the aggregator) into an `AggregatorRef`.
// Returns true on success and writes the resolved aggregator to `*out`.
// Returns false on failure and writes the appropriate scalar error to
// `*out_err`.
//
// Resolution order:
//   1. If the raw arg AST is a `NameRef`, try the name environment first.
//      A bound name shadows any registry function with the same identifier.
//      If the binding evaluates to a Lambda of arity 1, that is Form B.
//   2. If still unresolved AND the raw arg AST is a `NameRef`, look the
//      name up in the registry. A hit is Form C.
//   3. Otherwise (LAMBDA literal, LET-bound lambda the parser surfaced via
//      something other than NameRef, or any other expression), evaluate the
//      arg via `eval_node`. A Lambda value of arity 1 is Form A. Anything
//      else surfaces `#VALUE!`.
//
// A Lambda whose `param_count != 1` surfaces `#VALUE!` regardless of form.
bool resolve_aggregator(const parser::AstNode& arg, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, AggregatorRef* out, Value* out_err) {
  // Step 1 + 2: NameRef short-circuit. A name bound in scope evaluates as the
  // bound Value (Form B); an unbound name falls through to the registry
  // (Form C) before we ever touch `eval_node` (which would surface #NAME?).
  if (arg.kind() == parser::NodeKind::NameRef) {
    const std::string_view name = arg.as_name();
    const NameEnv* env = ctx.name_env();
    const Value* bound = (env != nullptr) ? env->lookup(name) : nullptr;
    if (bound != nullptr) {
      if (bound->is_error()) {
        *out_err = *bound;
        return false;
      }
      if (!bound->is_lambda()) {
        *out_err = Value::error(ErrorCode::Value);
        return false;
      }
      const LambdaValue* lv = bound->as_lambda();
      if (lv->param_count != 1U) {
        *out_err = Value::error(ErrorCode::Value);
        return false;
      }
      out->kind = AggregatorRef::Kind::Lambda;
      out->lambda = lv;
      return true;
    }
    // Not bound in scope; consult the registry. Hit -> Form C.
    if (const FunctionDef* def = registry.lookup(name); def != nullptr) {
      out->kind = AggregatorRef::Kind::Function;
      out->function_def = def;
      return true;
    }
    // Miss in both scopes: fall through to general eval, which produces
    // `#NAME?`. The caller surfaces that verbatim.
  }

  // Step 3: evaluate the arg expression normally. Arity-1 Lambda -> Form A.
  const Value v = eval_node(arg, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  if (!v.is_lambda()) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  const LambdaValue* lv = v.as_lambda();
  if (lv->param_count != 1U) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  out->kind = AggregatorRef::Kind::Lambda;
  out->lambda = lv;
  return true;
}

// ---------------------------------------------------------------------------
// Argument helpers
// ---------------------------------------------------------------------------

// Reads an array argument via the standard `eval_node_as_array` seam so a
// Ref / RangeOp / ArrayLiteral / OFFSET-call argument keeps its 2D shape.
// Returns nullptr and writes the appropriate scalar error to `*out_err` on
// failure paths.
const ArrayValue* read_array_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                                 const EvalContext& ctx, Value* out_err) {
  const Value v = eval_node_as_array(node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return nullptr;
  }
  if (!v.is_array()) {
    *out_err = Value::error(ErrorCode::Value);
    return nullptr;
  }
  return v.as_array();
}

// Reads a scalar integer argument truncated toward zero, then validates it
// is a member of `allowed`. Returns true on success; on failure writes the
// appropriate scalar error to `*out_err` and returns false.
//
// `allowed` is a small file-static set passed in as `(values, count)` rather
// than `std::initializer_list` to keep the call sites trivially constexpr-
// constructible without hauling in <initializer_list>.
bool read_int_in_set(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx, const int* allowed, std::size_t count, int* out, Value* out_err) {
  const Value v = eval_node(node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    *out_err = Value::error(coerced.error());
    return false;
  }
  const double n = coerced.value();
  if (std::isnan(n) || std::isinf(n)) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  const int truncated = static_cast<int>(std::trunc(n));
  for (std::size_t i = 0; i < count; ++i) {
    if (allowed[i] == truncated) {
      *out = truncated;
      return true;
    }
  }
  *out_err = Value::error(ErrorCode::Value);
  return false;
}

// Reads a scalar integer argument truncated toward zero, with no membership
// check (used for `sort_order` where any int is grammatically valid; the
// out-of-range check happens after we know the column count).
bool read_int(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
              int* out, Value* out_err) {
  const Value v = eval_node(node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    *out_err = Value::error(coerced.error());
    return false;
  }
  const double n = coerced.value();
  if (std::isnan(n) || std::isinf(n)) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  *out = static_cast<int>(std::trunc(n));
  return true;
}

// ---------------------------------------------------------------------------
// Group-key equality
// ---------------------------------------------------------------------------

// Excel-canonical cell equality for GROUPBY group keys. Mirrors UNIQUE's
// rules with one difference: Text comparison runs through `fold_jp_text`
// first so `ｱ` (half-width katakana) folds to `ア` (full-width), matching
// Mac Excel COUNTIF / VLOOKUP ja-JP behaviour. Numbers compare bit-exact
// via `==` (so `0.0 == 0.0` and `1.0 == 1`, but the IEEE-754 `+0.0 != -0.0`
// distinction is preserved). Cross-kind pairs are never equal — `Number 0`
// and `Bool FALSE` form distinct groups.
bool group_cell_equal(const Value& a, const Value& b) {
  if (a.kind() != b.kind()) {
    return false;
  }
  switch (a.kind()) {
    case ValueKind::Blank:
      return true;
    case ValueKind::Number:
      return a.as_number() == b.as_number();
    case ValueKind::Bool:
      return a.as_boolean() == b.as_boolean();
    case ValueKind::Error:
      return a.as_error() == b.as_error();
    case ValueKind::Text:
      return fold_jp_text(a.as_text()) == fold_jp_text(b.as_text());
    default:
      // Array / Ref / Lambda are not produced by cell reads; treat as
      // not-equal defensively to avoid silent dedup of complex values.
      return false;
  }
}

// Multi-column key equality: walks each column of the keys and compares
// cellwise via `group_cell_equal`.
bool group_key_equal(const ArrayValue& keys, std::uint32_t row_a, std::uint32_t row_b) {
  for (std::uint32_t c = 0; c < keys.cols; ++c) {
    const Value& va = keys.cells[static_cast<std::size_t>(row_a) * keys.cols + c];
    const Value& vb = keys.cells[static_cast<std::size_t>(row_b) * keys.cols + c];
    if (!group_cell_equal(va, vb)) {
      return false;
    }
  }
  return true;
}

// True iff every column of the row's key is an Error value. Such rows form
// "error groups" that sort to the bottom (after all valid keys) — Mac Excel
// surfaces `#N/A` cells as a trailing group rather than aborting.
bool row_key_is_error(const ArrayValue& keys, std::uint32_t row) {
  for (std::uint32_t c = 0; c < keys.cols; ++c) {
    const Value& v = keys.cells[static_cast<std::size_t>(row) * keys.cols + c];
    if (v.is_error()) {
      return true;
    }
  }
  return false;
}

// ---------------------------------------------------------------------------
// Aggregator invocation
// ---------------------------------------------------------------------------

// Builds a 1-column array containing `column_indices` rows pulled from
// `values`'s `value_col`-th column. Used to construct the per-group slice
// passed to the aggregator.
const ArrayValue* build_group_slice(const ArrayValue& values, std::uint32_t value_col,
                                    const std::vector<std::uint32_t>& row_indices, Arena& arena) {
  const std::uint32_t n = static_cast<std::uint32_t>(row_indices.size());
  if (n == 0U) {
    return nullptr;
  }
  Value* cells = arena.create_array<Value>(n);
  if (cells == nullptr) {
    return nullptr;
  }
  for (std::uint32_t i = 0; i < n; ++i) {
    cells[i] = values.cells[static_cast<std::size_t>(row_indices[i]) * values.cols + value_col];
  }
  ArrayValue* arr = arena.create<ArrayValue>();
  if (arr == nullptr) {
    return nullptr;
  }
  arr->rows = n;
  arr->cols = 1U;
  arr->cells = cells;
  return arr;
}

// Builds a synthetic `ArrayLiteral` AST mirroring `arr`. This lets a Lambda
// body that expects a range-shaped argument (`SUM(v)`, `AVERAGE(v)`, ...)
// see the slice through the dispatcher's existing ArrayLiteral branch
// rather than as an opaque `Value::Array`. Mirrors the helper in
// `lambda_helpers_lazy.cpp`.
const parser::AstNode* build_array_literal_for_slice(const ArrayValue* arr, Arena& arena) {
  if (arr == nullptr || arr->rows == 0U || arr->cols == 0U) {
    return nullptr;
  }
  const std::size_t total = static_cast<std::size_t>(arr->rows) * static_cast<std::size_t>(arr->cols);
  const parser::AstNode** children = arena.create_array<const parser::AstNode*>(total);
  if (children == nullptr) {
    return nullptr;
  }
  for (std::size_t i = 0; i < total; ++i) {
    parser::AstNode* lit = parser::make_literal(arena, arr->cells[i]);
    if (lit == nullptr) {
      return nullptr;
    }
    children[i] = lit;
  }
  return parser::make_array_literal(arena, arr->rows, arr->cols, children);
}

// Invokes a Lambda aggregator for one group. Mirrors `invoke_lambda_with_values`
// in `lambda_helpers_lazy.cpp` but lives here so we can opt into the per-group
// error-isolation path (the helper there is shape-mismatch driven and short-
// circuits the whole call on the first error). The returned Value is whatever
// the lambda body produced — including errors, which the caller stores
// verbatim in the cell.
Value invoke_lambda_for_group(const LambdaValue* lv, const ArrayValue* slice, const parser::AstNode* slice_ast,
                              Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx) {
  NameEnv env;
  if (lv->captured_env != nullptr) {
    env = *lv->captured_env;
  }
  const Value slice_v = Value::array(slice);
  env = env.extend(lv->params[0], slice_v, slice_ast, arena);
  const EvalContext body_ctx = ctx.with_name_env(&env);
  return eval_node(*lv->body, arena, registry, body_ctx);
}

// Invokes a registry-backed aggregator (Form C) for one group. The slice is
// flattened into the args vector cellwise so a SUM-style impl sees the same
// shape it would see from `=SUM({1;2;3})`.
Value invoke_function_for_group(const FunctionDef* def, const ArrayValue* slice, Arena& arena) {
  if (slice == nullptr || slice->rows == 0U) {
    // No values to aggregate; conservatively surface the aggregator's
    // empty-input behaviour by passing zero args. Most aggregate impls
    // (SUM, MIN, MAX, ...) check arity >= min_arity and return #VALUE!.
    return Value::error(ErrorCode::Calc);
  }
  const std::uint32_t n = slice->rows;
  Value* args = arena.create_array<Value>(n);
  if (args == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::uint32_t i = 0; i < n; ++i) {
    args[i] = slice->cells[i];
  }
  if (n < def->min_arity || (def->max_arity != kVariadic && n > def->max_arity)) {
    return Value::error(ErrorCode::Value);
  }
  return def->impl(args, n, arena);
}

// Invokes the resolved aggregator for one group's column slice and returns
// whatever it produced. Errors are returned verbatim (the per-group
// error-isolation seam). A multi-cell array return surfaces as `#CALC!`
// matching BYROW / MAP — the cell has no slot to spill into.
Value invoke_aggregator_for_group(const AggregatorRef& agg, const ArrayValue* slice, Arena& arena,
                                  const FunctionRegistry& registry, const EvalContext& ctx) {
  Value res = Value::blank();
  if (agg.kind == AggregatorRef::Kind::Lambda) {
    const parser::AstNode* slice_ast = build_array_literal_for_slice(slice, arena);
    if (slice_ast == nullptr) {
      return Value::error(ErrorCode::Num);
    }
    res = invoke_lambda_for_group(agg.lambda, slice, slice_ast, arena, registry, ctx);
  } else {
    res = invoke_function_for_group(agg.function_def, slice, arena);
  }
  if (res.is_array()) {
    return Value::error(ErrorCode::Calc);
  }
  if (res.is_lambda()) {
    return Value::error(ErrorCode::Calc);
  }
  return res;
}

// ---------------------------------------------------------------------------
// Output assembly
// ---------------------------------------------------------------------------

// Renders one row of cells to the buffer at `out[row * cols + 0..cols-1]`.
// Used to materialise headers, total rows, and per-group rows.
void emit_row(std::vector<std::vector<Value>>* rows, const std::vector<Value>& row) {
  rows->push_back(row);
}

// Comparator helper for sort tie-breaking: ascending compare on a single
// scalar Value. Numbers compare by value; text compares by Mac-folded UTF-8
// bytes; cross-kind pairs use the kind() ordinal so ordering is stable but
// unspecified in detail. Errors and Blanks are pushed to the end (Excel's
// "blanks last" rule).
int cmp_value_asc(const Value& a, const Value& b) {
  // Map kinds to ordering buckets: Number(0) < Text(1) < Bool(2) < Error(3) < Blank(4).
  auto bucket = [](const Value& v) -> int {
    switch (v.kind()) {
      case ValueKind::Number:
        return 0;
      case ValueKind::Text:
        return 1;
      case ValueKind::Bool:
        return 2;
      case ValueKind::Error:
        return 3;
      case ValueKind::Blank:
        return 4;
      default:
        return 5;
    }
  };
  const int ba = bucket(a);
  const int bb = bucket(b);
  if (ba != bb) {
    return ba < bb ? -1 : 1;
  }
  switch (a.kind()) {
    case ValueKind::Number: {
      const double na = a.as_number();
      const double nb = b.as_number();
      if (na < nb) {
        return -1;
      }
      if (na > nb) {
        return 1;
      }
      return 0;
    }
    case ValueKind::Text: {
      const std::string ta = fold_jp_text(a.as_text());
      const std::string tb = fold_jp_text(b.as_text());
      if (ta < tb) {
        return -1;
      }
      if (ta > tb) {
        return 1;
      }
      return 0;
    }
    case ValueKind::Bool: {
      const bool ba2 = a.as_boolean();
      const bool bb2 = b.as_boolean();
      if (ba2 == bb2) {
        return 0;
      }
      return ba2 ? 1 : -1;
    }
    case ValueKind::Error: {
      const auto ea = static_cast<int>(a.as_error());
      const auto eb = static_cast<int>(b.as_error());
      if (ea < eb) {
        return -1;
      }
      if (ea > eb) {
        return 1;
      }
      return 0;
    }
    default:
      return 0;
  }
}

// Compares two group keys lexicographically across every column for the
// stable-sort tie-break. Returns -1 / 0 / 1.
int cmp_keys_asc(const ArrayValue& keys, std::uint32_t a_row, std::uint32_t b_row) {
  for (std::uint32_t c = 0; c < keys.cols; ++c) {
    const Value& va = keys.cells[static_cast<std::size_t>(a_row) * keys.cols + c];
    const Value& vb = keys.cells[static_cast<std::size_t>(b_row) * keys.cols + c];
    const int r = cmp_value_asc(va, vb);
    if (r != 0) {
      return r;
    }
  }
  return 0;
}

}  // namespace

// ---------------------------------------------------------------------------
// GROUPBY entry point
// ---------------------------------------------------------------------------

Value eval_groupby_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 3U || arity > 7U) {
    return Value::error(ErrorCode::Value);
  }

  Value err = Value::blank();

  // -- arg 0: row_fields ----------------------------------------------------
  const ArrayValue* row_fields = read_array_arg(call.as_call_arg(0), arena, registry, ctx, &err);
  if (row_fields == nullptr) {
    return err;
  }

  // -- arg 1: values --------------------------------------------------------
  const ArrayValue* values = read_array_arg(call.as_call_arg(1), arena, registry, ctx, &err);
  if (values == nullptr) {
    return err;
  }

  // Row-count consistency. Mac Excel surfaces #VALUE! when the two arrays
  // have different row counts.
  if (row_fields->rows != values->rows) {
    return Value::error(ErrorCode::Value);
  }
  if (row_fields->rows == 0U || row_fields->cols == 0U || values->cols == 0U) {
    return Value::error(ErrorCode::Calc);
  }

  // -- arg 2: aggregator ----------------------------------------------------
  AggregatorRef agg;
  if (!resolve_aggregator(call.as_call_arg(2), arena, registry, ctx, &agg, &err)) {
    return err;
  }

  // -- arg 3: field_headers ∈ {0,1,2,3} ------------------------------------
  int field_headers = 0;
  if (arity >= 4U) {
    static constexpr int kAllowed[] = {0, 1, 2, 3};
    if (!read_int_in_set(call.as_call_arg(3), arena, registry, ctx, kAllowed, sizeof(kAllowed) / sizeof(kAllowed[0]),
                         &field_headers, &err)) {
      return err;
    }
  }

  // -- arg 4: total_depth ∈ {-2,-1,0,1,2} ----------------------------------
  int total_depth = -1;
  if (arity >= 5U) {
    static constexpr int kAllowed[] = {-2, -1, 0, 1, 2};
    if (!read_int_in_set(call.as_call_arg(4), arena, registry, ctx, kAllowed, sizeof(kAllowed) / sizeof(kAllowed[0]),
                         &total_depth, &err)) {
      return err;
    }
  }

  // -- arg 5: sort_order ----------------------------------------------------
  int sort_order = 0;
  if (arity >= 6U) {
    if (!read_int(call.as_call_arg(5), arena, registry, ctx, &sort_order, &err)) {
      return err;
    }
  }

  // Determine header row layout. Inputs have a header row when
  // field_headers ∈ {1, 3}; outputs emit a header row when
  // field_headers ∈ {1, 2, 3}.
  const bool inputs_have_header = (field_headers == 1 || field_headers == 3);
  const bool output_emits_header = (field_headers == 1 || field_headers == 2 || field_headers == 3);

  if (inputs_have_header && row_fields->rows < 1U) {
    return Value::error(ErrorCode::Value);
  }

  const std::uint32_t data_start_row = inputs_have_header ? 1U : 0U;
  if (row_fields->rows < data_start_row) {
    return Value::error(ErrorCode::Calc);
  }
  const std::uint32_t data_row_count = row_fields->rows - data_start_row;

  // -- arg 6: filter_array --------------------------------------------------
  std::vector<bool> include_row(data_row_count, true);
  if (arity == 7U) {
    const ArrayValue* mask = read_array_arg(call.as_call_arg(6), arena, registry, ctx, &err);
    if (mask == nullptr) {
      return err;
    }
    // Excel allows a single-column or single-row 1D mask.
    const std::uint32_t mask_n = (mask->rows >= mask->cols) ? mask->rows : mask->cols;
    if (mask->rows != 1U && mask->cols != 1U) {
      return Value::error(ErrorCode::Value);
    }
    if (mask_n != data_row_count) {
      return Value::error(ErrorCode::Value);
    }
    for (std::uint32_t i = 0; i < data_row_count; ++i) {
      const Value& cell = mask->cells[i];
      if (cell.is_error()) {
        return cell;
      }
      auto coerced = coerce_to_bool(cell);
      if (!coerced) {
        return Value::error(coerced.error());
      }
      include_row[i] = coerced.value();
    }
  }

  // -- Build groups --------------------------------------------------------
  // Walk filtered data rows in input order; for each row, look up its group
  // key against the existing list of unique keys (linear scan via
  // `group_key_equal`). New keys append; matching keys add the row index to
  // their bucket. This preserves first-occurrence ordering for sort_order=0
  // for free.
  //
  // Group representatives are stored as row indices (into `row_fields`) so
  // that subsequent equality checks can re-use the same column-walk path.
  std::vector<std::uint32_t> group_repr;               // row index of representative
  std::vector<std::vector<std::uint32_t>> group_rows;  // row indices in each group
  std::vector<bool> group_is_error;

  for (std::uint32_t i = 0; i < data_row_count; ++i) {
    if (!include_row[i]) {
      continue;
    }
    const std::uint32_t row = data_start_row + i;
    bool matched = false;
    for (std::size_t g = 0; g < group_repr.size(); ++g) {
      if (group_key_equal(*row_fields, row, group_repr[g])) {
        group_rows[g].push_back(row);
        matched = true;
        break;
      }
    }
    if (!matched) {
      group_repr.push_back(row);
      group_rows.push_back(std::vector<std::uint32_t>{row});
      group_is_error.push_back(row_key_is_error(*row_fields, row));
    }
  }

  if (group_repr.empty()) {
    // No data rows survived the filter (or input was empty after header
    // stripping). Mac Excel surfaces #CALC! for "no values to return".
    return Value::error(ErrorCode::Calc);
  }

  // -- Aggregate per group -------------------------------------------------
  // Output has `row_fields->cols` key columns followed by `values->cols`
  // aggregated columns. Per-group error isolation: each invocation's error
  // is captured into its cell; the rest of the result is still computed.
  const std::uint32_t key_cols = row_fields->cols;
  const std::uint32_t val_cols = values->cols;
  const std::uint32_t out_cols = key_cols + val_cols;

  std::vector<std::vector<Value>> agg_rows;
  agg_rows.reserve(group_repr.size());
  for (std::size_t g = 0; g < group_repr.size(); ++g) {
    std::vector<Value> row(out_cols, Value::blank());
    // Key columns: copy the representative's key cells verbatim.
    const std::uint32_t repr_row = group_repr[g];
    for (std::uint32_t c = 0; c < key_cols; ++c) {
      row[c] = row_fields->cells[static_cast<std::size_t>(repr_row) * key_cols + c];
    }
    // Value columns: invoke the aggregator per column with the group's
    // slice. Errors land in the cell; they do NOT short-circuit the result.
    for (std::uint32_t vc = 0; vc < val_cols; ++vc) {
      const ArrayValue* slice = build_group_slice(*values, vc, group_rows[g], arena);
      if (slice == nullptr) {
        row[key_cols + vc] = Value::error(ErrorCode::Num);
        continue;
      }
      row[key_cols + vc] = invoke_aggregator_for_group(agg, slice, arena, registry, ctx);
    }
    agg_rows.push_back(std::move(row));
  }

  // -- Sort ---------------------------------------------------------------
  // sort_order semantics:
  //   * 0 -> preserve first-occurrence order (already true).
  //   * N>0 -> stable-sort ascending by the N-th aggregated value column
  //     (1-based; out-of-range -> #VALUE!). Tie-break on group key.
  //   * N<0 -> stable-sort descending by |N|-th column.
  // Error-keyed groups always sort to the bottom (after all valid groups).
  std::vector<std::size_t> order(agg_rows.size());
  for (std::size_t i = 0; i < order.size(); ++i) {
    order[i] = i;
  }
  if (sort_order != 0) {
    const int abs_sort = sort_order > 0 ? sort_order : -sort_order;
    if (abs_sort < 1 || static_cast<std::uint32_t>(abs_sort) > val_cols) {
      return Value::error(ErrorCode::Value);
    }
    const std::uint32_t value_col_idx = static_cast<std::uint32_t>(abs_sort - 1);
    const bool descending = (sort_order < 0);
    std::stable_sort(order.begin(), order.end(), [&](std::size_t a, std::size_t b) {
      // Error-keyed groups always go last regardless of
      // direction; this matches Mac Excel's "errors trail"
      // surface for the dynamic-array sort family.
      if (group_is_error[a] != group_is_error[b]) {
        return !group_is_error[a];
      }
      const Value& va = agg_rows[a][key_cols + value_col_idx];
      const Value& vb = agg_rows[b][key_cols + value_col_idx];
      const int c = cmp_value_asc(va, vb);
      if (c != 0) {
        return descending ? (c > 0) : (c < 0);
      }
      // Tie-break on group key (always ascending).
      return cmp_keys_asc(*row_fields, group_repr[a], group_repr[b]) < 0;
    });
  } else {
    // Even at sort_order=0, error-keyed groups sink to the bottom in stable
    // first-occurrence order (matching Mac Excel's UNIQUE / FILTER pattern
    // for cells whose evaluation produced an error).
    std::stable_sort(order.begin(), order.end(), [&](std::size_t a, std::size_t b) {
      if (group_is_error[a] != group_is_error[b]) {
        return !group_is_error[a];
      }
      return false;  // preserve original order within bucket
    });
  }

  // -- Compute grand total (if requested) ---------------------------------
  // The grand total aggregates over EVERY filtered data row. It is rendered
  // with the literal "Grand Total" label in the first key column and blank
  // cells in the remaining key columns.
  std::vector<Value> grand_total_row;
  bool emit_grand_total = (total_depth != 0);
  if (emit_grand_total) {
    grand_total_row.assign(out_cols, Value::blank());
    grand_total_row[0] = Value::text(arena.intern("Grand Total"));
    // Build the row-index list of every included data row.
    std::vector<std::uint32_t> all_rows;
    all_rows.reserve(data_row_count);
    for (std::uint32_t i = 0; i < data_row_count; ++i) {
      if (include_row[i]) {
        all_rows.push_back(data_start_row + i);
      }
    }
    for (std::uint32_t vc = 0; vc < val_cols; ++vc) {
      const ArrayValue* slice = build_group_slice(*values, vc, all_rows, arena);
      if (slice == nullptr) {
        grand_total_row[key_cols + vc] = Value::error(ErrorCode::Num);
        continue;
      }
      grand_total_row[key_cols + vc] = invoke_aggregator_for_group(agg, slice, arena, registry, ctx);
    }
  }
  // Subtotals (total_depth == ±2) require multi-column row_fields; with
  // single-column keys they silently degrade to the ±1 behaviour. For the
  // multi-column case in this implementation we currently emit only the
  // grand total — proper subtotal rows require an outer/inner split that
  // is deferred until the oracle confirms exact placement against Mac
  // Excel's actual surface.

  // -- Assemble output ----------------------------------------------------
  std::vector<std::vector<Value>> out_rows;
  out_rows.reserve(agg_rows.size() + 2U);

  // Header row (if requested).
  if (output_emits_header) {
    std::vector<Value> header(out_cols, Value::blank());
    if (field_headers == 1 || field_headers == 3) {
      // Inputs had a header row (row 0 of each input). Copy it verbatim.
      for (std::uint32_t c = 0; c < key_cols; ++c) {
        header[c] = row_fields->cells[c];
      }
      for (std::uint32_t c = 0; c < val_cols; ++c) {
        header[key_cols + c] = values->cells[c];
      }
    } else {
      // field_headers == 2: synthesize English defaults. Mac Excel may
      // emit Japanese labels in ja-JP; this divergence is logged for the
      // first oracle run.
      for (std::uint32_t c = 0; c < key_cols; ++c) {
        const std::string label = "Field " + std::to_string(c + 1U);
        header[c] = Value::text(arena.intern(label));
      }
      for (std::uint32_t c = 0; c < val_cols; ++c) {
        const std::string label = "Value " + std::to_string(c + 1U);
        header[key_cols + c] = Value::text(arena.intern(label));
      }
    }
    emit_row(&out_rows, header);
  }

  // Grand total at top (negative total_depth).
  if (emit_grand_total && total_depth < 0) {
    emit_row(&out_rows, grand_total_row);
  }

  // Per-group rows in sorted order.
  for (std::size_t i : order) {
    emit_row(&out_rows, agg_rows[i]);
  }

  // Grand total at bottom (positive total_depth).
  if (emit_grand_total && total_depth > 0) {
    emit_row(&out_rows, grand_total_row);
  }

  if (out_rows.empty()) {
    return Value::error(ErrorCode::Calc);
  }

  // Flatten into row-major arena buffer.
  const std::uint32_t out_rows_n = static_cast<std::uint32_t>(out_rows.size());
  const std::size_t total_cells = static_cast<std::size_t>(out_rows_n) * static_cast<std::size_t>(out_cols);
  Value* buffer = arena.create_array<Value>(total_cells);
  if (buffer == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::uint32_t r = 0; r < out_rows_n; ++r) {
    for (std::uint32_t c = 0; c < out_cols; ++c) {
      buffer[static_cast<std::size_t>(r) * out_cols + c] = out_rows[r][c];
    }
  }
  ArrayValue* arr = arena.create<ArrayValue>();
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  arr->rows = out_rows_n;
  arr->cols = out_cols;
  arr->cells = buffer;
  return Value::array(arr);
}

// ---------------------------------------------------------------------------
// PIVOTBY entry point
// ---------------------------------------------------------------------------

namespace {

// Maps a single row's group key to its group index, appending a new group on
// first occurrence. `keys` is the source key array (row_fields or
// col_fields); `row` is the absolute row index. Returns the group index.
//
// A composite copy of `representative_rows` is kept so callers can later
// pull the canonical key cells via `keys.cells[representative_rows[g] *
// keys.cols + c]`. Mirrors the inline group-build loop in
// `eval_groupby_lazy`; factored here so both the row and column axes share
// the same first-occurrence semantics.
std::size_t find_or_add_group(const ArrayValue& keys, std::uint32_t row,
                              std::vector<std::uint32_t>* representative_rows,
                              std::vector<std::vector<std::uint32_t>>* member_rows, std::vector<bool>* is_error_group) {
  for (std::size_t g = 0; g < representative_rows->size(); ++g) {
    if (group_key_equal(keys, row, (*representative_rows)[g])) {
      (*member_rows)[g].push_back(row);
      return g;
    }
  }
  representative_rows->push_back(row);
  member_rows->push_back(std::vector<std::uint32_t>{row});
  is_error_group->push_back(row_key_is_error(keys, row));
  return representative_rows->size() - 1U;
}

}  // namespace

Value eval_pivotby_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 4U || arity > 10U) {
    return Value::error(ErrorCode::Value);
  }

  Value err = Value::blank();

  // -- arg 0: row_fields ----------------------------------------------------
  const ArrayValue* row_fields = read_array_arg(call.as_call_arg(0), arena, registry, ctx, &err);
  if (row_fields == nullptr) {
    return err;
  }

  // -- arg 1: col_fields ----------------------------------------------------
  const ArrayValue* col_fields = read_array_arg(call.as_call_arg(1), arena, registry, ctx, &err);
  if (col_fields == nullptr) {
    return err;
  }

  // -- arg 2: values --------------------------------------------------------
  const ArrayValue* values = read_array_arg(call.as_call_arg(2), arena, registry, ctx, &err);
  if (values == nullptr) {
    return err;
  }

  // Row-count consistency across all three rectangles. Mac Excel surfaces
  // `#VALUE!` when any pair has different row counts.
  if (row_fields->rows != values->rows || col_fields->rows != values->rows) {
    return Value::error(ErrorCode::Value);
  }
  if (row_fields->rows == 0U || row_fields->cols == 0U || col_fields->cols == 0U || values->cols == 0U) {
    return Value::error(ErrorCode::Calc);
  }

  // First-commit scope: single-column row_fields / col_fields / values.
  // Multi-column shapes are deferred (the cross-product output layout needs
  // an oracle pass to confirm header / total placement). They surface
  // `#VALUE!` here. TODO(pivotby-multi-col): lift this restriction once the
  // oracle has captured Mac Excel's surface for the multi-column variants.
  if (row_fields->cols != 1U || col_fields->cols != 1U || values->cols != 1U) {
    return Value::error(ErrorCode::Value);
  }

  // -- arg 3: aggregator ----------------------------------------------------
  AggregatorRef agg;
  if (!resolve_aggregator(call.as_call_arg(3), arena, registry, ctx, &agg, &err)) {
    return err;
  }

  // -- arg 4: field_headers ∈ {0,1,2,3}, default 3 -------------------------
  // PIVOTBY's default differs from GROUPBY's (0): pivot output typically
  // wants both the input row to be treated as a header AND a header to be
  // emitted on the output's left/top edges.
  int field_headers = 3;
  if (arity >= 5U) {
    static constexpr int kAllowed[] = {0, 1, 2, 3};
    if (!read_int_in_set(call.as_call_arg(4), arena, registry, ctx, kAllowed, sizeof(kAllowed) / sizeof(kAllowed[0]),
                         &field_headers, &err)) {
      return err;
    }
  }

  // -- arg 5: row_total_depth ∈ {-2,-1,0,1,2}, default -1 ------------------
  // The grand-total row (showing column totals) defaults to the TOP of the
  // result. ±2 (subtotal rows) is deferred and silently degrades to ±1 in
  // the single-column row_fields scope of this commit.
  int row_total_depth = -1;
  if (arity >= 6U) {
    static constexpr int kAllowed[] = {-2, -1, 0, 1, 2};
    if (!read_int_in_set(call.as_call_arg(5), arena, registry, ctx, kAllowed, sizeof(kAllowed) / sizeof(kAllowed[0]),
                         &row_total_depth, &err)) {
      return err;
    }
  }

  // -- arg 6: row_sort_order, default 0 ------------------------------------
  // Sort the row groups by their row totals (`SUM`-like aggregation over
  // every (row_group, col_group) cell of the row). 0 preserves first-
  // occurrence order; positive means ascending; negative descending.
  int row_sort_order = 0;
  if (arity >= 7U) {
    if (!read_int(call.as_call_arg(6), arena, registry, ctx, &row_sort_order, &err)) {
      return err;
    }
  }

  // -- arg 7: col_total_depth ∈ {-2,-1,0,1,2}, default 1 -------------------
  // The grand-total column (showing row totals) defaults to the RIGHT of
  // the result. ±2 silently degrades to ±1 in the single-column col_fields
  // scope of this commit.
  int col_total_depth = 1;
  if (arity >= 8U) {
    static constexpr int kAllowed[] = {-2, -1, 0, 1, 2};
    if (!read_int_in_set(call.as_call_arg(7), arena, registry, ctx, kAllowed, sizeof(kAllowed) / sizeof(kAllowed[0]),
                         &col_total_depth, &err)) {
      return err;
    }
  }

  // -- arg 8: col_sort_order, default 0 ------------------------------------
  int col_sort_order = 0;
  if (arity >= 9U) {
    if (!read_int(call.as_call_arg(8), arena, registry, ctx, &col_sort_order, &err)) {
      return err;
    }
  }

  // Determine header row layout. Same as GROUPBY but the header / output
  // emission flags drive both the row-axis labels (left edge) and the
  // col-axis labels (top edge).
  const bool inputs_have_header = (field_headers == 1 || field_headers == 3);
  const bool output_emits_header = (field_headers == 1 || field_headers == 2 || field_headers == 3);

  if (inputs_have_header && row_fields->rows < 1U) {
    return Value::error(ErrorCode::Value);
  }

  const std::uint32_t data_start_row = inputs_have_header ? 1U : 0U;
  if (row_fields->rows < data_start_row) {
    return Value::error(ErrorCode::Calc);
  }
  const std::uint32_t data_row_count = row_fields->rows - data_start_row;

  // -- arg 9: filter_array --------------------------------------------------
  std::vector<bool> include_row(data_row_count, true);
  if (arity == 10U) {
    const ArrayValue* mask = read_array_arg(call.as_call_arg(9), arena, registry, ctx, &err);
    if (mask == nullptr) {
      return err;
    }
    const std::uint32_t mask_n = (mask->rows >= mask->cols) ? mask->rows : mask->cols;
    if (mask->rows != 1U && mask->cols != 1U) {
      return Value::error(ErrorCode::Value);
    }
    if (mask_n != data_row_count) {
      return Value::error(ErrorCode::Value);
    }
    for (std::uint32_t i = 0; i < data_row_count; ++i) {
      const Value& cell = mask->cells[i];
      if (cell.is_error()) {
        return cell;
      }
      auto coerced = coerce_to_bool(cell);
      if (!coerced) {
        return Value::error(coerced.error());
      }
      include_row[i] = coerced.value();
    }
  }

  // -- Build row and column groups ----------------------------------------
  // For each filtered data row: assign the row to one row-group (by
  // row_fields key) and one col-group (by col_fields key). The (row_group,
  // col_group) pair is later used to aggregate the values column. We also
  // keep per-group full row-index lists for row totals and col totals.
  std::vector<std::uint32_t> row_repr;
  std::vector<std::vector<std::uint32_t>> row_members;  // rows in each row-group
  std::vector<bool> row_is_error;
  std::vector<std::uint32_t> col_repr;
  std::vector<std::vector<std::uint32_t>> col_members;  // rows in each col-group
  std::vector<bool> col_is_error;

  // Per-row tags so we can later compute (row_g, col_g) intersections by
  // walking the data rows once.
  std::vector<std::size_t> row_tag(data_row_count, 0);
  std::vector<std::size_t> col_tag(data_row_count, 0);

  for (std::uint32_t i = 0; i < data_row_count; ++i) {
    if (!include_row[i]) {
      continue;
    }
    const std::uint32_t row = data_start_row + i;
    row_tag[i] = find_or_add_group(*row_fields, row, &row_repr, &row_members, &row_is_error);
    col_tag[i] = find_or_add_group(*col_fields, row, &col_repr, &col_members, &col_is_error);
  }

  if (row_repr.empty() || col_repr.empty()) {
    return Value::error(ErrorCode::Calc);
  }

  const std::size_t n_rows = row_repr.size();
  const std::size_t n_cols = col_repr.size();

  // -- Aggregate per (row_group, col_group) -------------------------------
  // For each (rg, cg) pair, build the intersection row list and invoke the
  // aggregator. Per-cell error isolation: an aggregator failure for one
  // pair lands in that cell; the rest of the body is still computed.
  std::vector<std::vector<Value>> body(n_rows, std::vector<Value>(n_cols, Value::blank()));
  for (std::size_t rg = 0; rg < n_rows; ++rg) {
    for (std::size_t cg = 0; cg < n_cols; ++cg) {
      std::vector<std::uint32_t> intersection;
      // The intersection is small for typical workbooks; a linear walk is
      // simpler than building an index. (For large pivots this is the hot
      // loop and should be revisited with a hash bucket.)
      for (std::uint32_t i = 0; i < data_row_count; ++i) {
        if (!include_row[i]) {
          continue;
        }
        if (row_tag[i] == rg && col_tag[i] == cg) {
          intersection.push_back(data_start_row + i);
        }
      }
      if (intersection.empty()) {
        // No data points at this intersection. Mac Excel surfaces an empty
        // (Blank) cell here — the aggregator is not called for empty
        // groups in the pivot body.
        body[rg][cg] = Value::blank();
        continue;
      }
      const ArrayValue* slice = build_group_slice(*values, /*value_col=*/0U, intersection, arena);
      if (slice == nullptr) {
        body[rg][cg] = Value::error(ErrorCode::Num);
        continue;
      }
      body[rg][cg] = invoke_aggregator_for_group(agg, slice, arena, registry, ctx);
    }
  }

  // -- Compute row totals (one per row group) -----------------------------
  // A "row total" is the aggregate over an entire row group regardless of
  // col group — it lives in the GRAND-TOTAL COLUMN, whose presence is
  // governed by `col_total_depth` (the col-axis total knob).
  const bool emit_row_totals_col = (col_total_depth != 0);
  std::vector<Value> row_totals(n_rows, Value::blank());
  if (emit_row_totals_col) {
    for (std::size_t rg = 0; rg < n_rows; ++rg) {
      const ArrayValue* slice = build_group_slice(*values, /*value_col=*/0U, row_members[rg], arena);
      if (slice == nullptr) {
        row_totals[rg] = Value::error(ErrorCode::Num);
        continue;
      }
      row_totals[rg] = invoke_aggregator_for_group(agg, slice, arena, registry, ctx);
    }
  }

  // -- Compute column totals (one per col group) --------------------------
  // A "column total" is the aggregate over an entire col group regardless
  // of row group — it lives in the GRAND-TOTAL ROW, whose presence is
  // governed by `row_total_depth` (the row-axis total knob).
  const bool emit_col_totals_row = (row_total_depth != 0);
  std::vector<Value> col_totals(n_cols, Value::blank());
  if (emit_col_totals_row) {
    for (std::size_t cg = 0; cg < n_cols; ++cg) {
      const ArrayValue* slice = build_group_slice(*values, /*value_col=*/0U, col_members[cg], arena);
      if (slice == nullptr) {
        col_totals[cg] = Value::error(ErrorCode::Num);
        continue;
      }
      col_totals[cg] = invoke_aggregator_for_group(agg, slice, arena, registry, ctx);
    }
  }

  // -- Compute grand total ------------------------------------------------
  // The grand total cell exists only when both axes emit totals.
  Value grand_total = Value::blank();
  if (emit_row_totals_col && emit_col_totals_row) {
    std::vector<std::uint32_t> all_rows;
    all_rows.reserve(data_row_count);
    for (std::uint32_t i = 0; i < data_row_count; ++i) {
      if (include_row[i]) {
        all_rows.push_back(data_start_row + i);
      }
    }
    if (all_rows.empty()) {
      grand_total = Value::error(ErrorCode::Calc);
    } else {
      const ArrayValue* slice = build_group_slice(*values, /*value_col=*/0U, all_rows, arena);
      if (slice == nullptr) {
        grand_total = Value::error(ErrorCode::Num);
      } else {
        grand_total = invoke_aggregator_for_group(agg, slice, arena, registry, ctx);
      }
    }
  }

  // -- Sort row groups ----------------------------------------------------
  // Sort by row-total (the "first/only value column" reduces to the row
  // total in single-column-values scope). Error-keyed groups sink to the
  // bottom; sort_order=0 preserves first-occurrence order modulo that
  // sink.
  std::vector<std::size_t> row_order(n_rows);
  for (std::size_t i = 0; i < n_rows; ++i) {
    row_order[i] = i;
  }
  if (row_sort_order != 0) {
    if (row_sort_order != 1 && row_sort_order != -1) {
      // First-commit scope: only ±1 / 0 are supported. Multi-column values
      // would let |sort_order| index a value column; that ships when we
      // lift the multi-column restriction above.
      return Value::error(ErrorCode::Value);
    }
    const bool descending = (row_sort_order < 0);
    std::stable_sort(row_order.begin(), row_order.end(), [&](std::size_t a, std::size_t b) {
      if (row_is_error[a] != row_is_error[b]) {
        return !row_is_error[a];
      }
      // Compute on the fly if row totals weren't emitted (col_total_depth
      // == 0): aggregate just for the sort. This is rare; keep it simple.
      Value va = row_totals[a];
      Value vb = row_totals[b];
      if (!emit_row_totals_col) {
        const ArrayValue* sa = build_group_slice(*values, 0U, row_members[a], arena);
        const ArrayValue* sb = build_group_slice(*values, 0U, row_members[b], arena);
        va =
            (sa != nullptr) ? invoke_aggregator_for_group(agg, sa, arena, registry, ctx) : Value::error(ErrorCode::Num);
        vb =
            (sb != nullptr) ? invoke_aggregator_for_group(agg, sb, arena, registry, ctx) : Value::error(ErrorCode::Num);
      }
      const int c = cmp_value_asc(va, vb);
      if (c != 0) {
        return descending ? (c > 0) : (c < 0);
      }
      return cmp_keys_asc(*row_fields, row_repr[a], row_repr[b]) < 0;
    });
  } else {
    std::stable_sort(row_order.begin(), row_order.end(), [&](std::size_t a, std::size_t b) {
      if (row_is_error[a] != row_is_error[b]) {
        return !row_is_error[a];
      }
      return false;
    });
  }

  // -- Sort col groups ----------------------------------------------------
  std::vector<std::size_t> col_order(n_cols);
  for (std::size_t i = 0; i < n_cols; ++i) {
    col_order[i] = i;
  }
  if (col_sort_order != 0) {
    if (col_sort_order != 1 && col_sort_order != -1) {
      return Value::error(ErrorCode::Value);
    }
    const bool descending = (col_sort_order < 0);
    std::stable_sort(col_order.begin(), col_order.end(), [&](std::size_t a, std::size_t b) {
      if (col_is_error[a] != col_is_error[b]) {
        return !col_is_error[a];
      }
      Value va = col_totals[a];
      Value vb = col_totals[b];
      if (!emit_col_totals_row) {
        const ArrayValue* sa = build_group_slice(*values, 0U, col_members[a], arena);
        const ArrayValue* sb = build_group_slice(*values, 0U, col_members[b], arena);
        va =
            (sa != nullptr) ? invoke_aggregator_for_group(agg, sa, arena, registry, ctx) : Value::error(ErrorCode::Num);
        vb =
            (sb != nullptr) ? invoke_aggregator_for_group(agg, sb, arena, registry, ctx) : Value::error(ErrorCode::Num);
      }
      const int c = cmp_value_asc(va, vb);
      if (c != 0) {
        return descending ? (c > 0) : (c < 0);
      }
      return cmp_keys_asc(*col_fields, col_repr[a], col_repr[b]) < 0;
    });
  } else {
    std::stable_sort(col_order.begin(), col_order.end(), [&](std::size_t a, std::size_t b) {
      if (col_is_error[a] != col_is_error[b]) {
        return !col_is_error[a];
      }
      return false;
    });
  }

  // -- Assemble output ----------------------------------------------------
  // Layout:
  //   [optional header row]        : row_label_header | col_keys... | optional "Grand Total"
  //   [optional top totals row]    : "Grand Total"     | col_totals  | grand_total           (when row_total_depth < 0)
  //   per-row-group rows           : row_key           | body cells  | optional row_total
  //   [optional bottom totals row] : "Grand Total"     | col_totals  | grand_total           (when row_total_depth > 0)
  //
  // The position of the row_total ("Grand Total") column is governed by
  // col_total_depth: negative → leftmost data column (immediately after
  // the row-key column), positive → rightmost column.
  const bool grand_total_left = emit_row_totals_col && (col_total_depth < 0);
  const std::uint32_t out_cols =
      1U /*row label*/ + static_cast<std::uint32_t>(n_cols) + (emit_row_totals_col ? 1U : 0U);

  // Helper: render one body row (not header / total) into a vector.
  auto render_body_row = [&](std::size_t rg) {
    std::vector<Value> row(out_cols, Value::blank());
    // Row key cell (col 0).
    row[0] = row_fields->cells[static_cast<std::size_t>(row_repr[rg]) * row_fields->cols];
    // Body cells in sorted col order.
    for (std::size_t ci = 0; ci < n_cols; ++ci) {
      const std::size_t cg = col_order[ci];
      const std::uint32_t out_col_idx = (grand_total_left ? 2U : 1U) + static_cast<std::uint32_t>(ci);
      row[out_col_idx] = body[rg][cg];
    }
    if (emit_row_totals_col) {
      row[grand_total_left ? 1U : (out_cols - 1U)] = row_totals[rg];
    }
    return row;
  };

  // Helper: render the totals row (column totals + grand total).
  auto render_totals_row = [&]() {
    std::vector<Value> row(out_cols, Value::blank());
    row[0] = Value::text(arena.intern("Grand Total"));
    for (std::size_t ci = 0; ci < n_cols; ++ci) {
      const std::size_t cg = col_order[ci];
      const std::uint32_t out_col_idx = (grand_total_left ? 2U : 1U) + static_cast<std::uint32_t>(ci);
      row[out_col_idx] = col_totals[cg];
    }
    if (emit_row_totals_col) {
      row[grand_total_left ? 1U : (out_cols - 1U)] = grand_total;
    }
    return row;
  };

  // Helper: render the header row.
  auto render_header_row = [&]() {
    std::vector<Value> row(out_cols, Value::blank());
    if (field_headers == 1 || field_headers == 3) {
      // Inputs had a header row. Top-left = row_fields header label.
      row[0] = row_fields->cells[0];
    } else {
      // field_headers == 2: synthesise.
      row[0] = Value::text(arena.intern("Field 1"));
    }
    for (std::size_t ci = 0; ci < n_cols; ++ci) {
      const std::size_t cg = col_order[ci];
      // The col-key labels come from col_fields' representative data row.
      const std::uint32_t out_col_idx = (grand_total_left ? 2U : 1U) + static_cast<std::uint32_t>(ci);
      row[out_col_idx] = col_fields->cells[static_cast<std::size_t>(col_repr[cg]) * col_fields->cols];
    }
    if (emit_row_totals_col) {
      row[grand_total_left ? 1U : (out_cols - 1U)] = Value::text(arena.intern("Grand Total"));
    }
    return row;
  };

  std::vector<std::vector<Value>> out_rows;
  out_rows.reserve(n_rows + 3U);
  if (output_emits_header) {
    out_rows.push_back(render_header_row());
  }
  if (emit_col_totals_row && row_total_depth < 0) {
    out_rows.push_back(render_totals_row());
  }
  for (std::size_t ri = 0; ri < n_rows; ++ri) {
    out_rows.push_back(render_body_row(row_order[ri]));
  }
  if (emit_col_totals_row && row_total_depth > 0) {
    out_rows.push_back(render_totals_row());
  }

  if (out_rows.empty()) {
    return Value::error(ErrorCode::Calc);
  }

  const std::uint32_t out_rows_n = static_cast<std::uint32_t>(out_rows.size());
  const std::size_t total_cells = static_cast<std::size_t>(out_rows_n) * static_cast<std::size_t>(out_cols);
  Value* buffer = arena.create_array<Value>(total_cells);
  if (buffer == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::uint32_t r = 0; r < out_rows_n; ++r) {
    for (std::uint32_t c = 0; c < out_cols; ++c) {
      buffer[static_cast<std::size_t>(r) * out_cols + c] = out_rows[r][c];
    }
  }
  ArrayValue* arr = arena.create<ArrayValue>();
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  arr->rows = out_rows_n;
  arr->cols = out_cols;
  arr->cells = buffer;
  return Value::array(arr);
}

}  // namespace eval
}  // namespace formulon
