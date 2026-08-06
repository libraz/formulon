
#include "eval/groupby_pivotby/common.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/array_alloc.h"
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
#include "value_sort_order.h"

namespace formulon {
namespace eval {

namespace {

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

}  // namespace

std::string_view grand_total_label(const EvalContext& ctx) {
  if (ctx.excel_profile().locale == ExcelLocale::kJaJP) {
    return "合計";
  }
  return "Grand Total";
}

std::string subtotal_label(const Value& outer_key, const EvalContext& ctx) {
  // The suffix mirrors the grand-total label of the same locale rather than
  // the PivotTable layer's "総計 / 集計" pair, because GROUPBY / PIVOTBY use
  // their own "Grand Total" / "合計" wording. The ja-JP suffix has not been
  // confirmed against a live Excel run; it is isolated here so a single edit
  // corrects every subtotal row once the goldens are captured.
  const bool ja = ctx.excel_profile().locale == ExcelLocale::kJaJP;
  const std::string_view suffix = ja ? " 合計" : " Total";
  auto key_text = coerce_to_text(outer_key);
  std::string label = key_text ? key_text.value() : std::string();
  label.append(suffix);
  return label;
}

OuterGrouping build_outer_grouping(const ArrayValue& keys, const std::vector<std::uint32_t>& group_repr,
                                   const std::vector<std::vector<std::uint32_t>>& group_rows) {
  OuterGrouping out;
  out.outer_of_group.resize(group_repr.size(), 0U);
  // The outer level is the first key column alone. Outer groups are few
  // relative to the data rows, so a linear scan over the representatives
  // beats standing up a second hash index.
  const std::uint32_t key_cols = keys.cols;
  for (std::size_t g = 0; g < group_repr.size(); ++g) {
    const Value& key = keys.cells[static_cast<std::size_t>(group_repr[g]) * key_cols];
    std::size_t outer = out.repr_of_outer.size();
    for (std::size_t o = 0; o < out.repr_of_outer.size(); ++o) {
      const Value& existing = keys.cells[static_cast<std::size_t>(out.repr_of_outer[o]) * key_cols];
      if (group_cell_equal(existing, key)) {
        outer = o;
        break;
      }
    }
    if (outer == out.repr_of_outer.size()) {
      out.repr_of_outer.push_back(group_repr[g]);
      out.rows_of_outer.emplace_back();
    }
    out.outer_of_group[g] = outer;
    out.rows_of_outer[outer].insert(out.rows_of_outer[outer].end(), group_rows[g].begin(), group_rows[g].end());
  }
  return out;
}

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

bool read_optional_int_in_set(const parser::AstNode& call, std::uint32_t arg_index, std::uint32_t arity,
                              int default_value, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                              const int* allowed, std::size_t count, int* out, Value* out_err) {
  *out = default_value;
  if (arity <= arg_index) {
    return true;
  }
  return read_int_in_set(call.as_call_arg(arg_index), arena, registry, ctx, allowed, count, out, out_err);
}

bool read_optional_int(const parser::AstNode& call, std::uint32_t arg_index, std::uint32_t arity, int default_value,
                       Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx, int* out,
                       Value* out_err) {
  *out = default_value;
  if (arity <= arg_index) {
    return true;
  }
  return read_int(call.as_call_arg(arg_index), arena, registry, ctx, out, out_err);
}

Expected<HeaderLayout, ErrorCode> resolve_header_layout(int field_headers, std::uint32_t input_rows) {
  HeaderLayout layout;
  layout.inputs_have_header = (field_headers == 1 || field_headers == 3);
  layout.output_emits_header = (field_headers == 1 || field_headers == 2 || field_headers == 3);
  if (layout.inputs_have_header && input_rows < 1U) {
    return Expected<HeaderLayout, ErrorCode>::Err(ErrorCode::Value);
  }
  layout.data_start_row = layout.inputs_have_header ? 1U : 0U;
  if (input_rows < layout.data_start_row) {
    return Expected<HeaderLayout, ErrorCode>::Err(ErrorCode::Calc);
  }
  layout.data_row_count = input_rows - layout.data_start_row;
  return Expected<HeaderLayout, ErrorCode>::Ok(layout);
}

bool read_filter_mask(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx, std::uint32_t data_row_count, std::vector<bool>* include_row,
                      Value* out_err) {
  const ArrayValue* mask = read_array_arg(node, arena, registry, ctx, out_err);
  if (mask == nullptr) {
    return false;
  }
  const std::uint32_t mask_n = (mask->rows >= mask->cols) ? mask->rows : mask->cols;
  if (mask->rows != 1U && mask->cols != 1U) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  if (mask_n != data_row_count) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  for (std::uint32_t i = 0; i < data_row_count; ++i) {
    const Value& cell = mask->cells[i];
    if (cell.is_error()) {
      *out_err = cell;
      return false;
    }
    auto coerced = coerce_to_bool(cell);
    if (!coerced) {
      *out_err = Value::error(coerced.error());
      return false;
    }
    (*include_row)[i] = coerced.value();
  }
  return true;
}

std::vector<std::uint32_t> collect_included_rows(const std::vector<bool>& include_row, std::uint32_t data_start_row) {
  std::vector<std::uint32_t> rows;
  rows.reserve(include_row.size());
  for (std::uint32_t i = 0; i < include_row.size(); ++i) {
    if (include_row[i]) {
      rows.push_back(data_start_row + i);
    }
  }
  return rows;
}

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

std::string normalized_group_key(const ArrayValue& keys, std::uint32_t row) {
  std::string out;
  out.reserve(static_cast<std::size_t>(keys.cols) * 12U);
  for (std::uint32_t c = 0; c < keys.cols; ++c) {
    const Value& v = keys.cells[static_cast<std::size_t>(row) * keys.cols + c];
    out.push_back(static_cast<char>(v.kind()));
    switch (v.kind()) {
      case ValueKind::Blank:
        break;
      case ValueKind::Number: {
        double number = v.as_number();
        if (std::isnan(number)) {
          // NaN never compares equal, including to itself.
          out.append("row");
          out.append(std::to_string(row));
          break;
        }
        number = number == 0.0 ? 0.0 : number;
        std::uint64_t bits = 0;
        static_assert(sizeof(bits) == sizeof(number));
        std::memcpy(&bits, &number, sizeof(bits));
        out.append(reinterpret_cast<const char*>(&bits), sizeof(bits));
        break;
      }
      case ValueKind::Bool:
        out.push_back(v.as_boolean() ? '\x01' : '\x00');
        break;
      case ValueKind::Error:
        out.append(std::to_string(static_cast<std::uint16_t>(v.as_error())));
        out.push_back('\0');
        break;
      case ValueKind::Text: {
        const std::string folded = fold_jp_text(v.as_text());
        out.append(std::to_string(folded.size()));
        out.push_back(':');
        out.append(folded);
        break;
      }
      default:
        // Array / Ref / Lambda are intentionally never equal in
        // group_cell_equal, so each row must form its own group.
        out.append("row");
        out.append(std::to_string(row));
        break;
    }
    out.push_back('\xff');
  }
  return out;
}

bool row_key_is_error(const ArrayValue& keys, std::uint32_t row) {
  for (std::uint32_t c = 0; c < keys.cols; ++c) {
    const Value& v = keys.cells[static_cast<std::size_t>(row) * keys.cols + c];
    if (v.is_error()) {
      return true;
    }
  }
  return false;
}

const ArrayValue* build_group_slice(const ArrayValue& values, std::uint32_t value_col,
                                    const std::vector<std::uint32_t>& row_indices, Arena& arena) {
  const std::uint32_t n = static_cast<std::uint32_t>(row_indices.size());
  if (n == 0U) {
    return nullptr;
  }
  Value* cells = nullptr;
  ArrayValue* arr = allocate_array_value(n, 1U, arena, cells, kMaxDerivedArrayCells);
  if (arr == nullptr) {
    return nullptr;
  }
  for (std::uint32_t i = 0; i < n; ++i) {
    cells[i] = values.cells[static_cast<std::size_t>(row_indices[i]) * values.cols + value_col];
  }
  return arr;
}

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

std::vector<Value> aggregate_value_columns(const ArrayValue& values, std::uint32_t val_cols,
                                           const std::vector<std::uint32_t>& row_indices, const AggregatorRef& agg,
                                           Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                                           ErrorCode empty_error) {
  std::vector<Value> cells(val_cols, Value::blank());
  if (row_indices.empty()) {
    for (std::uint32_t v = 0; v < val_cols; ++v) {
      cells[v] = Value::error(empty_error);
    }
    return cells;
  }
  for (std::uint32_t v = 0; v < val_cols; ++v) {
    const ArrayValue* slice = build_group_slice(values, v, row_indices, arena);
    if (slice == nullptr) {
      cells[v] = Value::error(ErrorCode::Num);
      continue;
    }
    cells[v] = invoke_aggregator_for_group(agg, slice, arena, registry, ctx);
  }
  return cells;
}

void emit_row(std::vector<std::vector<Value>>* rows, const std::vector<Value>& row) {
  rows->push_back(row);
}

Value rows_to_array_value(const std::vector<std::vector<Value>>& rows, std::uint32_t out_cols, Arena& arena) {
  if (rows.empty()) {
    return Value::error(ErrorCode::Calc);
  }
  const std::uint32_t out_rows_n = static_cast<std::uint32_t>(rows.size());
  Value* buffer = nullptr;
  ArrayValue* arr = allocate_array_value(out_rows_n, out_cols, arena, buffer, kMaxDerivedArrayCells);
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::uint32_t r = 0; r < out_rows_n; ++r) {
    for (std::uint32_t c = 0; c < out_cols; ++c) {
      buffer[static_cast<std::size_t>(r) * out_cols + c] = rows[r][c];
    }
  }
  return Value::array(arr);
}

int cmp_value_asc(const Value& a, const Value& b) {
  // Cross-kind ordering buckets come from the shared Excel rank
  // (Number < Text < Bool < Error < Blank), so GROUPBY / SORT and the pivot
  // comparator cannot diverge on the relative position of Bool vs Text.
  const int ba = excel_kind_rank(a.kind());
  const int bb = excel_kind_rank(b.kind());
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

}  // namespace eval
}  // namespace formulon
