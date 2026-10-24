
#include "eval/groupby_pivotby/common.h"

#include <algorithm>
#include <cassert>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/coerce.h"
#include "eval/collation.h"
#include "eval/dynamic_array/common.h"
#include "eval/eval_context.h"
#include "eval/eval_profile_scope.h"
#include "eval/function_registry.h"
#include "eval/lambda_value.h"
#include "eval/lazy_impls.h"
#include "eval/name_env_resolve.h"
#include "eval/omitted_arg.h"
#include "eval/range_args.h"
#include "eval/shape_ops_lazy.h"
#include "eval/tree_walker/dispatch.h"
#include "eval/tree_walker_lazy_table.h"
#include "excel_locale.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"
#include "value_sort_order.h"

namespace formulon {
namespace eval {

namespace {

// Arguments an aggregator receives per group: two, `(group_values,
// all_values)`, for a LAMBDA or built-in that requires exactly two; else one.
std::uint32_t aggregator_arity(const LambdaValue& lv, const FunctionRegistry& registry) {
  std::uint32_t required = lv.param_count - lv.optional_count;
  if (!lv.builtin.empty()) {
    if (find_lazy_impl(lv.builtin) != nullptr) {
      required = find_lazy_min_arity(lv.builtin);
    } else {
      const FunctionDef* def = registry.lookup(lv.builtin);
      required = def != nullptr ? def->min_arity : 0U;
    }
  }
  return required == 2U ? 2U : 1U;
}

}  // namespace

std::string_view grand_total_label(const EvalContext& ctx) {
  return locale_facts(ctx.excel_profile()).grand_total;
}

std::string row_field_label(const EvalContext& ctx, std::uint32_t n) {
  const std::string_view prefix = locale_facts(ctx.excel_profile()).row_field_prefix;
  return std::string(prefix) + std::to_string(n);
}

std::string column_field_label(const EvalContext& ctx, std::uint32_t n) {
  const std::string_view prefix = locale_facts(ctx.excel_profile()).column_field_prefix;
  return std::string(prefix) + std::to_string(n);
}

std::string value_label(const EvalContext& ctx, std::uint32_t n) {
  const std::string_view prefix = locale_facts(ctx.excel_profile()).value_field_prefix;
  return std::string(prefix) + std::to_string(n);
}

std::string_view hierarchy_grand_total_label(const EvalContext& ctx) {
  return locale_facts(ctx.excel_profile()).hierarchy_grand_total;
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

bool resolve_aggregator(const parser::AstNode& arg, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, Aggregator* out, Value* out_err) {
  // The arity depends on the callable, so read it before resolving at that
  // arity; `resolve_callable` reports every error the peek skips.
  const Value peek = eval_node(arg, arena, registry, ctx);
  const std::uint32_t arity = peek.is_lambda() ? aggregator_arity(*peek.as_lambda(), registry) : 1U;
  out->fn = resolve_callable(arg, arity, arena, registry, ctx, out_err);
  out->arity = arity;
  out->all_by_column.clear();
  return out->fn != nullptr;
}

bool bind_aggregator_totals(const ArrayValue& values, const std::vector<std::uint32_t>& all_rows, Arena& arena,
                            const Aggregator& agg, Aggregator* out) {
  *out = agg;
  if (agg.arity < 2U) {
    return true;
  }
  out->all_by_column.assign(values.cols, nullptr);
  for (std::uint32_t c = 0; c < values.cols; ++c) {
    out->all_by_column[c] = build_group_slice(values, c, all_rows, arena);
    if (out->all_by_column[c] == nullptr) {
      return false;
    }
  }
  return true;
}

const ArrayValue* read_array_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                                 const EvalContext& ctx, Value* out_err) {
  const ArrayValue* out = nullptr;
  return resolve_array_value(node, arena, registry, ctx, &out, out_err) ? out : nullptr;
}

bool read_int(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
              int* out, Value* out_err) {
  double n = 0.0;
  if (!dynamic_array::eval_number_arg(node, arena, registry, ctx, n, *out_err)) {
    return false;
  }
  if (std::isnan(n) || std::isinf(n)) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  *out = truncate_saturated_int(n);
  return true;
}

bool read_int_in_set(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx, const int* allowed, std::size_t count, int* out, Value* out_err) {
  int truncated = 0;
  if (!read_int(node, arena, registry, ctx, &truncated, out_err)) {
    return false;
  }
  for (std::size_t i = 0; i < count; ++i) {
    if (allowed[i] == truncated) {
      *out = truncated;
      return true;
    }
  }
  *out_err = Value::error(ErrorCode::Value);
  return false;
}

bool read_optional_int_in_set(const parser::AstNode& call, std::uint32_t arg_index, std::uint32_t arity,
                              int default_value, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                              const int* allowed, std::size_t count, int* out, Value* out_err) {
  *out = default_value;
  if (arity <= arg_index || is_omitted_arg(call.as_call_arg(arg_index))) {
    return true;
  }
  return read_int_in_set(call.as_call_arg(arg_index), arena, registry, ctx, allowed, count, out, out_err);
}

bool read_field_headers(const parser::AstNode& call, std::uint32_t arg_index, std::uint32_t arity, int default_value,
                        Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx, int* out,
                        Value* out_err) {
  static constexpr int kFieldHeaders[] = {0, 1, 2, 3};
  return read_optional_int_in_set(call, arg_index, arity, default_value, arena, registry, ctx, kFieldHeaders,
                                  sizeof(kFieldHeaders) / sizeof(kFieldHeaders[0]), out, out_err);
}

bool read_total_depth(const parser::AstNode& call, std::uint32_t arg_index, std::uint32_t arity, int default_value,
                      Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx, int* out,
                      Value* out_err) {
  static constexpr int kTotalDepths[] = {-2, -1, 0, 1, 2};
  return read_optional_int_in_set(call, arg_index, arity, default_value, arena, registry, ctx, kTotalDepths,
                                  sizeof(kTotalDepths) / sizeof(kTotalDepths[0]), out, out_err);
}

bool read_optional_sort_spec(const parser::AstNode& call, std::uint32_t arg_index, std::uint32_t arity, Arena& arena,
                             const FunctionRegistry& registry, const EvalContext& ctx, std::vector<int>* out,
                             Value* out_err) {
  out->clear();
  if (arity <= arg_index || is_omitted_arg(call.as_call_arg(arg_index))) {
    return true;
  }
  const ArrayValue* spec = nullptr;
  if (!resolve_array_value(call.as_call_arg(arg_index), arena, registry, ctx, &spec, out_err)) {
    return false;
  }
  const std::size_t n = static_cast<std::size_t>(spec->rows) * spec->cols;
  out->reserve(n);
  for (std::size_t i = 0; i < n; ++i) {
    const Value& cell = spec->cells[i];
    if (cell.is_error()) {
      *out_err = cell;
      return false;
    }
    auto number = coerce_to_number(cell);
    if (!number) {
      *out_err = Value::error(number.error());
      return false;
    }
    const int index = std::isfinite(number.value()) ? truncate_saturated_int(number.value()) : 0;
    if (index == 0) {
      *out_err = Value::error(ErrorCode::Value);
      return false;
    }
    out->push_back(index);
  }
  return true;
}

bool plan_group_sort(const std::vector<int>& spec, std::uint32_t key_cols, std::uint32_t val_cols, GroupSort* out) {
  out->key_descending.assign(key_cols, false);
  out->value_col = 0U;
  out->value_descending = false;
  for (const int n : spec) {
    const std::int64_t magnitude = n > 0 ? static_cast<std::int64_t>(n) : -static_cast<std::int64_t>(n);
    if (magnitude > static_cast<std::int64_t>(key_cols) + val_cols || (magnitude > key_cols && spec.size() != 1U)) {
      return false;
    }
    const auto column = static_cast<std::uint32_t>(magnitude);
    if (column > key_cols) {
      out->value_col = column - key_cols;
      out->value_descending = n < 0;
    } else {
      out->key_descending[column - 1U] = n < 0;
    }
  }
  return true;
}

void order_groups(const GroupSort& plan, const ArrayValue& keys, const std::vector<std::uint32_t>& repr,
                  const std::vector<bool>& is_error, const std::vector<Value>& value_keys,
                  std::vector<std::size_t>* order) {
  order->resize(repr.size());
  for (std::size_t i = 0; i < order->size(); ++i) {
    (*order)[i] = i;
  }
  const auto cmp_keys = [&](std::size_t a, std::size_t b) {
    for (std::uint32_t c = 0; c < keys.cols; ++c) {
      const int r = cmp_value_asc(keys.cells[static_cast<std::size_t>(repr[a]) * keys.cols + c],
                                  keys.cells[static_cast<std::size_t>(repr[b]) * keys.cols + c]);
      if (r != 0) {
        return plan.key_descending[c] ? -r : r;
      }
    }
    return 0;
  };
  sort_group_order(*order, [&](std::size_t a, std::size_t b) {
    // Error-keyed groups always go last regardless of direction, matching
    // Mac Excel's "errors trail" surface for the dynamic-array sort family.
    if (is_error[a] != is_error[b]) {
      return !is_error[a];
    }
    if (plan.value_col != 0U) {
      const int c = cmp_value_asc(value_keys[a], value_keys[b]);
      if (c != 0) {
        return plan.value_descending ? (c > 0) : (c < 0);
      }
      return cmp_keys_asc(keys, repr[a], repr[b]) < 0;
    }
    return cmp_keys(a, b) < 0;
  });
}

int resolve_auto_field_headers(int field_headers, const ArrayValue& values) {
  if (field_headers != kFieldHeadersAuto) {
    return field_headers;
  }
  if (values.rows < 2U) {
    return 0;
  }
  for (std::uint32_t c = 0; c < values.cols; ++c) {
    if (values.cells[c].is_text() && !values.cells[values.cols + c].is_text()) {
      return 1;
    }
  }
  return 0;
}

Expected<HeaderLayout, ErrorCode> resolve_header_layout(int field_headers, std::uint32_t input_rows) {
  HeaderLayout layout;
  layout.inputs_have_header = (field_headers == 1 || field_headers == 3);
  layout.output_emits_header = (field_headers == 2 || field_headers == 3);
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
  // Decide 1D-ness from `node`'s DECLARED shape when it is a static
  // reference, not from the walked array `read_array_arg` returns: a
  // multi-column whole reference populated in only one row walks as a
  // 1-row array, which would misread it as a vector even though its
  // declared shape is 2D. Mirrors the same rule IFS / XLOOKUP / MATCH /
  // LOOKUP already apply via `declared_range_rect`.
  std::uint32_t declared_rows = mask->rows;
  std::uint32_t declared_cols = mask->cols;
  (void)static_reference_shape(node, ctx, &declared_rows, &declared_cols);
  if (declared_rows != 1U && declared_cols != 1U) {
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

bool read_layout_and_mask(const parser::AstNode& call, std::uint32_t filter_arg_index, std::uint32_t arity,
                          int field_headers, std::uint32_t input_rows, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx, HeaderLayout* layout, std::vector<bool>* include_row,
                          Value* out_err) {
  auto layout_result = resolve_header_layout(field_headers, input_rows);
  if (!layout_result) {
    *out_err = Value::error(layout_result.error());
    return false;
  }
  *layout = layout_result.take();
  include_row->assign(layout->data_row_count, true);
  if (arity == filter_arg_index + 1U) {
    return read_filter_mask(call.as_call_arg(filter_arg_index), arena, registry, ctx, layout->data_row_count,
                            include_row, out_err);
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

// Excel-canonical cell equality for GROUPBY group keys. Text compares at the
// dynamic collation strength of the evaluating profile. Numbers compare with
// `==`, so `+0.0` and `-0.0` fall in one group. Cross-kind pairs are never
// equal: `Number 0` and `Bool FALSE` form distinct groups.
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
      assert(eval_profile_scope_active());
      return collate(a.as_text(), b.as_text(), current_eval_profile(), CollationStrength::kDynamic) == 0;
    default:
      // Array / Ref / Lambda never merge, so each row forms its own group.
      return false;
  }
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
        assert(eval_profile_scope_active());
        const std::string key = collation_key(v.as_text(), current_eval_profile(), CollationStrength::kDynamic);
        out.append(std::to_string(key.size()));
        out.push_back(':');
        out.append(key);
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

std::size_t find_or_add_group(const ArrayValue& keys, std::uint32_t row,
                              std::vector<std::uint32_t>* representative_rows,
                              std::vector<std::vector<std::uint32_t>>* member_rows, std::vector<bool>* is_error_group,
                              std::unordered_map<std::string, std::size_t>* index) {
  const std::string key = normalized_group_key(keys, row);
  const auto existing = index->find(key);
  if (existing != index->end()) {
    (*member_rows)[existing->second].push_back(row);
    return existing->second;
  }
  const std::size_t group = representative_rows->size();
  index->emplace(key, group);
  representative_rows->push_back(row);
  member_rows->push_back(std::vector<std::uint32_t>{row});
  is_error_group->push_back(row_key_is_error(keys, row));
  return group;
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

Value invoke_aggregator_for_group(const Aggregator& agg, std::uint32_t value_col, const ArrayValue* slice, Arena& arena,
                                  const FunctionRegistry& registry, const EvalContext& ctx) {
  const ArrayValue* all = agg.arity == 2U ? agg.all_by_column[value_col] : nullptr;
  const parser::AstNode* slice_ast = array_literal_ast(slice, arena);
  const parser::AstNode* all_ast = all != nullptr ? array_literal_ast(all, arena) : nullptr;
  if (slice_ast == nullptr || (agg.arity == 2U && all_ast == nullptr)) {
    return Value::error(ErrorCode::Num);
  }
  // The slice is bound with both its Value and a synthetic ArrayLiteral
  // AST so a body written as `SUM(v)` -- or the eta-expanded `SUM` --
  // flattens it through the dispatcher's ArrayLiteral branch. Errors are
  // not filtered: whatever the body produced -- including an error -- lands
  // in this group's cell.
  const Value args[2] = {Value::array(slice), all != nullptr ? Value::array(all) : Value::blank()};
  const parser::AstNode* ast_args[2] = {slice_ast, all_ast};
  const Value res = invoke_lambda_values_with_ast(agg.fn, agg.arity, args, ast_args, arena, registry, ctx);
  if (res.is_array()) {
    return Value::error(ErrorCode::Calc);
  }
  if (res.is_lambda()) {
    return Value::error(ErrorCode::Calc);
  }
  return res;
}

std::vector<Value> aggregate_value_columns(const ArrayValue& values, std::uint32_t val_cols,
                                           const std::vector<std::uint32_t>& row_indices, const Aggregator& agg,
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
    cells[v] = invoke_aggregator_for_group(agg, v, slice, arena, registry, ctx);
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
      return na > nb ? 1 : 0;
    }
    case ValueKind::Text:
      assert(eval_profile_scope_active());
      return collate(a.as_text(), b.as_text(), current_eval_profile(), CollationStrength::kDynamic);
    case ValueKind::Bool:
      if (a.as_boolean() == b.as_boolean()) {
        return 0;
      }
      return a.as_boolean() ? 1 : -1;
    case ValueKind::Error: {
      const auto ea = static_cast<int>(a.as_error());
      const auto eb = static_cast<int>(b.as_error());
      if (ea < eb) {
        return -1;
      }
      return ea > eb ? 1 : 0;
    }
    default:
      // Blank, and kinds with no in-kind order.
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
