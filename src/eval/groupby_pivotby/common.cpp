
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
#include "eval/scalar_args.h"
#include "eval/shape_ops_lazy.h"
#include "eval/tree_walker/dispatch.h"
#include "eval/tree_walker_lazy_table.h"
#include "excel_locale.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/index_sort.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace {

// Arguments an aggregator receives per group: two, `(group_values,
// all_values)`, for a LAMBDA that requires exactly two or a built-in whose
// minimum is two; else one. Excel calls IF with the group values alone.
std::uint32_t aggregator_arity(const LambdaValue& lv, const FunctionRegistry& registry) {
  std::uint32_t required = lv.param_count - lv.optional_count;
  if (!lv.builtin.empty()) {
    std::uint32_t max_arity = 0U;
    if (strings::case_insensitive_eq(lv.builtin, "IF")) {
      required = 1U;
    } else if (!find_lazy_arity(lv.builtin, &required, &max_arity)) {
      const FunctionDef* def = registry.lookup(lv.builtin);
      required = def != nullptr ? def->min_arity : 0U;
    }
  }
  return required == 2U ? 2U : 1U;
}

std::string numbered_label(std::string_view prefix, std::uint32_t n) {
  return std::string(prefix) + std::to_string(n);
}

// Appends the canonical key of one key cell. Cells of one kind get equal keys
// when they are equal (text at dynamic collation strength, `+0.0` and `-0.0`
// alike); cross-kind cells never match. `row` makes the key unique for cells
// that never compare equal (NaN, Array / Ref / Lambda).
void append_cell_group_key(std::string* out, const Value& v, std::uint32_t row) {
  out->push_back(static_cast<char>(v.kind()));
  switch (v.kind()) {
    case ValueKind::Blank:
      break;
    case ValueKind::Number: {
      double number = v.as_number();
      if (std::isnan(number)) {
        // NaN never compares equal, including to itself.
        out->append("row");
        out->append(std::to_string(row));
        break;
      }
      number = number == 0.0 ? 0.0 : number;
      std::uint64_t bits = 0;
      static_assert(sizeof(bits) == sizeof(number));
      std::memcpy(&bits, &number, sizeof(bits));
      out->append(reinterpret_cast<const char*>(&bits), sizeof(bits));
      break;
    }
    case ValueKind::Bool:
      out->push_back(v.as_boolean() ? '\x01' : '\x00');
      break;
    case ValueKind::Error:
      out->append(std::to_string(static_cast<std::uint16_t>(v.as_error())));
      out->push_back('\0');
      break;
    case ValueKind::Text: {
      assert(eval_profile_scope_active());
      const std::string key = collation_key(v.as_text(), current_eval_profile(), CollationStrength::kDynamic);
      out->append(std::to_string(key.size()));
      out->push_back(':');
      out->append(key);
      break;
    }
    default:
      // Array / Ref / Lambda never merge, so each row forms its own group.
      out->append("row");
      out->append(std::to_string(row));
      break;
  }
  out->push_back('\xff');
}

}  // namespace

std::string_view grand_total_label(const EvalContext& ctx) {
  return locale_facts(ctx.excel_profile()).grand_total;
}

namespace {

// Generated header labels for `field_headers == 2`, e.g. "行フィールド 1" in
// ja-JP; `n` is 1-based.
std::string row_field_label(const EvalContext& ctx, std::uint32_t n) {
  return numbered_label(locale_facts(ctx.excel_profile()).row_field_prefix, n);
}

std::string column_field_label(const EvalContext& ctx, std::uint32_t n) {
  return numbered_label(locale_facts(ctx.excel_profile()).column_field_prefix, n);
}

std::string value_label(const EvalContext& ctx, std::uint32_t n) {
  return numbered_label(locale_facts(ctx.excel_profile()).value_field_prefix, n);
}

}  // namespace

Value header_label(Arena& arena, const EvalContext& ctx, int field_headers, HeaderField kind,
                   const Value* input_header_cells, std::uint32_t index) {
  if (field_headers == 3) {
    return input_header_cells[index];
  }
  switch (kind) {
    case HeaderField::kRow:
      return Value::text(arena.intern(row_field_label(ctx, index + 1U)));
    case HeaderField::kColumn:
      return Value::text(arena.intern(column_field_label(ctx, index + 1U)));
    case HeaderField::kValue:
      break;
  }
  return Value::text(arena.intern(value_label(ctx, index + 1U)));
}

namespace {

// ja-JP names the two levels apart ("総計" above the subtotals' "合計");
// locales that draw no such distinction reuse `grand_total_label`.
std::string_view hierarchy_grand_total_label(const EvalContext& ctx) {
  return locale_facts(ctx.excel_profile()).hierarchy_grand_total;
}

}  // namespace

OuterGrouping build_prefix_grouping(const ArrayValue& keys, const std::vector<std::uint32_t>& group_repr,
                                    const std::vector<std::vector<std::uint32_t>>& group_rows,
                                    std::uint32_t prefix_cols) {
  OuterGrouping out;
  out.outer_of_group.resize(group_repr.size(), 0U);
  const std::uint32_t key_cols = keys.cols;
  std::unordered_map<std::string, std::size_t> outer_index;
  std::string key;
  for (std::size_t g = 0; g < group_repr.size(); ++g) {
    key.clear();
    for (std::uint32_t c = 0; c < prefix_cols; ++c) {
      append_cell_group_key(&key, keys.cells[static_cast<std::size_t>(group_repr[g]) * key_cols + c], group_repr[g]);
    }
    const auto [it, inserted] = outer_index.try_emplace(key, out.repr_of_outer.size());
    const std::size_t outer = it->second;
    if (inserted) {
      out.repr_of_outer.push_back(group_repr[g]);
      out.first_group_of_outer.push_back(g);
      out.rows_of_outer.emplace_back();
    }
    out.outer_of_group[g] = outer;
    out.rows_of_outer[outer].insert(out.rows_of_outer[outer].end(), group_rows[g].begin(), group_rows[g].end());
  }
  return out;
}

AxisHierarchy build_axis_hierarchy(const ArrayValue& keys, const std::vector<std::uint32_t>& group_repr,
                                   const std::vector<std::vector<std::uint32_t>>& group_rows, bool with_prefixes) {
  AxisHierarchy out;
  out.keys = keys.cols;
  if (with_prefixes) {
    for (std::uint32_t level = 1U; level < keys.cols; ++level) {
      out.prefix.push_back(build_prefix_grouping(keys, group_repr, group_rows, level));
    }
  }
  return out;
}

std::size_t axis_group_at(const AxisHierarchy& hierarchy, std::uint32_t level, std::size_t flat_group) {
  if (level == 0U) {
    return 0U;
  }
  return level >= hierarchy.keys ? flat_group : hierarchy.prefix[level - 1U].outer_of_group[flat_group];
}

AxisEntity axis_parent(const AxisHierarchy& hierarchy, AxisEntity e) {
  if (e.level <= 1U) {
    return AxisEntity{};
  }
  const std::size_t flat_group =
      e.level >= hierarchy.keys ? e.index : hierarchy.prefix[e.level - 1U].first_group_of_outer[e.index];
  return AxisEntity{e.level - 1U, axis_group_at(hierarchy, e.level - 1U, flat_group)};
}

std::uint32_t axis_entity_repr_row(const AxisHierarchy& hierarchy, const std::vector<std::uint32_t>& group_repr,
                                   AxisEntity e) {
  return e.level >= hierarchy.keys ? group_repr[e.index] : hierarchy.prefix[e.level - 1U].repr_of_outer[e.index];
}

const std::vector<std::uint32_t>& axis_entity_rows(const AxisHierarchy& hierarchy,
                                                   const std::vector<std::vector<std::uint32_t>>& group_rows,
                                                   const std::vector<std::uint32_t>& all_rows, AxisEntity e) {
  if (e.level == 0U) {
    return all_rows;
  }
  return e.level >= hierarchy.keys ? group_rows[e.index] : hierarchy.prefix[e.level - 1U].rows_of_outer[e.index];
}

namespace {

// Appends the nodes of the flat groups `groups` (in output order, sharing a
// prefix of `level - 1` keys): one block per `level`-prefix group.
void plan_axis_block(const AxisHierarchy& hierarchy, const std::vector<std::size_t>& groups, std::uint32_t level,
                     int depth, std::vector<AxisEntity>* out) {
  if (level >= hierarchy.keys) {
    for (const std::size_t g : groups) {
      out->push_back(AxisEntity{hierarchy.keys, g});
    }
    return;
  }
  const OuterGrouping& grouping = hierarchy.prefix[level - 1U];
  std::unordered_map<std::size_t, std::size_t> block_of;
  std::vector<std::size_t> block_ids;
  std::vector<std::vector<std::size_t>> blocks;
  for (const std::size_t g : groups) {
    const std::size_t id = grouping.outer_of_group[g];
    const auto [it, inserted] = block_of.try_emplace(id, blocks.size());
    if (inserted) {
      block_ids.push_back(id);
      blocks.emplace_back();
    }
    blocks[it->second].push_back(g);
  }
  const bool subtotal = static_cast<int>(level) < (depth < 0 ? -depth : depth);
  for (std::size_t b = 0; b < blocks.size(); ++b) {
    if (subtotal && depth < 0) {
      out->push_back(AxisEntity{level, block_ids[b]});
    }
    plan_axis_block(hierarchy, blocks[b], level + 1U, depth, out);
    if (subtotal && depth > 0) {
      out->push_back(AxisEntity{level, block_ids[b]});
    }
  }
}

}  // namespace

std::vector<AxisEntity> plan_axis_slots(const AxisHierarchy& hierarchy, const std::vector<std::size_t>& order,
                                        int depth) {
  std::vector<AxisEntity> slots;
  slots.reserve(order.size());
  if (depth > -2 && depth < 2) {
    for (const std::size_t g : order) {
      slots.push_back(AxisEntity{hierarchy.keys, g});
    }
  } else {
    plan_axis_block(hierarchy, order, 1U, depth, &slots);
  }
  return slots;
}

bool total_depth_fits(int depth, std::uint32_t levels) {
  const std::int64_t magnitude = depth < 0 ? -static_cast<std::int64_t>(depth) : depth;
  return magnitude <= 1 || magnitude <= static_cast<std::int64_t>(levels);
}

std::string_view total_label(const EvalContext& ctx, bool with_subtotals) {
  return with_subtotals ? hierarchy_grand_total_label(ctx) : grand_total_label(ctx);
}

bool resolve_aggregator(const parser::AstNode& arg, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, Aggregator* out, Value* out_err) {
  // The arity depends on the callable, so read it before resolving at that
  // arity; `resolve_callable_value` reports every error the peek skips.
  const Value peek = eval_node(arg, arena, registry, ctx);
  const std::uint32_t arity = peek.is_lambda() ? aggregator_arity(*peek.as_lambda(), registry) : 1U;
  out->fn = resolve_callable_value(peek, arity, arena, out_err);
  out->arity = arity;
  out->all_by_column.clear();
  out->all_ast_by_column.clear();
  return out->fn != nullptr;
}

bool bind_aggregator_totals(const ArrayValue& values, const std::vector<std::uint32_t>& all_rows, Arena& arena,
                            const Aggregator& agg, Aggregator* out) {
  *out = agg;
  if (agg.arity < 2U) {
    return true;
  }
  out->all_by_column.assign(values.cols, nullptr);
  out->all_ast_by_column.assign(values.cols, nullptr);
  for (std::uint32_t c = 0; c < values.cols; ++c) {
    out->all_by_column[c] = build_group_slice(values, c, all_rows, arena);
    if (out->all_by_column[c] == nullptr) {
      return false;
    }
    out->all_ast_by_column[c] = array_literal_ast(out->all_by_column[c], arena);
  }
  return true;
}

const ArrayValue* read_array_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                                 const EvalContext& ctx, Value* out_err) {
  const ArrayValue* out = nullptr;
  return resolve_array_value(node, arena, registry, ctx, &out, out_err) ? out : nullptr;
}

namespace {

// Reads a scalar integer argument truncated toward zero, then validates it is
// a member of `allowed`; on failure writes the scalar error to `*out_err`.
bool read_int_in_set(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx, const int* allowed, std::size_t count, int* out, Value* out_err) {
  double n = 0.0;
  if (!dynamic_array::eval_number_arg(node, arena, registry, ctx, n, *out_err)) {
    return false;
  }
  if (std::isnan(n) || std::isinf(n)) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  const int truncated = truncate_saturated_int(n);
  for (std::size_t i = 0; i < count; ++i) {
    if (allowed[i] == truncated) {
      *out = truncated;
      return true;
    }
  }
  *out_err = Value::error(ErrorCode::Value);
  return false;
}

}  // namespace

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
  *out = default_value;
  if (arity <= arg_index || is_omitted_arg(call.as_call_arg(arg_index))) {
    return true;
  }
  double n = 0.0;
  if (!dynamic_array::eval_number_arg(call.as_call_arg(arg_index), arena, registry, ctx, n, *out_err)) {
    return false;
  }
  if (!std::isfinite(n) || n != std::trunc(n)) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  *out = truncate_saturated_int(n);
  return true;
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

bool plan_group_sort(const std::vector<int>& spec, std::uint32_t key_cols, std::uint32_t val_cols, bool flat_table,
                     GroupSort* out) {
  out->key_descending.assign(key_cols, false);
  out->key_order.clear();
  out->value_col = 0U;
  out->value_descending = false;
  std::vector<bool> named(key_cols, false);
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
      if (flat_table && !named[column - 1U]) {
        named[column - 1U] = true;
        out->key_order.push_back(column - 1U);
      }
    }
  }
  for (std::uint32_t c = 0; c < key_cols; ++c) {
    if (!named[c]) {
      out->key_order.push_back(c);
    }
  }
  return true;
}

std::vector<OuterSortLevel> build_outer_sort_levels(const ArrayValue& keys, const std::vector<std::uint32_t>& repr,
                                                    const std::vector<std::vector<std::uint32_t>>& members,
                                                    const ArrayValue& values, std::uint32_t value_col,
                                                    const std::vector<const Aggregator*>& agg_of_group, Arena& arena,
                                                    const FunctionRegistry& registry, const EvalContext& ctx) {
  std::vector<OuterSortLevel> levels;
  for (std::uint32_t prefix = 1U; prefix < keys.cols; ++prefix) {
    const OuterGrouping grouping = build_prefix_grouping(keys, repr, members, prefix);
    const std::size_t outer_count = grouping.rows_of_outer.size();
    std::vector<const Aggregator*> agg_of_outer(outer_count, nullptr);
    for (std::size_t g = 0; g < repr.size(); ++g) {
      const std::size_t outer = grouping.outer_of_group[g];
      if (agg_of_outer[outer] == nullptr) {
        agg_of_outer[outer] = agg_of_group[g];
      }
    }
    std::vector<Value> outer_values;
    outer_values.reserve(outer_count);
    for (std::size_t outer = 0; outer < outer_count; ++outer) {
      const ArrayValue* slice = build_group_slice(values, value_col, grouping.rows_of_outer[outer], arena);
      outer_values.push_back(
          slice != nullptr ? invoke_aggregator_for_group(*agg_of_outer[outer], value_col, slice, arena, registry, ctx)
                           : Value::error(ErrorCode::Num));
    }
    OuterSortLevel level;
    level.prefix_of_group = grouping.outer_of_group;
    level.value_of_group.reserve(repr.size());
    for (const std::size_t outer : grouping.outer_of_group) {
      level.value_of_group.push_back(outer_values[outer]);
    }
    levels.push_back(std::move(level));
  }
  return levels;
}

void order_groups(const GroupSort& plan, const ArrayValue& keys, const std::vector<std::uint32_t>& repr,
                  const std::vector<bool>& is_error, const std::vector<Value>& value_keys,
                  const std::vector<OuterSortLevel>& outer_levels, std::vector<std::size_t>* order) {
  // Collation keys are computed once per representative cell rather than on
  // every comparison.
  const auto key_cell = [&](std::size_t g, std::uint32_t c) -> const Value& {
    return keys.cells[static_cast<std::size_t>(repr[g]) * keys.cols + c];
  };
  std::vector<std::string> key_text(repr.size() * keys.cols);
  for (std::size_t g = 0; g < repr.size(); ++g) {
    for (std::uint32_t c = 0; c < keys.cols; ++c) {
      key_text[g * keys.cols + c] = dynamic_array::sort_text_key(key_cell(g, c));
    }
  }
  std::vector<std::string> value_text(value_keys.size());
  for (std::size_t g = 0; g < value_keys.size(); ++g) {
    value_text[g] = dynamic_array::sort_text_key(value_keys[g]);
  }
  std::vector<std::vector<std::string>> outer_text(outer_levels.size());
  for (std::size_t level = 0; level < outer_levels.size(); ++level) {
    outer_text[level].reserve(outer_levels[level].value_of_group.size());
    for (const Value& v : outer_levels[level].value_of_group) {
      outer_text[level].push_back(dynamic_array::sort_text_key(v));
    }
  }
  const auto cmp_key_cell = [&](std::size_t a, std::size_t b, std::uint32_t c) {
    return dynamic_array::cmp_value_keyed(key_cell(a, c), key_cell(b, c), &key_text[a * keys.cols + c],
                                          &key_text[b * keys.cols + c]);
  };
  // First differing key column's order; `directional` applies each column's descending flag.
  const auto cmp_keys = [&](std::size_t a, std::size_t b, bool directional) {
    for (const std::uint32_t c : plan.key_order) {
      const int r = cmp_key_cell(a, b, c);
      if (r != 0) {
        return directional && plan.key_descending[c] ? -r : r;
      }
    }
    return 0;
  };
  const auto group_less = [&](std::size_t a, std::size_t b) {
    // Error-keyed groups always go last regardless of direction, matching
    // Mac Excel's "errors trail" surface for the dynamic-array sort family.
    if (is_error[a] != is_error[b]) {
      return !is_error[a];
    }
    if (plan.value_col != 0U) {
      for (std::size_t level = 0; level < outer_levels.size(); ++level) {
        const OuterSortLevel& outer = outer_levels[level];
        if (outer.prefix_of_group[a] == outer.prefix_of_group[b]) {
          continue;
        }
        const Value& va = outer.value_of_group[a];
        const Value& vb = outer.value_of_group[b];
        const int c = dynamic_array::cmp_value_keyed(va, vb, &outer_text[level][a], &outer_text[level][b]);
        if (c != 0) {
          return plan.value_descending ? (c > 0) : (c < 0);
        }
        // Equal aggregates: the prefix groups first differ at this key column.
        const int k = cmp_key_cell(a, b, static_cast<std::uint32_t>(level));
        if (k != 0) {
          return k < 0;
        }
      }
      const int c = dynamic_array::cmp_value_keyed(value_keys[a], value_keys[b], &value_text[a], &value_text[b]);
      if (c != 0) {
        return plan.value_descending ? (c > 0) : (c < 0);
      }
      return cmp_keys(a, b, false) < 0;
    }
    return cmp_keys(a, b, true) < 0;
  };
  std::vector<std::uint32_t> sorted;
  sorted_index_order(sorted, static_cast<std::uint32_t>(repr.size()), make_index_less(group_less));
  order->assign(sorted.begin(), sorted.end());
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

namespace {

// Computes input/output header layout from `field_headers` in {0,1,2,3} and
// the input row count.
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

// Reads the optional filter_array argument: a 1D column-shaped boolean mask
// over the data rows.
bool read_filter_mask(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx, std::uint32_t data_row_count, std::vector<bool>* include_row,
                      Value* out_err) {
  const ArrayValue* mask = read_array_arg(node, arena, registry, ctx, out_err);
  if (mask == nullptr) {
    return false;
  }
  const std::uint32_t mask_n = (mask->rows >= mask->cols) ? mask->rows : mask->cols;
  if (!dynamic_array::is_declared_vector(node, ctx, mask->rows, mask->cols)) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  if (mask_n != data_row_count) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  for (std::uint32_t i = 0; i < data_row_count; ++i) {
    bool keep = false;
    if (!bool_from_value(mask->cells[i], keep, *out_err)) {
      return false;
    }
    (*include_row)[i] = keep;
  }
  return true;
}

}  // namespace

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
  if (arity > filter_arg_index && !is_omitted_arg(call.as_call_arg(filter_arg_index))) {
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

namespace {

// Hashable canonical key for one row; see `append_cell_group_key`.
std::string normalized_group_key(const ArrayValue& keys, std::uint32_t row) {
  std::string out;
  out.reserve(static_cast<std::size_t>(keys.cols) * 12U);
  for (std::uint32_t c = 0; c < keys.cols; ++c) {
    append_cell_group_key(&out, keys.cells[static_cast<std::size_t>(row) * keys.cols + c], row);
  }
  return out;
}

// True iff any column of the row's key is an Error value.
bool row_key_is_error(const ArrayValue& keys, std::uint32_t row) {
  for (std::uint32_t c = 0; c < keys.cols; ++c) {
    const Value& v = keys.cells[static_cast<std::size_t>(row) * keys.cols + c];
    if (v.is_error()) {
      return true;
    }
  }
  return false;
}

}  // namespace

std::size_t find_or_add_group(const ArrayValue& keys, std::uint32_t row,
                              std::vector<std::uint32_t>* representative_rows,
                              std::vector<std::vector<std::uint32_t>>* member_rows, std::vector<bool>* is_error_group,
                              std::unordered_map<std::string, std::size_t>* index) {
  const std::size_t group = representative_rows->size();
  const auto [it, inserted] = index->try_emplace(normalized_group_key(keys, row), group);
  if (!inserted) {
    (*member_rows)[it->second].push_back(row);
    return it->second;
  }
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
  ArrayValue* arr = allocate_derived_array_value(n, 1U, arena, cells);
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
  const parser::AstNode* all_ast = all != nullptr ? agg.all_ast_by_column[value_col] : nullptr;
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
  if (res.is_array() || res.is_lambda()) {
    return Value::error(ErrorCode::Calc);
  }
  return res;
}

std::vector<Value> aggregate_value_columns(const ArrayValue& values, std::uint32_t val_cols,
                                           const std::vector<std::uint32_t>& row_indices, const Aggregator& agg,
                                           Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx) {
  std::vector<Value> cells(val_cols, Value::blank());
  if (row_indices.empty()) {
    for (std::uint32_t v = 0; v < val_cols; ++v) {
      cells[v] = Value::error(ErrorCode::Calc);
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

Value rows_to_array_value(const std::vector<std::vector<Value>>& rows, std::uint32_t out_cols, Arena& arena) {
  if (rows.empty()) {
    return Value::error(ErrorCode::Calc);
  }
  const std::uint32_t out_rows_n = static_cast<std::uint32_t>(rows.size());
  Value* buffer = nullptr;
  ArrayValue* arr = allocate_derived_array_value(out_rows_n, out_cols, arena, buffer);
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

}  // namespace eval
}  // namespace formulon
