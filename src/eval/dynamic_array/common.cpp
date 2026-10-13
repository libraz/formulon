
#include "eval/dynamic_array/common.h"

#include <algorithm>
#include <cassert>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include "eval/coerce.h"
#include "eval/collation.h"
#include "eval/eval_profile_scope.h"
#include "eval/lazy_impls.h"
#include "eval/omitted_arg.h"
#include "eval/range_args.h"
#include "eval/value_identity.h"
#include "parser/ast.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/checked_mul.h"
#include "utils/error.h"
#include "utils/index_sort.h"
#include "utils/resource_budget.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace dynamic_array {

const ArrayValue* resolve_array_prelude(const parser::AstNode& call, std::uint32_t min_arity, std::uint32_t max_arity,
                                        bool reject_empty, Arena& arena, const FunctionRegistry& registry,
                                        const EvalContext& ctx, Value& error_out) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < min_arity || arity > max_arity) {
    error_out = Value::error(ErrorCode::Value);
    return nullptr;
  }
  const ArrayValue* array = nullptr;
  if (!resolve_array_value(call.as_call_arg(0), arena, registry, ctx, &array, &error_out)) {
    return nullptr;
  }
  if (reject_empty && (array->rows == 0U || array->cols == 0U)) {
    error_out = Value::error(ErrorCode::Calc);
    return nullptr;
  }
  return array;
}

bool eval_number_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx, double& out, Value& error_out) {
  const Value v = eval_node(node, arena, registry, ctx);
  if (v.is_error()) {
    error_out = v;
    return false;
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    error_out = Value::error(coerced.error());
    return false;
  }
  out = coerced.value();
  return true;
}

bool eval_truncated_number_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                               const EvalContext& ctx, double& out, Value& error_out) {
  double n = 0.0;
  if (!eval_number_arg(node, arena, registry, ctx, n, error_out)) {
    return false;
  }
  out = std::trunc(n);
  return true;
}

ArrayValue* materialise_selected_lanes(const ArrayValue& src, const std::vector<std::uint32_t>& indices, bool by_col,
                                       Arena& arena) {
  const std::uint32_t out_rows = by_col ? src.rows : static_cast<std::uint32_t>(indices.size());
  const std::uint32_t out_cols = by_col ? static_cast<std::uint32_t>(indices.size()) : src.cols;
  Value* buffer = nullptr;
  ArrayValue* out = allocate_array_value(out_rows, out_cols, arena, buffer, kMaxDerivedArrayCells);
  if (out == nullptr) {
    return nullptr;
  }
  if (by_col) {
    for (std::uint32_t r = 0; r < out_rows; ++r) {
      for (std::uint32_t i = 0; i < out_cols; ++i) {
        const std::uint32_t src_col = indices[i];
        buffer[static_cast<std::size_t>(r) * out_cols + i] =
            src.cells[static_cast<std::size_t>(r) * src.cols + src_col].promote_reference_blank_to_value_array();
      }
    }
  } else {
    for (std::uint32_t i = 0; i < out_rows; ++i) {
      const std::uint32_t src_row = indices[i];
      for (std::uint32_t c = 0; c < out_cols; ++c) {
        buffer[static_cast<std::size_t>(i) * out_cols + c] =
            src.cells[static_cast<std::size_t>(src_row) * src.cols + c].promote_reference_blank_to_value_array();
      }
    }
  }

  return out;
}

ArrayValue* materialise_slice(const ArrayValue& src, std::uint32_t row_lo, std::uint32_t row_hi, std::uint32_t col_lo,
                              std::uint32_t col_hi, Arena& arena) {
  const std::uint32_t out_rows = row_hi - row_lo;
  const std::uint32_t out_cols = col_hi - col_lo;
  Value* buffer = nullptr;
  ArrayValue* out = allocate_array_value(out_rows, out_cols, arena, buffer, kMaxDerivedArrayCells);
  if (out == nullptr) {
    return nullptr;
  }
  for (std::uint32_t r = 0; r < out_rows; ++r) {
    for (std::uint32_t c = 0; c < out_cols; ++c) {
      buffer[static_cast<std::size_t>(r) * out_cols + c] =
          src.cells[static_cast<std::size_t>(row_lo + r) * src.cols + (col_lo + c)]
              .promote_reference_blank_to_value_array();
    }
  }
  return out;
}

ArrayValue* materialise_vector(std::vector<Value>&& cells, bool as_column, Arena& arena) {
  const auto n = static_cast<std::uint32_t>(cells.size());
  const std::uint32_t rows = as_column ? n : 1U;
  const std::uint32_t cols = as_column ? 1U : n;
  Value* buffer = nullptr;
  ArrayValue* out = allocate_array_value(rows, cols, arena, buffer, kMaxDerivedArrayCells);
  if (out == nullptr) {
    return nullptr;
  }
  for (std::size_t i = 0; i < cells.size(); ++i) {
    buffer[i] = cells[i].promote_reference_blank_to_value_array();
  }
  return out;
}

bool unique_lane_key(const ArrayValue& arr, std::uint32_t lane, bool by_col, std::string* out) {
  out->clear();
  const ExcelProfile profile = current_eval_profile();
  const std::uint32_t n = by_col ? arr.rows : arr.cols;
  for (std::uint32_t k = 0; k < n; ++k) {
    const std::size_t index =
        by_col ? static_cast<std::size_t>(k) * arr.cols + lane : static_cast<std::size_t>(lane) * arr.cols + k;
    if (!append_value_identity_key(arr.cells[index], profile, CollationStrength::kDynamic, out)) {
      return false;
    }
  }
  return true;
}

namespace {

// Excel-canonical kind rank for SORT. Lower rank sorts earlier under
// ascending order. Blank deliberately sits at rank 5 -- the SORT impl
// special-cases blanks so they always sink to the end regardless of
// `sort_order`. Numbers / Text / Bool / Error participate normally.
int sort_kind_rank(const Value& v) {
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
      return 5;
    default:
      // Array / Ref / Lambda are not produced by cell reads. Slot them
      // between Error and Blank so they participate predictably without
      // colliding with real-world cells.
      return 4;
  }
}

}  // namespace

bool ascii_ci_less(std::string_view a, std::string_view b) {
  return collate(a, b, current_eval_profile(), CollationStrength::kDynamic) < 0;
}

std::string sort_text_key(const Value& v) {
  if (!v.is_text()) {
    return std::string();
  }
  return collation_key(v.as_text(), current_eval_profile(), CollationStrength::kDynamic);
}

namespace {

bool sort_cell_less_keyed(const Value& a, const Value& b, const std::string* key_a, const std::string* key_b) {
  const int ra = sort_kind_rank(a);
  const int rb = sort_kind_rank(b);
  if (ra != rb) {
    return ra < rb;
  }
  switch (a.kind()) {
    case ValueKind::Number:
      return a.as_number() < b.as_number();
    case ValueKind::Text:
      if (key_a != nullptr && key_b != nullptr) {
        return collate_key_compare(*key_a, *key_b) < 0;
      }
      return ascii_ci_less(a.as_text(), b.as_text());
    case ValueKind::Bool:
      return !a.as_boolean() && b.as_boolean();
    case ValueKind::Error:
      return static_cast<int>(a.as_error()) < static_cast<int>(b.as_error());
    case ValueKind::Blank:
    default:
      return false;
  }
}

}  // namespace

bool sort_cell_less_asc(const Value& a, const Value& b) {
  return sort_cell_less_keyed(a, b, nullptr, nullptr);
}

bool sort_lane_less_keyed(const Value& key_a, const std::string* text_key_a, const Value& key_b,
                          const std::string* text_key_b, bool descending) {
  const bool a_blank = key_a.is_blank();
  const bool b_blank = key_b.is_blank();
  if (a_blank != b_blank) {
    // Non-blank wins (sorts earlier) regardless of direction.
    return !a_blank;
  }
  if (a_blank && b_blank) {
    return false;
  }
  return descending ? sort_cell_less_keyed(key_b, key_a, text_key_b, text_key_a)
                    : sort_cell_less_keyed(key_a, key_b, text_key_a, text_key_b);
}

bool sort_lane_less(const Value& key_a, const Value& key_b, bool descending) {
  return sort_lane_less_keyed(key_a, nullptr, key_b, nullptr, descending);
}

std::uint32_t collation_classes(const Value* values, std::size_t count, ExcelProfile profile,
                                CollationStrength strength, std::uint32_t* out_ids) {
  if (count == 0) {
    return 0;
  }
  // Text keys follow the kind tag and the 8-byte length in the typed key.
  constexpr std::size_t kTextKeyOffset = 1 + sizeof(std::uint64_t);
  std::unordered_map<std::string, std::uint32_t> provisional;
  std::vector<std::uint32_t> value_class(count);
  std::vector<const Value*> reps;
  std::vector<std::string> text_keys;
  std::string typed;
  for (std::size_t i = 0; i < count; ++i) {
    typed.clear();
    const bool ok = append_value_identity_key(values[i], profile, strength, &typed);
    assert(ok);
    (void)ok;
    const auto inserted = provisional.emplace(typed, static_cast<std::uint32_t>(reps.size()));
    if (inserted.second) {
      reps.push_back(&values[i]);
      text_keys.push_back(values[i].is_text() ? typed.substr(kTextKeyOffset) : std::string());
    }
    value_class[i] = inserted.first->second;
  }

  const auto rep_less = [&](std::uint32_t a, std::uint32_t b) {
    return sort_cell_less_keyed(*reps[a], *reps[b], &text_keys[a], &text_keys[b]);
  };
  std::vector<std::uint32_t> order;
  sorted_index_order(order, static_cast<std::uint32_t>(reps.size()), make_index_less(rep_less));
  std::vector<std::uint32_t> final_id(reps.size());
  for (std::uint32_t rank = 0; rank < order.size(); ++rank) {
    final_id[order[rank]] = rank;
  }
  for (std::size_t i = 0; i < count; ++i) {
    out_ids[i] = final_id[value_class[i]];
  }
  return static_cast<std::uint32_t>(reps.size());
}

bool resolve_sort_order_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx, bool& descending, Value& error_out) {
  if (is_omitted_arg(node)) {
    descending = false;
    return true;
  }
  double n = 0.0;
  if (!eval_number_arg(node, arena, registry, ctx, n, error_out)) {
    return false;
  }
  if (n == 1.0) {
    descending = false;
    return true;
  }
  if (n == -1.0) {
    descending = true;
    return true;
  }
  error_out = Value::error(ErrorCode::Value);
  return false;
}

bool resolve_take_drop_range(const parser::AstNode* node, std::uint32_t axis_size, bool take, Arena& arena,
                             const FunctionRegistry& registry, const EvalContext& ctx, std::uint32_t& lo,
                             std::uint32_t& hi, Value& error_out) {
  if (node == nullptr) {
    lo = 0;
    hi = axis_size;
    return true;
  }
  double count = 0.0;
  if (!eval_truncated_number_arg(*node, arena, registry, ctx, count, error_out)) {
    return false;
  }
  if (take) {
    if (count == 0.0) {
      error_out = Value::error(ErrorCode::Calc);
      return false;
    }
    if (count > 0.0) {
      const auto take_n = (count >= static_cast<double>(axis_size)) ? axis_size : static_cast<std::uint32_t>(count);
      lo = 0;
      hi = take_n;
      return true;
    }
    const double abs_count = -count;
    const auto take_n =
        (abs_count >= static_cast<double>(axis_size)) ? axis_size : static_cast<std::uint32_t>(abs_count);
    lo = axis_size - take_n;
    hi = axis_size;
    return true;
  }
  if (count >= 0.0) {
    if (count >= static_cast<double>(axis_size)) {
      error_out = Value::error(ErrorCode::Calc);
      return false;
    }
    lo = static_cast<std::uint32_t>(count);
    hi = axis_size;
    return true;
  }
  const double abs_count = -count;
  if (abs_count >= static_cast<double>(axis_size)) {
    error_out = Value::error(ErrorCode::Calc);
    return false;
  }
  lo = 0;
  hi = axis_size - static_cast<std::uint32_t>(abs_count);
  return true;
}

}  // namespace dynamic_array
}  // namespace eval
}  // namespace formulon
