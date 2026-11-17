#include "eval/scalar_args.h"

#include <algorithm>
#include <cmath>

#include "eval/array_alloc.h"
#include "eval/coerce.h"
#include "eval/lazy_impls.h"
#include "eval/omitted_arg.h"
#include "parser/ast.h"
#include "utils/arena.h"

namespace formulon {
namespace eval {

bool number_from_value(const Value& v, double& out, Value& error_out) {
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

bool bool_from_value(const Value& v, bool& out, Value& error_out) {
  if (v.is_error()) {
    error_out = v;
    return false;
  }
  auto coerced = coerce_to_bool(v);
  if (!coerced) {
    error_out = Value::error(coerced.error());
    return false;
  }
  out = coerced.value();
  return true;
}

bool eval_number_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx, double& out, Value& error_out) {
  return number_from_value(eval_node(node, arena, registry, ctx), out, error_out);
}

bool eval_bool_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                   bool& out, Value& error_out) {
  return bool_from_value(eval_node(node, arena, registry, ctx), out, error_out);
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

Value eval_scalar_call_args(const parser::AstNode& call, std::uint32_t min_arity, std::uint32_t max_arity,
                            bool atp_rule, std::uint32_t logical_args, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx, std::vector<Value>& out) {
  const std::uint32_t evaluated = atp_rule ? atp_evaluated_arity(call, min_arity, max_arity) : call.as_call_arity();
  out.clear();
  out.reserve(evaluated);
  for (std::uint32_t i = 0; i < evaluated; ++i) {
    Value v = eval_node(call.as_call_arg(i), arena, registry, ctx);
    if (v.is_error()) {
      return v;
    }
    if (atp_rule) {
      const bool logical = ((logical_args >> i) & 1U) != 0U;
      const Value atp = atp_arg_error(call.as_call_arg(i), v, atp_slot_required(i, min_arity, max_arity), logical);
      if (atp.is_error()) {
        return atp;
      }
    }
    out.push_back(v);
  }
  return Value::blank();
}

namespace {

// The rows x cols of a lifting argument, or false for one that does not lift.
bool lift_dims(const Value& arg, std::uint32_t* rows, std::uint32_t* cols) {
  if (arg.is_array()) {
    *rows = arg.as_array()->rows;
    *cols = arg.as_array()->cols;
    return true;
  }
  if (arg.is_ref() && arg.ref_form() == RefForm::kGrid) {
    const RefView view = arg.as_ref();
    *rows = view.rows;
    *cols = view.cols;
    return true;
  }
  return false;
}

// Element (r, c) of the lifting argument `arg` (rows x cols), a size-1 axis
// stretching; false when it has no such element.
bool lift_element(const Value& arg, std::uint32_t rows, std::uint32_t cols, std::uint32_t r, std::uint32_t c,
                  Arena& arena, Value* out) {
  const std::uint32_t rr = rows == 1U ? 0U : r;
  const std::uint32_t cc = cols == 1U ? 0U : c;
  if (rr >= rows || cc >= cols) {
    return false;
  }
  const std::size_t idx = static_cast<std::size_t>(rr) * cols + cc;
  if (arg.is_array()) {
    *out = arg.as_array()->cells[idx];
    return true;
  }
  const RefView view = arg.as_ref();
  const RefArea& area = view.areas[idx];
  if (area.is_error_element()) {
    *out = Value::error(area.error_code());
    return true;
  }
  const std::uint16_t book = view.element_books != nullptr ? view.element_books[idx] : view.book;
  *out = make_ref(arena, RefForm::kArea, book, &area, 1U);
  return true;
}

}  // namespace

bool lift_extent(const Value* args, std::uint32_t arity, const AreaEnds* ends, std::uint32_t* rows,
                 std::uint32_t* cols) {
  bool lifts = false;
  *rows = 1U;
  *cols = 1U;
  for (std::uint32_t i = 0; i < arity; ++i) {
    std::uint32_t r = 0;
    std::uint32_t c = 0;
    if ((ends == nullptr || !ends[i].active) && lift_dims(args[i], &r, &c)) {
      *rows = std::max(*rows, r);
      *cols = std::max(*cols, c);
      lifts = true;
    }
  }
  return lifts;
}

void lift_positions(const Value* args, std::uint32_t arity, const AreaEnds* ends, std::uint32_t rows,
                    std::uint32_t cols, Arena& arena, LiftedCellFn cell, const void* state, Value* out) {
  std::vector<Value> cell_args(arity, Value::blank());
  std::size_t k = 0;
  for (std::uint32_t r = 0; r < rows; ++r) {
    for (std::uint32_t c = 0; c < cols; ++c, ++k) {
      bool missing = false;
      for (std::uint32_t i = 0; i < arity && !missing; ++i) {
        std::uint32_t ar = 0;
        std::uint32_t ac = 0;
        if (ends != nullptr && ends[i].active) {
          cell_args[i] = k == 0U ? Value::error(ErrorCode::Value) : (k == 1U ? ends[i].first : ends[i].last);
        } else if (lift_dims(args[i], &ar, &ac)) {
          missing = !lift_element(args[i], ar, ac, r, c, arena, &cell_args[i]);
        } else {
          cell_args[i] = args[i];
        }
      }
      if (missing) {
        out[k] = Value::error(ErrorCode::NA);
        continue;
      }
      const auto error = std::find_if(cell_args.begin(), cell_args.end(), [](const Value& v) { return v.is_error(); });
      out[k] = error != cell_args.end() ? *error : cell(cell_args.data(), arity, state);
    }
  }
}

Value lift_over_array_args(const Value* args, std::uint32_t arity, Arena& arena, LiftedCellFn cell, const void* state,
                           const AreaEnds* ends) {
  std::uint32_t rows = 1U;
  std::uint32_t cols = 1U;
  if (!lift_extent(args, arity, ends, &rows, &cols)) {
    const bool masked = ends != nullptr && std::any_of(ends, ends + arity, [](const AreaEnds& e) { return e.active; });
    return masked ? Value::error(ErrorCode::Value) : cell(args, arity, state);
  }
  Value* cells = nullptr;
  ArrayValue* out = allocate_derived_array_value(rows, cols, arena, cells);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  lift_positions(args, arity, ends, rows, cols, arena, cell, state, cells);
  return Value::array(out);
}

}  // namespace eval
}  // namespace formulon
