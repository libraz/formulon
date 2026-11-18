#include "eval/scalar_args.h"

#include <algorithm>
#include <cmath>

#include "eval/array_alloc.h"
#include "eval/coerce.h"
#include "eval/error_origin.h"
#include "eval/lazy_impls.h"
#include "eval/omitted_arg.h"
#include "eval/reference/ref_value.h"
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

void eval_scalar_call_args(const parser::AstNode& call, std::uint32_t min_arity, std::uint32_t max_arity, bool atp_rule,
                           std::uint32_t logical_args, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx, std::vector<Value>& out,
                           std::optional<ErrorOriginSnapshot>& first_error_origin) {
  const std::uint32_t evaluated = atp_rule ? atp_evaluated_arity(call, min_arity, max_arity) : call.as_call_arity();
  out.clear();
  out.reserve(evaluated);
  first_error_origin.reset();
  for (std::uint32_t i = 0; i < evaluated; ++i) {
    Value v = eval_node(call.as_call_arg(i), arena, registry, ctx);
    if (!v.is_error() && atp_rule) {
      const bool logical = ((logical_args >> i) & 1U) != 0U;
      const Value atp = atp_arg_error(call.as_call_arg(i), v, atp_slot_required(i, min_arity, max_arity), logical);
      if (atp.is_error()) {
        v = atp;
      }
    }
    if (v.is_error() && !first_error_origin.has_value()) {
      first_error_origin = snapshot_error_origin();
    }
    out.push_back(v);
  }
  if (first_error_origin.has_value()) {
    forget_error_origin();
  }
}

const Value* reported_error_arg(const Value* args, std::uint32_t arity, bool last_first) {
  const Value* found = nullptr;
  for (std::uint32_t i = 0; i < arity; ++i) {
    if (args[i].is_error()) {
      found = &args[i];
      if (!last_first) {
        break;
      }
    }
  }
  return found;
}

Value report_error_arg(const Value& result, const Value* args, std::uint32_t arity, bool last_first) {
  if (result.is_error()) {
    return result;
  }
  const Value* reported = reported_error_arg(args, arity, last_first);
  return reported != nullptr ? *reported : result;
}

void restore_reported_origin(const Value& result, const Value* args, std::uint32_t arity, bool last_first,
                             const std::optional<ErrorOriginSnapshot>& reported_origin) {
  if (!reported_origin.has_value() || !result.is_error() || error_origin() != ErrorOrigin::kNone) {
    return;
  }
  const Value* reported = reported_error_arg(args, arity, last_first);
  if (reported != nullptr && reported->as_error() == result.as_error()) {
    restore_error_origin(*reported_origin);
  }
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
  *out = grid_element(arg.as_ref(), idx, arena);
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
      // A position whose call answers an array holds that array's top-left element.
      if (out[k].is_array()) {
        out[k] = out[k].as_array()->cells[0];
      }
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
