//
// Tree-walker lazy impl for the date1904-sensitive calendar family. See
// `datetime_lazy.h` for the rationale (the workbook date epoch is only
// reachable via `EvalContext`, which the eager calling convention cannot
// carry). The actual calendar math lives in `builtins/datetime.cpp`; this
// file only evaluates the arguments and threads `ctx.date1904()` into the
// shared `DateEntry::impl`.

#include "eval/datetime_lazy.h"

#include <cstdint>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {

Value invoke_date_entry(const DateEntry& entry, const Value* args, std::uint32_t arity, Arena& arena, bool date1904) {
  std::uint32_t rows = 1U;
  std::uint32_t cols = 1U;
  bool has_array = false;
  for (std::uint32_t i = 0; i < arity; ++i) {
    if (args[i].is_array()) {
      const ArrayValue* a = args[i].as_array();
      rows = std::max(rows, a->rows);
      cols = std::max(cols, a->cols);
      has_array = true;
    }
  }
  if (!has_array) {
    return entry.impl(args, arity, arena, date1904);
  }
  Value* cells = nullptr;
  ArrayValue* out = allocate_array_value(rows, cols, arena, cells, kMaxDerivedArrayCells);
  if (out == nullptr)
    return Value::error(ErrorCode::Num);
  std::vector<Value> cell_args;
  cell_args.reserve(arity);
  for (std::uint32_t i = 0; i < arity; ++i) {
    cell_args.push_back(Value::blank());
  }
  std::size_t idx = 0;
  for (std::uint32_t r = 0; r < rows; ++r)
    for (std::uint32_t c = 0; c < cols; ++c, ++idx) {
      bool missing = false;
      for (std::uint32_t i = 0; i < arity; ++i) {
        if (!args[i].is_array()) {
          cell_args[i] = args[i];
          continue;
        }
        const ArrayValue* a = args[i].as_array();
        const std::uint32_t rr = a->rows == 1U ? 0U : r;
        const std::uint32_t cc = a->cols == 1U ? 0U : c;
        if (rr >= a->rows || cc >= a->cols) {
          missing = true;
          break;
        }
        cell_args[i] = a->cells[static_cast<std::size_t>(rr) * a->cols + cc];
      }
      if (missing) {
        cells[idx] = Value::error(ErrorCode::NA);
        continue;
      }
      for (const Value& v : cell_args)
        if (v.is_error()) {
          cells[idx] = v;
          goto next_cell;
        }
      cells[idx] = entry.impl(cell_args.data(), arity, arena, date1904);
    next_cell:;
    }
  return Value::array(out);
}

Value eval_datetime_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  const DateEntry* entry = find_date_entry(call.as_call_name());
  if (entry == nullptr) {
    // Registered names always resolve; a miss can only mean a table drift.
    return Value::error(ErrorCode::Name);
  }
  const std::uint32_t arity = call.as_call_arity();
  if (arity < entry->min_arity || arity > entry->max_arity) {
    return Value::error(ErrorCode::Value);
  }
  // Calendar functions are scalar-only and propagate the left-most argument
  // error (they never opt out of `propagate_errors`), matching the eager
  // dispatcher's pre-evaluation contract.
  std::vector<Value> args;
  args.reserve(arity);
  for (std::uint32_t i = 0; i < arity; ++i) {
    Value v = eval_node(call.as_call_arg(i), arena, registry, ctx);
    if (v.is_error()) {
      return v;
    }
    args.push_back(v);
  }
  return invoke_date_entry(*entry, args.empty() ? nullptr : args.data(), arity, arena, ctx.date1904());
}

}  // namespace eval
}  // namespace formulon
