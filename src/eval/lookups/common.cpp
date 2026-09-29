#include "eval/lookups/common.h"

#include <cstdint>
#include <vector>

#include "eval/declared_rect.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/name_env_resolve.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {

Value promote_array_result_cell(const Value& value) {
  return value.promote_reference_blank_to_value_array();
}

Expected<bool, ErrorCode> resolve_reference_table(const parser::AstNode& arg, const EvalContext& ctx,
                                                  ReferenceTable* out) {
  const parser::AstNode& node = resolve_name_ast(arg, ctx.name_env());
  if (!declared_rect_endpoint_pair(node, &out->lhs, &out->rhs)) {
    return false;
  }
  const auto declared = ctx.declared_range_rect(out->lhs, out->rhs);
  if (!declared) {
    return declared.error();
  }
  const auto walked = ctx.walked_range_rect(out->lhs, out->rhs);
  if (!walked) {
    return walked.error();
  }
  out->declared = declared.value();
  if (walked.value().has_value()) {
    out->walked = *walked.value();
  } else {
    out->walked = declared.value();
    if (out->walked.rows() == Sheet::kMaxRows) {
      out->walked.row_last = out->walked.row_first;
    } else {
      out->walked.col_last = out->walked.col_first;
    }
  }
  return true;
}

Expected<std::vector<Value>, ErrorCode> read_table_block(const ReferenceTable& table, std::uint32_t row_first,
                                                         std::uint32_t row_last, std::uint32_t col_first,
                                                         std::uint32_t col_last, Arena& arena,
                                                         const FunctionRegistry& registry, const EvalContext& ctx) {
  parser::Reference lhs = table.lhs;
  parser::Reference rhs = table.rhs;
  lhs.is_full_col = lhs.is_full_row = false;
  rhs.is_full_col = rhs.is_full_row = false;
  lhs.row = row_first;
  lhs.col = col_first;
  rhs.row = row_last;
  rhs.col = col_last;
  return ctx.expand_range(lhs, rhs, arena, registry);
}

Value read_table_cell(const ReferenceTable& table, std::uint32_t row, std::uint32_t col, Arena& arena,
                      const FunctionRegistry& registry, const EvalContext& ctx) {
  const std::uint32_t sheet_row = table.declared.row_first + row;
  const std::uint32_t sheet_col = table.declared.col_first + col;
  auto cell = read_table_block(table, sheet_row, sheet_row, sheet_col, sheet_col, arena, registry, ctx);
  if (!cell) {
    return Value::error(cell.error());
  }
  return cell.value().front();
}

}  // namespace eval
}  // namespace formulon
