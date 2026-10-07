#include "eval/lookups/common.h"

#include <cstddef>
#include <cstdint>
#include <vector>

#include "eval/declared_rect.h"
#include "eval/eval_context.h"
#include "eval/external_ref.h"
#include "eval/function_registry.h"
#include "eval/name_env_resolve.h"
#include "external_book.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"
#include "value.h"

namespace formulon {
namespace eval {

Value promote_array_result_cell(const Value& value) {
  return value.promote_reference_blank_to_value_array();
}

Expected<bool, ErrorCode> resolve_reference_table(const parser::AstNode& arg, const EvalContext& ctx,
                                                  ReferenceTable* out) {
  const parser::AstNode& node = resolve_name_ast(arg, ctx.name_env());
  const parser::Reference* external_lhs = nullptr;
  const parser::Reference* external_rhs = nullptr;
  if (external_ref_declared_endpoints(node, &external_lhs, &external_rhs)) {
    const auto external = resolve_external_rect(node, ctx);
    if (!external) {
      return external.error();
    }
    out->lhs = *external_lhs;
    out->rhs = *external_rhs;
    out->declared = external.value().declared;
    out->walked = external.value().walked;
    out->external_book = external.value().book;
    out->external_sheet = external.value().sheet;
    return true;
  }
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
  if (table.external_book != nullptr) {
    // The same ceiling `expand_range` refuses a local rectangle with.
    const std::uint64_t rows = static_cast<std::uint64_t>(row_last - row_first) + 1U;
    const std::uint64_t total = rows * (static_cast<std::uint64_t>(col_last - col_first) + 1U);
    if (total > kMaxRangeExpansionCells) {
      return ErrorCode::Calc;
    }
    std::vector<Value> cells;
    cells.reserve(static_cast<std::size_t>(total));
    for (std::uint32_t r = row_first; r <= row_last; ++r) {
      for (std::uint32_t c = col_first; c <= col_last; ++c) {
        cells.push_back(read_external_cell(*table.external_book, table.external_sheet, r, c, arena));
      }
    }
    return cells;
  }
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
