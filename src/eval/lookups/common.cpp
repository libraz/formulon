#include "eval/lookups/common.h"

#include <cstddef>
#include <cstdint>
#include <vector>

#include "eval/declared_rect.h"
#include "eval/eval_context.h"
#include "eval/external_ref.h"
#include "eval/function_registry.h"
#include "eval/jp_fold.h"
#include "eval/name_env_resolve.h"
#include "eval/wildcard.h"
#include "excel_locale.h"
#include "external_book.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "utils/rect_iterator.h"
#include "utils/resource_budget.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {

Value promote_array_result_cell(const Value& value) {
  return value.promote_reference_blank_to_value_array();
}

Expected<bool, ErrorCode> resolve_reference_table(const parser::AstNode& arg, const EvalContext& ctx,
                                                  ReferenceTable* out, TableWalk walk) {
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
  out->declared = declared.value();
  out->walk_pending = true;
  if (walk == TableWalk::kNow) {
    const auto walked = walk_reference_table(out, ctx);
    if (!walked) {
      return walked.error();
    }
  }
  return true;
}

Expected<void, ErrorCode> walk_reference_table(ReferenceTable* table, const EvalContext& ctx) {
  if (!table->walk_pending) {
    return Expected<void, ErrorCode>::Ok();
  }
  const auto walked = ctx.walked_range_rect(table->lhs, table->rhs);
  if (!walked) {
    return Expected<void, ErrorCode>::Err(walked.error());
  }
  if (walked.value().has_value()) {
    table->walked = *walked.value();
  } else {
    table->walked = table->declared;
    if (table->walked.rows() == Sheet::kMaxRows) {
      table->walked.row_last = table->walked.row_first;
    } else {
      table->walked.col_last = table->walked.col_first;
    }
  }
  table->walk_pending = false;
  return Expected<void, ErrorCode>::Ok();
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
    for (const utils::RectCoord cell : utils::RectRange(row_first, col_first, row_last, col_last)) {
      cells.push_back(read_external_cell(*table.external_book, table.external_sheet, cell.row, cell.col, arena));
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

std::string lookup_wildcard_key(std::string_view text, ExcelProfile profile) {
  if (width_folding(profile) == WidthFolding::kMac) {
    return fold_and_lower_per_code_point(text);
  }
  return strings::to_ascii_lower(text);
}

PreparedTextQuery prepare_text_query(std::string_view raw, ExcelProfile profile, CollationStrength strength) {
  if (scan_has_wildcard(raw)) {
    return PreparedTextQuery{true, lookup_wildcard_key(raw, profile), strength, std::nullopt};
  }
  PreparedTextQuery query{false, unescape_wildcard_literal(raw), strength, std::nullopt};
  query.literal.emplace(query.prepared, profile, strength);
  return query;
}

bool prepared_text_query_matches(const Value& cell, const PreparedTextQuery& query, ExcelProfile profile) {
  if (!cell.is_text()) {
    return false;
  }
  if (query.has_pattern) {
    return wildcard_match(query.prepared, lookup_wildcard_key(cell.as_text(), profile));
  }
  return query.literal->compare(cell.as_text()) == 0;
}

}  // namespace eval
}  // namespace formulon
