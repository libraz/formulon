//
// Implementation of the shared structured-reference projection. See
// `structured_ref_project.h` for the contract.

#include "eval/structured_ref_project.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <utility>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/declared_rect.h"
#include "eval/eval_context.h"
#include "eval/range_args.h"
#include "eval/structured_ref.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon {
namespace eval {

namespace {

// The rectangle `table_name[column_payload]` names for the formula `ctx` evaluates.
Expected<StructuredRefRange, ErrorCode> resolve_structured_rect(std::string_view table_name,
                                                                std::string_view column_payload,
                                                                const EvalContext& ctx) {
  const Workbook* wb = ctx.workbook();
  const Sheet* current = ctx.current_sheet();
  if (wb == nullptr || current == nullptr) {
    return ErrorCode::Name;
  }
  // The bracket payload was captured verbatim by the parser; re-parsing
  // it here lets multi-specifier and column-range forms flow through a
  // single AST shape.
  auto sel_or = parse_structured_ref_payload(column_payload);
  if (!sel_or) {
    return sel_or.error();
  }
  StructuredRefSelector sel = std::move(sel_or).value();
  sel.table_name = table_name;
  // Locate the current sheet's workbook index; falling back to 0 is fine
  // because the resolver only consults it for future cross-sheet
  // contracts (the row-implicit form uses the formula cell's row).
  std::uint32_t current_sheet_index = 0;
  for (std::size_t i = 0; i < wb->sheet_count(); ++i) {
    if (&wb->sheet(i) == current) {
      current_sheet_index = static_cast<std::uint32_t>(i);
      break;
    }
  }
  const std::uint32_t current_row = ctx.has_formula_cell() ? ctx.formula_row() : EvalContext::kNoFormulaCell;
  return resolve_structured_ref(sel, *wb, current_sheet_index, current_row);
}

}  // namespace

Value project_structured_ref(std::string_view table_name, std::string_view column_payload, Arena& arena,
                             const FunctionRegistry& registry, const EvalContext& ctx, bool* out_arena_exhausted) {
  *out_arena_exhausted = false;
  auto rect_or = resolve_structured_rect(table_name, column_payload, ctx);
  if (!rect_or) {
    return Value::error(rect_or.error());
  }
  const StructuredRefRange rect = std::move(rect_or).value();
  // Single-cell rectangle: read the value directly, mirroring `Ref`.
  if (rect.row_first == rect.row_last && rect.col_first == rect.col_last) {
    parser::Reference cell{};
    cell.sheet = rect.sheet_name;
    cell.row = rect.row_first;
    cell.col = rect.col_first;
    return ctx.resolve_ref(cell, arena, registry);
  }
  std::vector<Value> expanded;
  ErrorCode expand_err = ErrorCode::Ref;
  if (!expand_resolved_rect_cells(rect.sheet_name, rect.row_first, rect.col_first, rect.row_last, rect.col_last, arena,
                                  registry, ctx, &expanded, &expand_err, nullptr, nullptr)) {
    return Value::error(expand_err);
  }
  const std::uint32_t rows = rect.row_last - rect.row_first + 1U;
  const std::uint32_t cols = rect.col_last - rect.col_first + 1U;
  ArrayValue* arr = array_from_values(rows, cols, expanded.data(), expanded.size(), arena);
  if (arr == nullptr) {
    *out_arena_exhausted = true;
    return Value::error(ErrorCode::Num);
  }
  return Value::array(arr);
}

bool external_reference_rect(const parser::AstNode& node, std::uint32_t* top, std::uint32_t* left,
                             std::uint32_t* bottom, std::uint32_t* right) {
  if (node.kind() != parser::NodeKind::ExternalRef || !node.as_external_ref_name().empty()) {
    return false;
  }
  const parser::Reference& a = node.as_external_ref_cell();
  const parser::Reference& b = node.as_external_ref_cell_end();
  *top = std::min(a.row, b.row);
  *left = std::min(a.col, b.col);
  *bottom = std::max(a.row, b.row);
  *right = std::max(a.col, b.col);
  return true;
}

bool reference_endpoint_pair(const parser::AstNode& node, parser::Reference* out_lhs, parser::Reference* out_rhs,
                             const EvalContext& ctx, ErrorCode* out_err) {
  if (node.kind() != parser::NodeKind::StructuredRef) {
    return declared_rect_endpoint_pair(node, out_lhs, out_rhs, ctx.workbook());
  }
  const auto rect = resolve_structured_rect(node.as_structured_ref_table(), node.as_structured_ref_column(), ctx);
  if (!rect) {
    if (out_err != nullptr) {
      *out_err = rect.error();
    }
    return false;
  }
  // The sheet's own name outlives the evaluation, unlike the range's copy of it.
  const std::string_view sheet = ctx.workbook()->sheet(rect.value().sheet_index).name();
  *out_lhs = parser::Reference{};
  *out_rhs = parser::Reference{};
  out_lhs->sheet = sheet;
  out_rhs->sheet = sheet;
  out_lhs->row = rect.value().row_first;
  out_lhs->col = rect.value().col_first;
  out_rhs->row = rect.value().row_last;
  out_rhs->col = rect.value().col_last;
  return true;
}

}  // namespace eval
}  // namespace formulon
