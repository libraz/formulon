//
// Implementation of the shared structured-reference projection. See
// `structured_ref_project.h` for the contract.

#include "eval/structured_ref_project.h"

#include <cstddef>
#include <cstdint>
#include <utility>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/eval_context.h"
#include "eval/structured_ref.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"

namespace formulon {
namespace eval {

Value project_structured_ref(std::string_view table_name, std::string_view column_payload, Arena& arena,
                             const FunctionRegistry& registry, const EvalContext& ctx, bool* out_arena_exhausted) {
  *out_arena_exhausted = false;
  const Workbook* wb = ctx.workbook();
  const Sheet* current = ctx.current_sheet();
  if (wb == nullptr || current == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  // The bracket payload was captured verbatim by the parser; re-parsing
  // it here lets multi-specifier and column-range forms flow through a
  // single AST shape.
  auto sel_or = parse_structured_ref_payload(column_payload);
  if (!sel_or) {
    return Value::error(sel_or.error());
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
  auto rect_or = resolve_structured_ref(sel, *wb, current_sheet_index, current_row);
  if (!rect_or) {
    return Value::error(rect_or.error());
  }
  const StructuredRefRange rect = std::move(rect_or).value();
  parser::Reference lhs{};
  lhs.sheet = rect.sheet_name;
  lhs.row = rect.row_first;
  lhs.col = rect.col_first;
  parser::Reference rhs{};
  rhs.sheet = rect.sheet_name;
  rhs.row = rect.row_last;
  rhs.col = rect.col_last;
  // Single-cell rectangle: read the value directly, mirroring `Ref`.
  if (rect.row_first == rect.row_last && rect.col_first == rect.col_last) {
    return ctx.resolve_ref(lhs, arena, registry);
  }
  auto cells = ctx.expand_range(lhs, rhs, arena, registry);
  if (!cells) {
    return Value::error(cells.error());
  }
  const std::uint32_t rows = rect.row_last - rect.row_first + 1U;
  const std::uint32_t cols = rect.col_last - rect.col_first + 1U;
  Value* buffer = nullptr;
  ArrayValue* arr = allocate_array_value(rows, cols, arena, buffer, kMaxDerivedArrayCells);
  if (arr == nullptr) {
    *out_arena_exhausted = true;
    return Value::error(ErrorCode::Num);
  }
  // The seam hands back uninitialised storage, so every cell must be
  // written even when the expansion returned fewer values than the
  // rectangle covers (a clamped whole-row / whole-column endpoint).
  const std::size_t total = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  const std::vector<Value>& expanded = cells.value();
  for (std::size_t i = 0; i < total; ++i) {
    buffer[i] = i < expanded.size() ? expanded[i] : Value::blank();
  }
  return Value::array(arr);
}

}  // namespace eval
}  // namespace formulon
