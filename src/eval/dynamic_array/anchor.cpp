
#include "eval/dynamic_array/anchor.h"

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/name_env_resolve.h"
#include "eval/range_resolvers.h"
#include "eval/spill_anchor.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {

bool resolve_spill_anchor_node(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                               const EvalContext& ctx, std::string_view* out_sheet, std::uint32_t* out_row,
                               std::uint32_t* out_col, ErrorCode* out_err) {
  // A LET binding is transparent here: `=LET(x, A4, SUM(x#))` anchors on
  // A4, so walk the name through to the AST it was bound to before
  // dispatching on shape.
  const parser::AstNode& target = resolve_name_ast(node, ctx.name_env());
  *out_err = ErrorCode::Ref;
  switch (target.kind()) {
    case parser::NodeKind::Ref: {
      const parser::Reference& ref = target.as_ref();
      if (ref.is_full_col || ref.is_full_row) {
        return false;
      }
      *out_sheet = ref.sheet;
      *out_row = ref.row;
      *out_col = ref.col;
      return true;
    }
    case parser::NodeKind::SpillRef: {
      // `A1##` is not a spelling Excel accepts, but ANCHORARRAY can be
      // handed a SpillRef through a LET binding. The inner operator's own
      // anchor is the anchor.
      if (target.as_spill_ref_anchor_expr() != nullptr) {
        return resolve_spill_anchor_node(*target.as_spill_ref_anchor_expr(), arena, registry, ctx, out_sheet, out_row,
                                         out_col, out_err);
      }
      const parser::Reference& ref = target.as_spill_ref();
      *out_sheet = ref.sheet;
      *out_row = ref.row;
      *out_col = ref.col;
      return true;
    }
    case parser::NodeKind::Call: {
      std::uint32_t bottom_row = 0;
      std::uint32_t right_col = 0;
      bool is_range = false;
      if (!resolve_reference_call(target, arena, registry, ctx, out_sheet, out_row, out_col, &bottom_row, &right_col,
                                  &is_range, out_err)) {
        return false;
      }
      if (is_range) {
        // An anchor is a single cell, and Excel does not narrow a wider
        // reference down to its top-left corner: `=SUM(OFFSET(A1,0,0,2,2)#)`
        // is `#REF!` even when A1 anchors a live spill.
        *out_err = ErrorCode::Ref;
        return false;
      }
      return true;
    }
    default:
      // A shape that names no reference at all (a literal, an arithmetic
      // result). Excel refuses `1#` and `"A1"#` at entry, so it never
      // answers here; `#VALUE!` is what the argument-shape surface of every
      // other reference consumer reports.
      *out_err = ErrorCode::Value;
      return false;
  }
}

Value eval_anchorarray_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  std::string_view sheet_name;
  std::uint32_t anchor_row = 0;
  std::uint32_t anchor_col = 0;
  ErrorCode err = ErrorCode::Ref;
  if (!resolve_spill_anchor_node(call.as_call_arg(0), arena, registry, ctx, &sheet_name, &anchor_row, &anchor_col,
                                 &err)) {
    return Value::error(err);
  }
  ArrayValue* arr = project_spill_at_anchor(sheet_name, anchor_row, anchor_col, arena, ctx, &err);
  if (arr == nullptr) {
    return Value::error(err);
  }
  return Value::array(arr);
}

}  // namespace eval
}  // namespace formulon
