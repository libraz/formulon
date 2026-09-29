//
// Shared defined-name resolution. See `defined_name_resolve.h`.

#include "eval/defined_name_resolve.h"

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "defined_name.h"
#include "eval/eval_context.h"
#include "eval/formula_text_utils.h"
#include "eval/lazy_impls.h"        // eval_node
#include "eval/name_env_resolve.h"  // is_range_shaped_ast
#include "eval/shape_ops_lazy.h"    // eval_node_as_array
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {

const DefinedName* find_defined_name(const Workbook& workbook, std::uint16_t current_sheet_id,
                                     std::string_view name) noexcept {
  const auto& names = workbook.defined_names();
  const DefinedName* workbook_match = nullptr;
  for (const auto& entry : names) {
    if (!strings::case_insensitive_eq(entry.name, name)) {
      continue;
    }
    // Compare in the wider type: `local_sheet_id` comes from the file and is
    // not bounded by `sheet_count()`, so narrowing it to the 16-bit sheet id
    // first would let a crafted scope (0x10000) match sheet 0.
    if (entry.local_sheet_id >= 0 && entry.local_sheet_id == static_cast<std::int32_t>(current_sheet_id)) {
      // Sheet-scoped match for the current sheet wins immediately.
      return &entry;
    }
    if (entry.local_sheet_id < 0 && workbook_match == nullptr) {
      // Latch the first workbook-scoped match but keep scanning for a
      // sheet-scoped one that should take priority.
      workbook_match = &entry;
    }
  }
  return workbook_match;
}

namespace {

// Resolves the 0-based index of `ctx.current_sheet()` within its workbook.
// Returns false (writing nothing) when either binding is absent or the sheet
// is not owned by the workbook (defensive; the two should always agree).
bool current_sheet_index(const EvalContext& ctx, std::uint16_t* out) noexcept {
  const Workbook* wb = ctx.workbook();
  const Sheet* current = ctx.current_sheet();
  if (wb == nullptr || current == nullptr) {
    return false;
  }
  for (std::size_t i = 0; i < wb->sheet_count(); ++i) {
    if (&wb->sheet(i) == current) {
      *out = static_cast<std::uint16_t>(i);
      return true;
    }
  }
  return false;
}

}  // namespace

const DefinedName* find_defined_name(const EvalContext& ctx, std::string_view name) noexcept {
  const Workbook* wb = ctx.workbook();
  std::uint16_t sheet_id = 0;
  if (wb == nullptr || !current_sheet_index(ctx, &sheet_id)) {
    return nullptr;
  }
  if (ctx.name_scope_sheet() >= 0) {
    sheet_id = static_cast<std::uint16_t>(ctx.name_scope_sheet());
  }
  return find_defined_name(*wb, sheet_id, name);
}

const DefinedName* find_sheet_defined_name(const Workbook& workbook, std::string_view sheet,
                                           std::string_view name) noexcept {
  const std::size_t sheet_id = workbook.sheet_index_by_name(sheet);
  if (sheet_id >= workbook.sheet_count()) {
    return nullptr;
  }
  return find_defined_name(workbook, static_cast<std::uint16_t>(sheet_id), name);
}

const parser::AstNode* prepare_defined_name_body(const DefinedName* def, Arena& arena, const EvalContext& ctx,
                                                 DefinedNameFrame* frame, EvalContext* out_ctx, ErrorCode* out_err) {
  if (def == nullptr) {
    *out_err = ErrorCode::Name;
    return nullptr;
  }
  // Cycle guard: a name already being expanded on the active resolution chain
  // is a circular reference. Surface `#REF!` to match the cell-cycle policy in
  // `EvalContext::resolve_ref` rather than recursing until the stack blows.
  for (const DefinedNameFrame* f = ctx.defined_name_stack(); f != nullptr; f = f->prev) {
    if (f->definition == def) {
      *out_err = ErrorCode::Ref;
      return nullptr;
    }
  }
  const std::string_view src = strip_formula_prefix(def->formula);
  parser::AstNode* root = src.empty() ? nullptr : parser::parse_strict(src, arena);
  if (root == nullptr) {
    *out_err = ErrorCode::Name;
    return nullptr;
  }
  // Evaluate the body as a top-level formula: clear the using formula's
  // lexical scope (a defined name never sees LET / LAMBDA bindings) and push
  // this definition onto the cycle chain.
  *frame = DefinedNameFrame{def, ctx.defined_name_stack()};
  *out_ctx = ctx.with_name_env(nullptr).with_defined_name_frame(frame);
  // A sheet-local name's body resolves its own unqualified names in the
  // owning sheet's scope, whichever sheet uses it; a workbook name's body
  // keeps the scope it is used from.
  if (def->local_sheet_id >= 0 && static_cast<std::size_t>(def->local_sheet_id) < ctx.workbook()->sheet_count()) {
    *out_ctx = out_ctx->with_name_scope_sheet(def->local_sheet_id);
  }
  return root;
}

namespace {

// Evaluates the body of the already-located definition `def` in `ctx`.
Value evaluate_defined_name(const DefinedName* def, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx) {
  DefinedNameFrame frame;
  EvalContext def_ctx = ctx;
  ErrorCode err = ErrorCode::Name;
  const parser::AstNode* root = prepare_defined_name_body(def, arena, ctx, &frame, &def_ctx, &err);
  if (root == nullptr) {
    return Value::error(err);
  }
  // A range-shaped body (e.g. `Sheet1!$A$1:$A$5`) must surface as a
  // `Value::Array` so range-aware consumers (`SUM`, `COUNT`, `VLOOKUP`, ...)
  // and the spill committer pick up its full shape instead of collapsing it to
  // the scalar `eval_node` would produce via implicit intersection. This mirrors
  // how `INDIRECT` returns a range as an array. Scalar bodies (including a
  // single-cell `Ref`) keep the scalar `eval_node` path.
  if (is_range_shaped_ast(*root)) {
    return eval_node_as_array(*root, arena, registry, def_ctx);
  }
  return eval_node(*root, arena, registry, def_ctx);
}

}  // namespace

Value resolve_defined_name(std::string_view name, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx) {
  if (ctx.workbook() == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate_defined_name(find_defined_name(ctx, name), arena, registry, ctx);
}

Value resolve_sheet_defined_name(std::string_view sheet, std::string_view name, Arena& arena,
                                 const FunctionRegistry& registry, const EvalContext& ctx) {
  const Workbook* wb = ctx.workbook();
  if (wb == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  if (wb->sheet_index_by_name(sheet) >= wb->sheet_count()) {
    return Value::error(ErrorCode::Ref);
  }
  return evaluate_defined_name(find_sheet_defined_name(*wb, sheet, name), arena, registry, ctx);
}

}  // namespace eval
}  // namespace formulon
