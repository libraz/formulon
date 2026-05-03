// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the PHONETIC lazy form. See eval/phonetic_lazy.h
// for the public contract.

#include "eval/phonetic_lazy.h"

#include "cell.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Same helper as info_lazy.cpp. Inlined here rather than extracted to a
// shared header because PHONETIC is the only outside consumer today;
// promoting `resolve_ref_sheet` later is a bigger refactor than this
// commit warrants.
const Sheet* resolve_ref_sheet_for_phonetic(std::string_view ref_sheet, const EvalContext& ctx) noexcept {
  if (ref_sheet.empty()) {
    return ctx.current_sheet();
  }
  if (ctx.workbook() == nullptr) {
    return nullptr;
  }
  return ctx.workbook()->sheet_by_name(ref_sheet);
}

// Applies Mac's strict-text passthrough to a flattened scalar `v`:
//   * Text  -> the text itself (unchanged).
//   * Blank -> empty string (interned in the eval arena).
//   * Error -> propagated unchanged.
//   * Anything else -> #N/A.
// Used for the non-Ref eager-evaluation arm and as the fallback when a
// Ref-arg cell has no `<rPh>` annotation.
Value apply_passthrough_surface(const Value& v, Arena& arena) {
  if (v.is_error()) {
    return v;
  }
  if (v.is_text()) {
    return v;
  }
  if (v.is_blank()) {
    return Value::text(arena.intern(""));
  }
  return Value::error(ErrorCode::NA);
}

}  // namespace

Value eval_phonetic_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    // Mac surfaces #VALUE! for arity 0 / 2+; mirrors the registered
    // FunctionDef's `min_arity == max_arity == 1` contract.
    return Value::error(ErrorCode::Value);
  }
  const parser::AstNode& arg = call.as_call_arg(0);

  // Ref path: look up the target cell's phonetic_text directly. Mirrors
  // ISFORMULA / FORMULATEXT's handling in info_lazy.cpp, including the
  // whole-row / whole-column #VALUE! surface and the broken-qualifier
  // #REF! / #NAME? mapping.
  if (arg.kind() == parser::NodeKind::Ref) {
    const parser::Reference& r = arg.as_ref();
    if (r.is_full_col || r.is_full_row) {
      return Value::error(ErrorCode::Value);
    }
    const Sheet* target = resolve_ref_sheet_for_phonetic(r.sheet, ctx);
    if (target == nullptr) {
      // Unbound context: #NAME?. Missing qualified sheet: #REF!. Same
      // mapping as ISFORMULA / FORMULATEXT.
      return Value::error(ctx.current_sheet() == nullptr ? ErrorCode::Name : ErrorCode::Ref);
    }
    const Cell* cell = target->cell_at(r.row, r.col);
    if (cell != nullptr && !cell->phonetic_text.empty()) {
      // Annotated cell: surface the kana. Intern into the eval arena so
      // the returned Value's lifetime matches every other Text emitted
      // by the evaluator.
      return Value::text(arena.intern(cell->phonetic_text));
    }
    // No annotation: fall back to the value-based passthrough surface.
    // We don't recurse with a registry here because PHONETIC's argument
    // is required to be a literal Ref; the resolve does not need to
    // re-evaluate a formula cell, only to read its cached value.
    Value resolved = ctx.resolve_ref(r);
    return apply_passthrough_surface(resolved, arena);
  }

  // Non-Ref arg (literal text, arithmetic, function call, range, ...):
  // eagerly evaluate the subtree so error propagation and the "literal
  // text passes through" rule still apply. Mac accepts =PHONETIC("x")
  // and returns "x"; everything non-text yields #N/A.
  const Value v = eval_node(arg, arena, registry, ctx);
  return apply_passthrough_surface(v, arena);
}

}  // namespace eval
}  // namespace formulon
