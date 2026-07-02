// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Shared defined-name resolution. See `defined_name_resolve.h`.

#include "eval/defined_name_resolve.h"

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/formula_text_utils.h"
#include "eval/lazy_impls.h"  // eval_node
#include "io/defined_names.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {

const io::DefinedName* find_defined_name(const Workbook& workbook, std::uint16_t current_sheet_id,
                                         std::string_view name) noexcept {
  const auto& names = workbook.defined_names();
  const io::DefinedName* workbook_match = nullptr;
  for (const auto& entry : names) {
    if (!strings::case_insensitive_eq(entry.name, name)) {
      continue;
    }
    if (entry.local_sheet_id >= 0 && static_cast<std::uint16_t>(entry.local_sheet_id) == current_sheet_id) {
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

Value resolve_defined_name(std::string_view name, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx) {
  const Workbook* wb = ctx.workbook();
  if (wb == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  std::uint16_t sheet_id = 0;
  if (!current_sheet_index(ctx, &sheet_id)) {
    return Value::error(ErrorCode::Name);
  }
  const io::DefinedName* def = find_defined_name(*wb, sheet_id, name);
  if (def == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  // Cycle guard: a name already being expanded on the active resolution chain
  // is a circular reference. Surface `#REF!` to match the cell-cycle policy in
  // `EvalContext::resolve_ref` rather than recursing until the stack blows.
  for (const DefinedNameFrame* f = ctx.defined_name_stack(); f != nullptr; f = f->prev) {
    if (strings::case_insensitive_eq(f->name, def->name)) {
      return Value::error(ErrorCode::Ref);
    }
  }
  const std::string_view src = strip_formula_prefix(def->formula);
  if (src.empty()) {
    return Value::error(ErrorCode::Name);
  }
  parser::Parser parser(src, arena);
  parser::AstNode* root = parser.parse();
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  // Evaluate the body as a top-level formula: clear the using formula's
  // lexical scope (a defined name never sees LET / LAMBDA bindings) and push
  // this name onto the cycle chain. `def->name` is owned by the workbook and
  // outlives this call, so the frame's `string_view` stays valid.
  const DefinedNameFrame frame{def->name, ctx.defined_name_stack()};
  const EvalContext def_ctx = ctx.with_name_env(nullptr).with_defined_name_frame(&frame);
  return eval_node(*root, arena, registry, def_ctx);
}

}  // namespace eval
}  // namespace formulon
