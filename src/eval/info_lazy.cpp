// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the context-aware information predicates. See
// `eval/info_lazy.h` for the semantics each function advertises.
//
// The family deliberately rides the lazy seam because every entry
// inspects information the eager path would erase:
//   * ISFORMULA needs the bound Sheet's `formula_text` for the
//     referenced cell — an evaluated reference is just a `Value`.
//   * ISREF branches on the AST node kind (and, for calls, on whether
//     the call is in the reference-returning set).
//   * SHEET / SHEETS consult `Workbook::sheet_by_name` and the bound
//     current-sheet pointer held on `EvalContext`.

#include "eval/info_lazy.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string_view>

#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Case-insensitive match against the fixed set of Excel calls whose
// result is a reference (and therefore count as "ref" for ISREF).
// INDEX and CHOOSE ride here too because both can return refs when
// invoked with reference arguments; our MVP treats any static Call
// with these names as ref-producing.
bool is_reference_call_name(std::string_view name) noexcept {
  constexpr std::string_view kNames[] = {"INDIRECT", "OFFSET", "INDEX", "CHOOSE"};
  for (const auto& n : kNames) {
    if (strings::case_insensitive_eq(name, n)) {
      return true;
    }
  }
  return false;
}

// Returns true when `node` is a static reference-shaped AST. Does NOT
// include reference-returning calls (INDIRECT/OFFSET/INDEX/CHOOSE);
// those need evaluation to confirm they produce a ref at runtime.
bool is_static_reference_shape(const parser::AstNode& node) noexcept {
  switch (node.kind()) {
    case parser::NodeKind::Ref:
    case parser::NodeKind::RangeOp:
    case parser::NodeKind::ExternalRef:
    case parser::NodeKind::StructuredRef:
    case parser::NodeKind::NameRef:
      return true;
    default:
      return false;
  }
}

// Locates `sheet` inside `workbook` by pointer identity and returns
// the 1-based sheet number on success, 0 on failure. Pointer identity
// is sufficient because `EvalContext` holds a raw pointer into the
// workbook's own vector; if the caller ever constructs a context with
// a sheet owned by a *different* workbook this returns 0 and SHEET
// surfaces `#N/A`, which is a safe degradation.
std::size_t find_sheet_index(const Workbook& workbook, const Sheet& sheet) noexcept {
  const std::size_t count = workbook.sheet_count();
  for (std::size_t i = 0; i < count; ++i) {
    if (&workbook.sheet(i) == &sheet) {
      return i + 1;
    }
  }
  return 0;
}

// Returns the referenced cell's sheet, resolving qualifier via the
// bound workbook. Returns `nullptr` when the qualifier is present but
// no workbook is bound, or when the named sheet is missing. `ref_sheet`
// is the sheet string from a `parser::Reference`; empty means "current
// sheet".
const Sheet* resolve_ref_sheet(std::string_view ref_sheet, const EvalContext& ctx) noexcept {
  if (ref_sheet.empty()) {
    return ctx.current_sheet();
  }
  if (ctx.workbook() == nullptr) {
    return nullptr;
  }
  return ctx.workbook()->sheet_by_name(ref_sheet);
}

}  // namespace

Value eval_isformula_lazy(const parser::AstNode& call, Arena& /*arena*/, const FunctionRegistry& /*registry*/,
                          const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  const parser::AstNode& arg = call.as_call_arg(0);
  if (arg.kind() != parser::NodeKind::Ref) {
    // Excel: `ISFORMULA(1)` and `ISFORMULA("A1")` both surface `#VALUE!`.
    // RangeOp would need array-spill to answer cell-by-cell; not in the
    // MVP.
    return Value::error(ErrorCode::Value);
  }
  const parser::Reference& r = arg.as_ref();
  if (r.is_full_col || r.is_full_row) {
    return Value::error(ErrorCode::Value);
  }
  const Sheet* target = resolve_ref_sheet(r.sheet, ctx);
  if (target == nullptr) {
    // Unbound context or missing qualified sheet. Excel returns `#REF!`
    // for a broken sheet qualifier; `#NAME?` for an unbound context is
    // consistent with our `resolve_ref` treatment.
    return Value::error(ctx.current_sheet() == nullptr ? ErrorCode::Name : ErrorCode::Ref);
  }
  const Cell* cell = target->cell_at(r.row, r.col);
  const bool is_formula = (cell != nullptr) && !cell->formula_text.empty();
  return Value::boolean(is_formula);
}

Value eval_formulatext_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& /*registry*/,
                            const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  const parser::AstNode& arg = call.as_call_arg(0);
  if (arg.kind() != parser::NodeKind::Ref) {
    // Excel surfaces `#VALUE!` for non-reference arguments (literals,
    // RangeOp, function calls). Full-row / full-column refs are also
    // out of scope for this MVP.
    return Value::error(ErrorCode::Value);
  }
  const parser::Reference& r = arg.as_ref();
  if (r.is_full_col || r.is_full_row) {
    return Value::error(ErrorCode::Value);
  }
  const Sheet* target = resolve_ref_sheet(r.sheet, ctx);
  if (target == nullptr) {
    return Value::error(ctx.current_sheet() == nullptr ? ErrorCode::Name : ErrorCode::Ref);
  }
  const Cell* cell = target->cell_at(r.row, r.col);
  if (cell == nullptr || cell->formula_text.empty()) {
    // No formula in the target cell (blank, value-only, or missing
    // slot). Excel 365 answers `#N/A`.
    return Value::error(ErrorCode::NA);
  }
  // Copy the formula source into the evaluation arena, stripping the
  // xlsx-only `_xlfn.` / `_xlfn._xlws.` storage prefixes so the returned
  // text matches what the user typed (and what Excel 365 displays).
  //
  // Only strip when the prefix stands at a token-start position — i.e.
  // preceded by nothing, by an ASCII non-identifier byte, or by a quote
  // boundary — and not while we are inside a `"..."` string literal.
  // This avoids corrupting user text that happens to contain `_xlfn.`.
  const std::string_view src(cell->formula_text);
  std::string stripped;
  stripped.reserve(src.size());
  constexpr std::string_view kXlws = "_xlfn._xlws.";
  constexpr std::string_view kXlfn = "_xlfn.";
  bool in_string = false;
  for (std::size_t i = 0; i < src.size();) {
    const char c = src[i];
    if (c == '"') {
      in_string = !in_string;
      stripped.push_back(c);
      ++i;
      continue;
    }
    if (!in_string) {
      const char prev = stripped.empty() ? '\0' : stripped.back();
      const bool at_token_start =
          prev == '\0' || !((prev >= 'A' && prev <= 'Z') || (prev >= 'a' && prev <= 'z') ||
                             (prev >= '0' && prev <= '9') || prev == '_' || prev == '.');
      if (at_token_start) {
        if (i + kXlws.size() <= src.size() &&
            strings::case_insensitive_eq(src.substr(i, kXlws.size()), kXlws)) {
          i += kXlws.size();
          continue;
        }
        if (i + kXlfn.size() <= src.size() &&
            strings::case_insensitive_eq(src.substr(i, kXlfn.size()), kXlfn)) {
          i += kXlfn.size();
          continue;
        }
      }
    }
    stripped.push_back(c);
    ++i;
  }
  char* buf = static_cast<char*>(arena.allocate(stripped.size(), alignof(char)));
  if (buf == nullptr) {
    return Value::error(ErrorCode::Value);
  }
  std::copy(stripped.begin(), stripped.end(), buf);
  return Value::text(std::string_view(buf, stripped.size()));
}

Value eval_isref_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  const parser::AstNode& arg = call.as_call_arg(0);
  if (is_static_reference_shape(arg)) {
    return Value::boolean(true);
  }
  if (arg.kind() == parser::NodeKind::Call && is_reference_call_name(arg.as_call_name())) {
    // Reference-returning calls must actually produce a non-error
    // result to count as "ref". E.g. `ISREF(INDIRECT("not a ref"))`
    // is FALSE because INDIRECT surfaces `#REF!`.
    const Value v = eval_node(arg, arena, registry, ctx);
    return Value::boolean(!v.is_error());
  }
  return Value::boolean(false);
}

Value eval_sheet_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity > 1U) {
    return Value::error(ErrorCode::Value);
  }
  const Sheet* current = ctx.current_sheet();
  const Workbook* wb = ctx.workbook();

  // 0-arity: number of the currently-bound sheet.
  if (arity == 0U) {
    if (current == nullptr) {
      return Value::error(ErrorCode::Value);
    }
    if (wb == nullptr) {
      // Bound only to a Sheet, no workbook. Convention: single sheet = 1.
      return Value::number(1.0);
    }
    const std::size_t idx = find_sheet_index(*wb, *current);
    if (idx == 0) {
      return Value::error(ErrorCode::Value);
    }
    return Value::number(static_cast<double>(idx));
  }

  const parser::AstNode& arg = call.as_call_arg(0);

  // Reference AST: the answer is determined by the sheet qualifier
  // without needing to evaluate the reference.
  std::string_view sheet_qualifier;
  bool have_ref_shape = false;
  if (arg.kind() == parser::NodeKind::Ref) {
    sheet_qualifier = arg.as_ref().sheet;
    have_ref_shape = true;
  } else if (arg.kind() == parser::NodeKind::RangeOp) {
    const parser::AstNode& lhs = arg.as_range_lhs();
    if (lhs.kind() == parser::NodeKind::Ref) {
      sheet_qualifier = lhs.as_ref().sheet;
      have_ref_shape = true;
    }
  }
  if (have_ref_shape) {
    if (sheet_qualifier.empty()) {
      // Unqualified ref — same sheet as the host context.
      if (current == nullptr) {
        return Value::error(ErrorCode::Value);
      }
      if (wb == nullptr) {
        return Value::number(1.0);
      }
      const std::size_t idx = find_sheet_index(*wb, *current);
      return idx == 0 ? Value::error(ErrorCode::Value) : Value::number(static_cast<double>(idx));
    }
    // Qualified ref — must have a workbook to look up the qualifier.
    if (wb == nullptr) {
      return Value::error(ErrorCode::NA);
    }
    const Sheet* target = wb->sheet_by_name(sheet_qualifier);
    if (target == nullptr) {
      return Value::error(ErrorCode::NA);
    }
    const std::size_t idx = find_sheet_index(*wb, *target);
    return idx == 0 ? Value::error(ErrorCode::NA) : Value::number(static_cast<double>(idx));
  }

  // Non-reference AST: evaluate the argument. If it coerces to text,
  // treat as a sheet-name lookup. Otherwise surface `#VALUE!` (or the
  // propagated error).
  const Value v = eval_node(arg, arena, registry, ctx);
  if (v.is_error()) {
    return v;
  }
  auto text_e = coerce_to_text(v);
  if (!text_e) {
    return Value::error(ErrorCode::Value);
  }
  if (wb == nullptr) {
    return Value::error(ErrorCode::NA);
  }
  const Sheet* target = wb->sheet_by_name(text_e.value());
  if (target == nullptr) {
    return Value::error(ErrorCode::NA);
  }
  const std::size_t idx = find_sheet_index(*wb, *target);
  return idx == 0 ? Value::error(ErrorCode::NA) : Value::number(static_cast<double>(idx));
}

Value eval_sheets_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity > 1U) {
    return Value::error(ErrorCode::Value);
  }
  const Workbook* wb = ctx.workbook();

  if (arity == 0U) {
    // Total sheets in the workbook. No workbook -> degrade to 1, which
    // matches the "bound to a single Sheet" context shape.
    if (wb == nullptr) {
      return Value::number(1.0);
    }
    return Value::number(static_cast<double>(wb->sheet_count()));
  }

  const parser::AstNode& arg = call.as_call_arg(0);
  // Reference or range: MVP returns 1 (no 3D references yet). Errors
  // in the argument subtree still propagate if we evaluate.
  if (arg.kind() == parser::NodeKind::Ref || arg.kind() == parser::NodeKind::RangeOp ||
      arg.kind() == parser::NodeKind::ExternalRef || arg.kind() == parser::NodeKind::StructuredRef ||
      arg.kind() == parser::NodeKind::NameRef) {
    return Value::number(1.0);
  }
  // Fall through: evaluate to surface errors, then reject non-references.
  // Mac Excel 365 surfaces `#N/A` here (not `#VALUE!`), matching Excel's
  // treatment of SHEETS("not a ref") as a missing-sheet lookup.
  const Value v = eval_node(arg, arena, registry, ctx);
  if (v.is_error()) {
    return v;
  }
  return Value::error(ErrorCode::NA);
}

}  // namespace eval
}  // namespace formulon
