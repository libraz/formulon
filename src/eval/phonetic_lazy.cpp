//
// Implementation of the PHONETIC lazy form. See eval/phonetic_lazy.h
// for the public contract.

#include "eval/phonetic_lazy.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/range_resolvers.h"
#include "excel_locale.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "phonetic.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/text_ops.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Applies ja-JP's strict-text passthrough to a flattened scalar `v`:
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

// The reading of cell `r` of `target`: its annotated spans substituted, the
// surface text otherwise, and the empty text for a number, boolean or error.
Value cell_phonetic(const Sheet& target, const parser::Reference& r, const EvalContext& text_ctx, Arena& arena) {
  const Cell* cell = target.cell_at(r.row, r.col);
  const Value surface = text_ctx.resolve_ref(r);
  if (cell != nullptr && !cell->phonetic_runs.empty()) {
    const std::string composed =
        compose_phonetic(surface.is_text() ? surface.as_text() : std::string_view{}, cell->phonetic_runs);
    return Value::text(arena.intern(composed));
  }
  if (surface.is_number() || surface.is_error() || surface.is_boolean()) {
    return Value::text(arena.intern(""));
  }
  return apply_passthrough_surface(surface, arena);
}

// Advances `*byte` past one UTF-8 sequence of `surface` and adds the
// UTF-16 cost of what it decoded to `*unit`. A malformed leading byte
// costs one unit, matching `utf16_units_in`, so the two agree on where
// an `<rPh>` offset lands even for input Excel would never emit.
void StepOneCodepoint(std::string_view surface, std::size_t* byte, std::uint32_t* unit) {
  std::size_t step = 0;
  const std::uint32_t cp = decode_utf8_step(surface, *byte, &step);
  if (step == 0) {
    step = 1;
  }
  *byte += step;
  *unit += cp > 0xFFFFu ? 2U : 1U;
}

}  // namespace

std::string compose_phonetic(std::string_view surface, const std::vector<PhoneticRun>& runs) {
  if (runs.empty()) {
    return std::string(surface);
  }
  std::string out;
  out.reserve(surface.size());

  std::size_t next_run = 0;
  std::size_t byte = 0;
  std::uint32_t unit = 0;
  while (byte < surface.size() || next_run < runs.size()) {
    // `sb <= unit` rather than `==` so a run whose span the walk has
    // already passed still contributes its kana instead of vanishing.
    if (next_run < runs.size() && runs[next_run].sb <= unit) {
      const PhoneticRun& run = runs[next_run];
      out.append(run.text);
      ++next_run;
      // Swallow the annotated span: those characters are represented by
      // the kana that was just emitted.
      while (unit < run.eb && byte < surface.size()) {
        StepOneCodepoint(surface, &byte, &unit);
      }
      continue;
    }
    if (byte >= surface.size()) {
      break;
    }
    const std::size_t start = byte;
    StepOneCodepoint(surface, &byte, &unit);
    out.append(surface.substr(start, byte - start));
  }
  return out;
}

Value eval_phonetic_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    // Mac surfaces #VALUE! for arity 0 / 2+; mirrors the registered
    // FunctionDef's `min_arity == max_arity == 1` contract.
    return Value::error(ErrorCode::Value);
  }
  if (!locale_facts(ctx.excel_profile()).phonetic) {
    return Value::error(ErrorCode::NA);
  }
  const parser::AstNode& arg = call.as_call_arg(0);

  // One cell: its reading directly, with ISFORMULA / FORMULATEXT's
  // broken-qualifier #REF! / #NAME? mapping.
  if (arg.kind() == parser::NodeKind::Ref && !arg.as_ref().is_full_col && !arg.as_ref().is_full_row) {
    const parser::Reference& r = arg.as_ref();
    const Sheet* target = ctx.sheet_for_qualifier(r.sheet);
    if (target == nullptr) {
      // Unbound context: #NAME?. Missing qualified sheet: #REF!. Same
      // mapping as ISFORMULA / FORMULATEXT.
      return Value::error(ctx.current_sheet() == nullptr ? ErrorCode::Name : ErrorCode::Ref);
    }
    // Excel does not count PHONETIC's read toward circularity:
    // `=PHONETIC(A1)` in A1 is not circular (Mac Excel 365).
    return cell_phonetic(*target, r, ctx.without_read_observer(), arena);
  }

  // A range, whole columns and rows included: the readings of its
  // populated cells joined row by row (a blank cell reads as "").
  if (arg.kind() == parser::NodeKind::RangeOp || arg.kind() == parser::NodeKind::Ref) {
    std::string_view qualifier;
    RefArea area{};
    ErrorCode err = ErrorCode::Value;
    if (!eval_reference_area(arg, arena, registry, ctx, &qualifier, &area, &err)) {
      return Value::error(err);
    }
    const Sheet* target = ctx.sheet_for_qualifier(qualifier);
    if (target == nullptr) {
      return Value::error(ctx.current_sheet() == nullptr ? ErrorCode::Name : ErrorCode::Ref);
    }
    std::vector<std::pair<std::uint32_t, std::uint32_t>> populated;
    std::uint64_t cursor = 0;
    while (cursor != Sheet::kCellCursorEnd) {
      cursor = target->cells_in_range(
          area.row_first, area.col_first, area.row_last, area.col_last, cursor, 4096U,
          [](const Sheet::RangeCell& cell, void* out) {
            static_cast<std::vector<std::pair<std::uint32_t, std::uint32_t>>*>(out)->emplace_back(cell.row, cell.col);
          },
          &populated);
    }
    const EvalContext text_ctx = ctx.without_read_observer();
    std::string joined;
    parser::Reference r{};
    r.sheet = qualifier;
    for (const auto& at : populated) {
      r.row = at.first;
      r.col = at.second;
      const Value part = cell_phonetic(*target, r, text_ctx, arena);
      if (!part.is_text()) {
        return part;
      }
      joined.append(part.as_text());
    }
    return Value::text(arena.intern(joined));
  }

  // Any other arg (literal text, arithmetic, function call, ...):
  // eagerly evaluate the subtree so error propagation and the "literal
  // text passes through" rule still apply in ja-JP. Everything non-text
  // yields #N/A.
  const Value v = eval_node(arg, arena, registry, ctx);
  return apply_passthrough_surface(v, arena);
}

}  // namespace eval
}  // namespace formulon
