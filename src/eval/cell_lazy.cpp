// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of `CELL(info_type, [reference])`. See `cell_lazy.h`
// for the contract and the explicit list of MVP stub keys.
//
// CELL rides the lazy seam because the reference argument is optional
// and, when supplied, may be a multi-cell range whose top-left we
// extract without dereferencing the whole rectangle. The eager path
// would flatten both shapes to a `Value` before the impl runs, losing
// the address information we need for "address" / "row" / "col".

#include "eval/cell_lazy.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/range_resolvers.h"
#include "io/ooxml_writer_cell.h"
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

// ASCII-lowercase fold for the info_type key. Excel matches CELL's first
// argument case-insensitively (`"ROW"` == `"row"` == `"Row"`); only ASCII
// is in the supported keyset so a byte-level fold suffices.
std::string ascii_tolower(std::string_view s) {
  std::string out;
  out.reserve(s.size());
  for (char c : s) {
    const auto u = static_cast<unsigned char>(c);
    if (u >= 'A' && u <= 'Z') {
      out.push_back(static_cast<char>(u + 32));
    } else {
      out.push_back(c);
    }
  }
  return out;
}

// Copies `s` into `arena` and returns a Text Value pointing at the
// arena-owned storage. Used for keys whose result is a constant string
// ("filename", "format", "prefix"). For genuinely empty strings we still
// return `Value::text({})` so we don't allocate a zero-byte buffer.
Value arena_text(Arena& arena, std::string_view s) {
  if (s.empty()) {
    return Value::text({});
  }
  char* buf = static_cast<char*>(arena.allocate(s.size(), alignof(char)));
  if (buf == nullptr) {
    return Value::error(ErrorCode::Value);
  }
  for (std::size_t i = 0; i < s.size(); ++i) {
    buf[i] = s[i];
  }
  return Value::text(std::string_view(buf, s.size()));
}

// Resolves the optional `reference` argument to the top-left cell of the
// rectangle it denotes, plus an optional sheet qualifier string when the
// reference targets a sheet other than the formula cell's sheet.
//
// Behaviour:
//   * `Ref` / `RangeOp` - delegate to `resolve_range_endpoint`, which
//     understands plain Refs and reference-returning calls (OFFSET /
//     INDIRECT). For RangeOp we resolve the lhs endpoint and then
//     normalise the rectangle's top-left as the smallest (row, col)
//     across both endpoints.
//   * Reference-returning Call (OFFSET / INDIRECT / nested CHOOSE / IF)
//     - `resolve_reference_call` produces the rectangle directly.
//   * Anything else - evaluate via `eval_node` so a subtree error
//     propagates with the correct code; if the result is a non-error
//     non-reference scalar, surface `#VALUE!` (CELL needs a reference,
//     not a Value).
//
// Returns `true` on success and writes the 0-based top-left coordinates
// to `*out_row` / `*out_col` plus the sheet qualifier (empty when the
// resolved cell is on the bound current sheet) to `*out_sheet`.
// Returns `false` on failure with the Excel error code surfaced via
// `*out_err`.
bool extract_topleft_ref(const parser::AstNode& ref_node, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, std::uint32_t* out_row, std::uint32_t* out_col,
                         std::string_view* out_sheet, Value* out_err) {
  const parser::NodeKind k = ref_node.kind();

  // Plain Ref or RangeOp: route through `resolve_range_endpoint` for the
  // lhs to get a normalised rectangle. For a RangeOp we also resolve the
  // rhs and union the rectangles so the top-left is the min(row,col)
  // pair, matching Excel's "top-left of the supplied range" rule.
  if (k == parser::NodeKind::Ref) {
    std::string_view sheet;
    std::uint32_t top = 0;
    std::uint32_t left = 0;
    std::uint32_t bottom = 0;
    std::uint32_t right = 0;
    ErrorCode err = ErrorCode::Ref;
    if (!resolve_range_endpoint(ref_node, arena, registry, ctx, &sheet, &top, &left, &bottom, &right, &err)) {
      *out_err = Value::error(err);
      return false;
    }
    *out_row = top;
    *out_col = left;
    *out_sheet = sheet;
    return true;
  }
  if (k == parser::NodeKind::RangeOp) {
    const parser::AstNode& lhs_ast = ref_node.as_range_lhs();
    const parser::AstNode& rhs_ast = ref_node.as_range_rhs();
    std::string_view lhs_sheet;
    std::string_view rhs_sheet;
    std::uint32_t lhs_top = 0;
    std::uint32_t lhs_left = 0;
    std::uint32_t lhs_bottom = 0;
    std::uint32_t lhs_right = 0;
    std::uint32_t rhs_top = 0;
    std::uint32_t rhs_left = 0;
    std::uint32_t rhs_bottom = 0;
    std::uint32_t rhs_right = 0;
    ErrorCode err = ErrorCode::Ref;
    if (!resolve_range_endpoint(lhs_ast, arena, registry, ctx, &lhs_sheet, &lhs_top, &lhs_left, &lhs_bottom, &lhs_right,
                                &err) ||
        !resolve_range_endpoint(rhs_ast, arena, registry, ctx, &rhs_sheet, &rhs_top, &rhs_left, &rhs_bottom, &rhs_right,
                                &err)) {
      *out_err = Value::error(err);
      return false;
    }
    *out_row = lhs_top < rhs_top ? lhs_top : rhs_top;
    *out_col = lhs_left < rhs_left ? lhs_left : rhs_left;
    // Prefer the lhs sheet qualifier when present; falls back to rhs's
    // qualifier so `Sheet2!A1:B2` (parser's typical RangeOp shape) and
    // the defensive `A1:Sheet2!B2` shape both surface the right sheet.
    *out_sheet = lhs_sheet.empty() ? rhs_sheet : lhs_sheet;
    return true;
  }

  // Reference-returning call: OFFSET / INDIRECT / nested CHOOSE / IF.
  // `resolve_reference_call` walks these and produces a rectangle.
  if (k == parser::NodeKind::Call) {
    std::string_view sheet;
    std::uint32_t top = 0;
    std::uint32_t left = 0;
    std::uint32_t bottom = 0;
    std::uint32_t right = 0;
    bool is_range = false;
    ErrorCode err = ErrorCode::Value;
    if (resolve_reference_call(ref_node, arena, registry, ctx, &sheet, &top, &left, &bottom, &right, &is_range, &err)) {
      *out_row = top;
      *out_col = left;
      *out_sheet = sheet;
      return true;
    }
    // Fall through: not a reference-returning call. Evaluate so any
    // subtree error propagates verbatim, then reject the scalar result.
  }

  // Anything else: evaluate the subtree and propagate a subtree error.
  // A non-error, non-reference scalar surfaces `#VALUE!` because CELL
  // requires a reference for the address / row / col / contents / type
  // info_types.
  const Value v = eval_node(ref_node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  *out_err = Value::error(ErrorCode::Value);
  return false;
}

// Resolves the top-left cell's value through the bound context. Used by
// the "contents" and "type" keys. The sheet qualifier may be empty (use
// current sheet) or non-empty (look up via workbook); when no current
// sheet is bound at all we surface `#NAME?` to mirror
// `EvalContext::resolve_ref`'s degradation rule.
Value resolve_topleft_value(std::string_view sheet, std::uint32_t row, std::uint32_t col, Arena& arena,
                            const FunctionRegistry& registry, const EvalContext& ctx) {
  parser::Reference r{};
  r.sheet = sheet;
  r.row = row;
  r.col = col;
  return ctx.resolve_ref(r, arena, registry);
}

// Builds the `"width"` MVP stub: a 1x2 array containing the default
// column width (8) and the auto-fit flag (TRUE). No column-width
// metadata yet; the result is fixed for every cell.
Value build_width_stub(Arena& arena) {
  Value* cells = arena.create_array<Value>(2);
  if (cells == nullptr) {
    return Value::error(ErrorCode::Value);
  }
  cells[0] = Value::number(8.0);
  cells[1] = Value::boolean(true);
  ArrayValue* arr = arena.create<ArrayValue>();
  if (arr == nullptr) {
    return Value::error(ErrorCode::Value);
  }
  arr->rows = 1U;
  arr->cols = 2U;
  arr->cells = cells;
  return Value::array(arr);
}

// Builds the `"address"` answer: an A1 absolute address (`$Col$Row`)
// optionally prefixed with a sheet qualifier. Sheet prefixing follows
// Excel's bare-name rule (no quoting for plain identifiers); we keep
// the formatting minimal because the workbook has no filename to wrap
// in `[Workbook]` brackets.
Value build_address(Arena& arena, std::string_view sheet, std::uint32_t row, std::uint32_t col) {
  // `EncodeA1` returns an unanchored address (e.g. "A1"); we splice the
  // `$` anchors in manually because the caller wants an absolute form.
  // The helper accepts 0-based (row, col) and produces 1-based output.
  const std::string a1 = io::EncodeA1(row, col);
  // Find the boundary between the column letters and the row digits so
  // we can insert the leading `$` and the inner `$` for absolute form.
  std::size_t row_start = 0;
  while (row_start < a1.size() && a1[row_start] >= 'A' && a1[row_start] <= 'Z') {
    ++row_start;
  }
  // Compose `[sheet!]$Col$Row` directly into a std::string, then copy
  // the bytes into the arena via `arena_text`.
  std::string result;
  result.reserve(sheet.size() + a1.size() + 4);
  if (!sheet.empty()) {
    result.append(sheet);
    result.push_back('!');
  }
  result.push_back('$');
  result.append(a1.substr(0, row_start));
  result.push_back('$');
  result.append(a1.substr(row_start));
  return arena_text(arena, result);
}

// Resolves the reference argument or, when omitted, the formula cell's
// own coords. Returns `true` on success and writes the top-left into
// `*out_row` / `*out_col` plus the sheet qualifier (empty for the
// current sheet). Failure returns `false` with the surfaced Value
// (error or otherwise) in `*out_result`.
bool resolve_topleft_or_formula_cell(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                     const EvalContext& ctx, std::uint32_t* out_row, std::uint32_t* out_col,
                                     std::string_view* out_sheet, Value* out_result) {
  if (call.as_call_arity() == 1U) {
    // No reference supplied: anchor to the formula cell. Without an
    // anchor (CLI ad-hoc eval) there's nothing meaningful to answer
    // about row / col / address, so surface `#REF!` defensively.
    if (!ctx.has_formula_cell()) {
      *out_result = Value::error(ErrorCode::Ref);
      return false;
    }
    *out_row = ctx.formula_row();
    *out_col = ctx.formula_col();
    *out_sheet = std::string_view{};
    return true;
  }
  // Two-argument form: extract from the supplied reference AST.
  Value err = Value::blank();
  if (!extract_topleft_ref(call.as_call_arg(1), arena, registry, ctx, out_row, out_col, out_sheet, &err)) {
    *out_result = err;
    return false;
  }
  return true;
}

}  // namespace

Value eval_cell_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 2U) {
    return Value::error(ErrorCode::Value);
  }

  // info_type: evaluate, propagate any error, then coerce to text. An
  // Array argument fails the coercion and surfaces `#VALUE!`. An empty
  // string is also `#VALUE!` because no Excel-defined key is empty.
  const Value info_v = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (info_v.is_error()) {
    return info_v;
  }
  auto info_text = coerce_to_text(info_v);
  if (!info_text) {
    return Value::error(info_text.error());
  }
  if (info_text.value().empty()) {
    return Value::error(ErrorCode::Value);
  }
  const std::string key = ascii_tolower(info_text.value());

  // Reference-dependent keys: address / col / row / contents / type.
  // Resolve the top-left first so any reference-side error propagates
  // before we branch on the key.
  if (key == "address" || key == "col" || key == "row" || key == "contents" || key == "type") {
    std::uint32_t row = 0;
    std::uint32_t col = 0;
    std::string_view sheet;
    Value early_result = Value::blank();
    if (!resolve_topleft_or_formula_cell(call, arena, registry, ctx, &row, &col, &sheet, &early_result)) {
      return early_result;
    }
    if (key == "address") {
      return build_address(arena, sheet, row, col);
    }
    if (key == "col") {
      return Value::number(static_cast<double>(col + 1U));
    }
    if (key == "row") {
      return Value::number(static_cast<double>(row + 1U));
    }
    // "contents" / "type" both need the resolved cell's value.
    const Value resolved = resolve_topleft_value(sheet, row, col, arena, registry, ctx);
    if (key == "contents") {
      // Excel's blank-as-zero rule for CELL("contents", blank_cell).
      if (resolved.is_blank()) {
        return Value::number(0.0);
      }
      return resolved;
    }
    // "type": "b" for blank or empty-string text, "l" for non-empty text,
    // "v" for everything else (number, bool, error). Errors short-circuit
    // only when they came from the reference argument itself; an error
    // *value* sitting in a cell still classifies as "v". Mac Excel folds
    // an empty string to "b" rather than "l".
    if (resolved.is_blank()) {
      return arena_text(arena, "b");
    }
    if (resolved.is_text()) {
      return arena_text(arena, resolved.as_text().empty() ? "b" : "l");
    }
    return arena_text(arena, "v");
  }

  // Reference-independent keys. These ignore the supplied reference
  // entirely (Excel observed: passes through without dereferencing,
  // even when the reference would otherwise error). If the reference
  // argument itself is a top-level error literal that path is rare;
  // the spec calls for fixed stubs here so we don't pre-evaluate.
  if (key == "filename") {
    // No filesystem path on the Workbook yet. Mac returns blank when the
    // workbook has never been saved; we surface empty text instead so
    // the top-level blank-as-zero rule (`evaluate()` in tree_walker)
    // does not collapse the result to 0. Oracle verification accepts the
    // xlwings "" read-back artifact via empty_string_readback.
    return arena_text(arena, "");
  }
  if (key == "format") {
    // No style metadata yet; "G" is Excel's "general format" code.
    return arena_text(arena, "G");
  }
  if (key == "color") {
    // No negative-number color flag yet.
    return Value::number(0.0);
  }
  if (key == "parentheses") {
    // No parenthesis-format flag yet.
    return Value::number(0.0);
  }
  if (key == "prefix") {
    // No text-alignment metadata yet (apostrophe / caret / quote /
    // backslash). Mac returns blank when no prefix character is set;
    // we surface empty text instead so the top-level blank-as-zero rule
    // does not collapse the result to 0. Oracle verification accepts the
    // xlwings "" read-back artifact via empty_string_readback.
    return arena_text(arena, "");
  }
  if (key == "protect") {
    // Default-locked until the style subsystem lands. Excel's CELL
    // returns 1 for locked cells.
    return Value::number(1.0);
  }
  if (key == "width") {
    // No column-width metadata yet; return the {default_width, autofit}
    // pair as a 1x2 array.
    return build_width_stub(arena);
  }

  // Unknown info_type after lowercase fold.
  return Value::error(ErrorCode::Value);
}

}  // namespace eval
}  // namespace formulon
