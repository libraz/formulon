// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the reference-manipulation lazy builtins `INDIRECT`
// and `OFFSET`, plus the shared `parse_a1_ref` / `column_letters`
// helpers they (and `ADDRESS`) consume. See `reference_lazy.h` for the
// dispatch-table contract and the design notes on why these two forms
// ride the lazy seam.

#include "eval/reference_lazy.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string_view>
#include <vector>

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

namespace refs_internal {

namespace {

// ASCII letter / digit predicates — intentionally local to avoid pulling
// in `<cctype>` which is locale-sensitive on some platforms.
constexpr bool is_letter(char ch) noexcept {
  return (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z');
}

constexpr bool is_digit(char ch) noexcept {
  return ch >= '0' && ch <= '9';
}

// Parses letters (A..XFD) at `text[*i]` into a 1-based column index, then
// advances `*i` past the consumed bytes. Returns 0 on malformed input
// (no letters, too many letters, or overflow past XFD = 16384).
std::uint32_t parse_column_letters(std::string_view text, std::size_t* i) {
  std::uint32_t col = 0;
  std::size_t letters_seen = 0;
  while (*i < text.size() && is_letter(text[*i])) {
    char ch = text[*i];
    if (ch >= 'a' && ch <= 'z') {
      ch = static_cast<char>(ch - ('a' - 'A'));
    }
    col = col * 26u + static_cast<std::uint32_t>(ch - 'A' + 1);
    ++(*i);
    ++letters_seen;
    if (letters_seen > 3 || col > Sheet::kMaxCols) {
      return 0;
    }
  }
  if (letters_seen == 0) {
    return 0;
  }
  return col;
}

// Parses digits at `text[*i]` into a 1-based row index, advancing `*i`
// past them. Returns 0 on no digits, too many digits, or row > kMaxRows.
std::uint32_t parse_row_digits(std::string_view text, std::size_t* i) {
  std::uint64_t row = 0;
  std::size_t digits_seen = 0;
  while (*i < text.size() && is_digit(text[*i])) {
    row = row * 10u + static_cast<std::uint32_t>(text[*i] - '0');
    ++(*i);
    ++digits_seen;
    if (digits_seen > 7 || row > Sheet::kMaxRows) {
      return 0;
    }
  }
  if (digits_seen == 0 || row == 0) {
    return 0;
  }
  return static_cast<std::uint32_t>(row);
}

// Attempts to parse `text[start..]` as a full-column shape
// `[$]?<letters>:[$]?<letters>` consuming the entire remainder. On
// success sets `out->is_full_col` / `out->is_range`, populates
// `col`/`col2` with the min/max 0-based columns, and fills
// `row`/`row2` with the full-column row span, then returns true. On
// failure leaves `*out` untouched and returns false.
bool try_parse_full_col(std::string_view text, std::size_t start, A1Parse* out) {
  std::size_t i = start;
  if (i < text.size() && text[i] == '$') {
    ++i;
  }
  const std::uint32_t c1 = parse_column_letters(text, &i);
  if (c1 == 0) {
    return false;
  }
  if (i >= text.size() || text[i] != ':') {
    return false;
  }
  ++i;
  if (i < text.size() && text[i] == '$') {
    ++i;
  }
  const std::uint32_t c2 = parse_column_letters(text, &i);
  if (c2 == 0) {
    return false;
  }
  if (i != text.size()) {
    return false;
  }
  const std::uint32_t lo = std::min(c1, c2) - 1U;
  const std::uint32_t hi = std::max(c1, c2) - 1U;
  out->col = lo;
  out->col2 = hi;
  out->row = 0;
  out->row2 = Sheet::kMaxRows - 1U;
  out->is_full_col = true;
  out->is_range = true;
  return true;
}

// Attempts to parse `text[start..]` as a full-row shape
// `[$]?<digits>:[$]?<digits>` consuming the entire remainder. On
// success sets `out->is_full_row` / `out->is_range`, populates
// `row`/`row2` with the min/max 0-based rows, and fills `col`/`col2`
// with the full-row column span, then returns true.
bool try_parse_full_row(std::string_view text, std::size_t start, A1Parse* out) {
  std::size_t i = start;
  if (i < text.size() && text[i] == '$') {
    ++i;
  }
  const std::uint32_t r1 = parse_row_digits(text, &i);
  if (r1 == 0) {
    return false;
  }
  if (i >= text.size() || text[i] != ':') {
    return false;
  }
  ++i;
  if (i < text.size() && text[i] == '$') {
    ++i;
  }
  const std::uint32_t r2 = parse_row_digits(text, &i);
  if (r2 == 0) {
    return false;
  }
  if (i != text.size()) {
    return false;
  }
  const std::uint32_t lo = std::min(r1, r2) - 1U;
  const std::uint32_t hi = std::max(r1, r2) - 1U;
  out->row = lo;
  out->row2 = hi;
  out->col = 0;
  out->col2 = Sheet::kMaxCols - 1U;
  out->is_full_row = true;
  out->is_range = true;
  return true;
}

// Parses a single A1 endpoint (optional `$` markers, letters, digits).
// Returns `false` on any malformed shape; on success writes 0-based
// row/col to `*out_row` / `*out_col` and advances `*i`.
bool parse_a1_endpoint(std::string_view text, std::size_t* i, std::uint32_t* out_row, std::uint32_t* out_col) {
  // Optional leading `$` on the column.
  if (*i < text.size() && text[*i] == '$') {
    ++(*i);
  }
  const std::uint32_t col_1based = parse_column_letters(text, i);
  if (col_1based == 0) {
    return false;
  }
  // Optional `$` between column and row.
  if (*i < text.size() && text[*i] == '$') {
    ++(*i);
  }
  const std::uint32_t row_1based = parse_row_digits(text, i);
  if (row_1based == 0) {
    return false;
  }
  *out_col = col_1based - 1U;
  *out_row = row_1based - 1U;
  return true;
}

// Copies `quoted` (everything between the surrounding single quotes)
// into `out`, collapsing each `''` pair into a single `'`. Returns the
// new size. `quoted` already excludes the outer quotes.
std::size_t unescape_quoted_sheet(std::string_view quoted, char* out, std::size_t cap) {
  std::size_t w = 0;
  for (std::size_t r = 0; r < quoted.size() && w < cap; ++r) {
    if (quoted[r] == '\'' && r + 1 < quoted.size() && quoted[r + 1] == '\'') {
      out[w++] = '\'';
      ++r;
      continue;
    }
    out[w++] = quoted[r];
  }
  return w;
}

}  // namespace

std::size_t column_letters(std::uint32_t col, char* out) {
  // Excel columns encode as 1-based bijective base-26 (no "zero letter"):
  //   1  -> A
  //   26 -> Z
  //   27 -> AA
  // The classic "divmod-26" loop has to subtract 1 each round to account
  // for the missing zero digit.
  char buf[4] = {0, 0, 0, 0};
  std::size_t n = 0;
  while (col > 0 && n < 3) {
    const std::uint32_t rem = (col - 1) % 26u;
    buf[n++] = static_cast<char>('A' + rem);
    col = (col - 1) / 26u;
  }
  // Written LSB-first; reverse into the caller's buffer.
  for (std::size_t i = 0; i < n; ++i) {
    out[i] = buf[n - 1 - i];
  }
  return n;
}

A1Parse parse_a1_ref(std::string_view text) {
  A1Parse out;
  if (text.empty()) {
    return out;
  }
  std::size_t i = 0;

  // Sheet qualifier detection. Two shapes:
  //   'Sheet Name'!A1 — quoted (supports spaces / punctuation, doubled `'`
  //                     is an escaped apostrophe inside the name).
  //   Sheet1!A1       — bare (letters / digits / `_`).
  if (text[0] == '\'') {
    // Scan for the closing `'` that is NOT followed by another `'` (the
    // doubled form is an escaped apostrophe and stays inside the name).
    std::size_t j = 1;
    while (j < text.size()) {
      if (text[j] == '\'') {
        if (j + 1 < text.size() && text[j + 1] == '\'') {
          j += 2;
          continue;
        }
        break;
      }
      ++j;
    }
    if (j >= text.size() || text[j] != '\'') {
      return out;  // unterminated
    }
    // Inside content is `text.substr(1, j - 1)`; unescape `''` -> `'`.
    // We don't own backing storage here, so write into a static-sized
    // local buffer; Excel sheet-name limit is 31 chars (we allow up to
    // 255 defensively, capped by the source view size).
    static thread_local char scratch[256];
    const std::size_t content_len = j - 1;
    const std::size_t used = unescape_quoted_sheet(text.substr(1, content_len), scratch, sizeof(scratch));
    // This view points into thread_local storage — callers keep it alive
    // only until the next `parse_a1_ref` call on the same thread. The
    // consuming code (`indirect`) copies before storing.
    out.sheet = std::string_view(scratch, used);
    i = j + 1;
    if (i >= text.size() || text[i] != '!') {
      return out;  // missing `!` after quoted sheet
    }
    ++i;
  } else {
    // Bare sheet name: run of letters/digits/underscore followed by `!`.
    // We only commit to treating it as a sheet qualifier if we find the
    // `!` — otherwise the run is part of the reference itself.
    std::size_t j = 0;
    while (j < text.size() && (is_letter(text[j]) || is_digit(text[j]) || text[j] == '_')) {
      ++j;
    }
    if (j > 0 && j < text.size() && text[j] == '!') {
      out.sheet = text.substr(0, j);
      i = j + 1;
    }
  }

  // Full-column / full-row shapes (`D:D`, `$FF:FG`, `5:5`, `$12:$23`)
  // are tried before the single-endpoint path because they never share a
  // prefix with a valid single-cell reference (the latter always has a
  // digit immediately after the letter run, never a `:`).
  if (try_parse_full_col(text, i, &out)) {
    out.valid = true;
    return out;
  }
  if (try_parse_full_row(text, i, &out)) {
    out.valid = true;
    return out;
  }

  // Parse the first endpoint.
  if (!parse_a1_endpoint(text, &i, &out.row, &out.col)) {
    return out;
  }

  // Optional `:` + second endpoint for ranges.
  if (i < text.size() && text[i] == ':') {
    ++i;
    if (!parse_a1_endpoint(text, &i, &out.row2, &out.col2)) {
      return out;
    }
    out.is_range = true;
  }

  // Trailing garbage -> invalid.
  if (i != text.size()) {
    return out;
  }
  out.valid = true;
  return out;
}

// Converts a signed offset plus a non-negative base (both measured in
// row / column units) into a 0-based grid coordinate, returning `false`
// if the result falls outside [0, max). `max` is `Sheet::kMaxRows` or
// `Sheet::kMaxCols`.
bool apply_offset(std::uint32_t base, int offset, std::uint32_t max, std::uint32_t* out) {
  const long long sum = static_cast<long long>(base) + static_cast<long long>(offset);
  if (sum < 0 || sum >= static_cast<long long>(max)) {
    return false;
  }
  *out = static_cast<std::uint32_t>(sum);
  return true;
}

// Reads an integer arg via truncation. `#VALUE!` on coercion failure,
// `#NUM!` on NaN/Inf.
Expected<int, ErrorCode> read_int(const Value& v) {
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    return coerced.error();
  }
  const double d = coerced.value();
  if (std::isnan(d) || std::isinf(d)) {
    return ErrorCode::Num;
  }
  return static_cast<int>(std::trunc(d));
}

// Shape of the base `reference` argument to OFFSET. Populated by
// `resolve_offset_base` from either a bare Ref (1x1) or a literal
// `RangeOp` rectangle. Stored as 0-based indices.
struct OffsetBase {
  std::string_view sheet;
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::uint32_t rows = 1;
  std::uint32_t cols = 1;
};

}  // namespace refs_internal

// Forward-declared for `refs_internal::resolve_offset_base`'s Call-branch.
bool resolve_reference_call(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx, std::string_view* out_sheet, std::uint32_t* out_top_row,
                            std::uint32_t* out_left_col, std::uint32_t* out_bottom_row, std::uint32_t* out_right_col,
                            bool* out_is_range, ErrorCode* out_err);

namespace refs_internal {

// Normalises the base reference from OFFSET's first argument. Literal
// Ref and `Ref:Ref` RangeOp shapes are accepted directly; INDIRECT /
// OFFSET nested calls go through `resolve_reference_call` so
// `OFFSET(INDIRECT("A1"), …)` and `OFFSET(OFFSET(A1,0,0,2,2), …)`
// resolve without dereferencing the base. Cross-sheet qualification is
// preserved so `OFFSET(Sheet2!A1, …)` routes through the right
// workbook sheet.
bool resolve_offset_base(const parser::AstNode& arg, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, OffsetBase* out, ErrorCode* out_err) {
  const parser::NodeKind k = arg.kind();
  if (k == parser::NodeKind::Ref) {
    const parser::Reference& r = arg.as_ref();
    if (r.is_full_col || r.is_full_row) {
      *out_err = ErrorCode::Value;
      return false;
    }
    out->sheet = r.sheet;
    out->row = r.row;
    out->col = r.col;
    out->rows = 1U;
    out->cols = 1U;
    return true;
  }
  if (k == parser::NodeKind::RangeOp) {
    const parser::AstNode& lhs_ast = arg.as_range_lhs();
    const parser::AstNode& rhs_ast = arg.as_range_rhs();
    if (lhs_ast.kind() != parser::NodeKind::Ref || rhs_ast.kind() != parser::NodeKind::Ref) {
      *out_err = ErrorCode::Ref;
      return false;
    }
    const parser::Reference& lhs = lhs_ast.as_ref();
    const parser::Reference& rhs = rhs_ast.as_ref();
    if (lhs.is_full_col || lhs.is_full_row || rhs.is_full_col || rhs.is_full_row) {
      *out_err = ErrorCode::Value;
      return false;
    }
    // The effective sheet qualifier mirrors `expand_range`: whichever
    // endpoint carries it wins, and mismatched qualifiers are `#REF!`.
    if (!lhs.sheet.empty() && !rhs.sheet.empty()) {
      if (!strings::case_insensitive_eq(lhs.sheet, rhs.sheet)) {
        *out_err = ErrorCode::Ref;
        return false;
      }
      out->sheet = lhs.sheet;
    } else if (!lhs.sheet.empty()) {
      out->sheet = lhs.sheet;
    } else if (!rhs.sheet.empty()) {
      out->sheet = rhs.sheet;
    }
    const std::uint32_t r_lo = std::min(lhs.row, rhs.row);
    const std::uint32_t r_hi = std::max(lhs.row, rhs.row);
    const std::uint32_t c_lo = std::min(lhs.col, rhs.col);
    const std::uint32_t c_hi = std::max(lhs.col, rhs.col);
    out->row = r_lo;
    out->col = c_lo;
    out->rows = r_hi - r_lo + 1U;
    out->cols = c_hi - c_lo + 1U;
    return true;
  }
  if (k == parser::NodeKind::Call) {
    // Nested INDIRECT / OFFSET as OFFSET's base: resolve to a rectangle
    // without dereferencing, then adopt it as the base shape.
    std::string_view sheet;
    std::uint32_t top = 0;
    std::uint32_t left = 0;
    std::uint32_t bottom = 0;
    std::uint32_t right = 0;
    bool is_range = false;
    ErrorCode err = ErrorCode::Value;
    if (!eval::resolve_reference_call(arg, arena, registry, ctx, &sheet, &top, &left, &bottom, &right, &is_range,
                                      &err)) {
      *out_err = err;
      return false;
    }
    out->sheet = sheet;
    out->row = top;
    out->col = left;
    out->rows = bottom - top + 1U;
    out->cols = right - left + 1U;
    return true;
  }
  // Anything else (literal, scalar expr, array literal, named ref) is
  // not a valid reference shape for OFFSET. Excel returns `#VALUE!`.
  *out_err = ErrorCode::Value;
  return false;
}

// Computes the (0-based) top-left corner and the (height, width) of the
// rectangle produced by `OFFSET(base, rows, cols, [height], [width])`.
// Returns `false` on any error, writing the Excel-visible code into
// `*out_err`.
bool compute_offset_rect(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, OffsetBase* out_base, std::uint32_t* out_top_row,
                         std::uint32_t* out_left_col, std::uint32_t* out_height, std::uint32_t* out_width,
                         ErrorCode* out_err) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 3U || arity > 5U) {
    *out_err = ErrorCode::Value;
    return false;
  }
  if (!resolve_offset_base(call.as_call_arg(0), arena, registry, ctx, out_base, out_err)) {
    return false;
  }

  // Evaluate rows / cols and optional height / width in turn. Any error
  // propagates with its original code.
  auto eval_int = [&](std::uint32_t idx, int* out_val) -> bool {
    const Value v = eval_node(call.as_call_arg(idx), arena, registry, ctx);
    if (v.is_error()) {
      *out_err = v.as_error();
      return false;
    }
    auto parsed = read_int(v);
    if (!parsed) {
      *out_err = parsed.error();
      return false;
    }
    *out_val = parsed.value();
    return true;
  };

  // Height / width get a slightly different coercion than rows_off /
  // cols_off: Mac Excel 365 truncates the fractional part toward zero,
  // but a non-zero magnitude < 1 (e.g. `0.9`, `-0.5`) is bumped up to
  // ±1 instead of collapsing to 0 (which would otherwise misfire the
  // `height == 0 || width == 0 -> #REF!` guard below). The bump is
  // sign-preserving so that negative-fractional widths still extend in
  // the negative direction.
  auto eval_dim = [&](std::uint32_t idx, int* out_val) -> bool {
    const Value v = eval_node(call.as_call_arg(idx), arena, registry, ctx);
    if (v.is_error()) {
      *out_err = v.as_error();
      return false;
    }
    auto coerced = coerce_to_number(v);
    if (!coerced) {
      *out_err = coerced.error();
      return false;
    }
    const double d = coerced.value();
    if (std::isnan(d) || std::isinf(d)) {
      *out_err = ErrorCode::Num;
      return false;
    }
    int truncated = static_cast<int>(std::trunc(d));
    if (truncated == 0 && d != 0.0) {
      truncated = (d > 0.0) ? 1 : -1;
    }
    *out_val = truncated;
    return true;
  };

  int rows_off = 0;
  int cols_off = 0;
  if (!eval_int(1U, &rows_off) || !eval_int(2U, &cols_off)) {
    return false;
  }

  int height_i = static_cast<int>(out_base->rows);
  int width_i = static_cast<int>(out_base->cols);
  if (arity >= 4U) {
    if (!eval_dim(3U, &height_i)) {
      return false;
    }
  }
  if (arity >= 5U) {
    if (!eval_dim(4U, &width_i)) {
      return false;
    }
  }
  // Zero height or width -> `#REF!`. Excel allows negative height / width
  // meaning the rectangle extends in the negative direction from the
  // anchor (anchor is the bottom-right corner of the rectangle instead
  // of the top-left). We normalise the absolute magnitude here and
  // adjust the anchor position below.
  if (height_i == 0 || width_i == 0) {
    *out_err = ErrorCode::Ref;
    return false;
  }
  const bool neg_height = height_i < 0;
  const bool neg_width = width_i < 0;
  const std::uint32_t abs_height = static_cast<std::uint32_t>(neg_height ? -height_i : height_i);
  const std::uint32_t abs_width = static_cast<std::uint32_t>(neg_width ? -width_i : width_i);

  // Apply the (rows, cols) offset to the base's top-left corner.
  std::uint32_t anchor_row = 0;
  std::uint32_t anchor_col = 0;
  if (!apply_offset(out_base->row, rows_off, Sheet::kMaxRows, &anchor_row) ||
      !apply_offset(out_base->col, cols_off, Sheet::kMaxCols, &anchor_col)) {
    *out_err = ErrorCode::Ref;
    return false;
  }

  // For negative height / width the anchor is the rectangle's bottom
  // (or right) edge: walk `abs_dim - 1` units back to find the top-left
  // corner. For positive dimensions the anchor is already the top-left.
  long long top_row = static_cast<long long>(anchor_row);
  long long left_col = static_cast<long long>(anchor_col);
  if (neg_height) {
    top_row -= static_cast<long long>(abs_height - 1);
  }
  if (neg_width) {
    left_col -= static_cast<long long>(abs_width - 1);
  }
  const long long bottom_row = top_row + static_cast<long long>(abs_height) - 1;
  const long long right_col = left_col + static_cast<long long>(abs_width) - 1;
  if (top_row < 0 || bottom_row >= static_cast<long long>(Sheet::kMaxRows) || left_col < 0 ||
      right_col >= static_cast<long long>(Sheet::kMaxCols)) {
    *out_err = ErrorCode::Ref;
    return false;
  }

  *out_top_row = static_cast<std::uint32_t>(top_row);
  *out_left_col = static_cast<std::uint32_t>(left_col);
  *out_height = abs_height;
  *out_width = abs_width;
  return true;
}

}  // namespace refs_internal

Value eval_indirect_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 2U) {
    return Value::error(ErrorCode::Value);
  }

  // Evaluate `ref_text` first so errors propagate per the dispatcher's
  // left-most-wins rule.
  const Value ref_val = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (ref_val.is_error()) {
    return ref_val;
  }
  auto text_exp = coerce_to_text(ref_val);
  if (!text_exp) {
    return Value::error(text_exp.error());
  }

  // `a1` flag: truthy / absent -> A1; falsy -> R1C1. Excel accepts any
  // coercible value. R1C1 style is deferred in this MVP.
  bool a1_style = true;
  if (arity == 2U) {
    const Value a1_val = eval_node(call.as_call_arg(1), arena, registry, ctx);
    if (a1_val.is_error()) {
      return a1_val;
    }
    auto b = coerce_to_bool(a1_val);
    if (!b) {
      return Value::error(b.error());
    }
    a1_style = b.value();
  }
  if (!a1_style) {
    // R1C1 path deferred — INDIRECT can return a reference today only via
    // the A1 parser. Document in `divergence.yaml` when this bites.
    return Value::error(ErrorCode::Ref);
  }

  const std::string& src = text_exp.value();
  if (src.empty()) {
    return Value::error(ErrorCode::Ref);
  }

  refs_internal::A1Parse parsed = refs_internal::parse_a1_ref(src);
  if (!parsed.valid) {
    return Value::error(ErrorCode::Ref);
  }
  // Range INDIRECT is deferred until `Value::Array` lands. `=SUM(INDIRECT(
  // "A1:B2"))` therefore surfaces as `#REF!` in scalar context today.
  if (parsed.is_range) {
    return Value::error(ErrorCode::Ref);
  }

  // The sheet view may point at thread-local scratch owned by
  // `parse_a1_ref`; intern it into the evaluation arena so the
  // `parser::Reference` we build outlives the next call.
  parser::Reference ref{};
  if (!parsed.sheet.empty()) {
    ref.sheet = arena.intern(parsed.sheet);
  }
  ref.row = parsed.row;
  ref.col = parsed.col;
  return ctx.resolve_ref(ref, arena, registry);
}

Value eval_offset_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  refs_internal::OffsetBase base{};
  std::uint32_t top_row = 0;
  std::uint32_t left_col = 0;
  std::uint32_t height = 0;
  std::uint32_t width = 0;
  ErrorCode err = ErrorCode::Value;
  if (!refs_internal::compute_offset_rect(call, arena, registry, ctx, &base, &top_row, &left_col, &height, &width,
                                          &err)) {
    return Value::error(err);
  }
  // Scalar context for a multi-cell OFFSET: Excel 365 dynamic-array
  // semantics spill the rectangle, and a reader that samples only the
  // anchor cell (as xlwings does in the oracle pipeline) sees the
  // top-left value. Aggregators hit a different path (they expand the
  // rectangle via `expand_offset_call` wired into `resolve_range_arg`),
  // so returning the top-left here only matters for direct scalar
  // consumption, where it reproduces Mac Excel 365's observable output.
  parser::Reference target{};
  target.sheet = base.sheet;
  target.row = top_row;
  target.col = left_col;
  return ctx.resolve_ref(target, arena, registry);
}

bool expand_offset_call(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                        std::uint32_t* out_rows, std::uint32_t* out_cols) {
  refs_internal::OffsetBase base{};
  std::uint32_t top_row = 0;
  std::uint32_t left_col = 0;
  std::uint32_t height = 0;
  std::uint32_t width = 0;
  ErrorCode err = ErrorCode::Value;
  if (!refs_internal::compute_offset_rect(call, arena, registry, ctx, &base, &top_row, &left_col, &height, &width,
                                          &err)) {
    *out_err_code = err;
    return false;
  }
  // Build two synthetic endpoint references delimiting the rectangle and
  // hand them off to `EvalContext::expand_range`, which already handles
  // cross-sheet routing, cycle detection, and per-cell recursion.
  parser::Reference lhs{};
  parser::Reference rhs{};
  lhs.sheet = base.sheet;
  lhs.row = top_row;
  lhs.col = left_col;
  rhs.sheet = base.sheet;
  rhs.row = top_row + height - 1U;
  rhs.col = left_col + width - 1U;
  auto expanded = ctx.expand_range(lhs, rhs, arena, registry);
  if (!expanded) {
    *out_err_code = expanded.error();
    return false;
  }
  *out_cells = std::move(expanded.value());
  if (out_rows != nullptr) {
    *out_rows = height;
  }
  if (out_cols != nullptr) {
    *out_cols = width;
  }
  return true;
}

bool resolve_reference_call(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx, std::string_view* out_sheet, std::uint32_t* out_top_row,
                            std::uint32_t* out_left_col, std::uint32_t* out_bottom_row, std::uint32_t* out_right_col,
                            bool* out_is_range, ErrorCode* out_err) {
  // Callers are expected to handle Ref / RangeOp / ArrayLiteral before
  // falling through here; anything that isn't a call is out of scope.
  if (node.kind() != parser::NodeKind::Call) {
    *out_err = ErrorCode::Value;
    return false;
  }
  const std::string_view name = node.as_call_name();
  if (strings::case_insensitive_eq(name, "INDIRECT")) {
    const std::uint32_t arity = node.as_call_arity();
    if (arity < 1U || arity > 2U) {
      *out_err = ErrorCode::Value;
      return false;
    }
    // Evaluate `ref_text` first so errors in the subtree propagate.
    const Value ref_val = eval_node(node.as_call_arg(0), arena, registry, ctx);
    if (ref_val.is_error()) {
      *out_err = ref_val.as_error();
      return false;
    }
    auto text_exp = coerce_to_text(ref_val);
    if (!text_exp) {
      *out_err = text_exp.error();
      return false;
    }
    bool a1_style = true;
    if (arity == 2U) {
      const Value a1_val = eval_node(node.as_call_arg(1), arena, registry, ctx);
      if (a1_val.is_error()) {
        *out_err = a1_val.as_error();
        return false;
      }
      auto b = coerce_to_bool(a1_val);
      if (!b) {
        *out_err = b.error();
        return false;
      }
      a1_style = b.value();
    }
    if (!a1_style) {
      // R1C1 style not yet supported by the A1 parser.
      *out_err = ErrorCode::Ref;
      return false;
    }
    const std::string& src = text_exp.value();
    if (src.empty()) {
      *out_err = ErrorCode::Ref;
      return false;
    }
    refs_internal::A1Parse parsed = refs_internal::parse_a1_ref(src);
    if (!parsed.valid) {
      *out_err = ErrorCode::Ref;
      return false;
    }
    // A sheet qualifier is only valid if the workbook actually holds a
    // matching sheet. Without this check a caller like `ROW(INDIRECT(
    // "NonExistent!A1"))` would happily report row 1 for a sheet that
    // doesn't exist; Excel surfaces `#REF!` in that case.
    if (!parsed.sheet.empty()) {
      const Workbook* wb = ctx.workbook();
      if (wb == nullptr || wb->sheet_by_name(parsed.sheet) == nullptr) {
        *out_err = ErrorCode::Ref;
        return false;
      }
    }
    // Intern the sheet qualifier: `parse_a1_ref` parks quoted sheet
    // names in thread-local scratch that dies on the next call.
    if (!parsed.sheet.empty()) {
      *out_sheet = arena.intern(parsed.sheet);
    } else {
      *out_sheet = std::string_view{};
    }
    if (parsed.is_range) {
      const std::uint32_t r_lo = std::min(parsed.row, parsed.row2);
      const std::uint32_t r_hi = std::max(parsed.row, parsed.row2);
      const std::uint32_t c_lo = std::min(parsed.col, parsed.col2);
      const std::uint32_t c_hi = std::max(parsed.col, parsed.col2);
      *out_top_row = r_lo;
      *out_left_col = c_lo;
      *out_bottom_row = r_hi;
      *out_right_col = c_hi;
      *out_is_range = (r_lo != r_hi) || (c_lo != c_hi);
    } else {
      *out_top_row = parsed.row;
      *out_left_col = parsed.col;
      *out_bottom_row = parsed.row;
      *out_right_col = parsed.col;
      *out_is_range = false;
    }
    return true;
  }
  if (strings::case_insensitive_eq(name, "OFFSET")) {
    refs_internal::OffsetBase base{};
    std::uint32_t top_row = 0;
    std::uint32_t left_col = 0;
    std::uint32_t height = 0;
    std::uint32_t width = 0;
    ErrorCode err = ErrorCode::Value;
    if (!refs_internal::compute_offset_rect(node, arena, registry, ctx, &base, &top_row, &left_col, &height, &width,
                                            &err)) {
      *out_err = err;
      return false;
    }
    *out_sheet = base.sheet;
    *out_top_row = top_row;
    *out_left_col = left_col;
    *out_bottom_row = top_row + height - 1U;
    *out_right_col = left_col + width - 1U;
    *out_is_range = (height > 1U) || (width > 1U);
    return true;
  }
  if (strings::case_insensitive_eq(name, "CHOOSE")) {
    const std::uint32_t arity = node.as_call_arity();
    if (arity < 2U) {
      *out_err = ErrorCode::Value;
      return false;
    }
    // Evaluate the index argument; CHOOSE expects a 1-based integer
    // selector. Anything that fails coercion (text, blank-as-strict,
    // error) propagates with its original code.
    const Value idx_val = eval_node(node.as_call_arg(0), arena, registry, ctx);
    if (idx_val.is_error()) {
      *out_err = idx_val.as_error();
      return false;
    }
    auto idx_int = refs_internal::read_int(idx_val);
    if (!idx_int) {
      *out_err = idx_int.error();
      return false;
    }
    const int idx = idx_int.value();
    const std::uint32_t n_choices = arity - 1U;
    if (idx < 1 || static_cast<std::uint32_t>(idx) > n_choices) {
      // Excel: out-of-range index -> #VALUE!
      *out_err = ErrorCode::Value;
      return false;
    }
    // The picked choice is at slot `idx` (0 = index, 1..n = choices).
    // Recurse via `resolve_range_endpoint` so plain Ref / nested
    // OFFSET-INDIRECT-CHOOSE endpoints all reduce to a rectangle.
    const parser::AstNode& picked = node.as_call_arg(static_cast<std::uint32_t>(idx));
    if (!resolve_range_endpoint(picked, arena, registry, ctx, out_sheet, out_top_row, out_left_col, out_bottom_row,
                                out_right_col, out_err)) {
      return false;
    }
    *out_is_range = (*out_top_row != *out_bottom_row) || (*out_left_col != *out_right_col);
    return true;
  }
  // Any other call name is not a reference-returning builtin we know
  // how to handle here.
  *out_err = ErrorCode::Value;
  return false;
}

bool resolve_range_endpoint(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx, std::string_view* out_sheet, std::uint32_t* out_top_row,
                            std::uint32_t* out_left_col, std::uint32_t* out_bottom_row, std::uint32_t* out_right_col,
                            ErrorCode* out_err) {
  if (node.kind() == parser::NodeKind::Ref) {
    const parser::Reference& r = node.as_ref();
    // Whole-column / whole-row refs cannot anchor an endpoint composition
    // because the union rectangle would be unbounded; surface #VALUE! to
    // match `expand_range`'s existing degradation for these shapes.
    if (r.is_full_col || r.is_full_row) {
      *out_err = ErrorCode::Value;
      return false;
    }
    *out_sheet = r.sheet;
    *out_top_row = r.row;
    *out_left_col = r.col;
    *out_bottom_row = r.row;
    *out_right_col = r.col;
    return true;
  }
  if (node.kind() == parser::NodeKind::Call) {
    bool is_range_unused = false;
    return resolve_reference_call(node, arena, registry, ctx, out_sheet, out_top_row, out_left_col, out_bottom_row,
                                  out_right_col, &is_range_unused, out_err);
  }
  // Anything else (NameRef, ExternalRef, BinaryOp, ArrayLiteral, etc.) is
  // not a recognized range endpoint shape.
  *out_err = ErrorCode::Ref;
  return false;
}

}  // namespace eval
}  // namespace formulon
