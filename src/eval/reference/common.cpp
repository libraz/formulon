//
// Shared bodies for the reference-family lazy impls. Hosts:
//
//   * The A1 textual reference parser (`parse_a1_ref`, `column_letters`,
//     `A1Parse`) — public surface declared in `eval/a1_parse.h` and used
//     by INDIRECT (this subdirectory), INFO ("address"), CELL, and the
//     `ADDRESS` builtin.
//   * Small helpers `apply_offset` / `read_int` that the OFFSET and
//     CHOOSE expanders share with the intersection resolver.
//   * The rectangle-construction routines `resolve_indirect_reference`,
//     `resolve_offset_base`, and `compute_offset_rect`. These touch both
//     the INDIRECT and OFFSET pipelines (and are reached from the
//     intersection resolver too), so they live here to avoid a circular
//     dependency between the sibling TUs.

#include "eval/reference/common.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>

#include "eval/a1_parse.h"
#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/range_resolvers.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "sheet.h"
#include "sheet_name.h"
#include "utils/a1_column.h"
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
  if (col == 0U || out == nullptr) {
    return 0U;
  }
  std::string letters;
  if (!a1::append_column_letters(letters, col - 1U)) {
    return 0U;
  }
  std::memcpy(out, letters.data(), letters.size());
  return letters.size();
}

// Outcome of reading an optional leading `Sheet!` qualifier.
enum class SheetQualifier : std::uint8_t {
  kNone,       ///< No qualifier; `*i` is unmoved and the whole text is the reference.
  kParsed,     ///< Qualifier consumed; `*sheet` is set and `*i` points past the `!`.
  kMalformed,  ///< Unterminated quote or a quoted name with no `!`; the whole reference is invalid.
};

// Reads the optional sheet qualifier at the head of `text`, advancing `*i`
// past it. Shared by the A1 and R1C1 parsers, which differ only in how
// they read the reference body that follows.
SheetQualifier parse_sheet_qualifier(std::string_view text, std::size_t* i, std::string_view* sheet) {
  if (text.empty()) {
    return SheetQualifier::kNone;
  }
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
      return SheetQualifier::kMalformed;  // unterminated
    }
    // Inside content is `text.substr(1, j - 1)`; unescape `''` -> `'`.
    // We don't own backing storage here, so write into a static-sized
    // local buffer; Excel sheet-name limit is 31 chars (we allow up to
    // 255 defensively, capped by the source view size).
    static thread_local char scratch[256];
    const std::size_t content_len = j - 1;
    const std::size_t used = unescape_quoted_sheet(text.substr(1, content_len), scratch, sizeof(scratch));
    // This view points into thread_local storage — callers keep it alive
    // only until the next parse on the same thread. The consuming code
    // (`indirect`) copies before storing.
    *sheet = std::string_view(scratch, used);
    std::size_t after = j + 1;
    if (after >= text.size() || text[after] != '!') {
      return SheetQualifier::kMalformed;  // missing `!` after quoted sheet
    }
    *i = after + 1;
    return SheetQualifier::kParsed;
  }
  // Bare sheet name: run of letters/digits/underscore followed by `!`.
  // We only commit to treating it as a sheet qualifier if we find the
  // `!` — otherwise the run is part of the reference itself.
  std::size_t j = 0;
  while (j < text.size() && (is_letter(text[j]) || is_digit(text[j]) || text[j] == '_')) {
    ++j;
  }
  if (j > 0 && j < text.size() && text[j] == '!') {
    *sheet = text.substr(0, j);
    *i = j + 1;
    return SheetQualifier::kParsed;
  }
  return SheetQualifier::kNone;
}

A1Parse parse_a1_ref(std::string_view text) {
  A1Parse out;
  if (text.empty()) {
    return out;
  }
  std::size_t i = 0;
  if (parse_sheet_qualifier(text, &i, &out.sheet) == SheetQualifier::kMalformed) {
    return out;
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

namespace {

// One axis of an R1C1 endpoint: `R`/`C` on its own, `R5` (absolute), or
// `R[-2]` (relative). `present` distinguishes an axis that was written
// from one that was left out, which is what makes `R5` a whole row and
// `R5C2` a single cell.
struct R1C1Axis {
  bool present = false;
  bool relative = false;
  long long value = 0;  ///< 1-based index when absolute, signed offset when relative.
};

// Reads one `R`/`C` axis at `text[*i]` when the marker matches `marker`.
// Returns false only on malformed input; an absent axis leaves `*out`
// unset and still returns true.
bool parse_r1c1_axis(std::string_view text, std::size_t* i, char marker, R1C1Axis* out) {
  if (*i >= text.size()) {
    return true;
  }
  char head = text[*i];
  if (head >= 'a' && head <= 'z') {
    head = static_cast<char>(head - ('a' - 'A'));
  }
  if (head != marker) {
    return true;
  }
  ++*i;
  out->present = true;
  if (*i < text.size() && text[*i] == '[') {
    ++*i;
    bool negative = false;
    if (*i < text.size() && (text[*i] == '-' || text[*i] == '+')) {
      negative = text[*i] == '-';
      ++*i;
    }
    const std::size_t digits_start = *i;
    long long magnitude = 0;
    while (*i < text.size() && is_digit(text[*i])) {
      magnitude = magnitude * 10 + (text[*i] - '0');
      if (magnitude > Sheet::kMaxRows) {
        return false;  // saturate well before overflow; out of grid either way
      }
      ++*i;
    }
    if (*i == digits_start || *i >= text.size() || text[*i] != ']') {
      return false;
    }
    ++*i;
    out->relative = true;
    out->value = negative ? -magnitude : magnitude;
    return true;
  }
  // A bare `R` / `C` is the current row / column: a zero relative offset.
  const std::size_t digits_start = *i;
  long long absolute = 0;
  while (*i < text.size() && is_digit(text[*i])) {
    absolute = absolute * 10 + (text[*i] - '0');
    if (absolute > Sheet::kMaxRows) {
      return false;
    }
    ++*i;
  }
  if (*i == digits_start) {
    out->relative = true;
    out->value = 0;
    return true;
  }
  out->value = absolute;
  return true;
}

// Resolves one axis to a 0-based coordinate against `base`, rejecting
// anything outside `[0, max)`. A relative axis with no base has nothing
// to measure from and fails rather than assuming the origin.
bool resolve_r1c1_axis(const R1C1Axis& axis, bool base_present, std::uint32_t base, std::uint32_t max,
                       std::uint32_t* out) {
  if (axis.relative && !base_present) {
    return false;
  }
  const long long resolved = axis.relative ? static_cast<long long>(base) + axis.value : axis.value - 1;
  if (resolved < 0 || resolved >= static_cast<long long>(max)) {
    return false;
  }
  *out = static_cast<std::uint32_t>(resolved);
  return true;
}

// Parses one `R...C...` endpoint. At least one axis must be written.
bool parse_r1c1_endpoint(std::string_view text, std::size_t* i, R1C1Axis* row, R1C1Axis* col) {
  if (!parse_r1c1_axis(text, i, 'R', row)) {
    return false;
  }
  if (!parse_r1c1_axis(text, i, 'C', col)) {
    return false;
  }
  return row->present || col->present;
}

}  // namespace

A1Parse parse_r1c1_ref(std::string_view text, const R1C1Base& base) {
  A1Parse out;
  if (text.empty()) {
    return out;
  }
  std::size_t i = 0;
  if (parse_sheet_qualifier(text, &i, &out.sheet) == SheetQualifier::kMalformed) {
    return out;
  }

  R1C1Axis row1;
  R1C1Axis col1;
  if (!parse_r1c1_endpoint(text, &i, &row1, &col1)) {
    return out;
  }
  R1C1Axis row2 = row1;
  R1C1Axis col2 = col1;
  bool is_range = false;
  if (i < text.size() && text[i] == ':') {
    ++i;
    row2 = R1C1Axis{};
    col2 = R1C1Axis{};
    if (!parse_r1c1_endpoint(text, &i, &row2, &col2)) {
      return out;
    }
    // A range whose endpoints name different axes (`R2:R3C4`) has no
    // rectangle; Excel rejects it rather than guessing the missing bound.
    if (row1.present != row2.present || col1.present != col2.present) {
      return out;
    }
    is_range = true;
  }
  if (i != text.size()) {
    return out;  // trailing garbage
  }

  // An endpoint that names only one axis is unbounded along the other, the
  // same shape `5:5` and `D:D` take on the A1 side.
  if (!col1.present) {
    if (!resolve_r1c1_axis(row1, base.present, base.row, Sheet::kMaxRows, &out.row) ||
        !resolve_r1c1_axis(row2, base.present, base.row, Sheet::kMaxRows, &out.row2)) {
      return out;
    }
    out.col = 0;
    out.col2 = Sheet::kMaxCols - 1U;
    out.is_full_row = true;
    out.is_range = true;
    out.valid = true;
    return out;
  }
  if (!row1.present) {
    if (!resolve_r1c1_axis(col1, base.present, base.col, Sheet::kMaxCols, &out.col) ||
        !resolve_r1c1_axis(col2, base.present, base.col, Sheet::kMaxCols, &out.col2)) {
      return out;
    }
    out.row = 0;
    out.row2 = Sheet::kMaxRows - 1U;
    out.is_full_col = true;
    out.is_range = true;
    out.valid = true;
    return out;
  }

  if (!resolve_r1c1_axis(row1, base.present, base.row, Sheet::kMaxRows, &out.row) ||
      !resolve_r1c1_axis(col1, base.present, base.col, Sheet::kMaxCols, &out.col) ||
      !resolve_r1c1_axis(row2, base.present, base.row, Sheet::kMaxRows, &out.row2) ||
      !resolve_r1c1_axis(col2, base.present, base.col, Sheet::kMaxCols, &out.col2)) {
    return out;
  }
  out.is_range = is_range;
  out.valid = true;
  return out;
}

bool apply_offset(std::uint32_t base, int offset, std::uint32_t max, std::uint32_t* out) {
  const long long sum = static_cast<long long>(base) + static_cast<long long>(offset);
  if (sum < 0 || sum >= static_cast<long long>(max)) {
    return false;
  }
  *out = static_cast<std::uint32_t>(sum);
  return true;
}

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

bool resolve_indirect_reference(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                const EvalContext& ctx, IndirectReference* out, ErrorCode* out_err) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 2U) {
    *out_err = ErrorCode::Value;
    return false;
  }

  // Evaluate `ref_text` first so errors propagate per the dispatcher's
  // left-most-wins rule.
  const Value ref_val = eval_node(call.as_call_arg(0), arena, registry, ctx);
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
    const Value a1_val = eval_node(call.as_call_arg(1), arena, registry, ctx);
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
  const std::string& src = text_exp.value();
  if (src.empty()) {
    *out_err = ErrorCode::Ref;
    return false;
  }
  // The flag picks the grammar outright: A1 text under `a1 = FALSE` is as
  // invalid as R1C1 text under `a1 = TRUE`, and each parser rejects the
  // other's spelling on its own. Relative axes resolve against the cell
  // the formula sits in, which is the only reading `R[1]C[1]` has.
  R1C1Base base;
  if (ctx.has_formula_cell()) {
    base.present = true;
    base.row = ctx.formula_row();
    base.col = ctx.formula_col();
  }
  const A1Parse parsed = a1_style ? parse_a1_ref(src) : parse_r1c1_ref(src, base);
  if (!parsed.valid) {
    *out_err = ErrorCode::Ref;
    return false;
  }

  out->sheet = parsed.sheet.empty() ? std::string_view{} : arena.intern(parsed.sheet);
  out->range_syntax = parsed.is_range;
  out->is_full_col = parsed.is_full_col;
  out->is_full_row = parsed.is_full_row;
  if (parsed.is_range) {
    out->top_row = std::min(parsed.row, parsed.row2);
    out->left_col = std::min(parsed.col, parsed.col2);
    out->bottom_row = std::max(parsed.row, parsed.row2);
    out->right_col = std::max(parsed.col, parsed.col2);
    out->is_range = (out->top_row != out->bottom_row) || (out->left_col != out->right_col);
  } else {
    out->top_row = parsed.row;
    out->left_col = parsed.col;
    out->bottom_row = parsed.row;
    out->right_col = parsed.col;
    out->is_range = false;
  }
  return true;
}

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
      if (!sheet_names::equal(lhs.sheet, rhs.sheet)) {
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
    if (!resolve_reference_call(arg, arena, registry, ctx, &sheet, &top, &left, &bottom, &right, &is_range, &err)) {
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
}  // namespace eval
}  // namespace formulon
