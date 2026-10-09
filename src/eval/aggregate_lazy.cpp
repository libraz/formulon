
#include "eval/aggregate_lazy.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <iterator>
#include <string_view>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "eval/array_alloc.h"
#include "eval/builtin_names.h"
#include "eval/builtins/numeric_helpers.h"
#include "eval/builtins/subtotal.h"
#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/formula_text_utils.h"
#include "eval/lazy_impls.h"
#include "eval/name_env_resolve.h"
#include "eval/range_args.h"
#include "eval/tree_walker/dispatch.h"
#include "numeric_aggregate_kernels.h"
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

using builtins_detail::to_finite_value;

// 1..13 are the SUBTOTAL-aligned modes; 14..19 are the AGGREGATE-only "k-
// arg" modes. Storing the integer code rather than an enum keeps the
// k-arg validation switch easy to read against the Excel docs.
constexpr int kCodeAverage = 1;
constexpr int kCodeCount = 2;
constexpr int kCodeCountA = 3;
constexpr int kCodeMax = 4;
constexpr int kCodeMin = 5;
constexpr int kCodeProduct = 6;
constexpr int kCodeStdevS = 7;
constexpr int kCodeStdevP = 8;
constexpr int kCodeSum = 9;
constexpr int kCodeVarS = 10;
constexpr int kCodeVarP = 11;
constexpr int kCodeMedian = 12;
constexpr int kCodeModeSngl = 13;
constexpr int kCodeLarge = 14;
constexpr int kCodeSmall = 15;
constexpr int kCodePercentileInc = 16;
constexpr int kCodeQuartileInc = 17;
constexpr int kCodePercentileExc = 18;
constexpr int kCodeQuartileExc = 19;

constexpr int kFnMin = 1;
constexpr int kFnMax = 19;
constexpr int kFnKArgFirst = 14;  // 14..19 take a trailing k arg.

// Reads a required scalar metadata argument (function_num / options /
// k). Errors propagate verbatim; non-coercible values surface
// `#VALUE!`; non-finite results (e.g. coercion overflow) surface
// `#NUM!`. The Expected return type avoids the previous in-band
// `0.0`-on-error sentinel, which collided with legitimate zero
// arguments (e.g. `AGGREGATE(2, 0, range)` -> COUNT mode + clear-flags
// option) and risked silent-wrong-result bugs.
Expected<double, ErrorCode> scalar_number(const Value& v) {
  if (v.is_error()) {
    return v.as_error();
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    return coerced.error();
  }
  const double x = coerced.value();
  if (!std::isfinite(x)) {
    return ErrorCode::Num;
  }
  return x;
}

Expected<double, ErrorCode> read_scalar(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                                        const EvalContext& ctx) {
  return scalar_number(eval_node(node, arena, registry, ctx));
}

// --- Row visibility ------------------------------------------------------
//
// `SUBTOTAL(100+n)` and `AGGREGATE` with the hidden-row option bit skip cells
// that sit on any hidden row. Visibility is a property of the sheet
// (`Sheet::layout().row_overrides`), so it is only knowable for an argument
// that still carries the rows its cells came from.
//
// A plain `Ref` and a `Ref:Ref` RangeOp do carry that provenance. An inline
// array literal, a computed array (`SORT(...)`), a scalar and a spilled-range
// reference do not: their cells have no sheet row behind them, or the anchor
// is not resolved here. Those arguments contribute every cell, which is both
// what Excel does for a literal array and the conservative answer elsewhere —
// a cell whose row we cannot name is never silently dropped from a total.
//
// `SUBTOTAL(1..11)` skips only filter-hidden rows. OOXML stores no filtered
// flag, so the class is derived the way Excel does after a reload: while any
// AutoFilter on the sheet (its own or a table's) carries a criterion, every
// hidden row of that sheet counts as filtered; otherwise none does.

/// Which hidden rows an aggregate skips.
enum class HiddenScope : std::uint8_t {
  kAll,       ///< Every hidden row (`SUBTOTAL(101..111)`, AGGREGATE's hidden option).
  kFiltered,  ///< Only filter-hidden rows (`SUBTOTAL(1..11)`).
};

// Resolves the sheet a reference-shaped argument reads from, together with
// the 0-based row its first (top-left) cell occupies. Returns nullptr when
// the argument carries no row provenance.
//
// The top row mirrors `EvalContext::expand_range`: a whole-column reference
// keeps its natural origin at row 0, every other shape starts at the
// rectangle's smallest row index.
const Sheet* reference_arg_origin(const parser::AstNode& node, const EvalContext& ctx, std::uint32_t* out_top_row) {
  const parser::Reference* first = nullptr;
  const parser::Reference* second = nullptr;
  if (node.kind() == parser::NodeKind::Ref) {
    first = &node.as_ref();
  } else if (node.kind() == parser::NodeKind::RangeOp) {
    const parser::AstNode& lhs = node.as_range_lhs();
    const parser::AstNode& rhs = node.as_range_rhs();
    if (lhs.kind() != parser::NodeKind::Ref || rhs.kind() != parser::NodeKind::Ref) {
      return nullptr;  // OFFSET / INDIRECT endpoints: no static provenance.
    }
    first = &lhs.as_ref();
    second = &rhs.as_ref();
  } else {
    return nullptr;
  }

  const bool whole_column = first->is_full_col || (second != nullptr && second->is_full_col);
  std::uint32_t top = whole_column ? 0U : first->row;
  if (!whole_column && second != nullptr) {
    top = std::min(first->row, second->row);
  }

  // The parser keeps a `:` operator's qualifier on the left endpoint, so
  // `Sheet2!A1:B2` arrives as RangeOp(Ref{sheet=Sheet2}, Ref{sheet=""}) and
  // the right one inherits; `expand_range` also accepts the mirrored shape,
  // so read whichever endpoint carries a name. A pair that names two
  // different sheets is a `#REF!` the range resolver already rejected.
  std::string_view sheet_name = first->sheet;
  if (sheet_name.empty() && second != nullptr) {
    sheet_name = second->sheet;
  }
  const Sheet* sheet = ctx.sheet_for_qualifier(sheet_name);
  if (sheet == nullptr) {
    return nullptr;
  }
  *out_top_row = top;
  return sheet;
}

// Appends `count` visibility flags for the cells of one argument, given the
// argument's resolved shape. Cells on a hidden row within `scope` are marked
// true.
void append_visibility(const parser::AstNode& node, const EvalContext& ctx, HiddenScope scope, std::uint32_t rows,
                       std::uint32_t cols, std::vector<bool>* out_hidden) {
  const std::size_t count = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  std::uint32_t top = 0;
  const Sheet* sheet = reference_arg_origin(node, ctx, &top);
  if (sheet == nullptr || rows == 0U || cols == 0U ||
      (scope == HiddenScope::kFiltered && !sheet_has_filter_criteria(ctx.workbook(), *sheet))) {
    out_hidden->resize(out_hidden->size() + count, false);
    return;
  }
  // One pass over the sheet's overrides rather than a lookup per row: the
  // override list holds only rows that differ from the sheet default, so it
  // is short even on a large range.
  std::vector<bool> hidden_row(rows, false);
  for (const RowLayout& row : sheet->layout().row_overrides) {
    if (!row.hidden || row.row < top) {
      continue;
    }
    const std::uint32_t offset = row.row - top;
    if (offset < rows) {
      hidden_row[offset] = true;
    }
  }
  out_hidden->reserve(out_hidden->size() + count);
  for (std::uint32_t r = 0; r < rows; ++r) {
    out_hidden->resize(out_hidden->size() + cols, hidden_row[r]);
  }
}

// --- Nested SUBTOTAL / AGGREGATE exclusion --------------------------------
//
// SUBTOTAL (every function number) and AGGREGATE options 0..3 skip a cell
// whose own formula contains a SUBTOTAL or AGGREGATE call anywhere in its
// AST: inside LET / LAMBDA bodies, in an unevaluated IF branch, in any case.
// A defined name, a string literal, INDIRECT or a cell that merely references
// a subtotal cell does not count. Every cell of a spill whose anchor
// qualifies is skipped too.
//
// The check is two-staged. Formula text is a cheap case-insensitive substring
// candidate test; only candidates are parsed and walked. It runs once per
// formula cell inside a reference argument, so a range without subtotal text
// costs one substring scan per formula cell and no parsing.

bool is_nested_call_name(std::string_view name) noexcept {
  return strings::case_insensitive_eq(name, "SUBTOTAL") || strings::case_insensitive_eq(name, "AGGREGATE");
}

bool ast_has_nested_call(const parser::AstNode& node) {
  using parser::NodeKind;
  switch (node.kind()) {
    case NodeKind::SpillRef: {
      const parser::AstNode* anchor = node.as_spill_ref_anchor_expr();
      return anchor != nullptr && ast_has_nested_call(*anchor);
    }
    case NodeKind::UnaryOp:
      return ast_has_nested_call(node.as_unary_operand());
    case NodeKind::ImplicitIntersection:
      return ast_has_nested_call(node.as_implicit_intersection_operand());
    case NodeKind::BinaryOp:
      return ast_has_nested_call(node.as_binary_lhs()) || ast_has_nested_call(node.as_binary_rhs());
    case NodeKind::RangeOp:
      return ast_has_nested_call(node.as_range_lhs()) || ast_has_nested_call(node.as_range_rhs());
    case NodeKind::IntersectOp:
      return ast_has_nested_call(node.as_intersect_lhs()) || ast_has_nested_call(node.as_intersect_rhs());
    case NodeKind::UnionOp:
      for (std::uint32_t i = 0; i < node.as_union_arity(); ++i) {
        if (ast_has_nested_call(node.as_union_child(i))) {
          return true;
        }
      }
      return false;
    case NodeKind::Call:
      if (is_nested_call_name(node.as_call_name())) {
        return true;
      }
      for (std::uint32_t i = 0; i < node.as_call_arity(); ++i) {
        if (ast_has_nested_call(node.as_call_arg(i))) {
          return true;
        }
      }
      return false;
    case NodeKind::ArrayLiteral:
      for (std::uint32_t r = 0; r < node.as_array_rows(); ++r) {
        for (std::uint32_t c = 0; c < node.as_array_cols(); ++c) {
          if (ast_has_nested_call(node.as_array_element(r, c))) {
            return true;
          }
        }
      }
      return false;
    case NodeKind::Lambda:
      return ast_has_nested_call(node.as_lambda_body());
    case NodeKind::LetBinding:
      for (std::uint32_t i = 0; i < node.as_let_binding_count(); ++i) {
        if (ast_has_nested_call(node.as_let_binding_expr(i))) {
          return true;
        }
      }
      return ast_has_nested_call(node.as_let_body());
    case NodeKind::LambdaCall:
      if (ast_has_nested_call(node.as_lambda_call_callee())) {
        return true;
      }
      for (std::uint32_t i = 0; i < node.as_lambda_call_arity(); ++i) {
        if (ast_has_nested_call(node.as_lambda_call_arg(i))) {
          return true;
        }
      }
      return false;
    default:
      return false;  // Leaves: literals, references, names, errors.
  }
}

bool text_may_hold_nested_call(std::string_view text) noexcept {
  static constexpr std::string_view kNeedles[] = {"SUBTOTAL", "AGGREGATE"};
  for (const std::string_view needle : kNeedles) {
    for (std::size_t i = 0; i + needle.size() <= text.size(); ++i) {
      if (strings::case_insensitive_eq(text.substr(i, needle.size()), needle)) {
        return true;
      }
    }
  }
  return false;
}

// True when the formula cell at (row, col) of `sheet` calls SUBTOTAL or
// AGGREGATE anywhere in its AST.
bool cell_formula_has_nested_call(const Sheet& sheet, std::uint32_t row, std::uint32_t col) {
  Sheet::CellRead read;
  sheet.read_formula_cell(row, col, read);
  if (!read.exists() || !read.is_formula() || !text_may_hold_nested_call(read.formula_text())) {
    return false;
  }
  Arena arena;
  const parser::AstNode* root = parse_formula_entry(strip_formula_prefix(read.formula_text()), arena);
  return root != nullptr && ast_has_nested_call(*root);
}

// Appends one flag per cell of the reference-shaped argument `node`, true for
// cells that must be excluded as nested subtotals. Arguments without sheet
// provenance (literals, computed arrays, `A1#`) are never flagged.
void append_nested_flags(const parser::AstNode& node, const EvalContext& ctx, std::uint32_t rows, std::uint32_t cols,
                         std::vector<bool>* out_nested) {
  const std::size_t count = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  const std::size_t base = out_nested->size();
  out_nested->resize(base + count, false);

  const parser::Reference* first = nullptr;
  const parser::Reference* second = nullptr;
  if (node.kind() == parser::NodeKind::Ref) {
    first = &node.as_ref();
    second = first;
  } else if (node.kind() == parser::NodeKind::RangeOp && node.as_range_lhs().kind() == parser::NodeKind::Ref &&
             node.as_range_rhs().kind() == parser::NodeKind::Ref) {
    first = &node.as_range_lhs().as_ref();
    second = &node.as_range_rhs().as_ref();
  }
  if (first == nullptr || count == 0U) {
    return;
  }
  const Sheet* sheet = nullptr;
  const auto rect_or = ctx.walked_range_rect(*first, *second, &sheet);
  if (!rect_or || !rect_or.value().has_value() || sheet == nullptr) {
    return;
  }
  const DeclaredRect& rect = *rect_or.value();
  if (rect.rows() != rows || rect.cols() != cols) {
    return;  // Shape disagrees with the resolver: flag nothing.
  }
  const auto mark = [&](std::uint32_t row, std::uint32_t col) {
    (*out_nested)[base + static_cast<std::size_t>(row - rect.row_first) * cols + (col - rect.col_first)] = true;
  };

  for (const CellAddress addr : sheet->formula_cells_in(rect.row_first, rect.col_first, rect.row_last, rect.col_last)) {
    if (cell_formula_has_nested_call(*sheet, addr.row, addr.col)) {
      mark(addr.row, addr.col);
    }
  }
  // A spill's phantom cells carry no formula of their own; the anchor's
  // formula decides for the whole footprint, even when the anchor lies
  // outside the rectangle.
  for (const SpillFootprint& fp : sheet->committed_spill_footprints()) {
    if (fp.rows == 0U || fp.cols == 0U || fp.anchor_row > rect.row_last || fp.anchor_col > rect.col_last ||
        fp.anchor_row + fp.rows - 1U < rect.row_first || fp.anchor_col + fp.cols - 1U < rect.col_first) {
      continue;
    }
    if (!cell_formula_has_nested_call(*sheet, fp.anchor_row, fp.anchor_col)) {
      continue;
    }
    const std::uint32_t r0 = std::max(fp.anchor_row, rect.row_first);
    const std::uint32_t r1 = std::min(fp.anchor_row + fp.rows - 1U, rect.row_last);
    const std::uint32_t c0 = std::max(fp.anchor_col, rect.col_first);
    const std::uint32_t c1 = std::min(fp.anchor_col + fp.cols - 1U, rect.col_last);
    for (std::uint32_t r = r0; r <= r1; ++r) {
      for (std::uint32_t c = c0; c <= c1; ++c) {
        mark(r, c);
      }
    }
  }
}

// Appends every scalar Value produced by `arg_node` to `out_cells`, mirroring
// PERCENTOF's `sum_arg_for_percentof` provenance walk. LET-bound NameRefs
// resolve to their bound AST when range-shaped. On any expansion failure
// (e.g. `#REF!` from a missing sheet) returns false with the propagating
// error in `*out_err`. Returns true on a clean walk.
//
// `out_hidden` grows in lockstep with `out_cells`, one flag per appended
// cell (hidden rows within `scope`), so a later filter can drop them without
// re-deriving where each one came from. `out_nested` (nullable) grows the same
// way with the nested SUBTOTAL / AGGREGATE flags; null skips the detection.
//
// Unlike PERCENTOF this helper does NOT filter by Value kind: AGGREGATE's
// per-mode rules (numeric branches drop non-numerics; COUNTA counts them;
// the error-ignore bit decides whether errors short-circuit) are applied
// later by `apply_filters`.
bool collect_arg(const parser::AstNode& arg_node, Arena& arena, const FunctionRegistry& registry,
                 const EvalContext& ctx, HiddenScope scope, std::vector<Value>* out_cells,
                 std::vector<bool>* out_hidden, std::vector<bool>* out_nested, Value* out_err) {
  const parser::AstNode& node = resolve_range_binding(arg_node, ctx.name_env(), /*accept_ref=*/false);
  const parser::NodeKind k = node.kind();

  if (k == parser::NodeKind::UnionOp) {
    for (std::uint32_t i = 0; i < node.as_union_arity(); ++i) {
      if (!collect_arg(node.as_union_child(i), arena, registry, ctx, scope, out_cells, out_hidden, out_nested,
                       out_err)) {
        return false;
      }
    }
    return true;
  }

  // Range / Ref / SpillRef / RangeOp -> use the canonical resolver.
  if (k == parser::NodeKind::Ref || k == parser::NodeKind::RangeOp || k == parser::NodeKind::SpillRef) {
    auto resolved = resolve_range_arg(node, arena, registry, ctx);
    if (!resolved) {
      *out_err = Value::error(resolved.error());
      return false;
    }
    auto& rr = resolved.value();
    // Trust the resolver's own shape report only when it accounts for every
    // cell; anything else means the argument was reshaped on the way out and
    // the row mapping would be a guess.
    const std::size_t n = rr.cells.size();
    if (static_cast<std::size_t>(rr.rows) * static_cast<std::size_t>(rr.cols) == n) {
      append_visibility(node, ctx, scope, rr.rows, rr.cols, out_hidden);
      if (out_nested != nullptr) {
        append_nested_flags(node, ctx, rr.rows, rr.cols, out_nested);
      }
    } else {
      out_hidden->resize(out_hidden->size() + n, false);
      if (out_nested != nullptr) {
        out_nested->resize(out_nested->size() + n, false);
      }
    }
    out_cells->insert(out_cells->end(), std::make_move_iterator(rr.cells.begin()),
                      std::make_move_iterator(rr.cells.end()));
    return true;
  }

  // Inline array literal `{a;b;c}` walked in row-major order.
  if (k == parser::NodeKind::ArrayLiteral) {
    const std::uint32_t rows = node.as_array_rows();
    const std::uint32_t cols = node.as_array_cols();
    for (std::uint32_t r = 0; r < rows; ++r) {
      for (std::uint32_t c = 0; c < cols; ++c) {
        const Value v = eval_node(node.as_array_element(r, c), arena, registry, ctx);
        out_cells->push_back(v);
        out_hidden->push_back(false);
        if (out_nested != nullptr) {
          out_nested->push_back(false);
        }
      }
    }
    return true;
  }

  // Anything else (literal scalar, arithmetic expression, function call) is
  // evaluated normally. Array-valued calls are flattened in row-major order
  // so AGGREGATE's code-3 COUNTA path sees the same marker-bearing cells as
  // the eager COUNTA dispatcher; scalar values remain one direct argument.
  const Value v = eval_node(node, arena, registry, ctx);
  if (v.is_array()) {
    const ArrayValue* array = v.as_array();
    const std::size_t n = static_cast<std::size_t>(array->rows) * static_cast<std::size_t>(array->cols);
    out_cells->insert(out_cells->end(), array->cells, array->cells + n);
    out_hidden->resize(out_hidden->size() + n, false);
    if (out_nested != nullptr) {
      out_nested->resize(out_nested->size() + n, false);
    }
    return true;
  }
  out_cells->push_back(v);
  out_hidden->push_back(false);
  if (out_nested != nullptr) {
    out_nested->push_back(false);
  }
  return true;
}

// Drops the cells whose visibility flag is set. `hidden` must be the vector
// `collect_arg` grew alongside `cells`.
void drop_hidden_cells(std::vector<Value>* cells, const std::vector<bool>& hidden) {
  if (cells->size() != hidden.size()) {
    return;  // Shapes disagree: keep every cell rather than drop a wrong one.
  }
  std::vector<Value> kept;
  kept.reserve(cells->size());
  for (std::size_t i = 0; i < cells->size(); ++i) {
    if (!hidden[i]) {
      kept.push_back((*cells)[i]);
    }
  }
  *cells = std::move(kept);
}

// Builds the drop mask for one call: the hidden flags when `use_hidden`, OR
// the nested flags when `use_nested`. Both inputs grew in lockstep with the
// cells, so the mask drops them in a single pass and no vector goes stale.
std::vector<bool> make_drop_mask(std::size_t n, bool use_hidden, const std::vector<bool>& hidden, bool use_nested,
                                 const std::vector<bool>& nested) {
  std::vector<bool> mask(n, false);
  for (std::size_t i = 0; i < n; ++i) {
    mask[i] = (use_hidden && i < hidden.size() && hidden[i]) || (use_nested && i < nested.size() && nested[i]);
  }
  return mask;
}

// Filters `cells` in place according to the options bit and the function
// code's expected provenance:
//
//   * Errors: dropped silently when ignore_errors == true; the first error
//     short-circuits the call and is written to `*out_err` otherwise.
//   * COUNTA (code 3): keep every non-blank, plus blanks owned by a derived
//     value array; raw-reference blanks are dropped.
//   * Numeric modes (everything else): keep only Numbers; drop Bool / Text /
//     Blank.
//
// Returns true on a clean filter, false (with `*out_err` populated) when an
// un-ignored error short-circuits.
bool apply_filters(std::vector<Value>* cells, int code, bool ignore_errors, Value* out_err) {
  std::vector<Value> kept;
  kept.reserve(cells->size());
  for (const Value& v : *cells) {
    if (v.is_error()) {
      if (ignore_errors) {
        continue;
      }
      *out_err = v;
      return false;
    }
    if (code == kCodeCountA) {
      if (!v.is_blank() || v.blank_counts_for_counta()) {
        kept.push_back(v);
      }
      continue;
    }
    // Numeric branches: drop everything that is not a Number.
    if (v.is_number()) {
      kept.push_back(v);
    }
  }
  *cells = std::move(kept);
  return true;
}

// Helper: extract the numeric slice once we know every kept cell is a
// Number (true for codes 1, 2, 4..19).
std::vector<double> to_numbers(const std::vector<Value>& cells) {
  std::vector<double> out;
  out.reserve(cells.size());
  for (const Value& v : cells) {
    if (v.is_number()) {
      out.push_back(v.as_number());
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Mode runners (codes 1..13). The numeric-aggregator slots (SUM / PRODUCT /
// MIN / MAX / AVERAGE / VAR.* / STDEV.*) all delegate to the shared kernels
// in `numeric_aggregate_kernels.h` so SUBTOTAL and AGGREGATE cannot drift. Empty-
// range behaviour matches Excel's convention for SUBTOTAL / AGGREGATE:
// SUM/PRODUCT/MIN/MAX -> 0, AVERAGE/VAR/STDEV -> #DIV/0!, MEDIAN -> #DIV/0!,
// COUNT/COUNTA -> 0, MODE.SNGL -> #N/A.

// Lifts an `Expected<double, ErrorCode>` kernel result into the `Value`
// shape AGGREGATE's dispatcher expects.
Value lift_kernel_result(Expected<double, ErrorCode> result) {
  if (!result) {
    return Value::error(result.error());
  }
  return Value::number(result.value());
}

Value run_count(const std::vector<Value>& cells) {
  // After `apply_filters`, numeric branches retain only Numbers; this gives
  // the same answer as iterating the post-filter `cells` directly.
  std::uint32_t n = 0;
  for (const Value& v : cells) {
    if (v.is_number()) {
      ++n;
    }
  }
  return Value::number(static_cast<double>(n));
}

Value run_counta(const std::vector<Value>& cells) {
  // The COUNTA branch of `apply_filters` already dropped plain and
  // raw-reference Blanks; everything remaining contributes 1.
  return Value::number(static_cast<double>(cells.size()));
}

// LARGE / SMALL — k must be a positive integer in [1, n]. k is truncated.
Value run_large_small(std::vector<double> xs, double k_raw, bool want_large) {
  if (xs.empty()) {
    return Value::error(ErrorCode::Num);
  }
  const double k_trunc = std::trunc(k_raw);
  if (!std::isfinite(k_trunc) || k_trunc < 1.0 || k_trunc > static_cast<double>(xs.size())) {
    return Value::error(ErrorCode::Num);
  }
  const auto k = static_cast<std::size_t>(k_trunc);
  std::sort(xs.begin(), xs.end());
  // LARGE: k-th largest = xs[n - k]. SMALL: k-th smallest = xs[k - 1].
  const double picked = want_large ? xs[xs.size() - k] : xs[k - 1];
  return to_finite_value(picked);
}

// PERCENTILE.INC. Domain / position-formula logic lives in
// `numeric_aggregate_kernels::percentile_sorted_inc`; this wrapper sorts in place
// and lifts the kernel result to `Value`.
Value run_percentile_inc(std::vector<double> xs, double p) {
  std::sort(xs.begin(), xs.end());
  return lift_kernel_result(numeric_aggregate_kernels::percentile_sorted_inc(xs, p));
}

// PERCENTILE.EXC. Domain / position-formula logic lives in
// `numeric_aggregate_kernels::percentile_sorted_exc`; this wrapper sorts in place
// and lifts the kernel result to `Value`.
Value run_percentile_exc(std::vector<double> xs, double p) {
  std::sort(xs.begin(), xs.end());
  return lift_kernel_result(numeric_aggregate_kernels::percentile_sorted_exc(xs, p));
}

// Shared body of QUARTILE.INC / QUARTILE.EXC: `quart` is truncated, must lie
// in [lo, hi], and selects the percentile `quart / 4` (exact in binary).
Value run_quartile(std::vector<double> xs, double quart_raw, double lo, double hi, bool exclusive) {
  const double q_trunc = std::trunc(quart_raw);
  if (!std::isfinite(q_trunc) || q_trunc < lo || q_trunc > hi) {
    return Value::error(ErrorCode::Num);
  }
  const double p = q_trunc / 4.0;
  return exclusive ? run_percentile_exc(std::move(xs), p) : run_percentile_inc(std::move(xs), p);
}

// QUARTILE.INC delegates to PERCENTILE.INC at p in {0, 0.25, 0.5, 0.75, 1.0}.
// `quart` must be an integer in [0, 4]; truncated like the rest.
Value run_quartile_inc(std::vector<double> xs, double quart_raw) {
  return run_quartile(std::move(xs), quart_raw, 0.0, 4.0, /*exclusive=*/false);
}

// QUARTILE.EXC delegates to PERCENTILE.EXC at p in {0.25, 0.5, 0.75}.
// `quart` must be an integer in {1, 2, 3}; 0 and 4 are rejected.
Value run_quartile_exc(std::vector<double> xs, double quart_raw) {
  return run_quartile(std::move(xs), quart_raw, 1.0, 3.0, /*exclusive=*/true);
}

// SUBTOTAL's function code selects the aggregator in 1..11 and repeats it in
// 101..111 with hidden rows excluded. Returns false for anything outside
// those two windows; `subtotal_apply` rejects it again on the same rule, so
// the two cannot disagree about what is valid.
bool subtotal_code_skips_hidden(double raw) noexcept {
  return std::isfinite(raw) && raw >= 101.0 && raw < 112.0;
}

// `fn` applied to each element of `in`, keeping its shape; 1x1 unwraps.
template <typename Fn>
Value map_array(const ArrayValue& in, const Fn& fn, Arena& arena) {
  Value* cells = nullptr;
  ArrayValue* out = allocate_array_value(in.rows, in.cols, arena, cells, kMaxDerivedArrayCells);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  const std::size_t n = static_cast<std::size_t>(in.rows) * in.cols;
  for (std::size_t i = 0; i < n; ++i) {
    cells[i] = fn(in.cells[i]);
  }
  return n == 1U ? cells[0] : Value::array(out);
}

}  // namespace

Value eval_subtotal_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  // function_num + at least one data arg.
  if (arity < 2U) {
    return Value::error(ErrorCode::Value);
  }

  const Value code_value = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (code_value.is_error()) {
    return code_value;
  }
  // The code is coerced twice — here to learn whether hidden rows are in
  // play, and again inside `subtotal_apply` to pick the mode. Coercion is
  // pure, so the second read cannot disagree with the first.
  auto code = coerce_to_number(code_value);
  if (!code) {
    return Value::error(code.error());
  }
  const HiddenScope scope = subtotal_code_skips_hidden(code.value()) ? HiddenScope::kAll : HiddenScope::kFiltered;

  std::vector<Value> cells;
  std::vector<bool> hidden;
  std::vector<bool> nested;
  Value err = Value::blank();
  for (std::uint32_t i = 1; i < arity; ++i) {
    if (!collect_arg(call.as_call_arg(i), arena, registry, ctx, scope, &cells, &hidden, &nested, &err)) {
      return err;
    }
  }
  drop_hidden_cells(&cells, make_drop_mask(cells.size(), /*use_hidden=*/true, hidden, /*use_nested=*/true, nested));

  // Hand the mode dispatch the same shape the eager dispatcher would have
  // built: the function code followed by the flattened data cells.
  std::vector<Value> argv;
  argv.reserve(cells.size() + 1U);
  argv.push_back(code_value);
  argv.insert(argv.end(), std::make_move_iterator(cells.begin()), std::make_move_iterator(cells.end()));
  return subtotal_apply(argv.data(), static_cast<std::uint32_t>(argv.size()), arena);
}

Value eval_aggregate_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  // function_num + options + at least one data arg.
  if (arity < 3U) {
    return Value::error(ErrorCode::Value);
  }

  Value err = Value::blank();
  auto fn_raw_or = read_scalar(call.as_call_arg(0), arena, registry, ctx);
  if (!fn_raw_or) {
    return Value::error(fn_raw_or.error());
  }
  const int code = static_cast<int>(std::trunc(fn_raw_or.value()));
  if (code < kFnMin || code > kFnMax) {
    return Value::error(ErrorCode::Value);
  }

  auto opts_raw_or = read_scalar(call.as_call_arg(1), arena, registry, ctx);
  if (!opts_raw_or) {
    return Value::error(opts_raw_or.error());
  }
  const int options = static_cast<int>(std::trunc(opts_raw_or.value()));
  if (options < 0 || options > 7) {
    return Value::error(ErrorCode::Value);
  }
  // Bit 0 (mask 1) is the hidden-row bit and bit 1 (mask 2) the error-ignore
  // bit. Bit 2 set (options 4..7) keeps nested SUBTOTAL / AGGREGATE cells.
  const bool ignore_hidden = (options & 1) != 0;
  const bool ignore_errors = (options & 2) != 0;
  const bool ignore_nested = options < 4;

  std::vector<Value> cells;
  std::vector<bool> hidden;
  std::vector<bool> nested;
  std::vector<bool>* nested_out = ignore_nested ? &nested : nullptr;

  if (code >= kFnKArgFirst) {
    // 14..19 — Excel requires exactly one data range plus a trailing k.
    // Anything other than `(fn, options, data, k)` -> #VALUE!.
    if (arity != 4U) {
      return Value::error(ErrorCode::Value);
    }
    if (!collect_arg(call.as_call_arg(2), arena, registry, ctx, HiddenScope::kAll, &cells, &hidden, nested_out, &err)) {
      return err;
    }
    drop_hidden_cells(&cells, make_drop_mask(cells.size(), ignore_hidden, hidden, ignore_nested, nested));
    if (!apply_filters(&cells, code, ignore_errors, &err)) {
      return err;
    }
    // k is a scalar metadata arg: errors propagate regardless of the
    // options bit (matches the function_num / options contract). An array k
    // evaluates the function once per element.
    const std::vector<double> xs = to_numbers(cells);
    const auto at_k = [&](const Value& k_value) -> Value {
      auto k_raw_or = scalar_number(k_value);
      if (!k_raw_or) {
        return Value::error(k_raw_or.error());
      }
      const double k_raw = k_raw_or.value();
      switch (code) {
        case kCodeLarge:
          return run_large_small(xs, k_raw, /*want_large=*/true);
        case kCodeSmall:
          return run_large_small(xs, k_raw, /*want_large=*/false);
        case kCodePercentileInc:
          return run_percentile_inc(xs, k_raw);
        case kCodeQuartileInc:
          return run_quartile_inc(xs, k_raw);
        case kCodePercentileExc:
          return run_percentile_exc(xs, k_raw);
        case kCodeQuartileExc:
          return run_quartile_exc(xs, k_raw);
        default:
          // Unreachable: code is constrained to [14, 19] in this branch.
          return Value::error(ErrorCode::Value);
      }
    };
    const Value k_value = eval_node(call.as_call_arg(3), arena, registry, ctx);
    return k_value.is_array() ? map_array(*k_value.as_array(), at_k, arena) : at_k(k_value);
  }

  // Codes 1..13 take references only: a value argument is #VALUE!, and an
  // array fourth argument evaluates the call per element, each one a value.
  for (std::uint32_t i = 2; i < arity; ++i) {
    const parser::AstNode& arg = call.as_call_arg(i);
    const parser::AstNode* ref = is_reference_shape(arg) ? &arg : resolve_binding_reference(arg, arena, registry, ctx);
    if (ref == nullptr) {
      const Value v = i == 3U ? eval_node(arg, arena, registry, ctx) : Value::blank();
      if (v.is_array()) {
        return map_array(*v.as_array(), [](const Value&) { return Value::error(ErrorCode::Value); }, arena);
      }
      return Value::error(ErrorCode::Value);
    }
    if (!collect_arg(*ref, arena, registry, ctx, HiddenScope::kAll, &cells, &hidden, nested_out, &err)) {
      return err;
    }
  }
  drop_hidden_cells(&cells, make_drop_mask(cells.size(), ignore_hidden, hidden, ignore_nested, nested));
  if (!apply_filters(&cells, code, ignore_errors, &err)) {
    return err;
  }

  switch (code) {
    case kCodeAverage:
      return lift_kernel_result(numeric_aggregate_kernels::run_average(to_numbers(cells)));
    case kCodeCount:
      return run_count(cells);
    case kCodeCountA:
      return run_counta(cells);
    case kCodeMax:
      return lift_kernel_result(numeric_aggregate_kernels::run_max(to_numbers(cells)));
    case kCodeMin:
      return lift_kernel_result(numeric_aggregate_kernels::run_min(to_numbers(cells)));
    case kCodeProduct:
      return lift_kernel_result(numeric_aggregate_kernels::run_product(to_numbers(cells)));
    case kCodeStdevS:
      return lift_kernel_result(numeric_aggregate_kernels::run_stdev(to_numbers(cells), /*sample=*/true));
    case kCodeStdevP:
      return lift_kernel_result(numeric_aggregate_kernels::run_stdev(to_numbers(cells), /*sample=*/false));
    case kCodeSum:
      return lift_kernel_result(numeric_aggregate_kernels::run_sum(to_numbers(cells)));
    case kCodeVarS:
      return lift_kernel_result(numeric_aggregate_kernels::run_variance(to_numbers(cells), /*sample=*/true));
    case kCodeVarP:
      return lift_kernel_result(numeric_aggregate_kernels::run_variance(to_numbers(cells), /*sample=*/false));
    case kCodeMedian:
      // The shared kernel sorts internally; do NOT pre-sort. Its empty-slice
      // code is `#NUM!`, which is what standalone MEDIAN reports.
      return lift_kernel_result(numeric_aggregate_kernels::run_median(to_numbers(cells)));
    case kCodeModeSngl:
      // First-occurrence tie-break (Excel MODE.SNGL): the shared kernel
      // consumes the cells in input order, so do NOT sort first.
      return lift_kernel_result(numeric_aggregate_kernels::mode_first_occurrence(to_numbers(cells)));
    default:
      // Unreachable: code is constrained to [1, 13] in this branch.
      return Value::error(ErrorCode::Value);
  }
}

Value eval_countblank_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U) {
    return Value::error(ErrorCode::Value);
  }
  const auto is_empty = [](const Value& v) { return v.is_blank() || (v.is_text() && v.as_text().empty()); };
  double total = 0.0;
  for (std::uint32_t i = 0; i < arity; ++i) {
    const parser::AstNode& node = resolve_range_binding(call.as_call_arg(i), ctx.name_env(), /*accept_ref=*/true);
    if (node.kind() == parser::NodeKind::ExternalRef) {
      return Value::error(ErrorCode::Value);
    }
    if (node.kind() == parser::NodeKind::Ref || is_range_shaped_ast(node)) {
      auto resolved = resolve_range_arg(node, arena, registry, ctx);
      if (!resolved) {
        return Value::error(resolved.error());
      }
      const std::vector<Value>& cells = resolved.value().cells;
      std::uint32_t declared_rows = 0;
      std::uint32_t declared_cols = 0;
      double declared = static_cast<double>(cells.size());
      if (static_reference_shape(node, ctx, &declared_rows, &declared_cols)) {
        declared = std::max(declared, static_cast<double>(declared_rows) * static_cast<double>(declared_cols));
      }
      double occupied = 0.0;
      for (const Value& cell : cells) {
        if (!is_empty(cell)) {
          occupied += 1.0;
        }
      }
      total += declared - occupied;
      continue;
    }
    const Value v = eval_node(node, arena, registry, ctx);
    if (v.is_array()) {
      const ArrayValue* array = v.as_array();
      const std::size_t n = static_cast<std::size_t>(array->rows) * static_cast<std::size_t>(array->cols);
      for (std::size_t k = 0; k < n; ++k) {
        if (is_empty(array->cells[k])) {
          total += 1.0;
        }
      }
    } else if (is_empty(v)) {
      total += 1.0;
    }
  }
  return Value::number(total);
}

}  // namespace eval
}  // namespace formulon
