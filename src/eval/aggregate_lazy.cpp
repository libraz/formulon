
#include "eval/aggregate_lazy.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <iterator>
#include <optional>
#include <string_view>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "eval/aggregate_filter.h"
#include "eval/array_alloc.h"
#include "eval/builtin_names.h"
#include "eval/builtins/numeric_helpers.h"
#include "eval/builtins/subtotal.h"
#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/external_ref.h"
#include "eval/lazy_impls.h"
#include "eval/name_env_resolve.h"
#include "eval/range_args.h"
#include "eval/range_resolvers.h"
#include "eval/tree_walker/dispatch.h"
#include "numeric_aggregate_kernels.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/formula_text.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

using builtins_detail::to_finite_value;

// 1..13 are the SUBTOTAL-aligned modes; 14..19 are the AGGREGATE-only "k-
// arg" modes. Storing the integer code rather than an enum keeps the
// k-arg validation switch easy to read against the Excel docs. The plain
// numeric aggregators (1, 4..11) dispatch through `run_aggregate_code`.
constexpr int kCodeCount = 2;
constexpr int kCodeCountA = 3;
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

// Decodes a plain `Ref` or a `Ref:Ref` range into its endpoints (the same
// reference twice for a single `Ref`). False for any other node, including
// OFFSET / INDIRECT endpoints, which carry no static provenance.
bool static_ref_endpoints(const parser::AstNode& node, const parser::Reference** first,
                          const parser::Reference** second) {
  if (node.kind() == parser::NodeKind::Ref) {
    *first = &node.as_ref();
    *second = *first;
    return true;
  }
  if (node.kind() == parser::NodeKind::RangeOp && node.as_range_arity() == 2U &&
      node.as_range_operand(0).kind() == parser::NodeKind::Ref &&
      node.as_range_operand(1).kind() == parser::NodeKind::Ref) {
    *first = &node.as_range_operand(0).as_ref();
    *second = &node.as_range_operand(1).as_ref();
    return true;
  }
  return false;
}

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
  if (!static_ref_endpoints(node, &first, &second)) {
    return nullptr;
  }

  const bool whole_column = first->is_full_col || second->is_full_col;
  const std::uint32_t top = whole_column ? 0U : std::min(first->row, second->row);

  // The parser keeps a `:` operator's qualifier on the left endpoint, so
  // `Sheet2!A1:B2` arrives as RangeOp(Ref{sheet=Sheet2}, Ref{sheet=""}) and
  // the right one inherits; `expand_range` also accepts the mirrored shape,
  // so read whichever endpoint carries a name. A pair that names two
  // different sheets is a `#REF!` the range resolver already rejected.
  std::string_view sheet_name = first->sheet;
  if (sheet_name.empty()) {
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
  const bool default_hidden = sheet->format_defaults().zero_height;
  std::vector<bool> hidden_row(rows, default_hidden);
  for (const RowLayout& row : sheet->layout().row_overrides) {
    if (row.row < top) {
      continue;
    }
    const std::uint32_t offset = row.row - top;
    if (offset < rows) {
      hidden_row[offset] = row.hidden;
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

constexpr std::string_view kNestedCallNames[] = {"SUBTOTAL", "AGGREGATE"};

bool is_nested_call_name(std::string_view name) noexcept {
  for (const std::string_view nested : kNestedCallNames) {
    if (strings::case_insensitive_eq(name, nested)) {
      return true;
    }
  }
  return false;
}

bool ast_has_nested_call(const parser::AstNode& node) {
  if (node.kind() == parser::NodeKind::Call && is_nested_call_name(node.as_call_name())) {
    return true;
  }
  return parser::any_child_node(node, [](const parser::AstNode& child) { return ast_has_nested_call(child); });
}

bool text_may_hold_nested_call(std::string_view text) noexcept {
  for (const std::string_view needle : kNestedCallNames) {
    if (strings::case_insensitive_contains(text, needle)) {
      return true;
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
  if (count == 0U || !static_ref_endpoints(node, &first, &second)) {
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

// Appends `n` unflagged (visible, not nested) entries to the flag vectors.
void append_unflagged(std::vector<bool>* out_hidden, std::vector<bool>* out_nested, std::size_t n) {
  out_hidden->resize(out_hidden->size() + n, false);
  if (out_nested != nullptr) {
    out_nested->resize(out_nested->size() + n, false);
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
    if (const std::optional<ErrorCode> invalid = union_reference_error(node, ctx)) {
      *out_err = Value::error(*invalid);
      return false;
    }
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
      append_unflagged(out_hidden, out_nested, n);
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
        append_unflagged(out_hidden, out_nested, 1U);
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
    append_unflagged(out_hidden, out_nested, n);
    return true;
  }
  out_cells->push_back(v);
  append_unflagged(out_hidden, out_nested, 1U);
  return true;
}

// Moves cells[i] down to slot `*kept` and advances it.
void keep_cell(std::vector<Value>* cells, std::size_t* kept, std::size_t i) {
  if (*kept != i) {
    (*cells)[*kept] = std::move((*cells)[i]);
  }
  ++*kept;
}

void truncate_cells(std::vector<Value>* cells, std::size_t kept) {
  cells->erase(cells->begin() + static_cast<std::vector<Value>::difference_type>(kept), cells->end());
}

// Filters `cells` in place according to the options bit and the function
// code's expected provenance:
//
//   * Errors: dropped silently when ignore_errors == true. Otherwise COUNT
//     skips them, COUNTA counts them, and for every other code the first
//     error short-circuits the call and is written to `*out_err`.
//   * COUNTA (code 3): keep every non-blank, plus blanks owned by a derived
//     value array; raw-reference blanks are dropped.
//   * Numeric modes (everything else): keep only Numbers; drop Bool / Text /
//     Blank.
//
// Returns true on a clean filter, false (with `*out_err` populated) when an
// un-ignored error short-circuits.
bool apply_filters(std::vector<Value>* cells, int code, bool ignore_errors, Value* out_err) {
  std::size_t kept = 0;
  for (std::size_t i = 0; i < cells->size(); ++i) {
    const Value& v = (*cells)[i];
    if (v.is_error()) {
      if (ignore_errors || code == kCodeCount) {
        continue;
      }
      if (code == kCodeCountA) {
        keep_cell(cells, &kept, i);  // an error cell is a non-blank cell
        continue;
      }
      *out_err = v;
      return false;
    }
    if (code == kCodeCountA) {
      if (!v.is_blank() || v.blank_counts_for_counta()) {
        keep_cell(cells, &kept, i);
      }
      continue;
    }
    // Numeric branches: drop everything that is not a Number.
    if (v.is_number()) {
      keep_cell(cells, &kept, i);
    }
  }
  truncate_cells(cells, kept);
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

// LARGE / SMALL share the standalone functions' k rule
// (`builtins_detail::kth_ascending_position`).
Value run_large_small(std::vector<double> xs, double k_raw, bool want_large) {
  const auto idx = builtins_detail::kth_ascending_position(k_raw, xs.size(), want_large);
  if (!idx) {
    return Value::error(idx.error());
  }
  numeric_aggregate_kernels::select_kth_preserving_sort_bits(xs, idx.value() - 1U);
  return to_finite_value(xs[idx.value() - 1U]);
}

// PERCENTILE.INC / PERCENTILE.EXC. Domain / position-formula logic lives in
// `numeric_aggregate_kernels::percentile_sorted_{inc,exc}`; this wrapper sorts in
// place and lifts the kernel result to `Value`.
Value run_percentile(std::vector<double> xs, double p, bool exclusive) {
  std::sort(xs.begin(), xs.end());
  return numeric_aggregate_kernels::value_from_kernel_result(
      exclusive ? numeric_aggregate_kernels::percentile_sorted_exc(xs, p)
                : numeric_aggregate_kernels::percentile_sorted_inc(xs, p));
}

// Shared body of QUARTILE.INC / QUARTILE.EXC: `quart` is truncated, must lie
// in [lo, hi], and selects the percentile `quart / 4` (exact in binary).
// INC accepts [0, 4]; EXC accepts [1, 3].
Value run_quartile(std::vector<double> xs, double quart_raw, double lo, double hi, bool exclusive) {
  const double q_trunc = std::trunc(quart_raw);
  if (!std::isfinite(q_trunc) || q_trunc < lo || q_trunc > hi) {
    return Value::error(ErrorCode::Num);
  }
  const double p = q_trunc / 4.0;
  return run_percentile(std::move(xs), p, exclusive);
}

// SUBTOTAL's function code selects the aggregator in 1..11 and repeats it in
// 101..111 with hidden rows excluded. Returns false for anything outside
// those two windows; `subtotal_apply` rejects it again on the same rule, so
// the two cannot disagree about what is valid.
bool subtotal_code_skips_hidden(double raw) noexcept {
  return std::isfinite(raw) && raw >= 101.0 && raw < 112.0;
}

}  // namespace

void aggregate_filter_detail::compact_by_flags(std::vector<Value>* cells, const std::vector<bool>& hidden,
                                               bool use_hidden, const std::vector<bool>& nested, bool use_nested) {
  std::size_t kept = 0;
  for (std::size_t i = 0; i < cells->size(); ++i) {
    const bool drop_hidden = use_hidden && i < hidden.size() && hidden[i];
    const bool drop_nested = use_nested && i < nested.size() && nested[i];
    if (drop_hidden || drop_nested) {
      continue;
    }
    keep_cell(cells, &kept, i);
  }
  truncate_cells(cells, kept);
}

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
  aggregate_filter_detail::compact_by_flags(&cells, hidden, /*use_hidden=*/true, nested, /*use_nested=*/true);

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
  const double code_d = std::trunc(fn_raw_or.value());
  if (code_d < static_cast<double>(kFnMin) || code_d > static_cast<double>(kFnMax)) {
    return Value::error(ErrorCode::Value);
  }
  const int code = static_cast<int>(code_d);

  auto opts_raw_or = read_scalar(call.as_call_arg(1), arena, registry, ctx);
  if (!opts_raw_or) {
    return Value::error(opts_raw_or.error());
  }
  const double options_d = std::trunc(opts_raw_or.value());
  if (options_d < 0.0 || options_d > 7.0) {
    return Value::error(ErrorCode::Value);
  }
  const int options = static_cast<int>(options_d);
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
    aggregate_filter_detail::compact_by_flags(&cells, hidden, ignore_hidden, nested, ignore_nested);
    if (!apply_filters(&cells, code, ignore_errors, &err)) {
      // An array k fails per element with the data error.
      const Value k_for_error = eval_node(call.as_call_arg(3), arena, registry, ctx);
      if (k_for_error.is_array()) {
        return map_array_with(*k_for_error.as_array(), [&err](const Value&) { return err; }, arena);
      }
      return err;
    }
    // k is a scalar metadata arg: errors propagate regardless of the
    // options bit (matches the function_num / options contract). An array k
    // evaluates the function once per element.
    std::vector<double> xs = to_numbers(cells);
    // The runners sort their copy; the last (or only) evaluation hands `xs` over.
    const auto at_k = [&](const Value& k_value, bool last) -> Value {
      auto k_raw_or = scalar_number(k_value);
      if (!k_raw_or) {
        return Value::error(k_raw_or.error());
      }
      const double k_raw = k_raw_or.value();
      const auto data = [&]() -> std::vector<double> {
        if (last) {
          return std::move(xs);
        }
        return xs;
      };
      switch (code) {
        case kCodeLarge:
          return run_large_small(data(), k_raw, /*want_large=*/true);
        case kCodeSmall:
          return run_large_small(data(), k_raw, /*want_large=*/false);
        case kCodePercentileInc:
          return run_percentile(data(), k_raw, /*exclusive=*/false);
        case kCodeQuartileInc:
          return run_quartile(data(), k_raw, 0.0, 4.0, /*exclusive=*/false);
        case kCodePercentileExc:
          return run_percentile(data(), k_raw, /*exclusive=*/true);
        case kCodeQuartileExc:
          return run_quartile(data(), k_raw, 1.0, 3.0, /*exclusive=*/true);
        default:
          // Unreachable: code is constrained to [14, 19] in this branch.
          return Value::error(ErrorCode::Value);
      }
    };
    const Value k_value = eval_node(call.as_call_arg(3), arena, registry, ctx);
    if (k_value.is_array()) {
      return map_array_with(*k_value.as_array(), [&](const Value& k) { return at_k(k, /*last=*/false); }, arena);
    }
    return at_k(k_value, /*last=*/true);
  }

  // Codes 1..13 take references only: a value argument is #VALUE!, and an
  // array fourth argument evaluates the call per element, each one a value.
  for (std::uint32_t i = 2; i < arity; ++i) {
    const parser::AstNode& arg = call.as_call_arg(i);
    const parser::AstNode* ref = is_reference_shape(arg) ? &arg : resolve_binding_reference(arg, arena, registry, ctx);
    if (ref == nullptr) {
      const Value v = i == 3U ? eval_node(arg, arena, registry, ctx) : Value::blank();
      if (v.is_array()) {
        return map_array_with(*v.as_array(), [](const Value&) { return Value::error(ErrorCode::Value); }, arena);
      }
      return Value::error(ErrorCode::Value);
    }
    if (!collect_arg(*ref, arena, registry, ctx, HiddenScope::kAll, &cells, &hidden, nested_out, &err)) {
      return err;
    }
  }
  aggregate_filter_detail::compact_by_flags(&cells, hidden, ignore_hidden, nested, ignore_nested);
  if (!apply_filters(&cells, code, ignore_errors, &err)) {
    return err;
  }

  switch (code) {
    case kCodeCount:
      return run_count(cells);
    case kCodeCountA:
      return run_counta(cells);
    case kCodeMedian:
      // The shared kernel sorts internally; do NOT pre-sort. Its empty-slice
      // code is `#NUM!`, which is what standalone MEDIAN reports.
      return numeric_aggregate_kernels::value_from_kernel_result(
          numeric_aggregate_kernels::run_median(to_numbers(cells)));
    case kCodeModeSngl:
      // First-occurrence tie-break (Excel MODE.SNGL): the shared kernel
      // consumes the cells in input order, so do NOT sort first.
      return numeric_aggregate_kernels::value_from_kernel_result(
          numeric_aggregate_kernels::mode_first_occurrence(to_numbers(cells)));
    default:
      return numeric_aggregate_kernels::run_aggregate_code(code, to_numbers(cells));
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
    if (is_grouped_ref3d(node)) {
      return Value::error(ErrorCode::Name);
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
