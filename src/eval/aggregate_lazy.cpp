
#include "eval/aggregate_lazy.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <iterator>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/aggregate_kernels.h"
#include "eval/builtins/subtotal.h"
#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/name_env_resolve.h"
#include "eval/range_args.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

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
Expected<double, ErrorCode> read_scalar(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                                        const EvalContext& ctx) {
  const Value v = eval_node(node, arena, registry, ctx);
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

// --- Row visibility ------------------------------------------------------
//
// `SUBTOTAL(100+n)` and `AGGREGATE` with the hidden-row option bit skip cells
// that sit on a manually hidden row. Visibility is a property of the sheet
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
// Excel additionally distinguishes filter-hidden from manually hidden rows
// (codes 1..11 already skip filter-hidden rows). Formulon models only the
// `RowLayout::hidden` flag the OOXML reader and `fm_sheet_set_row_hidden`
// write, so both variants read that one flag and 1..11 include everything.

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
  const Sheet* sheet = nullptr;
  if (sheet_name.empty()) {
    sheet = ctx.current_sheet();
  } else if (ctx.workbook() != nullptr) {
    sheet = ctx.workbook()->sheet_by_name(sheet_name);
  }
  if (sheet == nullptr) {
    return nullptr;
  }
  *out_top_row = top;
  return sheet;
}

// Appends `count` visibility flags for the cells of one argument, given the
// argument's resolved shape. Cells on a hidden row are marked true.
void append_visibility(const parser::AstNode& node, const EvalContext& ctx, std::uint32_t rows, std::uint32_t cols,
                       std::vector<bool>* out_hidden) {
  const std::size_t count = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  std::uint32_t top = 0;
  const Sheet* sheet = reference_arg_origin(node, ctx, &top);
  if (sheet == nullptr || rows == 0U || cols == 0U) {
    out_hidden->insert(out_hidden->end(), count, false);
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
    out_hidden->insert(out_hidden->end(), cols, hidden_row[r]);
  }
}

// Appends every scalar Value produced by `arg_node` to `out_cells`, mirroring
// PERCENTOF's `sum_arg_for_percentof` provenance walk. LET-bound NameRefs
// resolve to their bound AST when range-shaped. On any expansion failure
// (e.g. `#REF!` from a missing sheet) returns false with the propagating
// error in `*out_err`. Returns true on a clean walk.
//
// `out_hidden` grows in lockstep with `out_cells`, one flag per appended
// cell, so a later filter can drop the cells that sit on hidden rows without
// re-deriving where each one came from.
//
// Unlike PERCENTOF this helper does NOT filter by Value kind: AGGREGATE's
// per-mode rules (numeric branches drop non-numerics; COUNTA counts them;
// the error-ignore bit decides whether errors short-circuit) are applied
// later by `apply_filters`.
bool collect_arg(const parser::AstNode& arg_node, Arena& arena, const FunctionRegistry& registry,
                 const EvalContext& ctx, std::vector<Value>* out_cells, std::vector<bool>* out_hidden, Value* out_err) {
  const parser::AstNode* effective = &arg_node;
  if (arg_node.kind() == parser::NodeKind::NameRef) {
    const parser::AstNode& resolved = resolve_name_ast(arg_node, ctx.name_env());
    if (&resolved != &arg_node && is_range_shaped_ast(resolved)) {
      effective = &resolved;
    }
  }
  const parser::AstNode& node = *effective;
  const parser::NodeKind k = node.kind();

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
      append_visibility(node, ctx, rr.rows, rr.cols, out_hidden);
    } else {
      out_hidden->insert(out_hidden->end(), n, false);
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
    out_hidden->insert(out_hidden->end(), n, false);
    return true;
  }
  out_cells->push_back(v);
  out_hidden->push_back(false);
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
// in `aggregate_kernels.h` so SUBTOTAL and AGGREGATE cannot drift. Empty-
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
  if (!std::isfinite(picked)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(picked);
}

// PERCENTILE.INC. Domain / position-formula logic lives in
// `aggregate_kernels::percentile_sorted_inc`; this wrapper sorts in place
// and lifts the kernel result to `Value`.
Value run_percentile_inc(std::vector<double> xs, double p) {
  std::sort(xs.begin(), xs.end());
  return lift_kernel_result(aggregate_kernels::percentile_sorted_inc(xs, p));
}

// PERCENTILE.EXC. Domain / position-formula logic lives in
// `aggregate_kernels::percentile_sorted_exc`; this wrapper sorts in place
// and lifts the kernel result to `Value`.
Value run_percentile_exc(std::vector<double> xs, double p) {
  std::sort(xs.begin(), xs.end());
  return lift_kernel_result(aggregate_kernels::percentile_sorted_exc(xs, p));
}

// QUARTILE.INC delegates to PERCENTILE.INC at p ∈ {0, 0.25, 0.5, 0.75, 1.0}.
// `quart` must be an integer in [0, 4]; truncated like the rest.
Value run_quartile_inc(std::vector<double> xs, double quart_raw) {
  const double q_trunc = std::trunc(quart_raw);
  if (!std::isfinite(q_trunc) || q_trunc < 0.0 || q_trunc > 4.0) {
    return Value::error(ErrorCode::Num);
  }
  static constexpr double kProb[] = {0.0, 0.25, 0.5, 0.75, 1.0};
  const auto idx = static_cast<std::size_t>(q_trunc);
  return run_percentile_inc(std::move(xs), kProb[idx]);
}

// QUARTILE.EXC delegates to PERCENTILE.EXC at p ∈ {0.25, 0.5, 0.75}.
// `quart` must be an integer in {1, 2, 3}; 0 and 4 are rejected.
Value run_quartile_exc(std::vector<double> xs, double quart_raw) {
  const double q_trunc = std::trunc(quart_raw);
  if (!std::isfinite(q_trunc) || q_trunc < 1.0 || q_trunc > 3.0) {
    return Value::error(ErrorCode::Num);
  }
  static constexpr double kProb[] = {0.25, 0.5, 0.75};
  const auto idx = static_cast<std::size_t>(q_trunc) - 1U;
  return run_percentile_exc(std::move(xs), kProb[idx]);
}

// SUBTOTAL's function code selects the aggregator in 1..11 and repeats it in
// 101..111 with hidden rows excluded. Returns false for anything outside
// those two windows; `subtotal_apply` rejects it again on the same rule, so
// the two cannot disagree about what is valid.
bool subtotal_code_skips_hidden(double raw) noexcept {
  return std::isfinite(raw) && raw >= 101.0 && raw < 112.0;
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
  const bool skip_hidden = subtotal_code_skips_hidden(code.value());

  std::vector<Value> cells;
  std::vector<bool> hidden;
  Value err = Value::blank();
  for (std::uint32_t i = 1; i < arity; ++i) {
    if (!collect_arg(call.as_call_arg(i), arena, registry, ctx, &cells, &hidden, &err)) {
      return err;
    }
  }
  if (skip_hidden) {
    drop_hidden_cells(&cells, hidden);
  }

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
  // bit. Bit 2 selects whether nested SUBTOTAL / AGGREGATE results are also
  // skipped, which is not yet observable; see the file header.
  const bool ignore_hidden = (options & 1) != 0;
  const bool ignore_errors = (options & 2) != 0;

  std::vector<Value> cells;
  std::vector<bool> hidden;

  if (code >= kFnKArgFirst) {
    // 14..19 — Excel requires exactly one data range plus a trailing k.
    // Anything other than `(fn, options, data, k)` -> #VALUE!.
    if (arity != 4U) {
      return Value::error(ErrorCode::Value);
    }
    if (!collect_arg(call.as_call_arg(2), arena, registry, ctx, &cells, &hidden, &err)) {
      return err;
    }
    if (ignore_hidden) {
      drop_hidden_cells(&cells, hidden);
    }
    if (!apply_filters(&cells, code, ignore_errors, &err)) {
      return err;
    }
    // k is a scalar metadata arg: errors propagate regardless of the
    // options bit (matches the function_num / options contract).
    auto k_raw_or = read_scalar(call.as_call_arg(3), arena, registry, ctx);
    if (!k_raw_or) {
      return Value::error(k_raw_or.error());
    }
    const double k_raw = k_raw_or.value();
    std::vector<double> xs = to_numbers(cells);
    switch (code) {
      case kCodeLarge:
        return run_large_small(std::move(xs), k_raw, /*want_large=*/true);
      case kCodeSmall:
        return run_large_small(std::move(xs), k_raw, /*want_large=*/false);
      case kCodePercentileInc:
        return run_percentile_inc(std::move(xs), k_raw);
      case kCodeQuartileInc:
        return run_quartile_inc(std::move(xs), k_raw);
      case kCodePercentileExc:
        return run_percentile_exc(std::move(xs), k_raw);
      case kCodeQuartileExc:
        return run_quartile_exc(std::move(xs), k_raw);
      default:
        // Unreachable: code is constrained to [14, 19] in this branch.
        return Value::error(ErrorCode::Value);
    }
  }

  // Codes 1..13 — every remaining positional arg is data.
  for (std::uint32_t i = 2; i < arity; ++i) {
    if (!collect_arg(call.as_call_arg(i), arena, registry, ctx, &cells, &hidden, &err)) {
      return err;
    }
  }
  if (ignore_hidden) {
    drop_hidden_cells(&cells, hidden);
  }
  if (!apply_filters(&cells, code, ignore_errors, &err)) {
    return err;
  }

  switch (code) {
    case kCodeAverage:
      return lift_kernel_result(aggregate_kernels::run_average(to_numbers(cells)));
    case kCodeCount:
      return run_count(cells);
    case kCodeCountA:
      return run_counta(cells);
    case kCodeMax:
      return lift_kernel_result(aggregate_kernels::run_max(to_numbers(cells)));
    case kCodeMin:
      return lift_kernel_result(aggregate_kernels::run_min(to_numbers(cells)));
    case kCodeProduct:
      return lift_kernel_result(aggregate_kernels::run_product(to_numbers(cells)));
    case kCodeStdevS:
      return lift_kernel_result(aggregate_kernels::run_stdev(to_numbers(cells), /*sample=*/true));
    case kCodeStdevP:
      return lift_kernel_result(aggregate_kernels::run_stdev(to_numbers(cells), /*sample=*/false));
    case kCodeSum:
      return lift_kernel_result(aggregate_kernels::run_sum(to_numbers(cells)));
    case kCodeVarS:
      return lift_kernel_result(aggregate_kernels::run_variance(to_numbers(cells), /*sample=*/true));
    case kCodeVarP:
      return lift_kernel_result(aggregate_kernels::run_variance(to_numbers(cells), /*sample=*/false));
    case kCodeMedian:
      // The shared kernel sorts internally; do NOT pre-sort. Its empty-slice
      // code is `#NUM!`, which is what standalone MEDIAN reports.
      return lift_kernel_result(aggregate_kernels::run_median(to_numbers(cells)));
    case kCodeModeSngl:
      // First-occurrence tie-break (Excel MODE.SNGL): the shared kernel
      // consumes the cells in input order, so do NOT sort first.
      return lift_kernel_result(aggregate_kernels::mode_first_occurrence(to_numbers(cells)));
    default:
      // Unreachable: code is constrained to [1, 13] in this branch.
      return Value::error(ErrorCode::Value);
  }
}

}  // namespace eval
}  // namespace formulon
