// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the tree-walk evaluator. See `tree_walker.h` for the
// public contract and the design references.

#include "eval/tree_walker.h"

#include <cmath>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/coerce.h"
#include "eval/conditional_aggregates.h"
#include "eval/criteria.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/lazy_impls.h"
#include "eval/lookups_classic.h"
#include "eval/range_args.h"
#include "eval/special_forms_lazy.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/expected.h"  // FM_CHECK
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// ---------------------------------------------------------------------------
// Cross-type comparison
// ---------------------------------------------------------------------------

// Excel cross-type comparison order: Number < Text < Bool. Blank coerces to
// numeric zero. Text equality and ordering are case-insensitive over ASCII
// letters; locale-aware comparison is deferred. NaN compares as "unordered":
// every relational operator returns FALSE except `<>`.
//
// `out_unordered` is set to true iff one of the operands is NaN; the caller
// uses it to short-circuit relational operators to FALSE while still
// returning TRUE for `<>`. The integer return value (-1/0/+1) is meaningful
// only when `out_unordered` is false.
int compare_values(const Value& lhs, const Value& rhs, bool* out_unordered) {
  *out_unordered = false;

  auto rank = [](ValueKind k) -> int {
    // Blank is treated as numeric (0) for comparison purposes.
    switch (k) {
      case ValueKind::Number:
      case ValueKind::Blank:
        return 0;
      case ValueKind::Text:
        return 1;
      case ValueKind::Bool:
        return 2;
      default:
        return 3;
    }
  };

  const int lr = rank(lhs.kind());
  const int rr = rank(rhs.kind());
  if (lr != rr) {
    return lr < rr ? -1 : 1;
  }

  switch (lr) {
    case 0: {
      const double a = lhs.is_blank() ? 0.0 : lhs.as_number();
      const double b = rhs.is_blank() ? 0.0 : rhs.as_number();
      if (std::isnan(a) || std::isnan(b)) {
        *out_unordered = true;
        return 0;
      }
      if (a < b) {
        return -1;
      }
      if (a > b) {
        return 1;
      }
      return 0;
    }
    case 1: {
      const std::string_view a = lhs.as_text();
      const std::string_view b = rhs.as_text();
      const std::size_t n = a.size() < b.size() ? a.size() : b.size();
      for (std::size_t i = 0; i < n; ++i) {
        const char ca = strings::ascii_to_lower(a[i]);
        const char cb = strings::ascii_to_lower(b[i]);
        if (ca != cb) {
          return ca < cb ? -1 : 1;
        }
      }
      if (a.size() != b.size()) {
        return a.size() < b.size() ? -1 : 1;
      }
      return 0;
    }
    case 2: {
      const bool a = lhs.as_boolean();
      const bool b = rhs.as_boolean();
      if (a == b) {
        return 0;
      }
      // FALSE < TRUE.
      return a ? 1 : -1;
    }
    default:
      return 0;
  }
}

// ---------------------------------------------------------------------------
// Per-operator helpers
// ---------------------------------------------------------------------------

Value finalize_arithmetic(double r) {
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

Value apply_unary(parser::UnaryOp op, const Value& operand) {
  if (operand.is_error()) {
    return operand;
  }
  auto coerced = coerce_to_number(operand);
  if (!coerced) {
    return Value::error(coerced.error());
  }
  const double x = coerced.value();
  switch (op) {
    case parser::UnaryOp::Plus:
      return finalize_arithmetic(x);
    case parser::UnaryOp::Minus:
      return finalize_arithmetic(-x);
    case parser::UnaryOp::Percent:
      return finalize_arithmetic(x / 100.0);
  }
  return Value::error(ErrorCode::Value);
}

Value apply_arithmetic(parser::BinOp op, double lhs, double rhs) {
  switch (op) {
    case parser::BinOp::Add:
      return finalize_arithmetic(lhs + rhs);
    case parser::BinOp::Sub:
      return finalize_arithmetic(lhs - rhs);
    case parser::BinOp::Mul:
      return finalize_arithmetic(lhs * rhs);
    case parser::BinOp::Div:
      // Excel reports #DIV/0! for any division whose divisor is exactly
      // zero, including 0/0 (no #NUM! tie-break).
      if (rhs == 0.0) {
        return Value::error(ErrorCode::Div0);
      }
      return finalize_arithmetic(lhs / rhs);
    case parser::BinOp::Pow: {
      // Delegates to the shared `apply_pow` helper so the `^` operator and
      // the `POWER()` builtin cannot drift apart on edge cases.
      auto r = apply_pow(lhs, rhs);
      if (!r) {
        return Value::error(r.error());
      }
      return Value::number(r.value());
    }
    default:
      // Caller guarantees op is arithmetic.
      FM_CHECK(false, "apply_arithmetic called with non-arithmetic op");
      return Value::error(ErrorCode::Value);
  }
}

Value apply_concat(const Value& lhs, const Value& rhs, Arena& arena) {
  auto lhs_text = coerce_to_text(lhs);
  if (!lhs_text) {
    return Value::error(lhs_text.error());
  }
  auto rhs_text = coerce_to_text(rhs);
  if (!rhs_text) {
    return Value::error(rhs_text.error());
  }
  std::string joined;
  joined.reserve(lhs_text.value().size() + rhs_text.value().size());
  joined.append(lhs_text.value());
  joined.append(rhs_text.value());
  const std::string_view interned = arena.intern(joined);
  // Empty input is fine: Arena::intern returns an empty view that is still
  // a valid Text payload.
  return Value::text(interned);
}

Value apply_comparison(parser::BinOp op, const Value& lhs, const Value& rhs) {
  bool unordered = false;
  const int cmp = compare_values(lhs, rhs, &unordered);
  switch (op) {
    case parser::BinOp::Eq:
      return Value::boolean(!unordered && cmp == 0);
    case parser::BinOp::NotEq:
      // NaN != anything is TRUE, matching IEEE-754 semantics.
      return Value::boolean(unordered || cmp != 0);
    case parser::BinOp::Lt:
      return Value::boolean(!unordered && cmp < 0);
    case parser::BinOp::LtEq:
      return Value::boolean(!unordered && cmp <= 0);
    case parser::BinOp::Gt:
      return Value::boolean(!unordered && cmp > 0);
    case parser::BinOp::GtEq:
      return Value::boolean(!unordered && cmp >= 0);
    default:
      FM_CHECK(false, "apply_comparison called with non-comparison op");
      return Value::error(ErrorCode::Value);
  }
}

// ---------------------------------------------------------------------------
// Recursive evaluator
// ---------------------------------------------------------------------------
//
// `eval_node` is declared in `eval/lazy_impls.h` with external linkage so
// lazy-impl translation units (e.g. `special_forms_lazy.cpp`) can reach
// it. Its definition lives at the bottom of this file, outside the
// anonymous namespace.

// ---------------------------------------------------------------------------
// Lazy (short-circuit) function impls
// ---------------------------------------------------------------------------
//
// Each lazy impl receives the full `Call` AST node so it can pull arguments
// out by index and decide which subtrees to evaluate. The eager path in
// `dispatch_call` is bypassed entirely: arity checks and error propagation
// belong inside each impl. On arity mismatch the impls return #VALUE! to
// match the eager dispatcher's behaviour.
//
// Current entries:
//   IF          - short-circuit branch: only the taken side is evaluated.
//   IFERROR     - evaluates fallback only when primary is any error.
//   IFNA        - evaluates fallback only when primary is exactly #N/A.
//   COUNTIF     - range-aware: arg 0 must be a range/Ref, arg 1 is a scalar
//                 criterion evaluated once; counts matching cells.
//   SUMIF       - range-aware: arg 0 is the criteria range, arg 2 (optional)
//                 is the parallel sum range; sums matching numeric cells.
//   AVERAGEIF   - like SUMIF, but returns the mean of matching numeric
//                 cells or #DIV/0! when nothing matches.
//   COUNTIFS    - multi-criteria AND across N (range, criterion) pairs.
//   SUMIFS      - like COUNTIFS, but with a result range as the leading arg.
//   AVERAGEIFS  - like SUMIFS, returns mean or #DIV/0! when no matches.
//   MAXIFS      - like SUMIFS, returns max of numerics (0 if no matches).
//   MINIFS      - like SUMIFS, returns min of numerics (0 if no matches).
//   CHOOSE      - index-selected argument; only the chosen subtree runs.
//   INDEX       - range-aware: shape (rows,cols) of arg 0 is used to pick
//                 a single cell by (row_num, col_num).
//   MATCH       - range-aware: lookup_array (arg 1) must be a 1-D range/Ref.
//
// The conditional aggregators (`*IF`/`*IFS`) cannot ride on the eager
// `accepts_ranges` path because arg 0 must reach the impl as AST (so a
// bare single-cell Ref can be treated as a 1-cell range) AND the parallel
// result / additional criteria ranges must iterate in lockstep rather than
// being flattened into a single values vector alongside the first.

// The lazy impls themselves live in per-family translation units:
//   IF / IFERROR / IFNA                        -> src/eval/special_forms_lazy.cpp
//   COUNTIF / SUMIF / AVERAGEIF / *IFS         -> src/eval/conditional_aggregates.cpp
//   CHOOSE / INDEX / MATCH / VLOOKUP / HLOOKUP -> src/eval/lookups_classic.cpp
// Each family publishes its externs via its own header
// (`eval/special_forms_lazy.h`, `eval/conditional_aggregates.h`,
// `eval/lookups_classic.h`), which the dispatch table below includes.

// ---------------------------------------------------------------------------
// XLOOKUP / XMATCH shared machinery
// ---------------------------------------------------------------------------

// Excel 365 match-mode codes for XLOOKUP / XMATCH. `Exact` is the default
// (case-insensitive ASCII equality; no wildcard metacharacters honoured).
// `Smaller` / `Larger` fall back to the closest cell below / above the
// lookup value when no exact hit is found. `Wildcard` is "exact + honour
// `*` / `?` / `~` in text patterns".
enum class XMatchMode : std::int8_t { Exact = 0, Smaller = -1, Larger = 1, Wildcard = 2 };

// Excel 365 search-mode codes for XLOOKUP / XMATCH. `FirstToLast` /
// `LastToFirst` are linear scans in the obvious direction. `BinaryAsc` /
// `BinaryDesc` assume the caller's array is already sorted (Excel does not
// validate the sort order — neither do we) and use a lower-bound style
// search.
enum class XSearchMode : std::int8_t { FirstToLast = 1, LastToFirst = -1, BinaryAsc = 2, BinaryDesc = -2 };

// Three-way compare of (cell - lookup) with XLOOKUP's accepted divergence:
// cross-type pairings are reported as incomparable rather than ordered by
// value-kind (Excel actually orders kinds). Returns `true` on success with
// the sign in `*out_cmp`; `false` when the pair cannot be compared.
bool xlookup_cmp(const Value& cell, const Value& lookup, int* out_cmp) {
  auto cmp_numeric = [](double x, double y) -> int {
    if (x < y) {
      return -1;
    }
    if (x > y) {
      return 1;
    }
    return 0;
  };
  if (lookup.is_text() && cell.is_text()) {
    *out_cmp = strings::case_insensitive_compare(cell.as_text(), lookup.as_text());
    return true;
  }
  if ((lookup.is_number() || lookup.is_blank()) && (cell.is_number() || cell.is_blank())) {
    const double lv = lookup.is_blank() ? 0.0 : lookup.as_number();
    const double cv = cell.is_blank() ? 0.0 : cell.as_number();
    *out_cmp = cmp_numeric(cv, lv);
    return true;
  }
  if (lookup.is_boolean() && cell.is_boolean()) {
    const int lb = lookup.as_boolean() ? 1 : 0;
    const int cb = cell.as_boolean() ? 1 : 0;
    *out_cmp = cmp_numeric(cb, lb);
    return true;
  }
  return false;
}

// Exact-equality test with optional DOS-style wildcard expansion on Text vs
// Text pairs. For non-text / cross-type kinds this is a literal equality
// compare with the same "Blank acts as numeric 0" rule MATCH / VLOOKUP use.
// When `wildcards` is true AND lookup_value is text, pattern matching is
// always routed through `wildcard_match` so that `~*` continues to mean
// "literal *". See the VLOOKUP comment at the top of `lookup_scan` for why
// we deliberately do not gate on `scan_has_wildcard`.
bool xlookup_exact_eq(const Value& cell, const Value& lookup, bool wildcards) {
  if (lookup.is_text()) {
    if (!cell.is_text()) {
      return false;
    }
    const std::string pat_lower = strings::to_ascii_lower(lookup.as_text());
    const std::string cell_lower = strings::to_ascii_lower(cell.as_text());
    if (wildcards) {
      return wildcard_match(pat_lower, cell_lower);
    }
    return pat_lower == cell_lower;
  }
  if (lookup.is_number() || lookup.is_blank()) {
    const double target = lookup.is_blank() ? 0.0 : lookup.as_number();
    if (cell.is_number()) {
      return cell.as_number() == target;
    }
    if (cell.is_blank()) {
      return target == 0.0;
    }
    return false;
  }
  if (lookup.is_boolean()) {
    return cell.is_boolean() && cell.as_boolean() == lookup.as_boolean();
  }
  return false;
}

// Scan `cells` (already flattened to a 1-D axis of the lookup_array) for
// `lookup_value`, honouring Excel 365's XLOOKUP match_mode / search_mode
// combinations. Returns the 0-based offset of the match, or `SIZE_MAX`
// when no match is found.
//
// For `BinaryAsc` / `BinaryDesc`, the caller must guarantee the array is
// sorted in the implied direction; we do not validate. Cross-type compares
// count as non-match (same accepted divergence the linear-mode MATCH path
// documents). Wildcards are only honoured when `match_mode ==
// XMatchMode::Wildcard`; in every other mode `*` / `?` in the pattern are
// treated as literal characters.
std::size_t xlookup_scan(const std::vector<Value>& cells, const Value& lookup_value, XMatchMode match_mode,
                         XSearchMode search_mode) {
  const std::size_t n = cells.size();
  if (n == 0) {
    return SIZE_MAX;
  }
  const bool wildcards = match_mode == XMatchMode::Wildcard;

  // Linear paths (forward / reverse). For Exact / Wildcard we return on the
  // first hit; for Smaller / Larger we track the closest candidate across
  // the entire scan (Excel does not require the caller's array to be
  // sorted) and still short-circuit on an exact hit.
  if (search_mode == XSearchMode::FirstToLast || search_mode == XSearchMode::LastToFirst) {
    const bool reverse = search_mode == XSearchMode::LastToFirst;
    auto idx_at = [&](std::size_t k) -> std::size_t { return reverse ? (n - 1U - k) : k; };

    if (match_mode == XMatchMode::Exact || match_mode == XMatchMode::Wildcard) {
      for (std::size_t k = 0; k < n; ++k) {
        const std::size_t i = idx_at(k);
        if (xlookup_exact_eq(cells[i], lookup_value, wildcards)) {
          return i;
        }
      }
      return SIZE_MAX;
    }

    // Smaller / Larger: walk the whole axis, track the closest candidate
    // under the mode's ordering, short-circuit on any exact hit. `cmp`
    // only gives us the sign vs. lookup, so to rank two candidate cells
    // against each other we re-use `xlookup_cmp` on the pair directly.
    std::size_t best = SIZE_MAX;
    for (std::size_t k = 0; k < n; ++k) {
      const std::size_t i = idx_at(k);
      int cmp = 0;
      if (!xlookup_cmp(cells[i], lookup_value, &cmp)) {
        continue;
      }
      if (cmp == 0) {
        return i;  // exact hit always wins
      }
      if (match_mode == XMatchMode::Smaller) {
        if (cmp >= 0) {
          continue;  // cell > lookup, not a candidate
        }
        if (best == SIZE_MAX) {
          best = i;
          continue;
        }
        int pair_cmp = 0;
        if (xlookup_cmp(cells[i], cells[best], &pair_cmp) && pair_cmp > 0) {
          // New candidate is larger than the current best (and still <= lookup).
          best = i;
        }
      } else {
        if (cmp <= 0) {
          continue;  // cell < lookup, not a candidate
        }
        if (best == SIZE_MAX) {
          best = i;
          continue;
        }
        int pair_cmp = 0;
        if (xlookup_cmp(cells[i], cells[best], &pair_cmp) && pair_cmp < 0) {
          // New candidate is smaller than the current best (and still >= lookup).
          best = i;
        }
      }
    }
    return best;
  }

  // Binary search. The caller promises the array is sorted. We bound-search
  // with a lower-bound-style loop on the ordering implied by the mode.
  // Incomparable (cross-type) cells are treated as "greater than" the
  // lookup so the search keeps narrowing down — this preserves the
  // invariant that a pure-kind prefix will still be found when the array
  // is clean.
  const bool descending = search_mode == XSearchMode::BinaryDesc;
  std::size_t lo = 0;
  std::size_t hi = n;  // half-open
  while (lo < hi) {
    const std::size_t mid = lo + ((hi - lo) / 2U);
    int cmp = 0;  // sign of (cells[mid] - lookup)
    if (!xlookup_cmp(cells[mid], lookup_value, &cmp)) {
      // Treat incomparable as "greater than" the lookup.
      cmp = 1;
    }
    if (descending) {
      cmp = -cmp;
    }
    if (cmp < 0) {
      lo = mid + 1U;
    } else {
      hi = mid;
    }
  }

  // `lo` is now the first index whose (adjusted) cell >= lookup.
  if (match_mode == XMatchMode::Exact || match_mode == XMatchMode::Wildcard) {
    if (lo < n && xlookup_exact_eq(cells[lo], lookup_value, wildcards)) {
      return lo;
    }
    return SIZE_MAX;
  }
  if (match_mode == XMatchMode::Smaller) {
    // Exact hit at `lo` wins; otherwise fall back to the cell just below it
    // (in the original ordering that's `lo - 1` for ascending, `lo - 1`
    // still for descending because we flipped the compare sign).
    if (lo < n) {
      int cmp = 0;
      if (xlookup_cmp(cells[lo], lookup_value, &cmp) && cmp == 0) {
        return lo;
      }
    }
    if (lo == 0) {
      return SIZE_MAX;
    }
    return lo - 1U;
  }
  // Larger: `lo` is already the first cell >= lookup in the adjusted
  // ordering. In ascending order that IS the first cell >= lookup; in
  // descending order it's the first cell <= lookup, which per the Desc
  // contract is the next-larger candidate (values decrease with index, so
  // "next larger" sits at a smaller index and was walked past — this means
  // for Desc + Larger we actually want `lo - 1` unless `cells[lo]` is an
  // exact hit).
  if (lo < n) {
    int cmp = 0;
    if (xlookup_cmp(cells[lo], lookup_value, &cmp) && cmp == 0) {
      return lo;
    }
  }
  if (descending) {
    if (lo == 0) {
      return SIZE_MAX;
    }
    return lo - 1U;
  }
  if (lo >= n) {
    return SIZE_MAX;
  }
  return lo;
}

// Coerce `v` to an int that must lie in the caller's whitelist. Truncates
// toward zero (matching Excel's MATCH coercion). Returns true when the
// coerced value was accepted; on rejection leaves `*out` untouched.
bool coerce_mode_int(const Value& v, const int* allowed, std::size_t n_allowed, int* out) {
  auto num = coerce_to_number(v);
  if (!num) {
    return false;
  }
  const double raw = std::trunc(num.value());
  const int as_int = static_cast<int>(raw);
  if (static_cast<double>(as_int) != raw) {
    return false;
  }
  for (std::size_t i = 0; i < n_allowed; ++i) {
    if (allowed[i] == as_int) {
      *out = as_int;
      return true;
    }
  }
  return false;
}

/// XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found],
///        [match_mode], [search_mode])
///
/// Modern replacement for VLOOKUP / HLOOKUP. Scans `lookup_array` (a 1-D
/// row or column range) for `lookup_value` using the specified match and
/// search modes, then returns the corresponding cell from `return_array`.
///
/// `if_not_found` is evaluated only when the scan misses (short-circuit),
/// matching Excel's lazy behaviour so `XLOOKUP(...,1/0)` does not raise
/// `#DIV/0!` on a hit. Invalid `match_mode` / `search_mode` codes yield
/// `#VALUE!`.
///
/// Accepted divergence: a 2-D `return_array` (e.g. `XLOOKUP(x, A1:A5,
/// B1:D5)` which Excel spills as a row) is not supported yet — scalar
/// evaluation only. We surface `#VALUE!` for that shape; full spill support
/// lands with dynamic arrays.
Value eval_xlookup_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 3U || arity > 6U) {
    return Value::error(ErrorCode::Value);
  }

  // 1) lookup_value — errors propagate; Blank acts as numeric 0 via the
  //    comparison primitives.
  const Value lookup = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (lookup.is_error()) {
    return lookup;
  }

  // 2) lookup_array — must be a 1-D range.
  std::vector<Value> lookup_cells;
  std::uint32_t l_rows = 0;
  std::uint32_t l_cols = 0;
  ErrorCode l_err = ErrorCode::Value;
  if (!resolve_range_arg(call.as_call_arg(1), arena, registry, ctx, &lookup_cells, &l_err, &l_rows, &l_cols)) {
    return Value::error(l_err);
  }
  if (l_rows != 1U && l_cols != 1U) {
    return Value::error(ErrorCode::Value);
  }
  if (lookup_cells.empty()) {
    return Value::error(ErrorCode::NA);
  }

  // 3) return_array — shape must be compatible with the match axis. Any
  //    mismatch or 2-D return (which would require spill) yields #VALUE!.
  std::vector<Value> return_cells;
  std::uint32_t r_rows = 0;
  std::uint32_t r_cols = 0;
  ErrorCode r_err = ErrorCode::Value;
  if (!resolve_range_arg(call.as_call_arg(2), arena, registry, ctx, &return_cells, &r_err, &r_rows, &r_cols)) {
    return Value::error(r_err);
  }
  if (r_rows != 1U && r_cols != 1U) {
    // 2-D return_array implies a spill result; scalar context only.
    return Value::error(ErrorCode::Value);
  }
  if (return_cells.size() != lookup_cells.size()) {
    return Value::error(ErrorCode::Value);
  }

  // 5) match_mode (optional, default Exact). Validate the whitelist before
  //    casting into the enum.
  static constexpr int kMatchModes[] = {-1, 0, 1, 2};
  int match_raw = 0;
  if (arity >= 5U) {
    const Value mm_val = eval_node(call.as_call_arg(4), arena, registry, ctx);
    if (mm_val.is_error()) {
      return mm_val;
    }
    if (!coerce_mode_int(mm_val, kMatchModes, sizeof(kMatchModes) / sizeof(kMatchModes[0]), &match_raw)) {
      return Value::error(ErrorCode::Value);
    }
  }

  // 6) search_mode (optional, default FirstToLast).
  static constexpr int kSearchModes[] = {-2, -1, 1, 2};
  int search_raw = 1;
  if (arity >= 6U) {
    const Value sm_val = eval_node(call.as_call_arg(5), arena, registry, ctx);
    if (sm_val.is_error()) {
      return sm_val;
    }
    if (!coerce_mode_int(sm_val, kSearchModes, sizeof(kSearchModes) / sizeof(kSearchModes[0]), &search_raw)) {
      return Value::error(ErrorCode::Value);
    }
  }

  const auto match_mode = static_cast<XMatchMode>(match_raw);
  const auto search_mode = static_cast<XSearchMode>(search_raw);
  const std::size_t off = xlookup_scan(lookup_cells, lookup, match_mode, search_mode);

  // 7) No match: evaluate if_not_found lazily, else #N/A.
  if (off == SIZE_MAX) {
    if (arity >= 4U) {
      return eval_node(call.as_call_arg(3), arena, registry, ctx);
    }
    return Value::error(ErrorCode::NA);
  }

  // 8) Translate offset -> return cell. With both arrays constrained to
  //    1-D of equal length, flat indexing works for row-into-column,
  //    column-into-row, or same-orientation cases alike.
  return return_cells[off];
}

/// XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])
///
/// Modern replacement for MATCH. Returns the 1-based position of
/// `lookup_value` in `lookup_array`, or `#N/A` when no match is found.
/// Share the `xlookup_scan` machinery with XLOOKUP; validation of match /
/// search mode codes is identical (invalid codes -> `#VALUE!`).
Value eval_xmatch_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 4U) {
    return Value::error(ErrorCode::Value);
  }

  const Value lookup = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (lookup.is_error()) {
    return lookup;
  }

  std::vector<Value> cells;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  ErrorCode range_err = ErrorCode::Value;
  if (!resolve_range_arg(call.as_call_arg(1), arena, registry, ctx, &cells, &range_err, &rows, &cols)) {
    return Value::error(range_err);
  }
  if (rows != 1U && cols != 1U) {
    return Value::error(ErrorCode::Value);
  }
  if (cells.empty()) {
    return Value::error(ErrorCode::NA);
  }

  static constexpr int kMatchModes[] = {-1, 0, 1, 2};
  int match_raw = 0;
  if (arity >= 3U) {
    const Value mm_val = eval_node(call.as_call_arg(2), arena, registry, ctx);
    if (mm_val.is_error()) {
      return mm_val;
    }
    if (!coerce_mode_int(mm_val, kMatchModes, sizeof(kMatchModes) / sizeof(kMatchModes[0]), &match_raw)) {
      return Value::error(ErrorCode::Value);
    }
  }

  static constexpr int kSearchModes[] = {-2, -1, 1, 2};
  int search_raw = 1;
  if (arity >= 4U) {
    const Value sm_val = eval_node(call.as_call_arg(3), arena, registry, ctx);
    if (sm_val.is_error()) {
      return sm_val;
    }
    if (!coerce_mode_int(sm_val, kSearchModes, sizeof(kSearchModes) / sizeof(kSearchModes[0]), &search_raw)) {
      return Value::error(ErrorCode::Value);
    }
  }

  const auto match_mode = static_cast<XMatchMode>(match_raw);
  const auto search_mode = static_cast<XSearchMode>(search_raw);
  const std::size_t off = xlookup_scan(cells, lookup, match_mode, search_mode);
  if (off == SIZE_MAX) {
    return Value::error(ErrorCode::NA);
  }
  return Value::number(static_cast<double>(off + 1U));
}

// `LazyImpl` is declared in `eval/lazy_impls.h` so translation units that
// own individual lazy impls can publish matching function pointers.
struct LazyEntry {
  const char* name;  // canonical UPPERCASE
  LazyImpl impl;
};

constexpr LazyEntry kLazyDispatch[] = {
    {"AVERAGEIF", &eval_averageif_lazy},
    {"AVERAGEIFS", &eval_averageifs_lazy},
    {"CHOOSE", &eval_choose_lazy},
    {"COUNTIF", &eval_countif_lazy},
    {"COUNTIFS", &eval_countifs_lazy},
    {"HLOOKUP", &eval_hlookup_lazy},
    {"IF", &eval_if_lazy},
    {"IFERROR", &eval_iferror_lazy},
    {"IFNA", &eval_ifna_lazy},
    {"INDEX", &eval_index_lazy},
    {"MATCH", &eval_match_lazy},
    {"MAXIFS", &eval_maxifs_lazy},
    {"MINIFS", &eval_minifs_lazy},
    {"SUMIF", &eval_sumif_lazy},
    {"SUMIFS", &eval_sumifs_lazy},
    {"VLOOKUP", &eval_vlookup_lazy},
    {"XLOOKUP", &eval_xlookup_lazy},
    {"XMATCH", &eval_xmatch_lazy},
};

const LazyEntry* find_lazy(std::string_view name) noexcept {
  for (const auto& e : kLazyDispatch) {
    if (strings::case_insensitive_eq(name, std::string_view(e.name))) {
      return &e;
    }
  }
  return nullptr;
}

// Special-cased function-call dispatch.
//
// Lazy entries (`IF`, `IFERROR`, `IFNA`, the `*IF`/`*IFS` aggregators) are
// routed through the table above;
// each impl owns its own arity check and chooses which subtrees to evaluate.
//
// All other names are routed through `registry`: unknown name -> #NAME?,
// arity violation -> #VALUE!, otherwise every argument is pre-evaluated in
// order. By default the left-most error short-circuits before the impl
// runs, but an entry whose `propagate_errors` flag is `false` (the IS*
// type-predicate family) opts out of that short-circuit and receives raw
// error values among its arguments.
Value dispatch_call(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx) {
  const std::string_view name = node.as_call_name();
  const std::uint32_t arity = node.as_call_arity();

  if (const LazyEntry* lazy = find_lazy(name); lazy != nullptr) {
    return lazy->impl(node, arena, registry, ctx);
  }

  const FunctionDef* def = registry.lookup(name);
  if (def == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  // The pre-expansion arity guards min_arity / max_arity. This happens to
  // align with Excel's behaviour for the range-aware aggregators:
  // `=SUM()` is rejected at parse time, and `=SUM(A1:A1)` passes the
  // `min_arity = 1` check even though its expansion might be empty (which
  // cannot happen with a finite valid rectangle today).
  if (arity < def->min_arity || arity > def->max_arity) {
    return Value::error(ErrorCode::Value);
  }

  // Pre-evaluate arguments left-to-right. By default the first error wins
  // and the impl is never invoked; functions that need to inspect error
  // arguments (e.g. `ISERROR`) clear `propagate_errors` to opt out. When
  // the function is range-aware (`accepts_ranges`), any argument whose AST
  // node is a simple RangeOp (Ref:Ref) is flattened into the values vector
  // in row-major order.
  std::vector<Value> values;
  values.reserve(arity);
  for (std::uint32_t i = 0; i < arity; ++i) {
    const parser::AstNode& arg_node = node.as_call_arg(i);
    if (def->accepts_ranges && arg_node.kind() == parser::NodeKind::RangeOp) {
      const parser::AstNode& lhs_ast = arg_node.as_range_lhs();
      const parser::AstNode& rhs_ast = arg_node.as_range_rhs();
      // Only the simplest form -- literal A1:B2 where both endpoints are
      // Refs -- is expanded. Anything else (INDIRECT(...), named ranges,
      // etc.) surfaces as #REF! here; full dynamic range resolution is
      // deferred.
      if (lhs_ast.kind() != parser::NodeKind::Ref || rhs_ast.kind() != parser::NodeKind::Ref) {
        const Value err = Value::error(ErrorCode::Ref);
        if (def->propagate_errors) {
          return err;
        }
        values.push_back(err);
        continue;
      }
      auto expanded = ctx.expand_range(lhs_ast.as_ref(), rhs_ast.as_ref(), arena, registry);
      if (!expanded) {
        const Value err = Value::error(expanded.error());
        if (def->propagate_errors) {
          return err;
        }
        values.push_back(err);
        continue;
      }
      for (const Value& v : expanded.value()) {
        if (def->propagate_errors && v.is_error()) {
          return v;
        }
        values.push_back(v);
      }
      continue;
    }
    Value v = eval_node(arg_node, arena, registry, ctx);
    if (def->propagate_errors && v.is_error()) {
      return v;
    }
    values.push_back(v);
  }
  // Hand the post-expansion size to the impl; aggregator bodies walk the
  // flattened vector directly.
  return def->impl(values.data(), static_cast<std::uint32_t>(values.size()), arena);
}

}  // namespace

// Defined with external linkage (declared in `eval/lazy_impls.h`) so the
// per-family lazy-impl TUs can recurse into the evaluator. The helpers it
// calls below — `apply_unary`, `apply_arithmetic`, `apply_concat`,
// `apply_comparison`, `dispatch_call` — live in the anonymous namespace
// above and remain reachable via ordinary unqualified lookup because that
// anonymous namespace is nested inside `formulon::eval`.
Value eval_node(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx) {
  switch (node.kind()) {
    case parser::NodeKind::Literal:
      return node.as_literal();

    case parser::NodeKind::ErrorLiteral:
      return Value::error(node.as_error_literal());

    case parser::NodeKind::ErrorPlaceholder:
      // Panic-mode skipped this subtree at parse time; we cannot do better
      // than #NAME? since the original tokens are unavailable.
      return Value::error(ErrorCode::Name);

    case parser::NodeKind::ImplicitIntersection:
      // Identity for scalars. Once arrays land this becomes the contraction
      // operator (1x1 selection from a column / row at the call site).
      return eval_node(node.as_implicit_intersection_operand(), arena, registry, ctx);

    case parser::NodeKind::UnaryOp:
      return apply_unary(node.as_unary_op(), eval_node(node.as_unary_operand(), arena, registry, ctx));

    case parser::NodeKind::BinaryOp: {
      const parser::BinOp op = node.as_binary_op();
      // Evaluate left first so error propagation honours the documented
      // left-most-wins rule from backup/plans/02-calc-engine.md §2.1.1.
      const Value lhs = eval_node(node.as_binary_lhs(), arena, registry, ctx);
      if (lhs.is_error()) {
        return lhs;
      }
      const Value rhs = eval_node(node.as_binary_rhs(), arena, registry, ctx);
      if (rhs.is_error()) {
        return rhs;
      }

      switch (op) {
        case parser::BinOp::Add:
        case parser::BinOp::Sub:
        case parser::BinOp::Mul:
        case parser::BinOp::Div:
        case parser::BinOp::Pow: {
          auto lhs_n = coerce_to_number(lhs);
          if (!lhs_n) {
            return Value::error(lhs_n.error());
          }
          auto rhs_n = coerce_to_number(rhs);
          if (!rhs_n) {
            return Value::error(rhs_n.error());
          }
          return apply_arithmetic(op, lhs_n.value(), rhs_n.value());
        }
        case parser::BinOp::Concat:
          return apply_concat(lhs, rhs, arena);
        case parser::BinOp::Eq:
        case parser::BinOp::NotEq:
        case parser::BinOp::Lt:
        case parser::BinOp::LtEq:
        case parser::BinOp::Gt:
        case parser::BinOp::GtEq:
          return apply_comparison(op, lhs, rhs);
      }
      return Value::error(ErrorCode::Value);
    }

    case parser::NodeKind::Call:
      return dispatch_call(node, arena, registry, ctx);

    case parser::NodeKind::Ref:
      return ctx.resolve_ref(node.as_ref(), arena, registry);

    // -- Unsupported: name resolution / closures --------------------------
    case parser::NodeKind::ExternalRef:
    case parser::NodeKind::StructuredRef:
    case parser::NodeKind::NameRef:
    case parser::NodeKind::LambdaCall:
    case parser::NodeKind::Lambda:
    case parser::NodeKind::LetBinding:
      return Value::error(ErrorCode::Name);

    // -- Unsupported: range-producing operators / array literals ----------
    case parser::NodeKind::RangeOp:
    case parser::NodeKind::UnionOp:
    case parser::NodeKind::IntersectOp:
    case parser::NodeKind::ArrayLiteral:
      return Value::error(ErrorCode::Value);
  }
  return Value::error(ErrorCode::Value);
}

Value evaluate(const parser::AstNode& node, Arena& arena) {
  return evaluate(node, arena, default_registry(), EvalContext{});
}

Value evaluate(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry) {
  return evaluate(node, arena, registry, EvalContext{});
}

Value evaluate(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx) {
  return eval_node(node, arena, registry, ctx);
}

}  // namespace eval
}  // namespace formulon
