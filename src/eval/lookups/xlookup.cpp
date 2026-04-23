// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the XLOOKUP / XMATCH family of lazy impls. See
// `lookups/xlookup.h` for the dispatch-table contract and
// `eval/lazy_impls.h` for the shared vocabulary.

#include "eval/lookups/xlookup.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "eval/coerce.h"
#include "eval/criteria.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/lazy_impls.h"
#include "eval/range_args.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

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

}  // namespace

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

}  // namespace eval
}  // namespace formulon
