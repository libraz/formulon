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
#include "eval/dynamic_array/common.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/jp_fold.h"
#include "eval/lazy_impls.h"
#include "eval/omitted_arg.h"
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
// validate the sort order — neither do we) and drive an upper-bound
// style search so duplicates resolve to the LAST matching index.
enum class XSearchMode : std::int8_t { FirstToLast = 1, LastToFirst = -1, BinaryAsc = 2, BinaryDesc = -2 };

// Three-way compare of (cell - lookup) honouring Excel 365's cross-type
// ordering for XLOOKUP approximate-match: kind-rank first, value second.
// Rank 0 = Number / Blank (Blank acts as 0.0), Rank 1 = Text (case-insensitive
// ASCII compare), Rank 2 = Bool (FALSE < TRUE). Anything else (Error, Array,
// Ref, Lambda) is incomparable and the function returns `false` so callers
// can decide how to handle it. On success writes the sign to `*out_cmp` and
// returns `true`.
bool xlookup_cmp(const Value& cell, const Value& lookup, ExcelProfile profile, int* out_cmp) {
  auto cmp_numeric = [](double x, double y) -> int {
    if (x < y) {
      return -1;
    }
    if (x > y) {
      return 1;
    }
    return 0;
  };
  auto rank_of = [](const Value& v, int* rank) -> bool {
    if (v.is_number() || v.is_blank()) {
      *rank = 0;
      return true;
    }
    if (v.is_text()) {
      *rank = 1;
      return true;
    }
    if (v.is_boolean()) {
      *rank = 2;
      return true;
    }
    return false;
  };
  int cell_rank = 0;
  int lookup_rank = 0;
  if (!rank_of(cell, &cell_rank) || !rank_of(lookup, &lookup_rank)) {
    return false;
  }
  if (cell_rank != lookup_rank) {
    *out_cmp = cmp_numeric(static_cast<double>(cell_rank), static_cast<double>(lookup_rank));
    return true;
  }
  // Same kind — compare values directly.
  if (cell_rank == 1) {
    // ja-JP fold (see classic.cpp::lookup_scan) before the ASCII
    // case-insensitive compare so kana variants order together in
    // XLOOKUP / XMATCH approximate paths. Full-width digits are NOT
    // folded for lookups (Mac asymmetry — see jp_fold.h).
    if (uses_mac_jp_text_folding(profile)) {
      *out_cmp = strings::case_insensitive_compare(fold_jp_text(cell.as_text(), /*fold_fullwidth_digits=*/false),
                                                   fold_jp_text(lookup.as_text(), /*fold_fullwidth_digits=*/false));
    } else {
      *out_cmp = strings::case_insensitive_compare(cell.as_text(), lookup.as_text());
    }
    return true;
  }
  if (cell_rank == 2) {
    const int lb = lookup.as_boolean() ? 1 : 0;
    const int cb = cell.as_boolean() ? 1 : 0;
    *out_cmp = cmp_numeric(cb, lb);
    return true;
  }
  // Rank 0: Number or Blank.
  const double lv = lookup.is_blank() ? 0.0 : lookup.as_number();
  const double cv = cell.is_blank() ? 0.0 : cell.as_number();
  *out_cmp = cmp_numeric(cv, lv);
  return true;
}

// Exact-equality test with optional DOS-style wildcard expansion on Text vs
// Text pairs. For non-text / cross-type kinds this is a literal equality
// compare with the same "Blank acts as numeric 0" rule MATCH / VLOOKUP use.
// When `wildcards` is true AND lookup_value is text, pattern matching is
// always routed through `wildcard_match` so that `~*` continues to mean
// "literal *". See the VLOOKUP comment at the top of `lookup_scan` for why
// we deliberately do not gate on `scan_has_wildcard`.
bool xlookup_exact_eq(const Value& cell, const Value& lookup, bool wildcards, ExcelProfile profile) {
  if (lookup.is_text()) {
    if (!cell.is_text()) {
      return false;
    }
    // ja-JP fold (see classic.cpp::lookup_scan) before lower-casing so
    // XLOOKUP / XMATCH exact mode treats kana / full-width variants
    // identically to Mac Excel. Full-width digits are NOT folded for
    // lookups (Mac asymmetry — see jp_fold.h).
    //
    // The Windows path keeps the broad kana fold OFF (Windows Excel does
    // not fold hiragana<->katakana or plain half/full width in lookups —
    // see lookup_kana_folding_probes) but still composes a half-width
    // voicing mark onto its base (`ｶﾞ` -> `ガ`): a base + standalone ﾞ /
    // ﾟ is a malformed encoding both Windows and Mac compose before the
    // exact-match compare.
    const bool jp_fold = uses_mac_jp_text_folding(profile);
    const std::string pat_lower = jp_fold ? fold_and_lower(lookup.as_text(), /*fold_fullwidth_digits=*/false)
                                          : strings::to_ascii_lower(compose_jp_halfwidth_voicing(lookup.as_text()));
    const std::string cell_lower = jp_fold ? fold_and_lower(cell.as_text(), /*fold_fullwidth_digits=*/false)
                                           : strings::to_ascii_lower(compose_jp_halfwidth_voicing(cell.as_text()));
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
// follow Excel 365's kind-rank ordering (see `xlookup_cmp`): Number/Blank <
// Text < Bool. Wildcards are only honoured when `match_mode ==
// XMatchMode::Wildcard`; in every other mode `*` / `?` in the pattern are
// treated as literal characters.
std::size_t xlookup_scan(const std::vector<Value>& cells, const Value& lookup_value, XMatchMode match_mode,
                         XSearchMode search_mode, ExcelProfile profile) {
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
        if (xlookup_exact_eq(cells[i], lookup_value, wildcards, profile)) {
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
      if (!xlookup_cmp(cells[i], lookup_value, profile, &cmp)) {
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
        if (xlookup_cmp(cells[i], cells[best], profile, &pair_cmp) && pair_cmp > 0) {
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
        if (xlookup_cmp(cells[i], cells[best], profile, &pair_cmp) && pair_cmp < 0) {
          // New candidate is smaller than the current best (and still >= lookup).
          best = i;
        }
      }
    }
    return best;
  }

  // Binary search. The caller promises the array is sorted in the direction
  // implied by `search_mode`. We run a single upper-bound pass in the
  // effective (flipped-for-desc) view: `hi_eff` becomes the first index
  // whose effective cell > target. Cells that are incomparable with the
  // lookup (Error / Array / Ref / Lambda) are treated as > target so the
  // search continues to narrow down; with all-clean arrays this is a
  // no-op and with a foreign cell it keeps the upper-bound monotone.
  const bool descending = search_mode == XSearchMode::BinaryDesc;
  std::size_t hi_eff = 0;
  {
    std::size_t lo = 0;
    std::size_t hi = n;  // half-open
    while (lo < hi) {
      const std::size_t mid = lo + ((hi - lo) / 2U);
      int cmp = 0;
      const bool cmp_ok = xlookup_cmp(cells[mid], lookup_value, profile, &cmp);
      if (!cmp_ok) {
        // Incomparable — treat as greater than target.
        cmp = 1;
      }
      if (descending) {
        cmp = -cmp;
      }
      // Upper-bound: advance `lo` while effective cell <= target.
      if (cmp <= 0) {
        lo = mid + 1U;
      } else {
        hi = mid;
      }
    }
    hi_eff = lo;
  }

  // Exact / Wildcard: candidate is the last cell with effective value <=
  // target, i.e. `hi_eff - 1`. Run the literal equality test; anything
  // else is a miss.
  if (match_mode == XMatchMode::Exact || match_mode == XMatchMode::Wildcard) {
    if (hi_eff == 0) {
      return SIZE_MAX;
    }
    if (xlookup_exact_eq(cells[hi_eff - 1U], lookup_value, wildcards, profile)) {
      return hi_eff - 1U;
    }
    return SIZE_MAX;
  }

  // Smaller / Larger: short-circuit on an exact hit at `hi_eff - 1` so the
  // LAST duplicate wins (matches Excel 365).
  if (hi_eff > 0) {
    int cmp = 0;
    if (xlookup_cmp(cells[hi_eff - 1U], lookup_value, profile, &cmp) && cmp == 0) {
      return hi_eff - 1U;
    }
  }

  if (match_mode == XMatchMode::Smaller) {
    if (!descending) {
      // Ascending with no exact match: the largest cell < target sits at
      // `hi_eff - 1` (lower_bound and upper_bound coincide in this case).
      if (hi_eff == 0) {
        return SIZE_MAX;
      }
      return hi_eff - 1U;
    }
    // Descending with no exact match: cells at `hi_eff..n-1` are < target
    // (in real terms); values decrease with index, so the LARGEST of those
    // sits at position `hi_eff`.
    if (hi_eff >= n) {
      return SIZE_MAX;
    }
    return hi_eff;
  }

  // Larger, no exact.
  if (!descending) {
    // Ascending: the smallest cell > target is at `hi_eff`.
    if (hi_eff >= n) {
      return SIZE_MAX;
    }
    return hi_eff;
  }
  // Descending: cells at `0..hi_eff-1` are > target (in real terms); values
  // decrease with index, so the SMALLEST of those sits at `hi_eff - 1`.
  if (hi_eff == 0) {
    return SIZE_MAX;
  }
  return hi_eff - 1U;
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
/// A multi-column `return_array` (vertical lookup, e.g. `XLOOKUP(x, A1:A5,
/// B1:D5)`) spills the matched row across all return columns; a multi-row
/// `return_array` (horizontal lookup) spills the matched column. The
/// return_array's match axis must equal the lookup axis length, else
/// `#VALUE!`.
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
  auto lookup_resolved = resolve_range_arg(call.as_call_arg(1), arena, registry, ctx);
  if (!lookup_resolved) {
    return Value::error(lookup_resolved.error());
  }
  const std::uint32_t l_rows = lookup_resolved.value().rows;
  const std::uint32_t l_cols = lookup_resolved.value().cols;
  std::vector<Value> lookup_cells = std::move(lookup_resolved.value().cells);
  if (l_rows != 1U && l_cols != 1U) {
    return Value::error(ErrorCode::Value);
  }
  if (lookup_cells.empty()) {
    return Value::error(ErrorCode::NA);
  }

  // 3) return_array — its match axis must equal the lookup axis length.
  //    The lookup_array is 1-D; a vertical lookup (l_cols == 1) matches
  //    rows of return_array and may carry multiple columns (the matched
  //    row spills horizontally), a horizontal lookup (l_rows == 1) matches
  //    columns and may carry multiple rows (the matched column spills
  //    vertically). A 1x1 lookup defaults to the vertical convention.
  auto return_resolved = resolve_range_arg(call.as_call_arg(2), arena, registry, ctx);
  if (!return_resolved) {
    return Value::error(return_resolved.error());
  }
  const std::uint32_t r_rows = return_resolved.value().rows;
  const std::uint32_t r_cols = return_resolved.value().cols;
  std::vector<Value> return_cells = std::move(return_resolved.value().cells);
  const bool horizontal_lookup = l_rows == 1U && l_cols != 1U;
  // The number of result lanes parallel to the lookup axis, and the width
  // of the slice returned per match.
  const std::uint32_t lookup_len = static_cast<std::uint32_t>(lookup_cells.size());
  const std::uint32_t match_extent = horizontal_lookup ? r_cols : r_rows;
  const std::uint32_t slice_extent = horizontal_lookup ? r_rows : r_cols;
  if (match_extent != lookup_len) {
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

  // Excel 365 rejects the Wildcard (+2) match_mode combined with either
  // Binary search_mode (±2) as #VALUE!: binary search has no defined
  // meaning when the pattern contains `*` / `?` metacharacters.
  if (match_mode == XMatchMode::Wildcard &&
      (search_mode == XSearchMode::BinaryAsc || search_mode == XSearchMode::BinaryDesc)) {
    return Value::error(ErrorCode::Value);
  }

  const std::size_t off = xlookup_scan(lookup_cells, lookup, match_mode, search_mode, ctx.excel_profile());

  // 7) No match: evaluate if_not_found lazily, else #N/A. A blank-literal
  //    slot (`XLOOKUP(..., ..., ..., , , -1)`) is treated as "no
  //    if_not_found supplied" per Excel 365 — the parser injects a
  //    Blank literal for the empty comma slot, but the function should
  //    fall through to #N/A rather than returning that blank.
  if (off == SIZE_MAX) {
    if (arity >= 4U) {
      const parser::AstNode& if_arg = call.as_call_arg(3);
      const bool is_empty_slot = is_omitted_arg(if_arg);
      if (!is_empty_slot) {
        return eval_node(if_arg, arena, registry, ctx);
      }
    }
    return Value::error(ErrorCode::NA);
  }

  // 8) Translate offset -> return slice. For a single return lane the
  //    result is one cell; for a multi-lane return_array the matched
  //    row (vertical lookup) or column (horizontal lookup) spills as an
  //    array. `return_cells` is row-major (r_rows x r_cols).
  if (slice_extent == 1U) {
    // Single-cell result. `off` indexes the match axis; with a 1-lane
    // return_array the flat index is `off` regardless of orientation.
    return return_cells[off];
  }
  if (horizontal_lookup) {
    // Matched column `off`: gather every row at that column into an
    // r_rows x 1 vertical array.
    Value* buffer = nullptr;
    ArrayValue* out = dynamic_array::allocate_array_value(r_rows, 1U, arena, buffer);
    if (out == nullptr) {
      return Value::error(ErrorCode::Num);
    }
    for (std::uint32_t row = 0; row < r_rows; ++row) {
      buffer[row] = return_cells[(static_cast<std::size_t>(row) * r_cols) + off];
    }
    return Value::array(out);
  }
  // Vertical lookup, matched row `off`: gather every column at that row
  // into a 1 x r_cols horizontal array.
  Value* buffer = nullptr;
  ArrayValue* out = dynamic_array::allocate_array_value(1U, r_cols, arena, buffer);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::uint32_t col = 0; col < r_cols; ++col) {
    buffer[col] = return_cells[(static_cast<std::size_t>(off) * r_cols) + col];
  }
  return Value::array(out);
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

  auto resolved = resolve_range_arg(call.as_call_arg(1), arena, registry, ctx);
  if (!resolved) {
    return Value::error(resolved.error());
  }
  const std::uint32_t rows = resolved.value().rows;
  const std::uint32_t cols = resolved.value().cols;
  std::vector<Value> cells = std::move(resolved.value().cells);
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

  // Excel 365 rejects the Wildcard (+2) match_mode combined with either
  // Binary search_mode (±2) as #VALUE!: binary search has no defined
  // meaning when the pattern contains `*` / `?` metacharacters.
  if (match_mode == XMatchMode::Wildcard &&
      (search_mode == XSearchMode::BinaryAsc || search_mode == XSearchMode::BinaryDesc)) {
    return Value::error(ErrorCode::Value);
  }

  const std::size_t off = xlookup_scan(cells, lookup, match_mode, search_mode, ctx.excel_profile());
  if (off == SIZE_MAX) {
    return Value::error(ErrorCode::NA);
  }
  return Value::number(static_cast<double>(off + 1U));
}

}  // namespace eval
}  // namespace formulon
