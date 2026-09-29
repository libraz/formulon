//
// Implementation of the classic lookup-family lazy impls that search by
// value (`MATCH`, `VLOOKUP`, `HLOOKUP`, `LOOKUP`); `CHOOSE` and `INDEX` live
// in `lookups/index.cpp`. See `lookups/classic.h` for the dispatch-table
// contract and `eval/lazy_impls.h` for the shared vocabulary.

#include "eval/lookups/classic.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/coerce.h"
#include "eval/criteria.h"
#include "eval/declared_rect.h"
#include "eval/dynamic_array/common.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/jp_fold.h"
#include "eval/lazy_impls.h"
#include "eval/lookups/common.h"
#include "eval/name_env_resolve.h"
#include "eval/range_args.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Lowercased, lookup-normalised form of `s` for exact / wildcard text
// matching. On the Mac ja-JP path this folds kana / full-width variants
// (`fold_and_lower`); on every other profile it still composes a half-width
// voicing mark onto its base (`ｶﾞ` -> `ガ`) before ASCII-lowercasing, matching
// XLOOKUP's `xlookup_exact_eq` so VLOOKUP / HLOOKUP / MATCH agree with it.
std::string lookup_text_key(std::string_view s, ExcelProfile profile) {
  if (uses_mac_jp_text_folding(profile)) {
    return fold_and_lower(s, /*fold_fullwidth_digits=*/false);
  }
  return strings::to_ascii_lower(compose_jp_halfwidth_voicing(s));
}

// Normalised form of `s` for the case-insensitive ordering compare used by
// approximate matching. Case folding is left to `case_insensitive_compare`,
// so this returns the composed / folded (not lowercased) form: the Mac path
// folds broadly (`fold_jp_text`); other profiles compose the half-width
// voicing mark, mirroring `lookup_text_key`.
std::string lookup_text_cmp_key(std::string_view s, ExcelProfile profile) {
  if (uses_mac_jp_text_folding(profile)) {
    return fold_jp_text(s, /*fold_fullwidth_digits=*/false);
  }
  return compose_jp_halfwidth_voicing(s);
}

// ---------------------------------------------------------------------------
// MATCH / VLOOKUP / HLOOKUP / LOOKUP (lookup & reference)
// ---------------------------------------------------------------------------

// Axis along which VLOOKUP / HLOOKUP scan their table_array rectangle for the
// lookup_value. `Column` means "walk top-down through the first column"
// (VLOOKUP); `Row` means "walk left-to-right through the first row"
// (HLOOKUP).
enum class LookupAxis : std::uint8_t { Column, Row };

// Resolve VLOOKUP / HLOOKUP's `table_array` argument with scalar fallback.
// A numeric or bool literal (e.g. `HLOOKUP(M, 3, 1)` or `HLOOKUP(M, TRUE, 1)`)
// is wrapped into a 1x1 table so the subsequent scan produces `#N/A` on
// mismatch rather than the `#VALUE!` that a strict range-only resolver
// returns. Text scalars remain `#VALUE!` per Excel (`HLOOKUP(M,"Nothing",1)`
// -> `#VALUE!`), and errors from the node evaluation propagate unchanged.
//
// The Text-vs-Number/Bool distinction is enforced AFTER `resolve_range_arg`
// returns: that helper's generic-scalar fallback wraps any scalar (including
// Text) into a 1x1 table, but Excel rejects a non-numeric/non-bool scalar
// table_array with `#VALUE!`. We re-inspect the resolved 1x1 cell when the
// caller AST is not range-shaped (Ref / RangeOp / SpillRef / OFFSET / CHOOSE /
// INDIRECT / IF) — those shapes legitimately produce a 1x1 Text cell and must
// not be downgraded to `#VALUE!`.
bool resolve_table_array(const parser::AstNode& arg_node, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                         std::uint32_t* out_rows, std::uint32_t* out_cols) {
  auto resolved = resolve_range_arg(arg_node, arena, registry, ctx);
  if (resolved) {
    auto& rr = resolved.value();
    *out_rows = rr.rows;
    *out_cols = rr.cols;
    *out_cells = std::move(rr.cells);
    // Reject scalar Text table_array: `=HLOOKUP(M,"Nothing",1)` -> `#VALUE!`.
    // A 1x1 Text cell that came from a real Ref / RangeOp / SpillRef or a
    // reference-producing call (OFFSET/CHOOSE/INDIRECT/IF) is allowed because
    // those nodes legitimately produce range-shaped values.
    if (*out_rows == 1U && *out_cols == 1U && !out_cells->empty() && out_cells->front().kind() == ValueKind::Text &&
        arg_node.kind() != parser::NodeKind::Ref && !is_range_shaped_ast(arg_node)) {
      *out_err_code = ErrorCode::Value;
      return false;
    }
    return true;
  }
  *out_err_code = resolved.error();
  if (*out_err_code != ErrorCode::Value) {
    return false;
  }
  const Value scalar = eval_node(arg_node, arena, registry, ctx);
  if (scalar.is_error()) {
    *out_err_code = scalar.as_error();
    return false;
  }
  if (scalar.kind() != ValueKind::Number && scalar.kind() != ValueKind::Bool) {
    *out_err_code = ErrorCode::Value;
    return false;
  }
  out_cells->clear();
  out_cells->push_back(scalar);
  *out_rows = 1U;
  *out_cols = 1U;
  return true;
}

// Linear scan for VLOOKUP / HLOOKUP. Walks the first column (axis=Column) or
// the first row (axis=Row) of the `flat` rectangle (rows x cols, row-major)
// for `lookup_value`, using approximate (largest <= value) or exact (first
// hit) matching.
//
// Returns the 0-based offset along the scanned axis on match, or `SIZE_MAX`
// when no match was found.
//
// In exact mode, text-vs-text matching always routes through
// `wildcard_match` so the caller doesn't need a separate "has wildcards?"
// branch. `~X` always means "literal X" and a pattern with no `*` / `?` /
// `~` degenerates to a byte-exact compare, matching Excel's rules.
//
// Cross-type comparisons (Number vs Text, Bool vs anything else) produce "no
// match" - the scanned cell is skipped. This is the same accepted divergence
// MATCH documents for its approximate path.
std::size_t lookup_scan(const std::vector<Value>& flat, std::uint32_t rows, std::uint32_t cols, LookupAxis axis,
                        const Value& lookup_value, bool approximate, ExcelProfile profile) {
  const std::size_t n = axis == LookupAxis::Column ? rows : cols;
  if (n == 0) {
    return SIZE_MAX;
  }
  // Index the i-th cell along the scan axis. For Column we walk (i, 0);
  // for Row we walk (0, i). The flat buffer is row-major so the linear
  // index is `i * cols + 0` (Column) or `0 * cols + i` (Row).
  auto cell_at = [&](std::size_t i) -> const Value& {
    const std::size_t flat_idx = axis == LookupAxis::Column ? (i * static_cast<std::size_t>(cols)) : i;
    return flat[flat_idx];
  };

  if (!approximate) {
    // Exact match: first hit wins. Text vs Text is routed through the
    // wildcard matcher unconditionally — with no metacharacters the match
    // degenerates to case-insensitive byte equality, and `~X` is always
    // treated as a literal X. Every other kind-pairing is a literal
    // equality compare.
    if (lookup_value.is_text()) {
      // Mac Excel ja-JP folds kana variants (hira<->kata, half<->full-width
      // katakana with voicing composition, full<->half-width ASCII letters
      // / punctuation / space) before text equality. Apply `fold_jp_text`
      // on both sides BEFORE ASCII-lowercasing so e.g. `ｶﾞ` -> `ガ`,
      // `Ａ` -> `a`. Full-width digits are deliberately NOT folded for
      // lookups (Mac asymmetry — see jp_fold.h).
      const std::string pat_lower = lookup_text_key(lookup_value.as_text(), profile);
      for (std::size_t i = 0; i < n; ++i) {
        const Value& cell = cell_at(i);
        if (!cell.is_text()) {
          continue;
        }
        const std::string cell_lower = lookup_text_key(cell.as_text(), profile);
        if (wildcard_match(pat_lower, cell_lower)) {
          return i;
        }
      }
      return SIZE_MAX;
    }
    if (lookup_value.is_number() || lookup_value.is_blank()) {
      const double target = lookup_value.is_blank() ? 0.0 : lookup_value.as_number();
      for (std::size_t i = 0; i < n; ++i) {
        const Value& cell = cell_at(i);
        if (cell.is_number() && cell.as_number() == target) {
          return i;
        }
        if (cell.is_blank() && target == 0.0) {
          return i;
        }
      }
      return SIZE_MAX;
    }
    if (lookup_value.is_boolean()) {
      const bool target = lookup_value.as_boolean();
      for (std::size_t i = 0; i < n; ++i) {
        const Value& cell = cell_at(i);
        if (cell.is_boolean() && cell.as_boolean() == target) {
          return i;
        }
      }
      return SIZE_MAX;
    }
    return SIZE_MAX;
  }

  // Approximate match: scan top-down (or left-right) recording the last
  // position whose value is <= lookup_value. We do NOT short-circuit when
  // a strictly-greater cell is seen — Mac Excel keeps scanning and
  // returns the running last <= match even on unsorted data. For
  // ascending-sorted input this produces the documented binary-search
  // answer because every post-match cell is strictly greater (and
  // skipped). Wildcards are NEVER honoured here (Excel treats them as
  // literal text in approximate mode).
  auto cmp_numeric = [](double x, double y) -> int {
    if (x < y) {
      return -1;
    }
    if (x > y) {
      return 1;
    }
    return 0;
  };
  // Hoisted out of the scan loop below: `lookup_value` and `profile` are
  // loop-invariant, so the text-mode branch normalises the lookup side
  // once rather than once per scanned cell.
  const std::string lookup_key =
      lookup_value.is_text() ? lookup_text_cmp_key(lookup_value.as_text(), profile) : std::string();

  std::size_t best = SIZE_MAX;
  for (std::size_t i = 0; i < n; ++i) {
    const Value& cell = cell_at(i);
    int cmp = 0;  // sign of (cell - lookup_value)
    bool comparable = false;
    if (lookup_value.is_text() && cell.is_text()) {
      // Normalise (see exact-mode branch above) before the ASCII
      // case-insensitive compare so kana / half-width voicing variants order
      // together.
      cmp = strings::case_insensitive_compare(lookup_text_cmp_key(cell.as_text(), profile), lookup_key);
      comparable = true;
    } else if ((lookup_value.is_number() || lookup_value.is_blank()) && cell.is_number()) {
      // Blank cells in the scanned axis are NOT treated as numeric 0 in
      // VLOOKUP/HLOOKUP approximate mode: Mac Excel returns `#N/A` for
      // `=HLOOKUP(<blank>, $C$1:$G$6, 1)` even when the first row contains
      // a blank cell at the matched offset, because blanks aren't part of
      // the comparable ordering. Same rule when the lookup_value itself is
      // blank — only real numeric cells participate in the approximate
      // ranking, mirroring Mac Excel's empirical behaviour.
      const double lv = lookup_value.is_blank() ? 0.0 : lookup_value.as_number();
      const double cv = cell.as_number();
      cmp = cmp_numeric(cv, lv);
      comparable = true;
    } else if (lookup_value.is_boolean() && cell.is_boolean()) {
      const int lb = lookup_value.as_boolean() ? 1 : 0;
      const int cb = cell.as_boolean() ? 1 : 0;
      cmp = cmp_numeric(cb, lb);
      comparable = true;
    }
    if (!comparable) {
      // Cross-type: skip. Accepted divergence from Excel's full ordering.
      continue;
    }
    if (cmp <= 0) {
      best = i;
    }
  }
  return best;
}

// Maps a scalar lookup result over the query array without routing through the
// eager broadcaster. The lookup family has to resolve its table / index / mode
// arguments exactly once, and an error in one query cell must remain local to
// that lane even when the shared validation has already produced an error for
// every ordinary lane.
template <typename Mapper>
Value map_lookup_query_array(const Value& lookup, Arena& arena, const Mapper& mapper) {
  const std::uint32_t query_rows = lookup.as_array_rows();
  const std::uint32_t query_cols = lookup.as_array_cols();
  Value* output_cells = nullptr;
  ArrayValue* output =
      dynamic_array::allocate_array_value(query_rows, query_cols, arena, output_cells, kMaxDerivedArrayCells);
  if (output == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  const Value* query_cells = lookup.as_array_cells();
  const std::size_t query_count = static_cast<std::size_t>(query_rows) * query_cols;
  for (std::size_t i = 0; i < query_count; ++i) {
    output_cells[i] = promote_array_result_cell(mapper(query_cells[i]));
  }
  return Value::array(output);
}

// Fails the whole call with `err`. An array lookup fails lane by lane, so a
// query cell holding its own error keeps it.
Value fail_lookup(const Value& lookup, Arena& arena, const Value& err) {
  if (lookup.is_array()) {
    return map_lookup_query_array(lookup, arena, [&err](const Value& query) { return query.is_error() ? query : err; });
  }
  return err;
}

Value eval_table_lookup_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx, LookupAxis axis) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity != 3 && arity != 4) {
    return Value::error(ErrorCode::Value);
  }

  const Value lookup = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (lookup.is_error()) {
    return lookup;
  }

  const bool array_lookup = lookup.is_array();

  std::vector<Value> cells;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  ErrorCode range_err = ErrorCode::Value;
  ReferenceTable table;
  const auto by_reference = resolve_reference_table(call.as_call_arg(1), ctx, &table);
  bool table_ok = false;
  if (!by_reference) {
    range_err = by_reference.error();
  } else if (by_reference.value()) {
    rows = table.declared.rows();
    cols = table.declared.cols();
    table_ok = true;
  } else {
    table_ok = resolve_table_array(call.as_call_arg(1), arena, registry, ctx, &cells, &range_err, &rows, &cols);
  }
  if (!table_ok) {
    return fail_lookup(lookup, arena, Value::error(range_err));
  }
  if (rows == 0U || cols == 0U) {
    return fail_lookup(lookup, arena, Value::error(ErrorCode::Ref));
  }

  const Value index_val = eval_node(call.as_call_arg(2), arena, registry, ctx);
  if (index_val.is_error()) {
    return fail_lookup(lookup, arena, index_val);
  }
  auto index_num = coerce_to_number(index_val);
  if (!index_num) {
    return fail_lookup(lookup, arena, Value::error(index_num.error()));
  }
  const double index_raw = truncate_index(index_num.value());
  if (index_raw < 1.0) {
    return fail_lookup(lookup, arena, Value::error(ErrorCode::Value));
  }
  const std::uint32_t result_extent = axis == LookupAxis::Column ? cols : rows;
  if (index_raw > static_cast<double>(result_extent)) {
    return fail_lookup(lookup, arena, Value::error(ErrorCode::Ref));
  }
  const auto result_index = static_cast<std::uint32_t>(index_raw);

  bool approximate = true;
  if (arity == 4) {
    const Value rl_val = eval_node(call.as_call_arg(3), arena, registry, ctx);
    if (rl_val.is_error()) {
      return fail_lookup(lookup, arena, rl_val);
    }
    auto rl_bool = coerce_to_bool(rl_val);
    if (!rl_bool) {
      return fail_lookup(lookup, arena, Value::error(rl_bool.error()));
    }
    approximate = rl_bool.value();
  }

  // A reference table is scanned along its walked extent -- the same cells
  // a full expansion would scan -- and only the scanned line is read.
  const std::vector<Value>* scan_cells = &cells;
  std::uint32_t scan_rows = rows;
  std::uint32_t scan_cols = cols;
  std::vector<Value> scan_line;
  if (by_reference && by_reference.value()) {
    const DeclaredRect& walked = table.walked;
    auto line = axis == LookupAxis::Column
                    ? read_table_block(table, walked.row_first, walked.row_last, table.declared.col_first,
                                       table.declared.col_first, arena, registry, ctx)
                    : read_table_block(table, table.declared.row_first, table.declared.row_first, walked.col_first,
                                       walked.col_last, arena, registry, ctx);
    if (!line) {
      return Value::error(line.error());
    }
    scan_line = std::move(line.value());
    scan_cells = &scan_line;
    scan_rows = axis == LookupAxis::Column ? walked.rows() : 1U;
    scan_cols = axis == LookupAxis::Column ? 1U : walked.cols();
  }
  const auto fetch = [&](std::size_t off) -> Value {
    if (scan_cells == &scan_line) {
      if (result_index == 1U) {
        return scan_line[off];
      }
      const auto along = static_cast<std::uint32_t>(off);
      return axis == LookupAxis::Column ? read_table_cell(table, along, result_index - 1U, arena, registry, ctx)
                                        : read_table_cell(table, result_index - 1U, along, arena, registry, ctx);
    }
    const std::size_t flat = axis == LookupAxis::Column
                                 ? (off * static_cast<std::size_t>(cols)) + static_cast<std::size_t>(result_index - 1U)
                                 : (static_cast<std::size_t>(result_index - 1U) * static_cast<std::size_t>(cols)) + off;
    if (flat >= cells.size()) {
      return Value::error(ErrorCode::Ref);
    }
    return cells[flat];
  };

  if (array_lookup) {
    return map_lookup_query_array(lookup, arena, [&](const Value& query) {
      if (query.is_error()) {
        return query;
      }
      const std::size_t off =
          lookup_scan(*scan_cells, scan_rows, scan_cols, axis, query, approximate, ctx.excel_profile());
      if (off == SIZE_MAX) {
        return Value::error(ErrorCode::NA);
      }
      return fetch(off);
    });
  }

  const std::size_t off =
      lookup_scan(*scan_cells, scan_rows, scan_cols, axis, lookup, approximate, ctx.excel_profile());
  if (off == SIZE_MAX) {
    return Value::error(ErrorCode::NA);
  }
  return fetch(off);
}

}  // namespace

Value match_lookup_one(const std::vector<Value>& cells, const Value& lookup, int match_type, ExcelProfile profile) {
  if (lookup.is_error()) {
    return lookup;
  }
  const std::size_t n = cells.size();
  if (n == 0) {
    return Value::error(ErrorCode::NA);
  }

  // Exact match (match_type == 0): honours wildcards for Text lookup,
  // case-insensitive ASCII equality otherwise. Non-matching cell kinds do
  // not count.
  if (match_type == 0) {
    if (lookup.is_text()) {
      // The wildcard matcher handles `*` / `?` as metacharacters and
      // `~X` as an escaped literal. Running it on a pattern with no real
      // wildcards is still correct — `~*` becomes a literal `*` compare,
      // `foo` becomes a byte-exact compare. Lowering both sides gives
      // Excel's case-insensitive ASCII equality.
      const std::string pat_lower = lookup_text_key(lookup.as_text(), profile);
      for (std::size_t i = 0; i < n; ++i) {
        const Value& cell = cells[i];
        if (!cell.is_text()) {
          continue;
        }
        const std::string cell_lower = lookup_text_key(cell.as_text(), profile);
        if (wildcard_match(pat_lower, cell_lower)) {
          return Value::number(static_cast<double>(i + 1));
        }
      }
      return Value::error(ErrorCode::NA);
    }
    if (lookup.is_number() || lookup.is_blank()) {
      // Blank as lookup_value never matches anything in exact mode (Excel
      // behaviour: blank search values short-circuit to #N/A even when the
      // array contains blank cells).
      if (lookup.is_blank()) {
        return Value::error(ErrorCode::NA);
      }
      const double target = lookup.as_number();
      for (std::size_t i = 0; i < n; ++i) {
        const Value& cell = cells[i];
        if (cell.is_number() && cell.as_number() == target) {
          return Value::number(static_cast<double>(i + 1));
        }
      }
      return Value::error(ErrorCode::NA);
    }
    if (lookup.is_boolean()) {
      const bool target = lookup.as_boolean();
      for (std::size_t i = 0; i < n; ++i) {
        const Value& cell = cells[i];
        if (cell.is_boolean() && cell.as_boolean() == target) {
          return Value::number(static_cast<double>(i + 1));
        }
      }
      return Value::error(ErrorCode::NA);
    }
    return Value::error(ErrorCode::NA);
  }

  // Approximate match (match_type == 1 or -1). Linear scan; we do NOT
  // honour wildcards here (Excel treats them as literals in approximate
  // mode). Cross-type comparisons are skipped (treated as non-match).
  auto cmp_numeric = [](double x, double y) -> int {
    if (x < y) {
      return -1;
    }
    if (x > y) {
      return 1;
    }
    return 0;
  };
  // `lookup` and `profile` are loop-invariant; normalise the lookup side
  // once rather than once per scanned cell (see classic.cpp::lookup_scan).
  const std::string lookup_key = lookup.is_text() ? lookup_text_cmp_key(lookup.as_text(), profile) : std::string();
  auto cmp_text = [&](std::string_view a) -> int {
    // Normalise (see classic.cpp::lookup_scan / lookup_text_cmp_key) so kana /
    // half-width voicing variants order together in MATCH approximate mode.
    return strings::case_insensitive_compare(lookup_text_cmp_key(a, profile), lookup_key);
  };

  // `last_valid_pos` is the running best position under the ordering rule.
  // For type=+1 we want the largest position whose value is <= target; for
  // type=-1 the largest position whose value is >= target.
  std::size_t best_pos = 0;  // 1-based; 0 means "not found yet".
  for (std::size_t i = 0; i < n; ++i) {
    const Value& cell = cells[i];
    int cmp = 0;  // sign of (cell - lookup)
    bool comparable = false;
    if (lookup.is_text() && cell.is_text()) {
      cmp = cmp_text(cell.as_text());
      comparable = true;
    } else if ((lookup.is_number() || lookup.is_blank()) && cell.is_number()) {
      // Blank cells in the scanned axis are NOT treated as numeric 0 in MATCH
      // approximate mode: only real numeric cells participate in the ordering,
      // mirroring VLOOKUP/HLOOKUP's `lookup_scan` (see classic.cpp). A blank
      // cell is skipped (non-comparable) so a blank slot inside an ascending
      // range does not become a spurious 0 match.
      const double lv = lookup.is_blank() ? 0.0 : lookup.as_number();
      const double cv = cell.as_number();
      cmp = cmp_numeric(cv, lv);
      comparable = true;
    } else if (lookup.is_boolean() && cell.is_boolean()) {
      const int lb = lookup.as_boolean() ? 1 : 0;
      const int cb = cell.as_boolean() ? 1 : 0;
      cmp = cmp_numeric(cb, lb);
      comparable = true;
    }
    if (!comparable) {
      // Cross-type: skip. Accepted divergence from Excel's full ordering.
      continue;
    }
    if (match_type == 1) {
      // Ascending: record the last position whose cell is <= lookup. We
      // do NOT short-circuit on a strictly-greater cell — Mac Excel's
      // empirical behaviour on unsorted data is to keep scanning and
      // simply return the running last <= match. For sorted ascending
      // data this still produces the documented answer because every
      // post-match cell is strictly greater (and skipped).
      if (cmp <= 0) {
        best_pos = i + 1;
      }
      continue;
    }
    // match_type == -1, descending: symmetric rule — record the last
    // position whose cell is >= lookup, and keep scanning past strictly
    // smaller cells without short-circuiting.
    if (cmp >= 0) {
      best_pos = i + 1;
    }
  }
  if (best_pos == 0) {
    return Value::error(ErrorCode::NA);
  }
  return Value::number(static_cast<double>(best_pos));
}

// MATCH(lookup_value, lookup_array, [match_type])
//
// Returns the 1-based position of `lookup_value` inside the 1-D
// `lookup_array`. `lookup_array` must be a `RangeOp(Ref, Ref)` or a single
// `Ref` with a 1-D shape (row vector or column vector); a 2-D range yields
// `#N/A`.
//
// match_type semantics:
//   *  1 (default) - ascending array; returns the largest position whose
//      value is <= lookup_value. Wildcards are NOT honoured.
//   *  0           - exact match with DOS-style wildcards (`*`, `?`, `~`)
//      for text targets; first hit wins. No match -> `#N/A`.
//   * -1           - descending array; returns the largest position whose
//      value is >= lookup_value.
//
// Cross-type comparison is not implemented beyond "same rank, ordered" for
// approximate modes: a cell whose kind doesn't match the lookup_value rank
// is treated as a non-match and never participates in the approximate
// ranking. This is a documented accepted divergence.
Value eval_match_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity != 2 && arity != 3) {
    return Value::error(ErrorCode::Value);
  }

  // lookup_value (scalar). Errors propagate; Blank is treated as 0 for the
  // numeric comparison path below, matching Excel.
  const Value lookup = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (lookup.is_error()) {
    return lookup;
  }
  const bool array_lookup = lookup.is_array();

  // lookup_array: must be a range / Ref with a 1-D shape. A static
  // reference is judged by its declared shape, and a 2-D one is not read.
  ReferenceTable table;
  const auto by_reference = resolve_reference_table(call.as_call_arg(1), ctx, &table);
  if (by_reference && by_reference.value() && table.declared.rows() != 1U && table.declared.cols() != 1U) {
    return fail_lookup(lookup, arena, Value::error(ErrorCode::NA));
  }
  auto resolved = resolve_range_arg(call.as_call_arg(1), arena, registry, ctx);
  if (!resolved) {
    return fail_lookup(lookup, arena, Value::error(resolved.error()));
  }
  const std::uint32_t rows = resolved.value().rows;
  const std::uint32_t cols = resolved.value().cols;
  std::vector<Value> cells = std::move(resolved.value().cells);
  if (rows != 1U && cols != 1U) {
    // 2-D array to MATCH is not supported and Excel reports #N/A.
    return fail_lookup(lookup, arena, Value::error(ErrorCode::NA));
  }

  // match_type: default 1. The scalar path retains its existing {-1, 0, 1}
  // validation. Mac Excel's array path instead coerces any finite positive
  // mode to ascending and any finite negative mode to descending; in
  // particular, MATCH(array, range, 2) is not a global invalid result.
  int match_type = 1;
  if (arity == 3) {
    const Value mt_val = eval_node(call.as_call_arg(2), arena, registry, ctx);
    if (mt_val.is_error()) {
      return fail_lookup(lookup, arena, mt_val);
    }
    auto mt_num = coerce_to_number(mt_val);
    if (!mt_num) {
      return fail_lookup(lookup, arena, Value::error(mt_num.error()));
    }
    const double mt_raw = truncate_index(mt_num.value());
    if (mt_raw == -1.0) {
      match_type = -1;
    } else if (mt_raw == 0.0) {
      match_type = 0;
    } else if (mt_raw == 1.0) {
      match_type = 1;
    } else if (array_lookup && mt_raw > 0.0) {
      match_type = 1;
    } else if (array_lookup && mt_raw < 0.0) {
      match_type = -1;
    } else {
      return fail_lookup(lookup, arena, Value::error(ErrorCode::NA));
    }
  }

  const std::size_t n = cells.size();
  if (n == 0) {
    return fail_lookup(lookup, arena, Value::error(ErrorCode::NA));
  }

  if (array_lookup) {
    return map_lookup_query_array(lookup, arena, [&](const Value& query) {
      return match_lookup_one(cells, query, match_type, ctx.excel_profile());
    });
  }
  return match_lookup_one(cells, lookup, match_type, ctx.excel_profile());
}

// VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
//
// Scans the first column of `table_array` for `lookup_value` and returns
// the cell at (matched_row, col_index_num - 1). `range_lookup` defaults to
// TRUE (approximate match); FALSE enables exact match with DOS-style
// wildcards on Text lookup values (`*`, `?`, `~` escape). Approximate mode
// expects the first column to be ascending and returns the largest row
// whose value is <= lookup_value; `#N/A` when every first-column cell is
// already greater. Wildcards are NOT honoured in approximate mode.
//
// Error lookup_value propagates unchanged. Cross-type comparisons skip
// (accepted divergence, same as MATCH).
Value eval_vlookup_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  return eval_table_lookup_lazy(call, arena, registry, ctx, LookupAxis::Column);
}

// HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])
//
// Symmetric to VLOOKUP: scans the first row of `table_array` left-to-right
// for `lookup_value` and returns the cell at (row_index_num - 1,
// matched_col). All other rules (wildcards, range_lookup semantics, edge
// cases) mirror VLOOKUP with rows/cols swapped.
Value eval_hlookup_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  return eval_table_lookup_lazy(call, arena, registry, ctx, LookupAxis::Row);
}

// LOOKUP(lookup_value, lookup_vector, [result_vector])   -- vector form.
//
// Legacy approximate-match search: the lookup vector is assumed to be in
// ascending order and LOOKUP returns the result-vector cell whose parallel
// position is the last one whose lookup cell is <= `lookup_value`. Exact
// mode does not exist for LOOKUP; the match is always approximate. If
// `result_vector` is omitted the result cells come from `lookup_vector`
// itself.
//
// Axis: Excel picks the longer dimension of `lookup_vector`. When the
// vector is taller than wide, scan is vertical (treat as column); when
// wider than tall, scan is horizontal (treat as row). A square or
// single-cell vector defaults to column orientation.
//
// Errors / edge cases:
//   * `lookup_value` is an error: propagate unchanged.
//   * `lookup_value` is blank or text that never sorts before any vector
//     cell: `#N/A`.
//   * `result_vector` shorter than `lookup_vector`: we still index by the
//     matched offset and fall back to `#N/A` if the index is out of range
//     (Excel's observable behaviour).
//
// The 2-argument call site has two flavours:
//
//   * Vector form: `lookup_vector` is 1-D (1 row OR 1 column). Scan the
//     vector for the largest value <= lookup_value and return that same
//     cell.
//   * Array form: `array` is 2-D. When taller than wide (rows >= cols),
//     scan the first column and return the corresponding cell of the LAST
//     column. When wider than tall, scan the first row and return the
//     corresponding cell of the LAST row. Square arrays default to the
//     taller-than-wide rule.
//
// The 3-argument form is always vector-shaped: `lookup_vector` and
// `result_vector` must be parallel 1-D ranges; the matched offset on
// `lookup_vector` indexes `result_vector`.
Value eval_lookup_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity != 2 && arity != 3) {
    return Value::error(ErrorCode::Value);
  }

  const Value lookup = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (lookup.is_error()) {
    return lookup;
  }
  if (lookup.is_blank()) {
    return Value::error(ErrorCode::NA);
  }

  auto lookup_resolved = resolve_range_arg(call.as_call_arg(1), arena, registry, ctx);
  if (!lookup_resolved) {
    return Value::error(lookup_resolved.error());
  }
  const std::uint32_t lrows = lookup_resolved.value().rows;
  const std::uint32_t lcols = lookup_resolved.value().cols;
  std::vector<Value> lookup_cells = std::move(lookup_resolved.value().cells);
  if (lrows == 0U || lcols == 0U) {
    return Value::error(ErrorCode::Ref);
  }

  // The orientation is the declared one: `A:C` is taller than wide however
  // few of its rows hold values.
  std::uint32_t lshape_rows = lrows;
  std::uint32_t lshape_cols = lcols;
  static_reference_shape(call.as_call_arg(1), ctx, &lshape_rows, &lshape_cols);
  const LookupAxis axis = lshape_rows >= lshape_cols ? LookupAxis::Column : LookupAxis::Row;
  const std::size_t off =
      lookup_scan(lookup_cells, lrows, lcols, axis, lookup, /*approximate=*/true, ctx.excel_profile());
  if (off == SIZE_MAX) {
    return Value::error(ErrorCode::NA);
  }

  if (arity == 2) {
    // Vector form (1-D input): result is the matched cell of the lookup
    // vector itself. Array form (2-D input): result is the corresponding
    // cell of the last column (taller-than-wide) or last row (wider-than-
    // tall). The two paths agree on a 1-D input because last-col == col 0
    // and last-row == row 0 in that case.
    std::size_t flat = 0;
    if (axis == LookupAxis::Column) {
      // Scan first column; return last column at the matched row offset.
      const std::size_t last_col = static_cast<std::size_t>(lcols) - 1U;
      flat = off * static_cast<std::size_t>(lcols) + last_col;
    } else {
      // Scan first row; return last row at the matched column offset.
      const std::size_t last_row = static_cast<std::size_t>(lrows) - 1U;
      flat = last_row * static_cast<std::size_t>(lcols) + off;
    }
    return flat < lookup_cells.size() ? lookup_cells[flat] : Value::error(ErrorCode::NA);
  }

  auto result_resolved = resolve_range_arg(call.as_call_arg(2), arena, registry, ctx);
  if (!result_resolved) {
    return Value::error(result_resolved.error());
  }
  const std::uint32_t rrows = result_resolved.value().rows;
  const std::uint32_t rcols = result_resolved.value().cols;
  std::vector<Value> result_cells = std::move(result_resolved.value().cells);
  if (rrows == 0U || rcols == 0U) {
    return Value::error(ErrorCode::Ref);
  }
  // Index the result vector along its own long axis. Excel treats the
  // result vector as parallel to the lookup vector, so the "axis position"
  // maps to the same offset regardless of orientation.
  std::uint32_t rshape_rows = rrows;
  std::uint32_t rshape_cols = rcols;
  static_reference_shape(call.as_call_arg(2), ctx, &rshape_rows, &rshape_cols);
  const LookupAxis raxis = rshape_rows >= rshape_cols ? LookupAxis::Column : LookupAxis::Row;
  const std::size_t flat = raxis == LookupAxis::Column ? (off * static_cast<std::size_t>(rcols)) : off;
  return flat < result_cells.size() ? result_cells[flat] : Value::error(ErrorCode::NA);
}

}  // namespace eval
}  // namespace formulon
