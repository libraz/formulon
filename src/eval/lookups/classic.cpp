// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the classic lookup-family lazy impls (`CHOOSE`,
// `INDEX`, `MATCH`, `VLOOKUP`, `HLOOKUP`). See `lookups/classic.h` for
// the dispatch-table contract and `eval/lazy_impls.h` for the shared
// vocabulary.

#include "eval/lookups/classic.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
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
// CHOOSE / INDEX / MATCH / VLOOKUP / HLOOKUP (lookup & reference)
// ---------------------------------------------------------------------------

// Axis along which VLOOKUP / HLOOKUP scan their table_array rectangle for the
// lookup_value. `Column` means "walk top-down through the first column"
// (VLOOKUP); `Row` means "walk left-to-right through the first row"
// (HLOOKUP).
enum class LookupAxis : std::uint8_t { Column, Row };

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
                        const Value& lookup_value, bool approximate) {
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
      const std::string pat_lower = strings::to_ascii_lower(lookup_value.as_text());
      for (std::size_t i = 0; i < n; ++i) {
        const Value& cell = cell_at(i);
        if (!cell.is_text()) {
          continue;
        }
        const std::string cell_lower = strings::to_ascii_lower(cell.as_text());
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

  // Approximate match: ascending-assumed scan. Record the last position
  // whose value is <= lookup_value and stop at the first strictly greater
  // cell. Wildcards are NEVER honoured here (Excel treats them as literal
  // text in approximate mode).
  auto cmp_numeric = [](double x, double y) -> int {
    if (x < y) {
      return -1;
    }
    if (x > y) {
      return 1;
    }
    return 0;
  };
  std::size_t best = SIZE_MAX;
  for (std::size_t i = 0; i < n; ++i) {
    const Value& cell = cell_at(i);
    int cmp = 0;  // sign of (cell - lookup_value)
    bool comparable = false;
    if (lookup_value.is_text() && cell.is_text()) {
      cmp = strings::case_insensitive_compare(cell.as_text(), lookup_value.as_text());
      comparable = true;
    } else if ((lookup_value.is_number() || lookup_value.is_blank()) && (cell.is_number() || cell.is_blank())) {
      const double lv = lookup_value.is_blank() ? 0.0 : lookup_value.as_number();
      const double cv = cell.is_blank() ? 0.0 : cell.as_number();
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
      continue;
    }
    break;
  }
  return best;
}

}  // namespace

// CHOOSE(index_num, value1, value2, ...)
//
// Evaluates `index_num`, truncates to int, and returns only the corresponding
// argument subtree (`CHOOSE(2, a, b, c)` returns `b` and never touches `a`
// or `c`). Out-of-range indices yield `#VALUE!`; a numeric coercion failure
// on `index_num` also yields `#VALUE!`. Errors in `index_num` propagate.
// Errors in the chosen value also propagate; unselected arguments are never
// evaluated.
Value eval_choose_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  // Need at least index_num plus one value.
  if (arity < 2) {
    return Value::error(ErrorCode::Value);
  }
  const Value idx_val = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (idx_val.is_error()) {
    return idx_val;
  }
  auto idx_num = coerce_to_number(idx_val);
  if (!idx_num) {
    return Value::error(idx_num.error());
  }
  // Excel truncates (toward zero) rather than rounds: CHOOSE(2.9, ...)
  // selects the 2nd value, not the 3rd.
  const double raw = std::floor(idx_num.value());
  if (!(raw >= 1.0 && raw <= static_cast<double>(arity - 1))) {
    return Value::error(ErrorCode::Value);
  }
  const auto n = static_cast<std::uint32_t>(raw);
  return eval_node(call.as_call_arg(n), arena, registry, ctx);
}

// INDEX(array, row_num, [column_num])
//
// Returns a single cell from `array` by 1-based (row_num, column_num). The
// source array must be a `RangeOp(Ref, Ref)` or a single `Ref`; anything
// else is `#VALUE!`. Out-of-bounds indices are `#REF!`. Negative or
// non-coercible indices are `#VALUE!`.
//
// Shape disambiguation for the 2-arg form: if the array is 1-D (rows == 1
// or cols == 1), the sole index selects along the non-singleton dimension.
// For a 2-D array with only two args provided, `row_num` selects the row
// and the "entire row" result would be needed for the column — unsupported
// today, so we return `#VALUE!` (documented divergence from Excel 365
// which spills in that case).
//
// Accepted divergence: a zero in either index dimension is "return the
// whole row/column" in Excel 365 via dynamic arrays. Scalar array results
// are not wired up yet, so INDEX returns `#VALUE!` for that shape.
Value eval_index_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity != 2 && arity != 3) {
    return Value::error(ErrorCode::Value);
  }
  std::vector<Value> cells;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  ErrorCode range_err = ErrorCode::Value;
  if (!resolve_range_arg(call.as_call_arg(0), arena, registry, ctx, &cells, &range_err, &rows, &cols)) {
    return Value::error(range_err);
  }
  if (rows == 0U || cols == 0U) {
    // Defensive: expand_range always produces a positive rectangle today.
    return Value::error(ErrorCode::Ref);
  }

  // row_num is required (arity 2 or 3), col_num is optional.
  const Value row_val = eval_node(call.as_call_arg(1), arena, registry, ctx);
  if (row_val.is_error()) {
    return row_val;
  }
  auto row_num_exp = coerce_to_number(row_val);
  if (!row_num_exp) {
    return Value::error(row_num_exp.error());
  }
  const double row_raw = std::floor(row_num_exp.value());
  if (row_raw < 0.0) {
    return Value::error(ErrorCode::Value);
  }
  const auto row_idx = static_cast<std::uint32_t>(row_raw);

  std::uint32_t col_idx = 0;
  bool col_explicit = false;
  if (arity == 3) {
    const Value col_val = eval_node(call.as_call_arg(2), arena, registry, ctx);
    if (col_val.is_error()) {
      return col_val;
    }
    auto col_num_exp = coerce_to_number(col_val);
    if (!col_num_exp) {
      return Value::error(col_num_exp.error());
    }
    const double col_raw = std::floor(col_num_exp.value());
    if (col_raw < 0.0) {
      return Value::error(ErrorCode::Value);
    }
    col_idx = static_cast<std::uint32_t>(col_raw);
    col_explicit = true;
  }

  // Resolve (row_idx, col_idx) into a (0-based) row / column within the
  // rectangle. The logic depends on shape and how many indices the caller
  // provided. Zero values are "whole dimension" in Excel's spill model —
  // unsupported here.
  std::uint32_t r = 0;
  std::uint32_t c = 0;
  if (!col_explicit) {
    // Two-arg form.
    if (rows == 1U && cols == 1U) {
      // 1x1 range: row_num must be 1 (or 0 "whole", which we reject).
      if (row_idx != 1U) {
        return Value::error(ErrorCode::Ref);
      }
      r = 0;
      c = 0;
    } else if (rows == 1U) {
      // Row vector: sole index selects the column.
      if (row_idx == 0U) {
        return Value::error(ErrorCode::Value);
      }
      if (row_idx > cols) {
        return Value::error(ErrorCode::Ref);
      }
      r = 0;
      c = row_idx - 1U;
    } else if (cols == 1U) {
      // Column vector: sole index selects the row.
      if (row_idx == 0U) {
        return Value::error(ErrorCode::Value);
      }
      if (row_idx > rows) {
        return Value::error(ErrorCode::Ref);
      }
      r = row_idx - 1U;
      c = 0;
    } else {
      // 2-D array with only row selector: Excel would spill the whole row;
      // we don't support scalar spill results yet.
      return Value::error(ErrorCode::Value);
    }
  } else {
    // Three-arg form.
    if (row_idx == 0U || col_idx == 0U) {
      // Whole-row / whole-column via dynamic-array spill — unsupported.
      return Value::error(ErrorCode::Value);
    }
    if (rows == 1U) {
      // Row vector: row_num must be 1.
      if (row_idx != 1U) {
        return Value::error(ErrorCode::Ref);
      }
      if (col_idx > cols) {
        return Value::error(ErrorCode::Ref);
      }
      r = 0;
      c = col_idx - 1U;
    } else if (cols == 1U) {
      // Column vector: col_num must be 1.
      if (col_idx != 1U) {
        return Value::error(ErrorCode::Ref);
      }
      if (row_idx > rows) {
        return Value::error(ErrorCode::Ref);
      }
      r = row_idx - 1U;
      c = 0;
    } else {
      if (row_idx > rows || col_idx > cols) {
        return Value::error(ErrorCode::Ref);
      }
      r = row_idx - 1U;
      c = col_idx - 1U;
    }
  }

  const std::size_t flat = (static_cast<std::size_t>(r) * static_cast<std::size_t>(cols)) + static_cast<std::size_t>(c);
  if (flat >= cells.size()) {
    return Value::error(ErrorCode::Ref);
  }
  return cells[flat];
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

  // lookup_array: must be a range / Ref with a 1-D shape.
  std::vector<Value> cells;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  ErrorCode range_err = ErrorCode::Value;
  if (!resolve_range_arg(call.as_call_arg(1), arena, registry, ctx, &cells, &range_err, &rows, &cols)) {
    return Value::error(range_err);
  }
  if (rows != 1U && cols != 1U) {
    // 2-D array to MATCH is not supported and Excel reports #N/A.
    return Value::error(ErrorCode::NA);
  }

  // match_type: default 1. Truncate toward zero and clamp into {-1, 0, 1};
  // any other value is rejected with #N/A (Excel's effective behaviour).
  int match_type = 1;
  if (arity == 3) {
    const Value mt_val = eval_node(call.as_call_arg(2), arena, registry, ctx);
    if (mt_val.is_error()) {
      return mt_val;
    }
    auto mt_num = coerce_to_number(mt_val);
    if (!mt_num) {
      return Value::error(mt_num.error());
    }
    const double mt_raw = std::floor(mt_num.value());
    if (mt_raw == -1.0) {
      match_type = -1;
    } else if (mt_raw == 0.0) {
      match_type = 0;
    } else if (mt_raw == 1.0) {
      match_type = 1;
    } else {
      return Value::error(ErrorCode::NA);
    }
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
      const std::string pat_lower = strings::to_ascii_lower(lookup.as_text());
      for (std::size_t i = 0; i < n; ++i) {
        const Value& cell = cells[i];
        if (!cell.is_text()) {
          continue;
        }
        const std::string cell_lower = strings::to_ascii_lower(cell.as_text());
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
  auto cmp_text = [](std::string_view a, std::string_view b) -> int { return strings::case_insensitive_compare(a, b); };

  // `last_valid_pos` is the running best position under the ordering rule.
  // For type=+1 we want the largest position whose value is <= target; for
  // type=-1 the largest position whose value is >= target.
  std::size_t best_pos = 0;  // 1-based; 0 means "not found yet".
  for (std::size_t i = 0; i < n; ++i) {
    const Value& cell = cells[i];
    int cmp = 0;  // sign of (cell - lookup)
    bool comparable = false;
    if (lookup.is_text() && cell.is_text()) {
      cmp = cmp_text(cell.as_text(), lookup.as_text());
      comparable = true;
    } else if ((lookup.is_number() || lookup.is_blank()) && (cell.is_number() || cell.is_blank())) {
      const double lv = lookup.is_blank() ? 0.0 : lookup.as_number();
      const double cv = cell.is_blank() ? 0.0 : cell.as_number();
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
      // Ascending: record last position with cell <= lookup. Stop at the
      // first cell strictly greater than lookup.
      if (cmp <= 0) {
        best_pos = i + 1;
        continue;
      }
      break;
    }
    // match_type == -1, descending: record last position with cell >=
    // lookup. Stop at the first cell strictly less than lookup.
    if (cmp >= 0) {
      best_pos = i + 1;
      continue;
    }
    break;
  }
  if (best_pos == 0) {
    return Value::error(ErrorCode::NA);
  }
  return Value::number(static_cast<double>(best_pos));
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
  const std::uint32_t arity = call.as_call_arity();
  if (arity != 3 && arity != 4) {
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
  if (rows == 0U || cols == 0U) {
    return Value::error(ErrorCode::Ref);
  }

  const Value col_val = eval_node(call.as_call_arg(2), arena, registry, ctx);
  if (col_val.is_error()) {
    return col_val;
  }
  auto col_num_exp = coerce_to_number(col_val);
  if (!col_num_exp) {
    return Value::error(col_num_exp.error());
  }
  const double col_raw = std::floor(col_num_exp.value());
  if (col_raw < 1.0) {
    return Value::error(ErrorCode::Value);
  }
  if (col_raw > static_cast<double>(cols)) {
    return Value::error(ErrorCode::Ref);
  }
  const auto col_idx = static_cast<std::uint32_t>(col_raw);

  // range_lookup defaults to TRUE when omitted.
  bool approximate = true;
  if (arity == 4) {
    const Value rl_val = eval_node(call.as_call_arg(3), arena, registry, ctx);
    if (rl_val.is_error()) {
      return rl_val;
    }
    auto rl_bool = coerce_to_bool(rl_val);
    if (!rl_bool) {
      return Value::error(rl_bool.error());
    }
    approximate = rl_bool.value();
  }

  const std::size_t off = lookup_scan(cells, rows, cols, LookupAxis::Column, lookup, approximate);
  if (off == SIZE_MAX) {
    return Value::error(ErrorCode::NA);
  }
  const std::size_t flat = (off * static_cast<std::size_t>(cols)) + static_cast<std::size_t>(col_idx - 1U);
  if (flat >= cells.size()) {
    // Defensive; the bounds checks above should already rule this out.
    return Value::error(ErrorCode::Ref);
  }
  return cells[flat];
}

// HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])
//
// Symmetric to VLOOKUP: scans the first row of `table_array` left-to-right
// for `lookup_value` and returns the cell at (row_index_num - 1,
// matched_col). All other rules (wildcards, range_lookup semantics, edge
// cases) mirror VLOOKUP with rows/cols swapped.
Value eval_hlookup_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity != 3 && arity != 4) {
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
  if (rows == 0U || cols == 0U) {
    return Value::error(ErrorCode::Ref);
  }

  const Value row_val = eval_node(call.as_call_arg(2), arena, registry, ctx);
  if (row_val.is_error()) {
    return row_val;
  }
  auto row_num_exp = coerce_to_number(row_val);
  if (!row_num_exp) {
    return Value::error(row_num_exp.error());
  }
  const double row_raw = std::floor(row_num_exp.value());
  if (row_raw < 1.0) {
    return Value::error(ErrorCode::Value);
  }
  if (row_raw > static_cast<double>(rows)) {
    return Value::error(ErrorCode::Ref);
  }
  const auto row_idx = static_cast<std::uint32_t>(row_raw);

  bool approximate = true;
  if (arity == 4) {
    const Value rl_val = eval_node(call.as_call_arg(3), arena, registry, ctx);
    if (rl_val.is_error()) {
      return rl_val;
    }
    auto rl_bool = coerce_to_bool(rl_val);
    if (!rl_bool) {
      return Value::error(rl_bool.error());
    }
    approximate = rl_bool.value();
  }

  const std::size_t off = lookup_scan(cells, rows, cols, LookupAxis::Row, lookup, approximate);
  if (off == SIZE_MAX) {
    return Value::error(ErrorCode::NA);
  }
  const std::size_t flat = (static_cast<std::size_t>(row_idx - 1U) * static_cast<std::size_t>(cols)) + off;
  if (flat >= cells.size()) {
    return Value::error(ErrorCode::Ref);
  }
  return cells[flat];
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
// The array form `LOOKUP(lookup_value, array)` is intentionally not yet
// implemented -- it is deprecated, and the IronCalc oracle only covers the
// vector form.
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

  std::vector<Value> lookup_cells;
  std::uint32_t lrows = 0;
  std::uint32_t lcols = 0;
  ErrorCode range_err = ErrorCode::Value;
  if (!resolve_range_arg(call.as_call_arg(1), arena, registry, ctx, &lookup_cells, &range_err, &lrows, &lcols)) {
    return Value::error(range_err);
  }
  if (lrows == 0U || lcols == 0U) {
    return Value::error(ErrorCode::Ref);
  }

  const LookupAxis axis = lrows >= lcols ? LookupAxis::Column : LookupAxis::Row;
  const std::size_t off = lookup_scan(lookup_cells, lrows, lcols, axis, lookup, /*approximate=*/true);
  if (off == SIZE_MAX) {
    return Value::error(ErrorCode::NA);
  }

  if (arity == 2) {
    // Result comes from lookup_vector itself at the matched axis position.
    const std::size_t flat = axis == LookupAxis::Column ? (off * static_cast<std::size_t>(lcols))
                                                        : off;  // Row -> first-row strip index
    return flat < lookup_cells.size() ? lookup_cells[flat] : Value::error(ErrorCode::NA);
  }

  std::vector<Value> result_cells;
  std::uint32_t rrows = 0;
  std::uint32_t rcols = 0;
  if (!resolve_range_arg(call.as_call_arg(2), arena, registry, ctx, &result_cells, &range_err, &rrows, &rcols)) {
    return Value::error(range_err);
  }
  if (rrows == 0U || rcols == 0U) {
    return Value::error(ErrorCode::Ref);
  }
  // Index the result vector along its own long axis. Excel treats the
  // result vector as parallel to the lookup vector, so the "axis position"
  // maps to the same offset regardless of orientation.
  const LookupAxis raxis = rrows >= rcols ? LookupAxis::Column : LookupAxis::Row;
  const std::size_t flat =
      raxis == LookupAxis::Column ? (off * static_cast<std::size_t>(rcols)) : off;
  return flat < result_cells.size() ? result_cells[flat] : Value::error(ErrorCode::NA);
}

}  // namespace eval
}  // namespace formulon
