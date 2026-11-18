//
// Implementation of the classic lookup-family lazy impls that search by
// value (`MATCH`, `VLOOKUP`, `HLOOKUP`, `LOOKUP`); `CHOOSE` and `INDEX` live
// in `lookups/index.cpp`. See `lookups/classic.h` for the dispatch-table
// contract and `eval/lazy_impls.h` for the shared vocabulary.

#include "eval/lookups/classic.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/coerce.h"
#include "eval/collation.h"
#include "eval/declared_rect.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/lazy_impls.h"
#include "eval/lookups/binary_search.h"
#include "eval/lookups/common.h"
#include "eval/name_env_resolve.h"
#include "eval/range_args.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

template <typename CellAt>
std::size_t scan_exact_text_lookup(std::string_view raw, ExcelProfile profile, std::size_t count, CellAt&& cell_at) {
  const PreparedTextQuery query = prepare_text_query(raw, profile, CollationStrength::kClassicExact);
  for (std::size_t i = 0; i < count; ++i) {
    if (prepared_text_query_matches(cell_at(i), query, profile)) {
      return i;
    }
  }
  return SIZE_MAX;
}

// First cell equal to a number / boolean `lookup_value`, or SIZE_MAX. Blank
// cells never match.
template <typename CellAt>
std::size_t scan_exact_typed(const Value& lookup_value, std::size_t count, CellAt&& cell_at) {
  for (std::size_t i = 0; i < count; ++i) {
    if (literal_equals(cell_at(i), lookup_value)) {
      return i;
    }
  }
  return SIZE_MAX;
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

// Orders a cell against the lookup value for the approximate searches: the
// sign of (cell - lookup), or `kForeignType` for a cell of another type.
// Only real numbers rank against a number lookup; blank cells are foreign. Wildcards are never honoured.
class ApproxComparer {
 public:
  ApproxComparer(const Value& lookup, ExcelProfile profile)
      : lookup_(lookup),
        text_query_(lookup.is_text() ? std::optional<CollateQuery>(std::in_place, lookup.as_text(), profile,
                                                                   CollationStrength::kCompare)
                                     : std::nullopt) {}

  int operator()(const Value& cell) const {
    if (lookup_.is_text()) {
      if (!cell.is_text()) {
        return kForeignType;
      }
      return text_query_->compare(cell.as_text());
    }
    if (lookup_.is_number()) {
      if (!cell.is_number()) {
        return kForeignType;
      }
      return compare_numbers(cell.as_number(), lookup_.as_number());
    }
    if (lookup_.is_boolean()) {
      if (!cell.is_boolean()) {
        return kForeignType;
      }
      return compare_numbers(cell.as_boolean() ? 1.0 : 0.0, lookup_.as_boolean() ? 1.0 : 0.0);
    }
    return kForeignType;
  }

 private:
  const Value& lookup_;
  std::optional<CollateQuery> text_query_;
};

// Lookup scan for VLOOKUP / HLOOKUP / LOOKUP. Walks the first column (axis=Column) or
// the first row (axis=Row) of the `flat` rectangle (rows x cols, row-major)
// for `lookup_value`, using approximate (binary search) or exact (first hit)
// matching.
//
// Returns the 0-based offset along the scanned axis on match, or `SIZE_MAX`
// when no match was found.
//
// In exact mode, a text lookup without unescaped `*` / `?` compares the
// unescaped literal through the collator; otherwise it goes through
// `wildcard_match` on the folded keys.
//
// Approximate mode follows `approximate_search`: a cell of another type than the
// lookup value ends the binary search and the rest is scanned linearly.
// `descending` (MATCH -1) replaces the binary search with `descending_scan`.
// A blank `raw_lookup` is the number 0 (`BlankLookup::kZero`).
// `declared` is the vector's declared length when it is longer than the cells
// read (a whole-axis reference read up to its last populated cell): the
// ascending binary search runs over that length with the rest blank, as Excel
// bisects the whole axis.
std::size_t lookup_scan(const std::vector<Value>& flat, std::uint32_t rows, std::uint32_t cols, LookupAxis axis,
                        const Value& raw_lookup, bool approximate, ExcelProfile profile, bool descending = false,
                        std::size_t declared = 0U) {
  const Value lookup_value = typed_lookup_value(raw_lookup, BlankLookup::kZero);
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
    // Exact match: first hit wins. Text uses collation for literal patterns
    // and the wildcard matcher when metacharacters are present. Every other
    // kind-pairing uses literal equality.
    if (lookup_value.is_text()) {
      const std::string_view raw = lookup_value.as_text();
      return scan_exact_text_lookup(raw, profile, n, cell_at);
    }
    return scan_exact_typed(lookup_value, n, cell_at);
  }

  // Approximate match: binary search, falling back to a left-to-right scan at
  // the first probe of another type.
  const ApproxComparer order(lookup_value, profile);
  if (descending) {
    return descending_scan(n, [&](std::size_t i) { return order(cell_at(i)); });
  }
  const std::size_t length = declared > n ? declared : n;
  const auto cmp = [&](std::size_t i) { return i < n ? order(cell_at(i)) : kForeignType; };
  return approximate_search(length, cmp, n);
}

// Fails the whole call with `err`. An array lookup fails lane by lane, so a
// query cell holding its own error keeps it.
Value fail_lookup(const Value& lookup, Arena& arena, const Value& err) {
  if (lookup.is_array()) {
    return map_query_lanes(lookup, arena, [&err](const Value& query) { return query.is_error() ? query : err; });
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

  // A one-cell table holding an error is no table at all.
  if (rows == 1U && cols == 1U) {
    const Value only =
        (by_reference && by_reference.value()) ? read_table_cell(table, 0U, 0U, arena, registry, ctx) : cells.front();
    if (only.is_error()) {
      return fail_lookup(lookup, arena, Value::error(ErrorCode::NA));
    }
  }

  // An error read from a cell does not propagate: it reads as an index the
  // table cannot serve, which is #REF! once the lookup finds a row. An error
  // from any other expression propagates.
  const Value index_val = eval_node(call.as_call_arg(2), arena, registry, ctx);
  const bool index_unusable = index_val.is_error() && call.as_call_arg(2).kind() == parser::NodeKind::Ref;
  const std::uint32_t result_extent = axis == LookupAxis::Column ? cols : rows;
  // The error one index fails with, or nullopt with `*out` set to it.
  const auto index_error = [result_extent](const Value& candidate, std::uint32_t* out) -> std::optional<Value> {
    if (candidate.is_error()) {
      return candidate;
    }
    auto index_num = coerce_to_number(candidate);
    if (!index_num) {
      return Value::error(index_num.error());
    }
    if (truncate_index(index_num.value()) < 1.0 && snap_near_integer(index_num.value()) < 1.0) {
      return Value::error(ErrorCode::Value);
    }
    const double index_raw = truncate_index(snap_near_integer(index_num.value()));
    if (index_raw > static_cast<double>(result_extent)) {
      return Value::error(ErrorCode::Ref);
    }
    *out = static_cast<std::uint32_t>(index_raw);
    return std::nullopt;
  };
  // A scalar lookup with an array index picks one column (row) per element.
  const bool index_array = !array_lookup && index_val.is_array();
  std::uint32_t result_index = 0U;
  if (!index_unusable && !index_array) {
    if (std::optional<Value> bad = index_error(index_val, &result_index)) {
      return fail_lookup(lookup, arena, *bad);
    }
  }
  const auto fail_all = [&](const Value& err) -> Value { return fail_lookup(lookup, arena, err); };

  bool approximate = true;
  if (arity == 4) {
    const Value rl_val = eval_node(call.as_call_arg(3), arena, registry, ctx);
    if (rl_val.is_error()) {
      return fail_all(rl_val);
    }
    auto rl_bool = coerce_to_bool(rl_val);
    if (!rl_bool) {
      return fail_all(Value::error(rl_bool.error()));
    }
    approximate = rl_bool.value();
  }

  // A reference table is scanned along its walked extent -- the same cells
  // a full expansion would scan -- and only the scanned line is read.
  const bool read_by_reference = by_reference && by_reference.value();
  const std::vector<Value>* scan_cells = &cells;
  std::uint32_t scan_rows = rows;
  std::uint32_t scan_cols = cols;
  std::vector<Value> scan_line;
  if (read_by_reference) {
    const DeclaredRect& walked = table.walked;
    auto line = axis == LookupAxis::Column
                    ? read_table_block(table, walked.row_first, walked.row_last, table.declared.col_first,
                                       table.declared.col_first, arena, registry, ctx)
                    : read_table_block(table, table.declared.row_first, table.declared.row_first, walked.col_first,
                                       walked.col_last, arena, registry, ctx);
    if (!line) {
      return fail_all(Value::error(line.error()));
    }
    scan_line = std::move(line.value());
    scan_cells = &scan_line;
    scan_rows = axis == LookupAxis::Column ? walked.rows() : 1U;
    scan_cols = axis == LookupAxis::Column ? 1U : walked.cols();
  }
  const auto fetch = [&](std::size_t off, std::uint32_t pick) -> Value {
    if (index_unusable) {
      return Value::error(ErrorCode::Ref);
    }
    if (read_by_reference) {
      if (pick == 1U) {
        return scan_line[off];
      }
      const auto along = static_cast<std::uint32_t>(off);
      return axis == LookupAxis::Column ? read_table_cell(table, along, pick - 1U, arena, registry, ctx)
                                        : read_table_cell(table, pick - 1U, along, arena, registry, ctx);
    }
    const std::size_t flat = axis == LookupAxis::Column
                                 ? (off * static_cast<std::size_t>(cols)) + static_cast<std::size_t>(pick - 1U)
                                 : (static_cast<std::size_t>(pick - 1U) * static_cast<std::size_t>(cols)) + off;
    if (flat >= cells.size()) {
      return Value::error(ErrorCode::Ref);
    }
    return cells[flat];
  };

  // The declared length of the scanned line, for the approximate search over a whole-axis table.
  const std::size_t declared_line =
      read_by_reference
          ? static_cast<std::size_t>(axis == LookupAxis::Column ? table.declared.rows() : table.declared.cols())
          : 0U;
  const auto scan = [&](const Value& query) {
    return lookup_scan(*scan_cells, scan_rows, scan_cols, axis, query, approximate, ctx.excel_profile(), false,
                       declared_line);
  };
  if (index_array) {
    const std::size_t off = scan(lookup);
    if (off == SIZE_MAX) {
      return Value::error(ErrorCode::NA);
    }
    const std::size_t count = static_cast<std::size_t>(index_val.as_array_rows()) * index_val.as_array_cols();
    const Value* elements = index_val.as_array_cells();
    const Value* error_element = nullptr;
    for (std::size_t i = 0; i < count; ++i) {
      std::uint32_t unused = 0U;
      std::optional<Value> bad = index_error(elements[i], &unused);
      if (!bad) {
        continue;
      }
      if (!elements[i].is_error()) {
        return *bad;
      }
      if (error_element == nullptr) {
        error_element = &elements[i];
      }
    }
    if (error_element != nullptr) {
      return *error_element;
    }
    return map_query_lanes(index_val, arena, [&](const Value& element) -> Value {
      std::uint32_t pick = 0U;
      index_error(element, &pick);
      return fetch(off, pick);
    });
  }
  const auto lane = [&](const Value& query) {
    const std::size_t off = scan(query);
    return off == SIZE_MAX ? Value::error(ErrorCode::NA) : fetch(off, result_index);
  };
  if (array_lookup) {
    return map_query_lanes(lookup, arena, [&](const Value& query) { return query.is_error() ? query : lane(query); });
  }
  return lane(lookup);
}

Value match_lookup_one(const std::vector<Value>& cells, const Value& lookup, int match_type, ExcelProfile profile,
                       std::size_t declared = 0U) {
  if (lookup.is_error()) {
    return lookup;
  }
  if (cells.empty()) {
    return Value::error(ErrorCode::NA);
  }
  // Exact mode honours wildcards for text; approximate treats them as literal.
  const std::size_t pos = lookup_scan(cells, static_cast<std::uint32_t>(cells.size()), 1U, LookupAxis::Column, lookup,
                                      match_type != 0, profile, match_type < 0, declared);
  return pos == SIZE_MAX ? Value::error(ErrorCode::NA) : Value::number(static_cast<double>(pos + 1));
}

}  // namespace

// MATCH(lookup_value, lookup_array, [match_type])
//
// Returns the 1-based position of `lookup_value` inside the 1-D
// `lookup_array`. `lookup_array` must be a `RangeOp(Ref, Ref)` or a single
// `Ref` with a 1-D shape (row vector or column vector); a 2-D range yields
// `#N/A`.
//
// match_type semantics:
//   *  1 (default) - ascending array, binary search (`approximate_search`).
//      Wildcards are NOT honoured.
//   *  0           - exact match with DOS-style wildcards (`*`, `?`, `~`)
//      for text targets; first hit wins. No match -> `#N/A`.
//   * -1           - descending array, linear scan (`descending_scan`).
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
  const auto by_reference = resolve_reference_table(call.as_call_arg(1), ctx, &table, TableWalk::kDeferred);
  if (by_reference && by_reference.value() && table.declared.rows() != 1U && table.declared.cols() != 1U) {
    return fail_lookup(lookup, arena, Value::error(ErrorCode::NA));
  }
  // A lone value arrives as the error the lazy table's lone-value policy answers (`MATCH(1,1,0)` is #N/A).
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
      return fail_lookup(lookup, arena, Value::error(ErrorCode::Ref));
    }
    auto mt_index = coerce_to_index_number(mt_val);
    if (!mt_index) {
      return fail_lookup(lookup, arena, Value::error(mt_index.error()));
    }
    const double mt_raw = mt_index.value();
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

  // A one-dimensional reference is searched over its declared length.
  const std::size_t declared = by_reference && by_reference.value()
                                   ? static_cast<std::size_t>(table.declared.rows()) * table.declared.cols()
                                   : 0U;
  if (array_lookup) {
    return map_query_lanes(lookup, arena, [&](const Value& query) {
      return match_lookup_one(cells, query, match_type, ctx.excel_profile(), declared);
    });
  }
  return match_lookup_one(cells, lookup, match_type, ctx.excel_profile(), declared);
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
//   * An array `lookup_value` is searched element by element, as VLOOKUP's
//     is: the result takes its shape and an error element stays in its lane.
//   * `lookup_value` that never sorts before any vector cell: `#N/A`. A
//     blank `lookup_value` is the number 0.
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

  auto lookup_resolved = resolve_range_arg(call.as_call_arg(1), arena, registry, ctx);
  if (!lookup_resolved) {
    return fail_lookup(lookup, arena, Value::error(lookup_resolved.error()));
  }
  const std::uint32_t lrows = lookup_resolved.value().rows;
  const std::uint32_t lcols = lookup_resolved.value().cols;
  std::vector<Value> lookup_cells = std::move(lookup_resolved.value().cells);

  // The orientation is the declared one: `A:C` is taller than wide however
  // few of its rows hold values.
  std::uint32_t lshape_rows = lrows;
  std::uint32_t lshape_cols = lcols;
  static_reference_shape(call.as_call_arg(1), ctx, &lshape_rows, &lshape_cols);
  if (lshape_rows == 0U || lshape_cols == 0U) {
    return fail_lookup(lookup, arena, Value::error(ErrorCode::Ref));
  }
  const LookupAxis axis = lshape_rows >= lshape_cols ? LookupAxis::Column : LookupAxis::Row;
  // A whole-axis vector with nothing populated walks no cells; every declared
  // cell is blank, so nothing matches.
  const std::size_t declared_line = axis == LookupAxis::Column ? lshape_rows : lshape_cols;
  const auto scan = [&](const Value& query) {
    return lookup_cells.empty() ? SIZE_MAX
                                : lookup_scan(lookup_cells, lrows, lcols, axis, query,
                                              /*approximate=*/true, ctx.excel_profile(), false, declared_line);
  };

  // The result vector is read once; a failure to read it surfaces only on a
  // query that matched.
  std::vector<Value> result_cells;
  std::uint32_t rrows = 0U;
  std::uint32_t rcols = 0U;
  std::optional<Value> result_error;
  LookupAxis raxis = LookupAxis::Column;
  std::uint32_t rlength = 0U;
  if (arity == 3) {
    auto result_resolved = resolve_range_arg(call.as_call_arg(2), arena, registry, ctx);
    if (!result_resolved) {
      result_error = Value::error(result_resolved.error());
    } else {
      rrows = result_resolved.value().rows;
      rcols = result_resolved.value().cols;
      result_cells = std::move(result_resolved.value().cells);
      // Index the result vector along its own long axis. Excel treats the
      // result vector as parallel to the lookup vector, so the "axis
      // position" maps to the same offset regardless of orientation.
      std::uint32_t rshape_rows = rrows;
      std::uint32_t rshape_cols = rcols;
      static_reference_shape(call.as_call_arg(2), ctx, &rshape_rows, &rshape_cols);
      if (rshape_rows == 0U || rshape_cols == 0U) {
        result_error = Value::error(ErrorCode::Ref);
      }
      raxis = rshape_rows >= rshape_cols ? LookupAxis::Column : LookupAxis::Row;
      rlength = raxis == LookupAxis::Column ? rshape_rows : rshape_cols;
    }
  }

  const auto lane = [&](const Value& query) -> Value {
    const std::size_t off = scan(query);
    if (off == SIZE_MAX) {
      return Value::error(ErrorCode::NA);
    }
    if (arity == 2) {
      // Vector form (1-D input): result is the matched cell of the lookup
      // vector itself. Array form (2-D input): result is the corresponding
      // cell of the last column (taller-than-wide) or last row (wider-than-
      // tall). The two paths agree on a 1-D input because last-col == col 0
      // and last-row == row 0 in that case.
      const std::size_t flat = axis == LookupAxis::Column
                                   ? (off * static_cast<std::size_t>(lcols)) + (static_cast<std::size_t>(lcols) - 1U)
                                   : ((static_cast<std::size_t>(lrows) - 1U) * static_cast<std::size_t>(lcols)) + off;
      return flat < lookup_cells.size() ? lookup_cells[flat] : Value::error(ErrorCode::NA);
    }
    if (result_error) {
      return *result_error;
    }
    const std::size_t flat = raxis == LookupAxis::Column ? (off * static_cast<std::size_t>(rcols)) : off;
    if (flat < result_cells.size()) {
      return result_cells[flat];
    }
    // Past the walked extent but inside the declared vector the cell is blank.
    return off < rlength ? Value::blank(BlankGridProjection::kReferenceGridZero) : Value::error(ErrorCode::NA);
  };
  if (lookup.is_array()) {
    return map_query_lanes(lookup, arena, [&](const Value& query) { return query.is_error() ? query : lane(query); });
  }
  return lane(lookup);
}

}  // namespace eval
}  // namespace formulon
