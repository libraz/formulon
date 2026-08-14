//
// Implementation of the classic lookup-family lazy impls (`CHOOSE`,
// `INDEX`, `MATCH`, `VLOOKUP`, `HLOOKUP`). See `lookups/classic.h` for
// the dispatch-table contract and `eval/lazy_impls.h` for the shared
// vocabulary.

#include "eval/lookups/classic.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <string_view>
#include <vector>

#include "eval/coerce.h"
#include "eval/criteria.h"
#include "eval/dynamic_array/common.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/jp_fold.h"
#include "eval/lazy_impls.h"
#include "eval/name_env_resolve.h"
#include "eval/range_args.h"
#include "parser/ast.h"
#include "utils/arena.h"
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

// Materialises the whole `col`-th column (0-based) of a row-major `cells`
// rectangle (`rows` x `cols`) as a vertical `rows` x 1 `Value::Array`. Used
// by `INDEX(array, 0, col)` which Excel 365 spills as a column. Returns
// `#NUM!` on arena exhaustion.
Value index_whole_column(const std::vector<Value>& cells, std::uint32_t rows, std::uint32_t cols, std::uint32_t col,
                         Arena& arena) {
  Value* buffer = nullptr;
  ArrayValue* out = dynamic_array::allocate_array_value(rows, 1U, arena, buffer, kMaxDerivedArrayCells);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::uint32_t r = 0; r < rows; ++r) {
    buffer[r] = cells[(static_cast<std::size_t>(r) * cols) + col];
  }
  return Value::array(out);
}

// Materialises a whole row-major `cells` rectangle (`rows` x `cols`) as a
// `Value::Array` of the same shape. Used by the INDEX forms that select no
// single axis — `INDEX(array, 0, 0)` and its two-argument spelling
// `INDEX(array, 0)`. Returns `#NUM!` on arena exhaustion.
Value index_whole_array(const std::vector<Value>& cells, std::uint32_t rows, std::uint32_t cols, Arena& arena) {
  Value* buffer = nullptr;
  ArrayValue* out = dynamic_array::allocate_array_value(rows, cols, arena, buffer, kMaxDerivedArrayCells);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  const std::size_t total = static_cast<std::size_t>(rows) * cols;
  for (std::size_t i = 0; i < total; ++i) {
    buffer[i] = cells[i];
  }
  return Value::array(out);
}

// Materialises the whole `row`-th row (0-based) of a row-major `cells`
// rectangle (`rows` x `cols`) as a horizontal 1 x `cols` `Value::Array`.
// Used by `INDEX(array, row, 0)` which Excel 365 spills as a row. Returns
// `#NUM!` on arena exhaustion.
Value index_whole_row(const std::vector<Value>& cells, std::uint32_t cols, std::uint32_t row, Arena& arena) {
  Value* buffer = nullptr;
  ArrayValue* out = dynamic_array::allocate_array_value(1U, cols, arena, buffer, kMaxDerivedArrayCells);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::uint32_t c = 0; c < cols; ++c) {
    buffer[c] = cells[(static_cast<std::size_t>(row) * cols) + c];
  }
  return Value::array(out);
}

// A selector is either a scalar or an already materialised value array. The
// lazy lookup implementations keep these two cases separate so that a
// scalar INDEX / CHOOSE call retains its historical short-circuit and result
// shape, while an array selector can use the same row-major broadcast rules
// as the dynamic-array dispatcher without re-evaluating its AST.
struct SelectorView {
  const Value* scalar = nullptr;
  const ArrayValue* array = nullptr;
  std::uint32_t rows = 1U;
  std::uint32_t cols = 1U;

  bool is_array() const noexcept { return array != nullptr; }
};

SelectorView make_selector_view(const Value& value) {
  if (!value.is_array()) {
    return SelectorView{&value, nullptr, 1U, 1U};
  }
  const ArrayValue* array = value.as_array();
  return SelectorView{nullptr, array, array->rows, array->cols};
}

// Returns the value supplied by a selector at an output coordinate. A
// non-1 selector axis shorter than the output rectangle is a lane-local
// #N/A, rather than a global shape error. `missing` distinguishes that
// synthetic #N/A from an actual error cell, which keeps the precedence logic
// in the INDEX compositor explicit.
Value selector_at(const SelectorView& selector, std::uint32_t row, std::uint32_t col, bool* missing) {
  *missing = false;
  if (!selector.is_array()) {
    return *selector.scalar;
  }
  const std::uint32_t source_row = selector.rows == 1U ? 0U : row;
  const std::uint32_t source_col = selector.cols == 1U ? 0U : col;
  if (source_row >= selector.rows || source_col >= selector.cols) {
    *missing = true;
    return Value::error(ErrorCode::NA);
  }
  return selector.array->cells[static_cast<std::size_t>(source_row) * selector.cols + source_col];
}

std::uint32_t selector_rows(const SelectorView& row, const SelectorView& col) {
  return std::max(row.rows, col.rows);
}

std::uint32_t selector_cols(const SelectorView& row, const SelectorView& col) {
  return std::max(row.cols, col.cols);
}

// The array compositor is a materialisation boundary. Raw-reference blanks
// copied from a source range must therefore become value-array blanks here so
// that downstream COUNTA sees the occupied cells, while scalar INDEX / CHOOSE
// paths continue to return the original reference blank unchanged.
Value promote_array_result_cell(const Value& value) {
  return value.promote_reference_blank_to_value_array();
}

ArrayValue* allocate_lookup_array(std::uint32_t rows, std::uint32_t cols, Arena& arena, Value*& buffer) {
  return dynamic_array::allocate_array_value(rows, cols, arena, buffer, kMaxDerivedArrayCells);
}

enum class IndexAxisState : std::uint8_t { kValid, kError };

struct DecodedIndex {
  IndexAxisState state = IndexAxisState::kValid;
  std::uint32_t index = 0U;
  ErrorCode error = ErrorCode::Value;
};

DecodedIndex decode_index_cell(const Value& value) {
  if (value.is_error()) {
    return DecodedIndex{IndexAxisState::kError, 0U, value.as_error()};
  }
  auto number = coerce_to_number(value);
  if (!number) {
    return DecodedIndex{IndexAxisState::kError, 0U, number.error()};
  }
  const double original = number.value();
  const double raw = truncate_index(original);
  // A selector below 1 never names a row / column: negatives and sub-1
  // fractions are both `#VALUE!`. Only an exact zero means "whole spanned
  // dimension", so the fraction check tests `original` rather than `raw`.
  if (original < 0.0 || (raw == 0.0 && original != 0.0)) {
    return DecodedIndex{IndexAxisState::kError, 0U, ErrorCode::Value};
  }
  // Avoid an implementation-defined narrowing conversion for a gigantic
  // finite selector. All admitted source dimensions are far below this
  // bound, so the eventual result is the ordinary INDEX #REF! domain error.
  if (raw > static_cast<double>(std::numeric_limits<std::uint32_t>::max())) {
    return DecodedIndex{IndexAxisState::kError, 0U, ErrorCode::Ref};
  }
  return DecodedIndex{IndexAxisState::kValid, static_cast<std::uint32_t>(raw), ErrorCode::Value};
}

std::vector<DecodedIndex> decode_selector(const SelectorView& selector) {
  const std::size_t count = static_cast<std::size_t>(selector.rows) * selector.cols;
  std::vector<DecodedIndex> decoded;
  decoded.reserve(count);
  if (!selector.is_array()) {
    decoded.push_back(decode_index_cell(*selector.scalar));
    return decoded;
  }
  for (std::size_t i = 0; i < count; ++i) {
    decoded.push_back(decode_index_cell(selector.array->cells[i]));
  }
  return decoded;
}

const DecodedIndex& decoded_selector_at(const SelectorView& selector, const std::vector<DecodedIndex>& decoded,
                                        std::uint32_t row, std::uint32_t col, bool* missing) {
  *missing = false;
  const std::uint32_t source_row = selector.rows == 1U ? 0U : row;
  const std::uint32_t source_col = selector.cols == 1U ? 0U : col;
  if (source_row >= selector.rows || source_col >= selector.cols) {
    *missing = true;
    static const DecodedIndex kMissing{IndexAxisState::kError, 0U, ErrorCode::NA};
    return kMissing;
  }
  return decoded[static_cast<std::size_t>(source_row) * selector.cols + source_col];
}

enum class IndexTileKind : std::uint8_t { kScalar, kRow, kColumn, kWhole };

struct IndexTile {
  IndexTileKind kind = IndexTileKind::kScalar;
  std::uint32_t rows = 1U;
  std::uint32_t cols = 1U;
  std::uint32_t source_row = 0U;
};

IndexTile index_tile_for(std::uint32_t rows, std::uint32_t cols, std::uint32_t row_idx, std::uint32_t col_idx,
                         bool col_explicit) {
  IndexTile tile;
  if (!col_explicit) {
    if (rows == 1U && cols == 1U) {
      return tile;
    }
    if (rows == 1U) {
      if (row_idx == 0U) {
        tile.kind = IndexTileKind::kRow;
        tile.cols = cols;
      }
      return tile;
    }
    if (cols == 1U) {
      if (row_idx == 0U) {
        tile.kind = IndexTileKind::kColumn;
        tile.rows = rows;
      }
      return tile;
    }
    // The two-argument form on a 2-D source reads the omitted column
    // argument as zero, so a row selector returns that complete row and a
    // zero selector spans every row as well -- the whole array, exactly as
    // the explicit `INDEX(src, 0, 0)`. Matches the scalar path.
    if (row_idx == 0U) {
      tile.kind = IndexTileKind::kWhole;
      tile.rows = rows;
      tile.cols = cols;
      return tile;
    }
    tile.kind = IndexTileKind::kRow;
    tile.cols = cols;
    tile.source_row = row_idx - 1U;
    return tile;
  }

  if (rows == 1U && cols == 1U) {
    return tile;
  }
  if (rows == 1U) {
    if (col_idx == 0U) {
      tile.kind = IndexTileKind::kRow;
      tile.cols = cols;
    }
    return tile;
  }
  if (cols == 1U) {
    if (row_idx == 0U) {
      tile.kind = IndexTileKind::kColumn;
      tile.rows = rows;
    }
    return tile;
  }
  if (row_idx == 0U && col_idx == 0U) {
    tile.kind = IndexTileKind::kWhole;
    tile.rows = rows;
    tile.cols = cols;
  } else if (row_idx == 0U) {
    tile.kind = IndexTileKind::kColumn;
    tile.rows = rows;
  } else if (col_idx == 0U) {
    tile.kind = IndexTileKind::kRow;
    tile.cols = cols;
  }
  return tile;
}

Value index_tile_cell(const std::vector<Value>& cells, std::uint32_t source_rows, std::uint32_t source_cols,
                      std::uint32_t row_idx, std::uint32_t col_idx, bool col_explicit, const IndexTile& tile,
                      std::uint32_t output_row, std::uint32_t output_col) {
  if (tile.kind == IndexTileKind::kScalar) {
    const std::uint32_t source_row = col_explicit ? row_idx - 1U : (source_rows == 1U ? 0U : row_idx - 1U);
    const std::uint32_t source_col = col_explicit ? col_idx - 1U : (source_rows == 1U ? row_idx - 1U : 0U);
    const std::size_t flat = static_cast<std::size_t>(source_row) * source_cols + source_col;
    return flat < cells.size() ? cells[flat] : Value::error(ErrorCode::Ref);
  }

  std::uint32_t source_row = 0U;
  std::uint32_t source_col = 0U;
  switch (tile.kind) {
    case IndexTileKind::kRow:
      source_row = col_explicit ? row_idx - 1U : tile.source_row;
      source_col = output_col;
      break;
    case IndexTileKind::kColumn:
      source_row = output_row;
      source_col = col_explicit ? col_idx - 1U : 0U;
      break;
    case IndexTileKind::kWhole:
      source_row = output_row;
      source_col = output_col;
      break;
    case IndexTileKind::kScalar:
      break;
  }
  if (source_row >= source_rows || source_col >= source_cols) {
    return Value::error(ErrorCode::NA);
  }
  return cells[static_cast<std::size_t>(source_row) * source_cols + source_col];
}

// ---------------------------------------------------------------------------
// CHOOSE / INDEX / MATCH / VLOOKUP / HLOOKUP (lookup & reference)
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
  std::size_t best = SIZE_MAX;
  for (std::size_t i = 0; i < n; ++i) {
    const Value& cell = cell_at(i);
    int cmp = 0;  // sign of (cell - lookup_value)
    bool comparable = false;
    if (lookup_value.is_text() && cell.is_text()) {
      // Normalise (see exact-mode branch above) before the ASCII
      // case-insensitive compare so kana / half-width voicing variants order
      // together.
      cmp = strings::case_insensitive_compare(lookup_text_cmp_key(cell.as_text(), profile),
                                              lookup_text_cmp_key(lookup_value.as_text(), profile));
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
  if (!resolve_table_array(call.as_call_arg(1), arena, registry, ctx, &cells, &range_err, &rows, &cols)) {
    if (array_lookup) {
      return map_lookup_query_array(lookup, arena, [range_err](const Value& query) {
        return query.is_error() ? query : Value::error(range_err);
      });
    }
    return Value::error(range_err);
  }
  if (rows == 0U || cols == 0U) {
    if (array_lookup) {
      return map_lookup_query_array(
          lookup, arena, [](const Value& query) { return query.is_error() ? query : Value::error(ErrorCode::Ref); });
    }
    return Value::error(ErrorCode::Ref);
  }

  const Value index_val = eval_node(call.as_call_arg(2), arena, registry, ctx);
  if (index_val.is_error()) {
    if (array_lookup) {
      return map_lookup_query_array(lookup, arena,
                                    [&index_val](const Value& query) { return query.is_error() ? query : index_val; });
    }
    return index_val;
  }
  auto index_num = coerce_to_number(index_val);
  if (!index_num) {
    if (array_lookup) {
      const ErrorCode index_err = index_num.error();
      return map_lookup_query_array(lookup, arena, [index_err](const Value& query) {
        return query.is_error() ? query : Value::error(index_err);
      });
    }
    return Value::error(index_num.error());
  }
  const double index_raw = truncate_index(index_num.value());
  if (index_raw < 1.0) {
    if (array_lookup) {
      return map_lookup_query_array(
          lookup, arena, [](const Value& query) { return query.is_error() ? query : Value::error(ErrorCode::Value); });
    }
    return Value::error(ErrorCode::Value);
  }
  const std::uint32_t result_extent = axis == LookupAxis::Column ? cols : rows;
  if (index_raw > static_cast<double>(result_extent)) {
    if (array_lookup) {
      return map_lookup_query_array(
          lookup, arena, [](const Value& query) { return query.is_error() ? query : Value::error(ErrorCode::Ref); });
    }
    return Value::error(ErrorCode::Ref);
  }
  const auto result_index = static_cast<std::uint32_t>(index_raw);

  bool approximate = true;
  if (arity == 4) {
    const Value rl_val = eval_node(call.as_call_arg(3), arena, registry, ctx);
    if (rl_val.is_error()) {
      if (array_lookup) {
        return map_lookup_query_array(lookup, arena,
                                      [&rl_val](const Value& query) { return query.is_error() ? query : rl_val; });
      }
      return rl_val;
    }
    auto rl_bool = coerce_to_bool(rl_val);
    if (!rl_bool) {
      if (array_lookup) {
        const ErrorCode range_lookup_err = rl_bool.error();
        return map_lookup_query_array(lookup, arena, [range_lookup_err](const Value& query) {
          return query.is_error() ? query : Value::error(range_lookup_err);
        });
      }
      return Value::error(rl_bool.error());
    }
    approximate = rl_bool.value();
  }

  if (array_lookup) {
    return map_lookup_query_array(lookup, arena, [&](const Value& query) {
      if (query.is_error()) {
        return query;
      }
      const std::size_t off = lookup_scan(cells, rows, cols, axis, query, approximate, ctx.excel_profile());
      if (off == SIZE_MAX) {
        return Value::error(ErrorCode::NA);
      }
      const std::size_t flat =
          axis == LookupAxis::Column
              ? (off * static_cast<std::size_t>(cols)) + static_cast<std::size_t>(result_index - 1U)
              : (static_cast<std::size_t>(result_index - 1U) * static_cast<std::size_t>(cols)) + off;
      if (flat >= cells.size()) {
        return Value::error(ErrorCode::Ref);
      }
      return cells[flat];
    });
  }

  const std::size_t off = lookup_scan(cells, rows, cols, axis, lookup, approximate, ctx.excel_profile());
  if (off == SIZE_MAX) {
    return Value::error(ErrorCode::NA);
  }
  const std::size_t flat = axis == LookupAxis::Column
                               ? (off * static_cast<std::size_t>(cols)) + static_cast<std::size_t>(result_index - 1U)
                               : (static_cast<std::size_t>(result_index - 1U) * static_cast<std::size_t>(cols)) + off;
  if (flat >= cells.size()) {
    return Value::error(ErrorCode::Ref);
  }
  return cells[flat];
}

bool index_domain_valid(std::uint32_t rows, std::uint32_t cols, std::uint32_t row_idx, std::uint32_t col_idx,
                        bool col_explicit) {
  if (!col_explicit) {
    if (rows == 1U && cols == 1U) {
      return row_idx == 0U || row_idx == 1U;
    }
    if (rows == 1U) {
      return row_idx <= cols;
    }
    if (cols == 1U) {
      return row_idx <= rows;
    }
    return row_idx <= rows;
  }
  if (rows == 1U && cols == 1U) {
    return (row_idx == 0U || row_idx == 1U) && (col_idx == 0U || col_idx == 1U);
  }
  if (rows == 1U) {
    return (row_idx == 0U || row_idx == 1U) && col_idx <= cols;
  }
  if (cols == 1U) {
    return row_idx <= rows && (col_idx == 0U || col_idx == 1U);
  }
  return row_idx <= rows && col_idx <= cols;
}

Value eval_index_array_selector(const std::vector<Value>& cells, std::uint32_t source_rows, std::uint32_t source_cols,
                                bool source_ok, ErrorCode source_error, const Value& row_value, const Value* col_value,
                                bool col_explicit, Arena& arena) {
  const SelectorView row_selector = make_selector_view(row_value);
  const Value implicit_col = Value::number(0.0);
  const SelectorView col_selector = col_explicit ? make_selector_view(*col_value) : make_selector_view(implicit_col);
  const std::vector<DecodedIndex> row_decoded = decode_selector(row_selector);
  const std::vector<DecodedIndex> col_decoded = decode_selector(col_selector);

  const std::uint32_t selector_out_rows = selector_rows(row_selector, col_selector);
  const std::uint32_t selector_out_cols = selector_cols(row_selector, col_selector);
  std::uint32_t out_rows = selector_out_rows;
  std::uint32_t out_cols = selector_out_cols;

  // A zero selector is a tile in scalar INDEX. In array-selector mode the
  // tile is composed directly into the one flat output rectangle; no nested
  // ArrayValue is allocated per lane. Scan the selector rectangle once to
  // discover the largest tile shape that contributes to that rectangle.
  if (source_ok) {
    for (std::uint32_t r = 0; r < selector_out_rows; ++r) {
      for (std::uint32_t c = 0; c < selector_out_cols; ++c) {
        bool row_missing = false;
        const DecodedIndex& row = decoded_selector_at(row_selector, row_decoded, r, c, &row_missing);
        if (row_missing || row.state == IndexAxisState::kError) {
          continue;
        }
        DecodedIndex col{IndexAxisState::kValid, 0U, ErrorCode::Value};
        bool col_missing = false;
        if (col_explicit) {
          col = decoded_selector_at(col_selector, col_decoded, r, c, &col_missing);
          if (col_missing || col.state == IndexAxisState::kError) {
            continue;
          }
        }
        if (!index_domain_valid(source_rows, source_cols, row.index, col.index, col_explicit)) {
          continue;
        }
        const IndexTile tile = index_tile_for(source_rows, source_cols, row.index, col.index, col_explicit);
        out_rows = std::max(out_rows, tile.rows);
        out_cols = std::max(out_cols, tile.cols);
      }
    }
  }

  Value* output_cells = nullptr;
  ArrayValue* output = allocate_lookup_array(out_rows, out_cols, arena, output_cells);
  if (output == nullptr) {
    return Value::error(ErrorCode::Num);
  }

  for (std::uint32_t r = 0; r < out_rows; ++r) {
    for (std::uint32_t c = 0; c < out_cols; ++c) {
      Value result = Value::error(ErrorCode::NA);
      bool row_missing = false;
      const DecodedIndex& row = decoded_selector_at(row_selector, row_decoded, r, c, &row_missing);
      if (row_missing) {
        output_cells[static_cast<std::size_t>(r) * out_cols + c] = result;
        continue;
      }
      if (row.state == IndexAxisState::kError) {
        output_cells[static_cast<std::size_t>(r) * out_cols + c] = Value::error(row.error);
        continue;
      }

      DecodedIndex col{IndexAxisState::kValid, 0U, ErrorCode::Value};
      bool col_missing = false;
      if (col_explicit) {
        col = decoded_selector_at(col_selector, col_decoded, r, c, &col_missing);
        if (col_missing) {
          output_cells[static_cast<std::size_t>(r) * out_cols + c] = result;
          continue;
        }
        if (col.state == IndexAxisState::kError) {
          output_cells[static_cast<std::size_t>(r) * out_cols + c] = Value::error(col.error);
          continue;
        }
      }

      // Source errors are global to INDEX, but an array selector still
      // determines the output rectangle. Selector errors above retain their
      // lane-local precedence over that source error.
      if (!source_ok) {
        output_cells[static_cast<std::size_t>(r) * out_cols + c] = Value::error(source_error);
        continue;
      }
      if (!index_domain_valid(source_rows, source_cols, row.index, col.index, col_explicit)) {
        output_cells[static_cast<std::size_t>(r) * out_cols + c] = Value::error(ErrorCode::Ref);
        continue;
      }
      const IndexTile tile = index_tile_for(source_rows, source_cols, row.index, col.index, col_explicit);
      result = index_tile_cell(cells, source_rows, source_cols, row.index, col.index, col_explicit, tile, r, c);
      output_cells[static_cast<std::size_t>(r) * out_cols + c] = promote_array_result_cell(result);
    }
  }
  return Value::array(output);
}

}  // namespace

// Array-index CHOOSE compositor. The index has already been evaluated by the
// caller; every branch is evaluated exactly once, left-to-right, and selected
// cells are composed into one value array.
Value eval_choose_array_index_lazy(const parser::AstNode& call, const Value& index_value, Arena& arena,
                                   const FunctionRegistry& registry, const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2) {
    return Value::error(ErrorCode::Value);
  }
  if (!index_value.is_array()) {
    return Value::error(ErrorCode::Value);
  }

  const SelectorView index_selector = make_selector_view(index_value);

  // Array-index CHOOSE evaluates every branch once, left-to-right, then
  // selects cached scalar/array cells lane-by-lane. This is deliberately a
  // separate path from scalar CHOOSE: only the scalar form may skip
  // unselected branches.
  std::vector<Value> branches;
  branches.reserve(arity - 1U);
  for (std::uint32_t i = 1U; i < arity; ++i) {
    branches.push_back(eval_node(call.as_call_arg(i), arena, registry, ctx));
  }

  std::uint32_t out_rows = index_selector.rows;
  std::uint32_t out_cols = index_selector.cols;
  std::vector<SelectorView> branch_views;
  branch_views.reserve(branches.size());
  for (const Value& branch : branches) {
    branch_views.push_back(make_selector_view(branch));
    out_rows = std::max(out_rows, branch_views.back().rows);
    out_cols = std::max(out_cols, branch_views.back().cols);
  }

  Value* output_cells = nullptr;
  ArrayValue* output = allocate_lookup_array(out_rows, out_cols, arena, output_cells);
  if (output == nullptr) {
    return Value::error(ErrorCode::Num);
  }

  const std::vector<DecodedIndex> decoded_indices = decode_selector(index_selector);
  for (std::uint32_t r = 0; r < out_rows; ++r) {
    for (std::uint32_t c = 0; c < out_cols; ++c) {
      const std::size_t output_index = static_cast<std::size_t>(r) * out_cols + c;
      bool index_missing = false;
      const DecodedIndex& decoded = decoded_selector_at(index_selector, decoded_indices, r, c, &index_missing);
      if (index_missing) {
        output_cells[output_index] = Value::error(ErrorCode::NA);
        continue;
      }
      if (decoded.state == IndexAxisState::kError) {
        output_cells[output_index] = Value::error(decoded.error);
        continue;
      }
      if (decoded.index < 1U || decoded.index > branches.size()) {
        output_cells[output_index] = Value::error(ErrorCode::Value);
        continue;
      }

      const SelectorView& branch = branch_views[decoded.index - 1U];
      bool branch_missing = false;
      const Value selected = selector_at(branch, r, c, &branch_missing);
      if (branch_missing) {
        output_cells[output_index] = Value::error(ErrorCode::NA);
      } else {
        // The compositor, rather than scalar CHOOSE, owns this value-array
        // boundary. Promote reference blanks only after the selected cell
        // has been chosen so an unselected branch's cells cannot leak into
        // the output.
        output_cells[output_index] = promote_array_result_cell(selected);
      }
    }
  }
  return Value::array(output);
}

// CHOOSE(index_num, value1, value2, ...)
//
// Evaluates `index_num`, truncates to int, and returns only the corresponding
// argument subtree (`CHOOSE(2, a, b, c)` returns `b` and never touches `a`
// or `c`). Out-of-range indices yield `#VALUE!`; a numeric coercion failure
// on `index_num` also yields `#VALUE!`. Errors in `index_num` propagate.
// Errors in the chosen value also propagate; unselected arguments are never
// evaluated for the scalar-index path.
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
  if (idx_val.is_array()) {
    return eval_choose_array_index_lazy(call, idx_val, arena, registry, ctx);
  }
  auto idx_num = coerce_to_number(idx_val);
  if (!idx_num) {
    return Value::error(idx_num.error());
  }
  // Excel truncates (toward zero) rather than rounds: CHOOSE(2.9, ...)
  // selects the 2nd value, not the 3rd.
  const double raw = truncate_index(idx_num.value());
  if (!(raw >= 1.0 && raw <= static_cast<double>(arity - 1))) {
    return Value::error(ErrorCode::Value);
  }
  const auto n = static_cast<std::uint32_t>(raw);
  return eval_node(call.as_call_arg(n), arena, registry, ctx);
}

// INDEX(array, row_num, [column_num])
//
// Returns a cell — or a whole row / column — from `array` by 1-based
// (row_num, column_num). The source array must be a `RangeOp(Ref, Ref)` or
// a single `Ref`; anything else is `#VALUE!`. Out-of-bounds indices are
// `#REF!`. Negative or non-coercible indices are `#VALUE!`.
//
// Shape disambiguation for the 2-arg form: if the array is 1-D (rows == 1
// or cols == 1), the sole index selects along the non-singleton dimension.
// For a 2-D array with only two args provided, `row_num` selects the row
// and Excel 365 spills the entire selected row — we materialise that row as
// a horizontal `Value::Array` so the spill committer places it on the sheet.
//
// Zero indices spill the whole spanned dimension, matching Excel 365:
//   * `INDEX(array, 0, col)` / `INDEX(col_vector, 0)` -> the whole column
//     `col` as a vertical array.
//   * `INDEX(array, row, 0)` / `INDEX(array, row)` (2-D, 2-arg) /
//     `INDEX(row_vector, 0)` -> the whole row `row` as a horizontal array.
//   * `INDEX(array, 0, 0)` -> the whole array.
// A 1x1 source collapses any zero index to its sole cell.
Value eval_index_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity != 2 && arity != 3) {
    return Value::error(ErrorCode::Value);
  }
  auto resolved = resolve_range_arg(call.as_call_arg(0), arena, registry, ctx);
  const bool source_ok = resolved.has_value();
  const ErrorCode source_error = source_ok ? ErrorCode::Value : resolved.error();
  std::uint32_t rows = 0U;
  std::uint32_t cols = 0U;
  std::vector<Value> cells;
  if (source_ok) {
    rows = resolved.value().rows;
    cols = resolved.value().cols;
    cells = std::move(resolved.value().cells);
    if (rows == 0U || cols == 0U) {
      // Defensive: expand_range always produces a positive rectangle today.
      rows = 0U;
      cols = 0U;
    }
  }

  // row_num is required (arity 2 or 3), col_num is optional.
  const Value row_val = eval_node(call.as_call_arg(1), arena, registry, ctx);
  Value col_val = Value::number(0.0);
  if (arity == 3) {
    col_val = eval_node(call.as_call_arg(2), arena, registry, ctx);
  }
  if (row_val.is_array() || col_val.is_array()) {
    return eval_index_array_selector(cells, rows, cols, source_ok && rows != 0U && cols != 0U, source_error, row_val,
                                     arity == 3 ? &col_val : nullptr, arity == 3, arena);
  }
  if (!source_ok || rows == 0U || cols == 0U) {
    return Value::error(source_error == ErrorCode::Value ? ErrorCode::Ref : source_error);
  }
  if (row_val.is_error()) {
    return row_val;
  }
  auto row_num_exp = coerce_to_number(row_val);
  if (!row_num_exp) {
    return Value::error(row_num_exp.error());
  }
  const double row_orig = row_num_exp.value();
  const double row_raw = truncate_index(row_orig);
  if (row_orig < 0.0) {
    return Value::error(ErrorCode::Value);
  }
  // Fractional sub-1 values (`row_num` in (0, 1)) truncate to 0 but Excel
  // rejects them with #VALUE!. The "whole-vector" / "whole-array" sentinel
  // meaning of row_num == 0 only applies when the user explicitly passed 0.
  if (row_raw == 0.0 && row_orig != 0.0) {
    return Value::error(ErrorCode::Value);
  }
  const auto row_idx = static_cast<std::uint32_t>(row_raw);

  std::uint32_t col_idx = 0;
  bool col_explicit = false;
  if (arity == 3) {
    if (col_val.is_error()) {
      return col_val;
    }
    auto col_num_exp = coerce_to_number(col_val);
    if (!col_num_exp) {
      return Value::error(col_num_exp.error());
    }
    const double col_orig = col_num_exp.value();
    const double col_raw = truncate_index(col_orig);
    if (col_orig < 0.0) {
      return Value::error(ErrorCode::Value);
    }
    // Symmetric guard for sub-1 fractional col_num.
    if (col_raw == 0.0 && col_orig != 0.0) {
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
      // 1x1 range: row_num must be 1 (or 0 "whole", which collapses to the
      // sole cell).
      if (row_idx == 0U) {
        return cells[0];
      }
      if (row_idx != 1U) {
        return Value::error(ErrorCode::Ref);
      }
      r = 0;
      c = 0;
    } else if (rows == 1U) {
      // Row vector: sole index selects the column. Index 0 spills the
      // whole vector (a 1xN horizontal array).
      if (row_idx == 0U) {
        return index_whole_row(cells, cols, 0U, arena);
      }
      if (row_idx > cols) {
        return Value::error(ErrorCode::Ref);
      }
      r = 0;
      c = row_idx - 1U;
    } else if (cols == 1U) {
      // Column vector: sole index selects the row. Index 0 spills the
      // whole vector (an Nx1 vertical array).
      if (row_idx == 0U) {
        return index_whole_column(cells, rows, cols, 0U, arena);
      }
      if (row_idx > rows) {
        return Value::error(ErrorCode::Ref);
      }
      r = row_idx - 1U;
      c = 0;
    } else {
      // 2-D array with only a row selector: the omitted column argument is
      // read as zero, so the selected row spills whole. A zero row selector
      // then spans both dimensions and spills the entire array, the same
      // result the explicit `INDEX(array, 0, 0)` produces below.
      if (row_idx == 0U) {
        return index_whole_array(cells, rows, cols, arena);
      }
      if (row_idx > rows) {
        return Value::error(ErrorCode::Ref);
      }
      return index_whole_row(cells, cols, row_idx - 1U, arena);
    }
  } else {
    // Three-arg form.
    if (rows == 1U) {
      // Row vector: row_num must be 1 (or 0 "whole row", which spans the
      // single row anyway).
      if (row_idx != 1U && row_idx != 0U) {
        return Value::error(ErrorCode::Ref);
      }
      if (col_idx == 0U) {
        // Whole row of a 1-row source -> spill the entire vector.
        return index_whole_row(cells, cols, 0U, arena);
      }
      if (col_idx > cols) {
        return Value::error(ErrorCode::Ref);
      }
      r = 0;
      c = col_idx - 1U;
    } else if (cols == 1U) {
      // Column vector: col_num must be 1 (or 0 "whole column", which spans
      // the single column anyway).
      if (col_idx != 1U && col_idx != 0U) {
        return Value::error(ErrorCode::Ref);
      }
      if (row_idx == 0U) {
        // Whole column of a 1-column source -> spill the entire vector.
        return index_whole_column(cells, rows, cols, 0U, arena);
      }
      if (row_idx > rows) {
        return Value::error(ErrorCode::Ref);
      }
      r = row_idx - 1U;
      c = 0;
    } else {
      // 2-D array. Zero indices spill the spanned dimension.
      if (row_idx == 0U && col_idx == 0U) {
        return index_whole_array(cells, rows, cols, arena);
      }
      if (row_idx == 0U) {
        // Whole column at col_idx -> spill the column as a vertical array.
        if (col_idx > cols) {
          return Value::error(ErrorCode::Ref);
        }
        return index_whole_column(cells, rows, cols, col_idx - 1U, arena);
      }
      if (col_idx == 0U) {
        // Whole row at row_idx -> spill the row as a horizontal array.
        if (row_idx > rows) {
          return Value::error(ErrorCode::Ref);
        }
        return index_whole_row(cells, cols, row_idx - 1U, arena);
      }
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
  auto cmp_text = [&](std::string_view a, std::string_view b) -> int {
    // Normalise (see classic.cpp::lookup_scan / lookup_text_cmp_key) so kana /
    // half-width voicing variants order together in MATCH approximate mode.
    return strings::case_insensitive_compare(lookup_text_cmp_key(a, profile), lookup_text_cmp_key(b, profile));
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
      cmp = cmp_text(cell.as_text(), lookup.as_text());
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

  // lookup_array: must be a range / Ref with a 1-D shape.
  auto resolved = resolve_range_arg(call.as_call_arg(1), arena, registry, ctx);
  if (!resolved) {
    if (array_lookup) {
      const ErrorCode lookup_array_err = resolved.error();
      return map_lookup_query_array(lookup, arena, [lookup_array_err](const Value& query) {
        return query.is_error() ? query : Value::error(lookup_array_err);
      });
    }
    return Value::error(resolved.error());
  }
  const std::uint32_t rows = resolved.value().rows;
  const std::uint32_t cols = resolved.value().cols;
  std::vector<Value> cells = std::move(resolved.value().cells);
  if (rows != 1U && cols != 1U) {
    // 2-D array to MATCH is not supported and Excel reports #N/A.
    if (array_lookup) {
      return map_lookup_query_array(
          lookup, arena, [](const Value& query) { return query.is_error() ? query : Value::error(ErrorCode::NA); });
    }
    return Value::error(ErrorCode::NA);
  }

  // match_type: default 1. The scalar path retains its existing {-1, 0, 1}
  // validation. Mac Excel's array path instead coerces any finite positive
  // mode to ascending and any finite negative mode to descending; in
  // particular, MATCH(array, range, 2) is not a global invalid result.
  int match_type = 1;
  if (arity == 3) {
    const Value mt_val = eval_node(call.as_call_arg(2), arena, registry, ctx);
    if (mt_val.is_error()) {
      if (array_lookup) {
        return map_lookup_query_array(lookup, arena,
                                      [&mt_val](const Value& query) { return query.is_error() ? query : mt_val; });
      }
      return mt_val;
    }
    auto mt_num = coerce_to_number(mt_val);
    if (!mt_num) {
      if (array_lookup) {
        const ErrorCode match_type_err = mt_num.error();
        return map_lookup_query_array(lookup, arena, [match_type_err](const Value& query) {
          return query.is_error() ? query : Value::error(match_type_err);
        });
      }
      return Value::error(mt_num.error());
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
      if (array_lookup) {
        return map_lookup_query_array(
            lookup, arena, [](const Value& query) { return query.is_error() ? query : Value::error(ErrorCode::NA); });
      }
      return Value::error(ErrorCode::NA);
    }
  }

  const std::size_t n = cells.size();
  if (n == 0) {
    if (array_lookup) {
      return map_lookup_query_array(
          lookup, arena, [](const Value& query) { return query.is_error() ? query : Value::error(ErrorCode::NA); });
    }
    return Value::error(ErrorCode::NA);
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

  const LookupAxis axis = lrows >= lcols ? LookupAxis::Column : LookupAxis::Row;
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
  const LookupAxis raxis = rrows >= rcols ? LookupAxis::Column : LookupAxis::Row;
  const std::size_t flat = raxis == LookupAxis::Column ? (off * static_cast<std::size_t>(rcols)) : off;
  return flat < result_cells.size() ? result_cells[flat] : Value::error(ErrorCode::NA);
}

}  // namespace eval
}  // namespace formulon
