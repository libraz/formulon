//
// Implementation of the classic lookup-family lazy impls that select by
// position (`CHOOSE`, `INDEX`). See `lookups/classic.h` for the
// dispatch-table contract and `eval/lazy_impls.h` for the shared
// vocabulary.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <optional>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/coerce.h"
#include "eval/declared_rect.h"
#include "eval/dynamic_array/common.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/lazy_impls.h"
#include "eval/lookups/classic.h"
#include "eval/lookups/common.h"
#include "eval/name_env_resolve.h"
#include "eval/omitted_arg.h"
#include "eval/range_args.h"
#include "eval/range_resolvers.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

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

// `rows` x `cols` is the source's shape, which decides the selection; a tile
// spans `span_rows` x `span_cols`, the part of the source that holds values.
IndexTile index_tile_for(std::uint32_t rows, std::uint32_t cols, std::uint32_t span_rows, std::uint32_t span_cols,
                         std::uint32_t row_idx, std::uint32_t col_idx, bool col_explicit) {
  IndexTile tile;
  if (!col_explicit) {
    if (rows == 1U && cols == 1U) {
      return tile;
    }
    if (rows == 1U) {
      if (row_idx == 0U) {
        tile.kind = IndexTileKind::kRow;
        tile.cols = span_cols;
      }
      return tile;
    }
    if (cols == 1U) {
      if (row_idx == 0U) {
        tile.kind = IndexTileKind::kColumn;
        tile.rows = span_rows;
      }
      return tile;
    }
    // The two-argument form on a 2-D source reads the omitted column
    // argument as zero, so a row selector returns that complete row and a
    // zero selector spans every row as well -- the whole array, exactly as
    // the explicit `INDEX(src, 0, 0)`. Matches the scalar path.
    if (row_idx == 0U) {
      tile.kind = IndexTileKind::kWhole;
      tile.rows = span_rows;
      tile.cols = span_cols;
      return tile;
    }
    tile.kind = IndexTileKind::kRow;
    tile.cols = span_cols;
    tile.source_row = row_idx - 1U;
    return tile;
  }

  if (rows == 1U && cols == 1U) {
    return tile;
  }
  if (rows == 1U) {
    if (col_idx == 0U) {
      tile.kind = IndexTileKind::kRow;
      tile.cols = span_cols;
    }
    return tile;
  }
  if (cols == 1U) {
    if (row_idx == 0U) {
      tile.kind = IndexTileKind::kColumn;
      tile.rows = span_rows;
    }
    return tile;
  }
  if (row_idx == 0U && col_idx == 0U) {
    tile.kind = IndexTileKind::kWhole;
    tile.rows = span_rows;
    tile.cols = span_cols;
  } else if (row_idx == 0U) {
    tile.kind = IndexTileKind::kColumn;
    tile.rows = span_rows;
  } else if (col_idx == 0U) {
    tile.kind = IndexTileKind::kRow;
    tile.cols = span_cols;
  }
  return tile;
}

template <typename CellAt>
Value index_tile_cell(const CellAt& cell_at, std::uint32_t source_rows, std::uint32_t span_rows,
                      std::uint32_t span_cols, std::uint32_t row_idx, std::uint32_t col_idx, bool col_explicit,
                      const IndexTile& tile, std::uint32_t output_row, std::uint32_t output_col) {
  if (tile.kind == IndexTileKind::kScalar) {
    const std::uint32_t source_row = col_explicit ? row_idx - 1U : (source_rows == 1U ? 0U : row_idx - 1U);
    const std::uint32_t source_col = col_explicit ? col_idx - 1U : (source_rows == 1U ? row_idx - 1U : 0U);
    return cell_at(source_row, source_col);
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
  if (source_row >= span_rows && tile.kind != IndexTileKind::kRow) {
    return Value::error(ErrorCode::NA);
  }
  if (source_col >= span_cols && tile.kind != IndexTileKind::kColumn) {
    return Value::error(ErrorCode::NA);
  }
  return cell_at(source_row, source_col);
}

// The rectangle INDEX selects from, in its declared shape. A static
// reference is read on demand, so a single-cell selection reads one cell; a
// whole row, column or array spans the walked extent, as a full expansion
// would. Every other argument arrives materialised.
class IndexSource {
 public:
  IndexSource(std::vector<Value> cells, std::uint32_t rows, std::uint32_t cols)
      : cells_(std::move(cells)), rows_(rows), cols_(cols), stored_rows_(rows), stored_cols_(cols) {}
  explicit IndexSource(const ReferenceTable& table)
      : table_(&table), rows_(table.declared.rows()), cols_(table.declared.cols()) {}

  std::uint32_t rows() const noexcept { return rows_; }
  std::uint32_t cols() const noexcept { return cols_; }
  // The part of the rectangle that can hold values.
  std::uint32_t span_rows() const noexcept { return table_ != nullptr ? table_->walked.rows() : rows_; }
  std::uint32_t span_cols() const noexcept { return table_ != nullptr ? table_->walked.cols() : cols_; }

  // Reads the whole walked rectangle, for selections that may touch any of it.
  bool load(Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx, ErrorCode* out_err) {
    if (table_ == nullptr || loaded_) {
      return true;
    }
    const DeclaredRect& walked = table_->walked;
    auto block = read_table_block(*table_, walked.row_first, walked.row_last, walked.col_first, walked.col_last, arena,
                                  registry, ctx);
    if (!block) {
      *out_err = block.error();
      return false;
    }
    cells_ = std::move(block.value());
    stored_rows_ = walked.rows();
    stored_cols_ = walked.cols();
    loaded_ = true;
    return true;
  }

  // The cell at (`row`, `col`) of a materialised or loaded source.
  Value at(std::uint32_t row, std::uint32_t col) const {
    if (table_ != nullptr && (row >= stored_rows_ || col >= stored_cols_)) {
      return Value::blank(BlankGridProjection::kReferenceGridZero);
    }
    const std::size_t flat = (static_cast<std::size_t>(row) * stored_cols_) + col;
    return flat < cells_.size() ? cells_[flat] : Value::error(ErrorCode::Ref);
  }

  Value cell(std::uint32_t row, std::uint32_t col, Arena& arena, const FunctionRegistry& registry,
             const EvalContext& ctx) const {
    if (table_ != nullptr && !loaded_) {
      return read_table_cell(*table_, row, col, arena, registry, ctx);
    }
    return at(row, col);
  }

  Value row(std::uint32_t row, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx) const {
    if (table_ == nullptr) {
      return index_whole_row(cells_, cols_, row, arena);
    }
    const std::uint32_t sheet_row = table_->declared.row_first + row;
    return block_array(sheet_row, sheet_row, table_->walked.col_first, table_->walked.col_last, arena, registry, ctx);
  }

  Value column(std::uint32_t col, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx) const {
    if (table_ == nullptr) {
      return index_whole_column(cells_, rows_, cols_, col, arena);
    }
    const std::uint32_t sheet_col = table_->declared.col_first + col;
    return block_array(table_->walked.row_first, table_->walked.row_last, sheet_col, sheet_col, arena, registry, ctx);
  }

  Value whole(Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx) const {
    if (table_ == nullptr) {
      return index_whole_array(cells_, rows_, cols_, arena);
    }
    const DeclaredRect& walked = table_->walked;
    return block_array(walked.row_first, walked.row_last, walked.col_first, walked.col_last, arena, registry, ctx);
  }

 private:
  Value block_array(std::uint32_t row_first, std::uint32_t row_last, std::uint32_t col_first, std::uint32_t col_last,
                    Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx) const {
    auto block = read_table_block(*table_, row_first, row_last, col_first, col_last, arena, registry, ctx);
    if (!block) {
      return Value::error(block.error());
    }
    return index_whole_array(block.value(), row_last - row_first + 1U, col_last - col_first + 1U, arena);
  }

  const ReferenceTable* table_ = nullptr;
  std::vector<Value> cells_;
  std::uint32_t rows_ = 0U;
  std::uint32_t cols_ = 0U;
  std::uint32_t stored_rows_ = 0U;
  std::uint32_t stored_cols_ = 0U;
  bool loaded_ = false;
};

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

Value eval_index_array_selector(const IndexSource& source, bool source_ok, ErrorCode source_error,
                                const Value& row_value, const Value* col_value, bool col_explicit, Arena& arena) {
  const std::uint32_t source_rows = source.rows();
  const std::uint32_t source_cols = source.cols();
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
        const IndexTile tile = index_tile_for(source_rows, source_cols, source.span_rows(), source.span_cols(),
                                              row.index, col.index, col_explicit);
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
      const IndexTile tile = index_tile_for(source_rows, source_cols, source.span_rows(), source.span_cols(), row.index,
                                            col.index, col_explicit);
      result = index_tile_cell(
          [&source](std::uint32_t row_at, std::uint32_t col_at) { return source.at(row_at, col_at); }, source_rows,
          source.span_rows(), source.span_cols(), row.index, col.index, col_explicit, tile, r, c);
      output_cells[static_cast<std::size_t>(r) * out_cols + c] = promote_array_result_cell(result);
    }
  }
  return Value::array(output);
}

// What a scalar (row, col) selection names inside an INDEX source: one cell,
// one whole row, one whole column, or the whole source.
enum class IndexPickKind : std::uint8_t { kCell, kRow, kColumn, kWhole };

struct IndexPick {
  IndexPickKind kind = IndexPickKind::kCell;
  std::uint32_t row = 0U;  // 0-based; meaningful for kCell / kRow.
  std::uint32_t col = 0U;  // 0-based; meaningful for kCell / kColumn.
};

// Resolves a scalar (row_idx, col_idx) pair against a `rows` x `cols` source.
// Zero indices are "whole dimension" in Excel's spill model. The value path
// reads the pick out of the source and the reference path maps it onto the
// source rectangle, so the two cannot disagree on what INDEX selects.
// `reference` marks the reference form, whose 2-D area needs both indices.
Expected<IndexPick, ErrorCode> index_pick(std::uint32_t rows, std::uint32_t cols, std::uint32_t row_idx,
                                          std::uint32_t col_idx, bool col_explicit, bool reference) {
  if (!col_explicit) {
    // Two-arg form.
    if (rows == 1U && cols == 1U) {
      // 1x1 range: row_num must be 1 (or 0 "whole", which collapses to the
      // sole cell).
      if (row_idx > 1U) {
        return ErrorCode::Ref;
      }
      return IndexPick{IndexPickKind::kCell, 0U, 0U};
    }
    if (rows == 1U) {
      // Row vector: sole index selects the column. Index 0 spills the
      // whole vector (a 1xN horizontal array).
      if (row_idx == 0U) {
        return IndexPick{IndexPickKind::kRow, 0U, 0U};
      }
      if (row_idx > cols) {
        return ErrorCode::Ref;
      }
      return IndexPick{IndexPickKind::kCell, 0U, row_idx - 1U};
    }
    if (cols == 1U) {
      // Column vector: sole index selects the row. Index 0 spills the
      // whole vector (an Nx1 vertical array).
      if (row_idx == 0U) {
        return IndexPick{IndexPickKind::kColumn, 0U, 0U};
      }
      if (row_idx > rows) {
        return ErrorCode::Ref;
      }
      return IndexPick{IndexPickKind::kCell, row_idx - 1U, 0U};
    }
    // 2-D array with only a row selector: the omitted column argument is
    // read as zero, so the selected row spills whole. A zero row selector
    // then spans both dimensions and spills the entire array, the same
    // result the explicit `INDEX(array, 0, 0)` produces below.
    if (row_idx == 0U) {
      return IndexPick{IndexPickKind::kWhole, 0U, 0U};
    }
    // A 2-D reference with a row number alone is #REF! (measured on Excel
    // 365); an array spills the row.
    if (row_idx > rows || reference) {
      return ErrorCode::Ref;
    }
    return IndexPick{IndexPickKind::kRow, row_idx - 1U, 0U};
  }
  // Three-arg form.
  if (rows == 1U) {
    // Row vector: row_num must be 1 (or 0 "whole row", which spans the
    // single row anyway).
    if (row_idx != 1U && row_idx != 0U) {
      return ErrorCode::Ref;
    }
    if (col_idx == 0U) {
      // Whole row of a 1-row source -> spill the entire vector.
      return IndexPick{IndexPickKind::kRow, 0U, 0U};
    }
    if (col_idx > cols) {
      return ErrorCode::Ref;
    }
    return IndexPick{IndexPickKind::kCell, 0U, col_idx - 1U};
  }
  if (cols == 1U) {
    // Column vector: col_num must be 1 (or 0 "whole column", which spans
    // the single column anyway).
    if (col_idx != 1U && col_idx != 0U) {
      return ErrorCode::Ref;
    }
    if (row_idx == 0U) {
      // Whole column of a 1-column source -> spill the entire vector.
      return IndexPick{IndexPickKind::kColumn, 0U, 0U};
    }
    if (row_idx > rows) {
      return ErrorCode::Ref;
    }
    return IndexPick{IndexPickKind::kCell, row_idx - 1U, 0U};
  }
  // 2-D array. Zero indices spill the spanned dimension.
  if (row_idx == 0U && col_idx == 0U) {
    return IndexPick{IndexPickKind::kWhole, 0U, 0U};
  }
  if (row_idx == 0U) {
    if (col_idx > cols) {
      return ErrorCode::Ref;
    }
    return IndexPick{IndexPickKind::kColumn, 0U, col_idx - 1U};
  }
  if (col_idx == 0U) {
    if (row_idx > rows) {
      return ErrorCode::Ref;
    }
    return IndexPick{IndexPickKind::kRow, row_idx - 1U, 0U};
  }
  if (row_idx > rows || col_idx > cols) {
    return ErrorCode::Ref;
  }
  return IndexPick{IndexPickKind::kCell, row_idx - 1U, col_idx - 1U};
}

// Appends the areas of a (possibly nested) parenthesised union to `out`, in
// source order.
void collect_union_areas(const parser::AstNode& node, const EvalContext& ctx,
                         std::vector<const parser::AstNode*>* out) {
  const parser::AstNode& resolved = resolve_name_ast(node, ctx.name_env());
  if (resolved.kind() != parser::NodeKind::UnionOp) {
    out->push_back(&resolved);
    return;
  }
  const std::uint32_t arity = resolved.as_union_arity();
  for (std::uint32_t i = 0; i < arity; ++i) {
    collect_union_areas(resolved.as_union_child(i), ctx, out);
  }
}

// The area INDEX reads from: the `area_num`-th area of its first argument
// (default 1), where a parenthesised union supplies several areas and any
// other argument is its only area. An area past the end is `#REF!`.
Expected<const parser::AstNode*, ErrorCode> select_index_area(const parser::AstNode& call, Arena& arena,
                                                              const FunctionRegistry& registry,
                                                              const EvalContext& ctx) {
  std::vector<const parser::AstNode*> areas;
  collect_union_areas(call.as_call_arg(0), ctx, &areas);
  std::uint32_t area_idx = 1U;
  if (call.as_call_arity() == 4U && !is_omitted_arg(call.as_call_arg(3))) {
    const DecodedIndex area = decode_index_cell(eval_node(call.as_call_arg(3), arena, registry, ctx));
    if (area.state == IndexAxisState::kError) {
      return area.error;
    }
    if (area.index == 0U) {
      return ErrorCode::Value;
    }
    area_idx = area.index;
  }
  if (area_idx > areas.size()) {
    return ErrorCode::Ref;
  }
  // A lone first argument keeps its own node so a LET binding reaches the
  // lookup seams through their usual name look-through.
  return areas.size() == 1U ? &call.as_call_arg(0) : areas[area_idx - 1U];
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
  if (arity < 2U || arity > 4U) {
    return Value::error(ErrorCode::Value);
  }
  const auto area = select_index_area(call, arena, registry, ctx);
  if (!area) {
    return Value::error(area.error());
  }
  const parser::AstNode& source_node = *area.value();
  ReferenceTable table;
  const auto by_reference = resolve_reference_table(source_node, ctx, &table);
  ErrorCode source_error = ErrorCode::Value;
  std::optional<IndexSource> source;
  if (!by_reference) {
    source_error = by_reference.error();
  } else if (by_reference.value()) {
    source.emplace(table);
  } else {
    auto resolved = resolve_range_arg(source_node, arena, registry, ctx);
    if (resolved) {
      source.emplace(std::move(resolved.value().cells), resolved.value().rows, resolved.value().cols);
    } else {
      source_error = resolved.error();
    }
  }
  const std::uint32_t rows = source.has_value() ? source->rows() : 0U;
  const std::uint32_t cols = source.has_value() ? source->cols() : 0U;
  bool source_ok = source.has_value();

  // row_num is required, col_num is optional.
  const bool col_explicit = arity >= 3U;
  const Value row_val = eval_node(call.as_call_arg(1), arena, registry, ctx);
  Value col_val = Value::number(0.0);
  if (col_explicit) {
    col_val = eval_node(call.as_call_arg(2), arena, registry, ctx);
  }
  if (row_val.is_array() || col_val.is_array()) {
    if (source_ok && !source->load(arena, registry, ctx, &source_error)) {
      source_ok = false;
    }
    const IndexSource no_source(std::vector<Value>{}, 0U, 0U);
    return eval_index_array_selector(source_ok ? *source : no_source, source_ok && rows != 0U && cols != 0U,
                                     source_error, row_val, col_explicit ? &col_val : nullptr, col_explicit, arena);
  }
  if (!source_ok || rows == 0U || cols == 0U) {
    return Value::error(source_error == ErrorCode::Value ? ErrorCode::Ref : source_error);
  }
  const DecodedIndex row = decode_index_cell(row_val);
  if (row.state == IndexAxisState::kError) {
    return Value::error(row.error);
  }
  DecodedIndex col{IndexAxisState::kValid, 0U, ErrorCode::Value};
  if (col_explicit) {
    col = decode_index_cell(col_val);
    if (col.state == IndexAxisState::kError) {
      return Value::error(col.error);
    }
  }
  const auto pick = index_pick(rows, cols, row.index, col.index, col_explicit, by_reference && by_reference.value());
  if (!pick) {
    return Value::error(pick.error());
  }
  switch (pick.value().kind) {
    case IndexPickKind::kRow:
      return source->row(pick.value().row, arena, registry, ctx);
    case IndexPickKind::kColumn:
      return source->column(pick.value().col, arena, registry, ctx);
    case IndexPickKind::kWhole:
      return source->whole(arena, registry, ctx);
    case IndexPickKind::kCell:
      break;
  }
  return source->cell(pick.value().row, pick.value().col, arena, registry, ctx);
}

bool resolve_index_reference(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx, std::string_view* out_sheet, std::uint32_t* out_top_row,
                             std::uint32_t* out_left_col, std::uint32_t* out_bottom_row, std::uint32_t* out_right_col,
                             ErrorCode* out_err) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 4U) {
    *out_err = ErrorCode::Value;
    return false;
  }
  const auto area = select_index_area(call, arena, registry, ctx);
  if (!area) {
    *out_err = area.error();
    return false;
  }
  std::string_view sheet;
  std::uint32_t top = 0U;
  std::uint32_t left = 0U;
  std::uint32_t bottom = 0U;
  std::uint32_t right = 0U;
  if (!resolve_reference_rect(*area.value(), arena, registry, ctx, &sheet, &top, &left, &bottom, &right, out_err)) {
    return false;
  }
  const bool col_explicit = arity >= 3U;
  const Value row_val = eval_node(call.as_call_arg(1), arena, registry, ctx);
  const Value col_val = col_explicit ? eval_node(call.as_call_arg(2), arena, registry, ctx) : Value::number(0.0);
  // An array selector yields an array of values, not one reference.
  if (row_val.is_array() || col_val.is_array()) {
    *out_err = ErrorCode::Value;
    return false;
  }
  const DecodedIndex row = decode_index_cell(row_val);
  if (row.state == IndexAxisState::kError) {
    *out_err = row.error;
    return false;
  }
  DecodedIndex col{IndexAxisState::kValid, 0U, ErrorCode::Value};
  if (col_explicit) {
    col = decode_index_cell(col_val);
    if (col.state == IndexAxisState::kError) {
      *out_err = col.error;
      return false;
    }
  }
  const auto pick = index_pick(bottom - top + 1U, right - left + 1U, row.index, col.index, col_explicit, true);
  if (!pick) {
    *out_err = pick.error();
    return false;
  }
  const IndexPick& p = pick.value();
  *out_sheet = sheet;
  *out_top_row = top;
  *out_left_col = left;
  *out_bottom_row = bottom;
  *out_right_col = right;
  if (p.kind == IndexPickKind::kCell || p.kind == IndexPickKind::kRow) {
    *out_top_row = *out_bottom_row = top + p.row;
  }
  if (p.kind == IndexPickKind::kCell || p.kind == IndexPickKind::kColumn) {
    *out_left_col = *out_right_col = left + p.col;
  }
  return true;
}

}  // namespace eval
}  // namespace formulon
