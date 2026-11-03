//
// Primitives shared by the classic lookup TUs: `lookups/classic.cpp`
// (`MATCH`, `VLOOKUP`, `HLOOKUP`, `LOOKUP`) and `lookups/index.cpp`
// (`CHOOSE`, `INDEX`). Both read a table argument spelled as a static
// reference by coordinate rather than materialising it whole, and both
// promote reference blanks when an array compositor materialises a result.

#ifndef FORMULON_EVAL_LOOKUPS_COMMON_H_
#define FORMULON_EVAL_LOOKUPS_COMMON_H_

#include <cstddef>
#include <cstdint>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/declared_rect.h"
#include "parser/reference.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {

struct ExternalBook;

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

// The array compositor is a materialisation boundary. Raw-reference blanks
// copied from a source range must therefore become value-array blanks here so
// that downstream COUNTA sees the occupied cells, while scalar INDEX / CHOOSE
// paths continue to return the original reference blank unchanged.
Value promote_array_result_cell(const Value& value);

// Spills `per_lane(query)` for every cell of an array-valued lookup_value, in
// the query array's shape, lanes in row-major order. The lookup family
// resolves its table / index / mode arguments once, and a query cell's own
// error stays local to its lane, so this stands in for the eager broadcaster.
template <typename PerLane>
Value map_query_lanes(const Value& lookup, Arena& arena, const PerLane& per_lane) {
  const std::uint32_t query_rows = lookup.as_array_rows();
  const std::uint32_t query_cols = lookup.as_array_cols();
  const Value* query_cells = lookup.as_array_cells();
  Value* output_cells = nullptr;
  ArrayValue* output = allocate_array_value(query_rows, query_cols, arena, output_cells, kMaxDerivedArrayCells);
  if (output == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  const std::size_t query_count = static_cast<std::size_t>(query_rows) * query_cols;
  for (std::size_t i = 0; i < query_count; ++i) {
    output_cells[i] = promote_array_result_cell(per_lane(query_cells[i]));
  }
  return Value::array(output);
}

// A table argument spelled as a static reference. The lookup family reads it
// by coordinate instead of materialising it whole: the declared rectangle is
// its shape, and only the lines a lookup reads are expanded, so a wide
// whole-column table costs what the lookup touches.
struct ReferenceTable {
  parser::Reference lhs;
  parser::Reference rhs;
  DeclaredRect declared;
  // What a full expansion would walk. A whole-axis reference with nothing
  // populated walks its first line. Every cell beyond it is blank. Unset
  // while `walk_pending`.
  DeclaredRect walked;
  bool walk_pending = false;
  // The supporting workbook a cross-workbook table reads from, and its sheet
  // there; null for a table on this workbook. Its cells come from the
  // external link's cache, where an uncached address reads 0.
  const ExternalBook* external_book = nullptr;
  std::uint32_t external_sheet = 0;
};

// Whether `resolve_reference_table` computes `walked` up front, or leaves it
// to `walk_reference_table` so a single-cell read never pays for the
// whole-axis extent walk.
enum class TableWalk : std::uint8_t { kNow, kDeferred };

// True with `*out` filled when `arg` (after LET passthrough) is a static
// reference, including a single-sheet cell-form `ExternalRef`; false for
// every other shape. A static reference that does not
// resolve reports the error its expansion would.
Expected<bool, ErrorCode> resolve_reference_table(const parser::AstNode& arg, const EvalContext& ctx,
                                                  ReferenceTable* out, TableWalk walk = TableWalk::kNow);

// Fills `table->walked` for a table resolved with `TableWalk::kDeferred`; a
// no-op once it is filled.
Expected<void, ErrorCode> walk_reference_table(ReferenceTable* table, const EvalContext& ctx);

// Expands the bounded sheet rectangle `[row_first..row_last] x
// [col_first..col_last]` of `table`'s sheet, row-major.
Expected<std::vector<Value>, ErrorCode> read_table_block(const ReferenceTable& table, std::uint32_t row_first,
                                                         std::uint32_t row_last, std::uint32_t col_first,
                                                         std::uint32_t col_last, Arena& arena,
                                                         const FunctionRegistry& registry, const EvalContext& ctx);

// Reads the cell at (`row`, `col`), 0-based within `table`'s declared
// rectangle.
Value read_table_cell(const ReferenceTable& table, std::uint32_t row, std::uint32_t col, Arena& arena,
                      const FunctionRegistry& registry, const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_LOOKUPS_COMMON_H_
