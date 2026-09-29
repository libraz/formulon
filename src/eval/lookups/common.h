//
// Primitives shared by the classic lookup TUs: `lookups/classic.cpp`
// (`MATCH`, `VLOOKUP`, `HLOOKUP`, `LOOKUP`) and `lookups/index.cpp`
// (`CHOOSE`, `INDEX`). Both read a table argument spelled as a static
// reference by coordinate rather than materialising it whole, and both
// promote reference blanks when an array compositor materialises a result.

#ifndef FORMULON_EVAL_LOOKUPS_COMMON_H_
#define FORMULON_EVAL_LOOKUPS_COMMON_H_

#include <cstdint>
#include <vector>

#include "eval/declared_rect.h"
#include "parser/reference.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {

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

// A table argument spelled as a static reference. The lookup family reads it
// by coordinate instead of materialising it whole: the declared rectangle is
// its shape, and only the lines a lookup reads are expanded, so a wide
// whole-column table costs what the lookup touches.
struct ReferenceTable {
  parser::Reference lhs;
  parser::Reference rhs;
  DeclaredRect declared;
  // What a full expansion would walk. A whole-axis reference with nothing
  // populated walks its first line. Every cell beyond it is blank.
  DeclaredRect walked;
};

// True with `*out` filled when `arg` (after LET passthrough) is a static
// reference; false for every other shape. A static reference that does not
// resolve reports the error its expansion would.
Expected<bool, ErrorCode> resolve_reference_table(const parser::AstNode& arg, const EvalContext& ctx,
                                                  ReferenceTable* out);

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
