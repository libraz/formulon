//
// R1C1-notation formula formatter for `AstNode`.
//
// Mirrors `format_formula` (ast_format.h) but prints every cell reference
// relative to a host cell, the way Excel's `Range.FormulaR1C1` does. Only the
// reference spelling differs; everything notation-independent is delegated to
// `format_formula`.

#ifndef FORMULON_PARSER_AST_FORMAT_R1C1_H_
#define FORMULON_PARSER_AST_FORMAT_R1C1_H_

#include <cstdint>
#include <string>

#include "parser/ast.h"
#include "parser/structured_ref_payload.h"

namespace formulon {
namespace parser {

/// Returns the R1C1 rendering of `node` as written in the cell at the 0-based
/// (`anchor_row`, `anchor_col`).
///
/// Relative parts print as `R[-1]C[2]` (`R` / `C` for offset zero, so the host
/// cell itself is `RC`), absolute parts as `R3C4`. A whole-column or whole-row
/// reference prints as a single axis (`C[-2]`, `R1`) and a range as
/// `<start>:<end>`. A `@` in front of a reference or range is dropped. The
/// result carries no leading `=`.
///
/// With `display`, structured references print in the formula-bar form for
/// a cell in `display->own_table`; without it, as stored.
std::string format_formula_r1c1(const AstNode& node, std::uint32_t anchor_row, std::uint32_t anchor_col,
                                const StructuredRefDisplay* display = nullptr);

/// Appends one axis of an R1C1 reference spelled with `letter` and the
/// `open` / `close` brackets: `letter` + the 1-based `index` when absolute,
/// else `letter` + the bracketed plain difference `index - host`, or the bare
/// letter at offset 0.
void append_r1c1_axis(char letter, char open, char close, std::int64_t index, std::int64_t host, bool absolute,
                      std::string& out);

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_AST_FORMAT_R1C1_H_
