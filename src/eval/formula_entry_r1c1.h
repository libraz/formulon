//
// R1C1 reading of entered cell formulas. Excel parses entered text as A1
// and, when that fails or the text names an unbound R1C1-shaped reference,
// parses it again as R1C1 relative to the entry cell. In R1C1 an A1 cell
// shape (`B1`) is no reference, so it is a LET / LAMBDA name or an
// undefined one. The fallback rewrites the text lexically into the stored
// A1 form, keeping whitespace and every other byte: R1C1 endpoints become
// A1 references, a cell-shaped name takes `_xlpm.` when a LET / LAMBDA in
// scope binds it and `_xlnm.` otherwise, and a column-shaped name right of
// `name:` is parenthesised (`A:(B)`) so it cannot read back as a column.
// The invariant letters `R` / `C` and `[ ]` are the only ones read.

#ifndef FORMULON_EVAL_FORMULA_ENTRY_R1C1_H_
#define FORMULON_EVAL_FORMULA_ENTRY_R1C1_H_

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>

namespace formulon {

class Workbook;

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

/// True when entered `text` holds `_xlpm.` before an A1 cell-shaped name
/// (`_xlpm.B1`) outside string literals and quoted sheet names, at an
/// identifier boundary. Excel refuses such entry.
bool has_entered_cell_shaped_param(std::string_view text);

/// True when entered `text`, whose A1 parse is `root`, is still to be read
/// as R1C1: `root` holds an unqualified name no enclosing LET / LAMBDA binds,
/// or a sheet-qualified name, that reads as an in-grid R1C1 reference
/// (`R1C1`, `RC`, `Sheet1!R2`); or `text` holds a bracketed R1C1 endpoint
/// (`RC[-1]`, `R[1]`) outside strings, quoted names and brackets, which A1
/// reads as a structured reference, whose leading letters name no table of
/// `book`.
bool needs_r1c1_entry(const parser::AstNode& root, std::string_view text, const Workbook& book);

/// The stored A1 text of R1C1 entry `text` (a leading `=` optional) entered
/// at the 0-based (`row`, `col`), or nullopt when it is no R1C1 entry: a `$`
/// outside a string, a number before `:`, a malformed bracketed endpoint, or
/// a rewrite that does not parse. A relative axis wraps around the grid edge.
std::optional<std::string> translate_r1c1_entry(std::string_view text, std::uint32_t row, std::uint32_t col);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_FORMULA_ENTRY_R1C1_H_
