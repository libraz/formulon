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
// Entry reads the invariant letters `R` / `C` and `[ ]` only. The same
// rewrite, with a locale's letters and without the entry rules, reads the
// localized R1C1 text that `canonicalize_formula_text_r1c1` takes.

#ifndef FORMULON_EVAL_FORMULA_ENTRY_R1C1_H_
#define FORMULON_EVAL_FORMULA_ENTRY_R1C1_H_

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>

#include "eval/a1_parse.h"
#include "eval/formula_localize.h"
#include "parser/parse_status.h"
#include "utils/expected.h"

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

/// How `rewrite_r1c1_text` reads its text.
struct R1C1RewriteOptions {
  /// Row / column letters and relative brackets.
  refs_internal::R1C1Letters letters;
  /// Inside `{}` this byte separates array rows and is never a range colon.
  char array_row_separator = ';';
  /// Cell-entry rules: a `$` or a number before `:` refuses the text. Without
  /// them a trim operator (`.:`, `:.`, `.:.`) joins two endpoints and a local
  /// sheet qualifier is requoted by the stored rule.
  bool entry = false;
  bool has_base = false;
  std::uint32_t base_row = 0;  ///< 0-based.
  std::uint32_t base_col = 0;  ///< 0-based.
};

/// Rewrites R1C1 `text` into A1 lexically, keeping every other byte: each
/// endpoint becomes an A1 reference with `$` on its absolute axes, a relative
/// axis measured from the base and wrapped around the grid edge. A
/// cell-shaped name takes `_xlpm.` and a column-shaped name right of `name:`
/// is parenthesised (`A:(B)`). Strings, quoted names and `[...]` groups are
/// skipped. Fails with `InvalidReference`
/// over the endpoint for a relative axis without a base or a malformed
/// endpoint, and with `UnexpectedToken` for text the entry rules refuse.
/// The result's spans map it back to `text`.
Expected<MappedFormulaText, parser::ParseFailure> rewrite_r1c1_text(std::string_view text,
                                                                    const R1C1RewriteOptions& options);

/// Respells, in place, the `_xlpm.` of every cell-shaped name in A1 `text`
/// (a leading `=` optional) that no enclosing LET / LAMBDA binds as `_xlnm.`.
/// `root` is the parse of `text` after its `=`.
void respell_unbound_cell_names(const parser::AstNode& root, std::string& text);

/// Calls `visit(ctx, node)` for every unqualified name reference of `root`
/// that no enclosing LET / LAMBDA binds.
void for_each_unbound_name(const parser::AstNode& root, void (*visit)(void* ctx, const parser::AstNode& node),
                           void* ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_FORMULA_ENTRY_R1C1_H_
