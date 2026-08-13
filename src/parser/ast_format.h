//
// Excel-compatible formula-text formatter for `AstNode`.
//
// `format_formula` is the inverse of `parser::Parser::parse()` for the
// purposes of round-trip transforms: every input AST that originated from a
// well-formed formula must format to text that re-parses to a structurally
// equivalent AST (tested via `dump_sexpr` golden equivalence). Callers use
// this to write formulas back out after applying a `RefTransform`
// (sheet rename, relative shift, future row/column insert/delete).
//
// The output is *not* required to be byte-stable across reorderings of the
// arena or input variants — the only contract is round-trip equivalence,
// which means the formatter MAY emit redundant parentheses around any
// subexpression where the parent context's precedence would otherwise pull
// children apart. Adding parens never changes Excel semantics; missing
// parens for a `<` placed inside `^` would.
//
// Operator precedence (high → low) follows Excel:
//   `:`, intersect (space), union (`,`), unary `-`/`+`, `%`, `^`,
//   `*`/`/`, `+`/`-`, `&`, comparisons.
//
// The formatter is stateless and dependency-free: it does not touch the
// arena, allocate any AST, or read tokens.

#ifndef FORMULON_PARSER_AST_FORMAT_H_
#define FORMULON_PARSER_AST_FORMAT_H_

#include <string>
#include <string_view>

#include "parser/ast.h"

namespace formulon {
namespace parser {

/// Returns an Excel-compatible textual rendering of `node`.
///
/// The returned string never carries a leading `=` — callers that need one
/// for cell-formula storage prepend it themselves. Output may contain
/// redundant parentheses to preserve operator precedence; round-trip
/// equivalence (parse → format → parse → equal AST) is the contract, not
/// byte-exact stability.
std::string format_formula(const AstNode& node);

/// Storage prefix a function name carries in the OOXML `<f>` element.
enum class StoragePrefixKind {
  None,      ///< Classic (pre-2007) function: no prefix.
  Xlfn,      ///< Post-2007 function: `_xlfn.` prefix.
  XlfnXlws,  ///< Worksheet-only dynamic-array function: `_xlfn._xlws.` prefix.
};

/// Classifies a canonical (unprefixed, upper-case) function name into the
/// storage prefix Excel writes for it. Supplied by the writer so this
/// header stays free of the function catalog.
using StoragePrefixClassifier = StoragePrefixKind (*)(std::string_view canonical_name);

/// Spells the exact function name a stored formula carries for a call
/// written as `name`: the storage prefix plus whatever spelling Excel
/// itself stores, which is not always the one the user typed (a
/// localised formula-bar alias resolves to the invariant name).
/// Supplied by the writer for the same reason as the classifier — the
/// parser holds the name as typed and does not own the catalog.
using StorageFunctionNameSpeller = std::string (*)(std::string_view name);

/// Like `format_formula`, but emits what a real Excel worksheet stores in
/// `<f>`:
///   * each function call's name is spelled by `spell`, which applies the
///     `_xlfn.` / `_xlfn._xlws.` prefix and resolves any spelling Excel
///     accepts but does not store;
///   * LET / LAMBDA are themselves future functions (spelled the same
///     way), and every LET binding name / LAMBDA parameter name — plus each
///     in-scope reference to one — is emitted with the `_xlpm.` prefix.
///
/// This is the inverse of `io::strip_storage_prefixes` for the shapes the
/// writer produces, so a save → load cycle round-trips the canonical text.
std::string format_formula_storage(const AstNode& node, StorageFunctionNameSpeller spell);

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_AST_FORMAT_H_
