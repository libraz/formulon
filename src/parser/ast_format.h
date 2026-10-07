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

#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

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

/// Excel's spelling of a binary operator (`+`, `<>`, `&`, ...).
const char* binop_token(BinOp op) noexcept;

/// Appends a sheet name, wrapped in single quotes (embedded quotes doubled)
/// when `force_quote` is set or the name needs quoting. No trailing `!`.
void append_sheet_name(std::string_view sheet, bool force_quote, std::string& out);

/// Appends the qualifier of the cross-workbook reference `node`, through
/// its `!`, as the formula bar spells it: `[Book.xlsx]Sheet!`, or one quoted
/// unit around directory, book and sheets (`'/dir/[Book.xlsx]Sheet'!`) when
/// a directory is written, the book or a sheet needs quoting, or, in A1
/// notation only, the reference spans sheets. A book-scope name prints its
/// file instead (`Book.xlsx!`, `'/dir/Book.xlsx'!`), except a decimal book
/// with no directory, which keeps the bracketed `[1]!` / `[0]!` spelling.
void append_external_qualifier(const AstNode& node, bool r1c1, std::string& out);

/// Maps a cross-workbook reference's directory and book to the 1-based
/// index of its external link, for the `[N]` a stored formula carries.
/// `index` returns 0 when no link matches.
struct ExternalBookIndexer {
  std::uint32_t (*index)(const void* ctx, std::string_view path, std::string_view book);
  const void* ctx;
};

/// Fills `out` with the parenthesis pairs each node of `root` prints with:
/// the pairs written around it (`AstNode::paren_depth`), or one where its
/// slot's precedence demands it. Both formatters print exactly these, and a
/// token encoder emits the same count of XLSB `PtgParen`s after the node.
void collect_parenthesized_nodes(const AstNode& root, std::unordered_map<const AstNode*, std::uint8_t>& out);

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
/// This is the inverse of `parser::strip_storage_prefixes` for the shapes the
/// writer produces, so a save → load cycle round-trips the canonical text.
/// A written `@` is stored as `_xlfn.SINGLE(...)`, or as nothing when it is
/// one of `omitted_at` (a legacy formula's implied ones).
///
/// With an `indexer`, each cross-workbook reference names its book by link
/// index and drops the directory (`[1]Sheet!A1`, `'[1]My Sheet'!A1`,
/// `'[1]S1:S2'!A1`, `[1]!Name`); every reference must bind to a link.
/// Without one, the book is spelled as `format_formula` spells it.
std::string format_formula_storage(const AstNode& node, StorageFunctionNameSpeller spell,
                                   const std::vector<const AstNode*>* omitted_at = nullptr,
                                   const ExternalBookIndexer* indexer = nullptr);

/// True when `root` holds a local sheet qualifier written bare that A1
/// notation requires quoted (`S2!A1`, `2024!A1`, `Data:S2!A1`): stored
/// verbatim, that text would not read back in a file, so the writer emits
/// the storage formatter's spelling instead.
bool formula_needs_storage_requote(const AstNode& root);

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_AST_FORMAT_H_
