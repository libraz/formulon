//
// Built-in function name resolution over every source the engine
// recognises: the eager `FunctionRegistry`, the tree walker's lazy-dispatch
// table, and the parser-integrated special forms (LET / LAMBDA).
//
// Also owns the entry rule that follows from that set: Excel refuses a
// sheet- or book-qualified call to a built-in (`=Sheet1!SUM(1)`,
// `=Book!SUM(1)`) at entry, while the same spelling with a defined name
// (`=Sheet1!MyFn(1)`) or an unknown name (`=Sheet1!NOSUCH(1)`, `#NAME?`)
// is accepted. The parser cannot tell the two apart without the name set,
// so the rule runs here over the parsed tree.

#ifndef FORMULON_EVAL_BUILTIN_NAMES_H_
#define FORMULON_EVAL_BUILTIN_NAMES_H_

#include <string_view>

#include "parser/ast.h"

namespace formulon {
class Arena;

namespace eval {

/// Returns the canonical UPPERCASE name of the built-in `name` refers to
/// (ASCII case-insensitive, no storage prefix), or nullptr when `name` is
/// not a built-in. The returned pointer has program lifetime.
const char* resolve_builtin_function_name(std::string_view name);

/// Returns the first callee in `root` that qualifies a built-in function
/// name with a sheet (`Sheet1!SUM`) or with the formula's own workbook
/// (`[0]!SUM`), or nullptr when there is none. A non-null result means
/// Excel would reject the formula at entry; its `range()` locates the
/// offending callee.
const parser::AstNode* find_qualified_builtin_call(const parser::AstNode& root);

/// Parses a formula body at a public formula-entry boundary. The parser
/// accepts qualified names syntactically because they may denote defined
/// names; a qualified built-in is nevertheless invalid formula entry and is
/// treated like any other parse failure by evaluation and indexing callers.
/// The returned AST is owned by `arena`, or nullptr when parsing fails or the
/// AST contains a qualified built-in call.
parser::AstNode* parse_formula_entry(std::string_view src, Arena& arena);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTIN_NAMES_H_
