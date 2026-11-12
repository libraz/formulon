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

#include <cstdint>
#include <string_view>

#include "parser/ast.h"
#include "parser/parse_status.h"

namespace formulon {
class Arena;

namespace eval {

/// Most arguments a call may be written with, whatever the callee.
// @limit: excel
inline constexpr std::uint32_t kMaxCallArguments = 255U;

/// Returns the canonical UPPERCASE name of the built-in `name` refers to
/// (ASCII case-insensitive, no storage prefix), or nullptr when `name` is
/// not a built-in. The returned pointer has program lifetime.
const char* resolve_builtin_function_name(std::string_view name);

/// Stores the inclusive bounds of the written argument count Excel accepts
/// for the built-in `name` (ASCII case-insensitive, no storage prefix;
/// `kVariadic` for an open list) and returns true, or returns false when
/// `name` is not a built-in. Lazy forms are consulted first, as the
/// evaluator dispatches them first.
bool builtin_arity(std::string_view name, std::uint32_t* min_arity, std::uint32_t* max_arity);

/// Returns the first callee in `root` that qualifies a built-in function
/// name with a sheet (`Sheet1!SUM`) or with the formula's own workbook
/// (`[0]!SUM`), or nullptr when there is none. A non-null result means
/// Excel would reject the formula at entry; its `range()` locates the
/// offending callee.
const parser::AstNode* find_qualified_builtin_call(const parser::AstNode& root);

/// Returns the first argument in `root` that Excel refuses at entry because
/// the built-in's parameter takes only a reference (`SUMIF`'s range,
/// `SUBTOTAL`'s refs, `OFFSET`'s base, ...) and the argument is not a
/// reference expression: it is empty, a constant, an operator, or a call of a
/// built-in that never returns a reference. Reference expressions are cell,
/// range, 3-D, external and structured references, names, `#REF!`, spill
/// references, the range operators over references, `@` over one, lambda
/// calls, unknown functions, and calls of the built-ins that may return a
/// reference (`INDEX`, `INDIRECT`, `OFFSET`, `IF`, `CHOOSE`, `XLOOKUP`, `IFS`,
/// `SWITCH`, `LET`, `TRIMRANGE`, `SINGLE`, `TAKE`, `DROP`) whatever their
/// arguments. nullptr when there is none.
const parser::AstNode* find_non_reference_argument(const parser::AstNode& root);

/// Returns the first call in `root` that Excel refuses at entry for its
/// argument count: more than 255 arguments for any callee, or, for a
/// built-in, a count outside its arity or one that breaks its argument
/// pairs (`COUNTIFS` takes range/criteria pairs, `SUMIFS` a range then
/// pairs). An omitted argument counts as written. nullptr when there is
/// none.
const parser::AstNode* find_wrong_arity_call(const parser::AstNode& root);

/// Returns the first error literal in `root` that Excel refuses when typed
/// (`#CALC!`, `#SPILL!`, `#FIELD!`, `#BLOCKED!`, `#CONNECT!`, `#BUSY!`,
/// `#UNKNOWN!`), wherever it sits in the formula. nullptr when there is none.
/// Meaningful for entered text only: files and cached values carry these
/// codes legitimately.
const parser::AstNode* find_refused_error_literal(const parser::AstNode& root);

/// Returns the first structured reference in `root` whose table name opens
/// with an R1C1 reference spelled with `row_letter` / `col_letter` (a
/// locale's R1C1 tokens) and for which `is_table(ctx, name)` holds, which
/// Excel refuses when typed under that locale. nullptr when there is none.
/// Meaningful for entered text only.
const parser::AstNode* find_refused_table_reference(const parser::AstNode& root, char row_letter, char col_letter,
                                                    bool (*is_table)(const void* ctx, std::string_view name),
                                                    const void* ctx);

/// Parses a formula body at a public formula-entry boundary. The parser
/// accepts qualified names syntactically because they may denote defined
/// names; a qualified built-in is nevertheless invalid formula entry and is
/// treated like any other parse failure by evaluation and indexing callers.
/// The returned AST is owned by `arena`, or nullptr when parsing fails or the
/// AST contains a qualified built-in call.
parser::AstNode* parse_formula_entry(std::string_view src, Arena& arena);

/// `parse_formula_entry` that reports why it failed: the parse diagnostic, or
/// `EntryQualifiedBuiltin` over the offending callee. Returns the root
/// exactly when `parse_formula_entry` would. `out` must not be null and is
/// left untouched on success.
parser::AstNode* parse_formula_entry_with_diagnostics(std::string_view src, Arena& arena, parser::ParseFailure* out);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTIN_NAMES_H_
