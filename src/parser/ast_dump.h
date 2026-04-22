// Copyright 2026 libraz. Licensed under the MIT License.
//
// Deterministic S-expression dumper for `AstNode`. The format is the golden
// contract for the parser corpus, so the output must remain stable across
// platforms and compiler versions: it is locale-free, uses neither
// `<iostream>` nor `printf`-style floating-point formatting, and trims
// trailing zeros from non-integer numerics.
//
// Identifier escaping is intentionally minimal: the parser is responsible
// for producing valid Excel identifiers; this dumper just prints them
// verbatim. Sheet names that need quoting are quoted by `format_a1`.

#ifndef FORMULON_PARSER_AST_DUMP_H_
#define FORMULON_PARSER_AST_DUMP_H_

#include <string>

#include "parser/ast.h"

namespace formulon {
namespace parser {

/// Recursively dumps `node` as a single-line S-expression.
///
/// See the table in `backup/plans/02-calc-engine.md` §2.3 for the per-kind
/// format. Output is deterministic and contains no leading or trailing
/// whitespace.
std::string dump_sexpr(const AstNode& node);

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_AST_DUMP_H_
