//
// Respells the number literals of formula text the way Excel stores a formula
// it has parsed from text.
//
// Each literal is printed from its value (`format_formula_number`), so
// `1.50`, `1E3` and `0.30000000000000004` read back as `1.5`, `1000` and
// `0.3`. A unary `+` or `-` written directly against the literal, with no
// whitespace in between, belongs to the literal: `+` disappears and `-`
// makes it negative, so `-0` prints as `0`. Only that one sign is taken;
// any sign before it, a binary operator, and a sign over a reference, a
// parenthesis, a call, an array or any non-number literal stay as written.
//
// The rewrite is not idempotent where Excel's is not either: a surviving
// sign touching a literal (`=+5` stored for `=++5`) is absorbed when the
// stored text is parsed again.

#ifndef FORMULON_PARSER_LITERAL_SPELLING_H_
#define FORMULON_PARSER_LITERAL_SPELLING_H_

#include <string>
#include <string_view>

namespace formulon {
namespace parser {

/// Returns invariant-spelled `formula` (a leading `=` optional) with every
/// number literal respelled and its touching unary sign absorbed as the
/// header describes. Every other byte, whitespace and name case included,
/// is copied verbatim. Literals inside structured-reference brackets, row
/// numbers of a `1:3` row range and runs the tokenizer rejects (`1E-310`)
/// are left alone; nothing is validated.
std::string spell_number_literals(std::string_view formula);

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_LITERAL_SPELLING_H_
