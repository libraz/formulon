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
// Entered text also loses the whitespace its parsed form has no place for
// (`drop_unkept_whitespace`); `spell_entered_formula` applies both steps.
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

/// Returns `formula` (a leading `=` optional) without the whitespace Excel
/// drops when it parses entered text: the trailing run, a run before a `,`,
/// a run on either side of a `:`, every run inside an array constant and,
/// in any other run, the spaces directly before a line feed. Every other
/// run is kept, the intersection operator, a leading run and line feeds
/// among them. Text inside structured-reference brackets and string
/// literals is left alone; nothing is validated.
std::string drop_unkept_whitespace(std::string_view formula);

/// True when entered `formula` (a leading `=` optional) holds whitespace
/// Excel refuses to store: a tab or a carriage return between tokens, or a
/// run between a function name (`is_function_name`) and its `(`. Text inside
/// structured-reference brackets and string literals is not looked at.
bool has_refused_whitespace(std::string_view formula, bool (*is_function_name)(std::string_view));

/// Stored spelling of entered formula text: `drop_unkept_whitespace`, then
/// `spell_number_literals`.
std::string spell_entered_formula(std::string_view formula);

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_LITERAL_SPELLING_H_
