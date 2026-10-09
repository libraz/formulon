// Rewrites formula text between the stored (en-invariant) form and a profile's spelling.

#ifndef FORMULON_EVAL_FORMULA_LOCALIZE_H_
#define FORMULON_EVAL_FORMULA_LOCALIZE_H_

#include <string>
#include <string_view>

#include "excel_profile.h"

namespace formulon {
namespace eval {

/// Spelling Excel shows for the canonical function `canonical` (ASCII
/// case-insensitive) in `locale`, or nullptr when it shows the English name.
const char* localized_function_name(ExcelLocale locale, std::string_view canonical) noexcept;

/// Canonical name whose `locale` spelling is `localized` (ASCII
/// case-insensitive), or nullptr when no localized name matches.
const char* canonical_function_name(ExcelLocale locale, std::string_view localized) noexcept;

/// Renders stored (en-invariant, storage-prefix-free) formula text as Excel
/// displays it under `profile`: localized function names, separators,
/// booleans and error names, and number literals in their normalized form
/// (`1.5E-3` reads back as `0.0015`). Everything else, including whitespace,
/// strings and structured-reference brackets, is kept verbatim. A leading `=`
/// is preserved. Pure text rewrite: the grammar is never validated.
std::string localize_formula_text(std::string_view formula, ExcelProfile profile);

/// Inverse of `localize_formula_text`: reads `formula` as `profile`'s
/// spelling and returns the en-invariant text. `canonicalize(localize(f, p),
/// p) == f` holds for any `f` without exponent number literals.
std::string canonicalize_formula_text(std::string_view formula, ExcelProfile profile);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_FORMULA_LOCALIZE_H_
