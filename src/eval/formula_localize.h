// Rewrites formula text between the stored (en-invariant) form and a profile's spelling.

#ifndef FORMULON_EVAL_FORMULA_LOCALIZE_H_
#define FORMULON_EVAL_FORMULA_LOCALIZE_H_

#include <cstddef>
#include <string>
#include <string_view>
#include <vector>

#include "excel_profile.h"
#include "parser/token.h"

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
/// strings and column names, is kept verbatim. Inside structured-reference
/// brackets the `#All`/`#Data`/`#Headers`/`#Totals` words and the separator
/// between bracketed items take the profile's spelling. A leading `=`
/// is preserved. Pure text rewrite: the grammar is never validated.
std::string localize_formula_text(std::string_view formula, ExcelProfile profile);

/// Inverse of `localize_formula_text`: reads `formula` as `profile`'s
/// spelling and returns the en-invariant text. `canonicalize(localize(f, p),
/// p) == f` holds for any `f` without exponent number literals.
std::string canonicalize_formula_text(std::string_view formula, ExcelProfile profile);

/// One piece of a rewritten formula and the input bytes it came from. Byte
/// offsets; a verbatim piece is a byte-for-byte copy of its input span.
struct FormulaSourceSpan {
  std::size_t out_begin = 0;
  std::size_t out_end = 0;
  std::size_t in_begin = 0;
  std::size_t in_end = 0;
  bool verbatim = false;
};

/// A rewritten formula with the map from its text back to the input text.
/// `spans` are ascending and tile `text`.
struct MappedFormulaText {
  std::string text;
  std::vector<FormulaSourceSpan> spans;

  /// Maps `range`, in UTF-16 units of `text`, onto `source` (the text the
  /// rewrite read): an offset inside a verbatim piece keeps its place, one
  /// inside a rewritten token widens to the whole input token. `line` and
  /// `column` are recomputed from `source` the way the tokenizer counts them.
  parser::TextRange source_range(std::string_view source, parser::TextRange range) const;
};

/// `localize_formula_text` that also returns the position map.
MappedFormulaText localize_formula_text_mapped(std::string_view formula, ExcelProfile profile);

/// `canonicalize_formula_text` that also returns the position map, so a parse
/// error in the canonical text can be reported at its place in `formula`.
MappedFormulaText canonicalize_formula_text_mapped(std::string_view formula, ExcelProfile profile);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_FORMULA_LOCALIZE_H_
