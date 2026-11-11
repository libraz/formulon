// Rewrites formula text between the stored (en-invariant) form and a profile's spelling.

#ifndef FORMULON_EVAL_FORMULA_LOCALIZE_H_
#define FORMULON_EVAL_FORMULA_LOCALIZE_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "excel_profile.h"
#include "parser/parse_status.h"
#include "parser/token.h"
#include "utils/expected.h"

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
/// (`1.5E-3` reads back as `0.0015`). Local sheet qualifiers are requoted by
/// the profile's display rule (`local_sheet_needs_quoting_display`), and a
/// defined name that reads as a reference in the profile's R1C1 letters is
/// quoted (de `z` -> `'z'`). Everything else, including whitespace, strings
/// and column names, is kept verbatim. Inside structured-reference brackets
/// the `#All`/`#Data`/`#Headers`/`#Totals` words and the separator between
/// bracketed items take the profile's spelling. A leading `=` is preserved.
/// The grammar is not validated; text that does not parse quotes no names.
std::string localize_formula_text(std::string_view formula, ExcelProfile profile);

/// Inverse of `localize_formula_text`: reads `formula` as `profile`'s
/// spelling and returns the en-invariant text, local sheet qualifiers
/// requoted by the stored rule (`local_sheet_needs_quoting_a1`).
/// `canonicalize(localize(f, p), p) == f` holds for any `f` without exponent
/// number literals whose qualifiers are quoted as stored.
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

/// How the localized side spells references: A1, or R1C1 measured from a
/// base cell with the profile's row / column letters and brackets.
struct FormulaNotation {
  bool r1c1 = false;
  bool has_base = false;
  std::uint32_t base_row = 0;  ///< 0-based.
  std::uint32_t base_col = 0;  ///< 0-based.
};

/// `localize_formula_text_mapped` from one parse of `formula`, in either
/// notation. With `notation.r1c1` it also spells every reference in R1C1:
/// absolute axes as the letter and the 1-based index, relative axes as the
/// plain difference from the base in the profile's brackets (the bare letter
/// at offset 0), a whole column or row as one axis. Qualifiers then follow
/// `local_sheet_needs_quoting_r1c1_display` and a cell-shaped name prints
/// bare (`_xlpm.B1` -> `B1`). Fails with the parse failure of `formula` (its
/// `=` counted; empty text passes), or with `InvalidReference` over a
/// reference holding a relative axis when there is no base. Ranges are in
/// UTF-16 units of `formula`.
Expected<MappedFormulaText, parser::ParseFailure> localize_formula_text_r1c1(std::string_view formula,
                                                                             ExcelProfile profile,
                                                                             const FormulaNotation& notation);

/// Inverse of `localize_formula_text_r1c1`: with `notation.r1c1`, reads the
/// profile's R1C1 endpoints back into A1 (a relative axis wrapping around the
/// grid edge), then canonicalizes the A1 text; a cell-shaped name takes
/// `_xlpm.` when a LET / LAMBDA binds it and `_xlnm.` otherwise. Fails with
/// `InvalidReference` over an endpoint that is malformed or holds a relative
/// axis with no base, and with the parse failure of the result, mapped onto
/// `formula`. The spans map the result to `formula`.
Expected<MappedFormulaText, parser::ParseFailure> canonicalize_formula_text_r1c1(std::string_view formula,
                                                                                 ExcelProfile profile,
                                                                                 const FormulaNotation& notation);

/// Appends a sheet qualifier without its `!`: `first`, or the 3-D span
/// `first:last` when `last` is non-empty, quoted as one unit when `quote`.
void append_sheet_qualifier(std::string& out, std::string_view first, std::string_view last, bool quote);

/// `[begin, end)`, byte offsets into `text`, as a range in UTF-16 units with
/// the line and column of `begin` counted as the tokenizer counts them.
parser::TextRange formula_text_range(std::string_view text, std::size_t begin, std::size_t end);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_FORMULA_LOCALIZE_H_
