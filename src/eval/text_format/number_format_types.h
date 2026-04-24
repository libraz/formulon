// Copyright 2026 libraz. Licensed under the MIT License.
//
// Internal header -- do not include outside `src/eval/text_format/`.
//
// Shared token / section type definitions and helper prototypes used by the
// `number_format_tokenizer.cpp` and `number_format_render.cpp` translation
// units. The public API lives in `number_format.h`; everything here is an
// implementation detail and may change without notice.

#ifndef FORMULON_EVAL_TEXT_FORMAT_NUMBER_FORMAT_TYPES_H_
#define FORMULON_EVAL_TEXT_FORMAT_NUMBER_FORMAT_TYPES_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

namespace formulon {
namespace eval {
namespace number_format_detail {

// --- Token representation ----------------------------------------------

enum class Tok : std::uint8_t {
  DigitZero,      // `0`
  DigitOpt,       // `#`
  DigitPad,       // `?`
  Point,          // `.`
  Comma,          // `,`
  Percent,        // `%`
  SciPlus,        // `E+` or `e+`
  SciMinus,       // `E-` or `e-`
  At,             // `@`
  DateY2,         // `yy`
  DateY4,         // `yyyy`
  DateMOrMin,     // `m` or `mm` (month or minute; disambiguated in pass 2)
  DateMMM,        // `mmm` (month name short)
  DateMMMM,       // `mmmm` (month name long)
  DateMMMMM,      // `mmmmm` (month name first letter; run length >= 5)
  DateD,          // `d`
  DateDD,         // `dd`
  DateDDD,        // `ddd` (weekday short)
  DateDDDD,       // `dddd` (weekday long)
  DateH,          // `h`
  DateHH,         // `hh`
  DateS,          // `s`
  DateSS,         // `ss`
  DateElapsedH,   // `[h]` / `[hh]` etc.
  DateElapsedM,   // `[m]` / `[mm]`
  DateElapsedS,   // `[s]` / `[ss]`
  AmPm,           // `AM/PM` / `am/pm`
  AP,             // `A/P` / `a/p`
  FracSecDigits,  // `.0` / `.00` / ... when following a second token
  Literal,        // Arbitrary passthrough bytes (quoted / escaped / other)
  Space,          // `_X` underscore-skip: emits a single space placeholder.
  GeneralNumber,  // `General` keyword: renders value via format_double().
  DateM,          // After disambiguation: month.
  DateMM,         // After disambiguation: 2-digit month.
  DateMin,        // After disambiguation: minute.
  DateMMMin,      // After disambiguation: 2-digit minute.
};

struct Token {
  Tok kind = Tok::Literal;
  // Length hint (0..4) used for date tokens to carry the run length encoded
  // in the format (e.g. yyyy vs yy, hh vs h). Also reused by FracSecDigits
  // to store the fractional-digit count.
  std::uint8_t width = 0;
  // Byte range inside the format string (Literal only). `lit_begin <=
  // lit_end`; empty slices are legal and simply contribute no output.
  std::size_t lit_begin = 0;
  std::size_t lit_end = 0;
};

struct Section {
  std::vector<Token> tokens;

  // Set by the tokenizer when a `[...]` bracket specifier is neither one of
  // the recognised elapsed-time markers (`[h]`, `[m]`, `[s]`) nor a
  // recognised locale currency code (`[$...]`). Mac Excel ja-JP rejects
  // TEXT with colour brackets (e.g. `[Red]0.00`) and other bracketed
  // qualifiers, so `apply_format` surfaces `#VALUE!` whenever this flag
  // is set on any section it would have rendered.
  bool has_invalid_bracket = false;

  // Precomputed summary for numeric classification. Populated by `classify`.
  bool is_date = false;
  bool is_text = false;
  // Set if the section contains at least one `General` keyword token. The
  // renderer emits `format_double(abs(value))` for such tokens; no digit
  // accounting is performed (General is self-contained).
  bool has_general = false;
  int integer_zero_digits = 0;
  int integer_opt_digits = 0;
  int integer_pad_digits = 0;
  int fraction_zero_digits = 0;
  int fraction_opt_digits = 0;
  int fraction_pad_digits = 0;
  bool has_point = false;
  bool has_percent = false;
  // Count of `,` tokens adjacent to the decimal point (trailing commas);
  // each divides the scaled value by 1000.
  int trailing_comma_scale = 0;
  // Whether any `,` appears between two digit tokens in the integer part
  // (enables thousands separators).
  bool thousands_separator = false;
  // Scientific exponent sign handling, if present.
  bool has_scientific = false;
  bool sci_plus = false;
  int sci_digits = 0;
  // Number of fractional-second digits (after `[s]/s/ss` + `.0...`).
  int frac_sec_digits = 0;
};

// --- Tokenizer entry points (implemented in number_format_tokenizer.cpp) ---

// Splits `fmt` into up to 4 sections on unquoted `;`. Each section is
// returned as a string_view over `fmt`. Quoted (`"..."`) and escaped (`\x`
// / `!x`) characters are skipped so that `;` inside a literal does not
// split.
std::vector<std::string_view> split_sections(std::string_view fmt);

// Tokenizes one section. Writes the token list into `out.tokens` and
// surfaces invalid bracket qualifiers (colour names, conditional tests,
// DBNum markers, etc.) through `out.has_invalid_bracket`.
void tokenize_section(std::string_view fmt, Section& out);

// Populate the numeric/date summary on `section`. Also detect fractional
// seconds `.0...` that immediately follow a second token (used to format
// sub-second precision).
void classify(Section& section) noexcept;

// --- Renderer entry points (implemented in number_format_render.cpp) ---

// Render one numeric section through the walk-tokens pipeline.
void render_numeric(const Section& section, std::string_view fmt, double value, std::string& out);

// Render the date/time section for the given serial.
void render_date(const Section& section, std::string_view fmt, double serial, std::string& out);

// Render the text section by walking the raw format bytes. `original` is
// substituted for every unquoted `@` token.
void render_text_section(const Section& section, std::string_view fmt, std::string_view original, std::string& out);

}  // namespace number_format_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_TEXT_FORMAT_NUMBER_FORMAT_TYPES_H_
