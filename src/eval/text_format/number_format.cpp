// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the Excel TEXT() format-string engine declared in
// `number_format.h`. The design follows the two-phase approach described
// in the scope memo: (1) tokenize the format, splitting on `;` into up to
// four sections; (2) render a value through the section selected by its
// sign/zero/text classification.
//
// Numeric rendering pre-computes the integer / fractional digit counts,
// scales the value for `%` and trailing-comma divides, rounds to the
// fractional precision with half-away-from-zero, then walks the tokens
// once interleaving the pre-formatted digits and literal runs.
//
// Date/time rendering converts the serial via the shared `date_time`
// helpers and substitutes each `y/m/d/h/s` token by its textual form.
// The `m` vs minute disambiguation is resolved in a small second pass
// over the token stream driven by the surrounding `h:` / `:s` context.

#include "eval/text_format/number_format.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <string_view>
#include <vector>

#include "eval/date_time.h"

namespace formulon {
namespace eval {
namespace text_format {
namespace {

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

// ---------------------------------------------------------------------------
// Tokenizer
// ---------------------------------------------------------------------------

// Returns true iff `c` is one of `yYmMdDhHsS`. Used to detect date-family
// tokens; AM/PM handled separately because they span multiple characters.
bool is_date_letter(char c) noexcept {
  return c == 'y' || c == 'Y' || c == 'm' || c == 'M' || c == 'd' || c == 'D' || c == 'h' || c == 'H' || c == 's' ||
         c == 'S';
}

// Parses a run of `letter` characters (case-insensitive for the given
// target). Advances `*i` past the run. Returns the run length (>= 1 given
// the caller already matched at least one character).
std::size_t scan_run(std::string_view fmt, std::size_t& i, char letter) noexcept {
  const char upper = letter >= 'a' && letter <= 'z' ? static_cast<char>(letter - 32) : letter;
  const char lower = letter >= 'A' && letter <= 'Z' ? static_cast<char>(letter + 32) : letter;
  std::size_t start = i;
  while (i < fmt.size() && (fmt[i] == upper || fmt[i] == lower)) {
    ++i;
  }
  return i - start;
}

// Splits `fmt` into up to 4 sections on unquoted `;`. Each section is
// returned as a string_view over `fmt`. Quoted (`"..."`) and escaped (`\x`
// / `!x`) characters are skipped so that `;` inside a literal does not
// split.
std::vector<std::string_view> split_sections(std::string_view fmt) {
  std::vector<std::string_view> out;
  std::size_t start = 0;
  for (std::size_t i = 0; i < fmt.size();) {
    const char c = fmt[i];
    if (c == '"') {
      // Skip to closing quote.
      ++i;
      while (i < fmt.size() && fmt[i] != '"') {
        ++i;
      }
      if (i < fmt.size()) {
        ++i;
      }
      continue;
    }
    if (c == '\\' || c == '!') {
      i += 2;  // Skip escape + next byte.
      continue;
    }
    if (c == '[') {
      // Skip to matching `]` so e.g. `[Red]` does not interact with `;`.
      while (i < fmt.size() && fmt[i] != ']') {
        ++i;
      }
      if (i < fmt.size()) {
        ++i;
      }
      continue;
    }
    if (c == ';') {
      out.emplace_back(fmt.substr(start, i - start));
      start = i + 1;
      ++i;
      continue;
    }
    ++i;
  }
  out.emplace_back(fmt.substr(start));
  return out;
}

// Tokenizes one section. Writes the token list into `out.tokens` and
// surfaces invalid bracket qualifiers (colour names, conditional tests,
// DBNum markers, etc.) through `out.has_invalid_bracket`.
void tokenize_section(std::string_view fmt, Section& out) {
  std::vector<Token>& toks = out.tokens;
  auto push_literal = [&](std::size_t b, std::size_t e) {
    if (b == e) {
      return;
    }
    Token t;
    t.kind = Tok::Literal;
    t.lit_begin = b;
    t.lit_end = e;
    toks.push_back(t);
  };

  std::size_t i = 0;
  while (i < fmt.size()) {
    const char c = fmt[i];
    // Quoted literal `"..."`.
    if (c == '"') {
      std::size_t j = i + 1;
      while (j < fmt.size() && fmt[j] != '"') {
        ++j;
      }
      push_literal(i + 1, j);
      i = j < fmt.size() ? j + 1 : j;
      continue;
    }
    // Escape `\x` or `!x` -> next byte is a literal.
    if ((c == '\\' || c == '!') && i + 1 < fmt.size()) {
      push_literal(i + 1, i + 2);
      i += 2;
      continue;
    }
    // Bracketed specifier. Recognised kinds:
    //   `[h]` / `[m]` / `[s]`   -> elapsed-time tokens (any run length).
    //   `[$...]`                -> locale-currency marker; silently dropped.
    // Anything else (`[Red]`, `[Blue]`, `[>100]`, `[DBNum1]`, ...) is a
    // rejection: Mac Excel ja-JP returns #VALUE! for TEXT with a colour
    // qualifier, so we flag the section as invalid and let `apply_format`
    // surface that as kValueError.
    if (c == '[') {
      std::size_t j = i + 1;
      while (j < fmt.size() && fmt[j] != ']') {
        ++j;
      }
      const std::string_view body = fmt.substr(i + 1, (j < fmt.size() ? j : fmt.size()) - (i + 1));
      i = j < fmt.size() ? j + 1 : j;
      // Elapsed time markers: any run of `h`, `m`, or `s` (case-insensitive).
      bool all_h = !body.empty();
      bool all_m = !body.empty();
      bool all_s = !body.empty();
      for (char ch : body) {
        const char lo = (ch >= 'A' && ch <= 'Z') ? static_cast<char>(ch + 32) : ch;
        if (lo != 'h') {
          all_h = false;
        }
        if (lo != 'm') {
          all_m = false;
        }
        if (lo != 's') {
          all_s = false;
        }
      }
      if (all_h) {
        Token t;
        t.kind = Tok::DateElapsedH;
        t.width = static_cast<std::uint8_t>(body.size());
        toks.push_back(t);
      } else if (all_m) {
        Token t;
        t.kind = Tok::DateElapsedM;
        t.width = static_cast<std::uint8_t>(body.size());
        toks.push_back(t);
      } else if (!body.empty() && body.front() == '$') {
        // Locale-currency marker (e.g. `[$-409]`, `[$JPY]`). Silently
        // discarded; the trailing numeric/date tokens supply the rendered
        // value.
      } else if (all_s) {
        Token t;
        t.kind = Tok::DateElapsedS;
        t.width = static_cast<std::uint8_t>(body.size());
        toks.push_back(t);
      } else {
        // Anything else - colour names, conditional tests, DBNum digits,
        // unknown qualifiers - trips the invalid-bracket flag.
        out.has_invalid_bracket = true;
      }
      continue;
    }
    // AM/PM (case-insensitive). Match the longest valid prefix. We treat
    // `AM/PM`, `am/pm`, `A/P`, `a/p` as indivisible markers.
    if (c == 'A' || c == 'a' || c == 'P' || c == 'p') {
      auto match_ci = [&](std::size_t start, const char* a) -> bool {
        std::size_t n = 0;
        while (a[n] != '\0') {
          ++n;
        }
        if (start + n > fmt.size()) {
          return false;
        }
        for (std::size_t k = 0; k < n; ++k) {
          const char fc = fmt[start + k];
          const char ac = a[k];
          const char fc_lower = (fc >= 'A' && fc <= 'Z') ? static_cast<char>(fc + 32) : fc;
          const char ac_lower = (ac >= 'A' && ac <= 'Z') ? static_cast<char>(ac + 32) : ac;
          if (fc_lower != ac_lower) {
            return false;
          }
        }
        return true;
      };
      if (match_ci(i, "AM/PM")) {
        Token t;
        t.kind = Tok::AmPm;
        toks.push_back(t);
        i += 5;
        continue;
      }
      if (match_ci(i, "A/P")) {
        Token t;
        t.kind = Tok::AP;
        toks.push_back(t);
        i += 3;
        continue;
      }
    }
    // Date letters.
    if (is_date_letter(c)) {
      const char lc = (c >= 'A' && c <= 'Z') ? static_cast<char>(c + 32) : c;
      const std::size_t run = scan_run(fmt, i, lc);
      Token t;
      t.width = static_cast<std::uint8_t>(run);
      switch (lc) {
        case 'y':
          t.kind = (run <= 2) ? Tok::DateY2 : Tok::DateY4;
          break;
        case 'm':
          // Month/minute disambiguation happens in pass 2.
          if (run >= 4) {
            t.kind = Tok::DateMMMM;
          } else if (run == 3) {
            t.kind = Tok::DateMMM;
          } else {
            t.kind = Tok::DateMOrMin;
          }
          break;
        case 'd':
          if (run >= 4) {
            t.kind = Tok::DateDDDD;
          } else if (run == 3) {
            t.kind = Tok::DateDDD;
          } else if (run == 2) {
            t.kind = Tok::DateDD;
          } else {
            t.kind = Tok::DateD;
          }
          break;
        case 'h':
          t.kind = (run >= 2) ? Tok::DateHH : Tok::DateH;
          break;
        case 's':
          t.kind = (run >= 2) ? Tok::DateSS : Tok::DateS;
          break;
        default:
          t.kind = Tok::Literal;
          t.lit_begin = i - run;
          t.lit_end = i;
          break;
      }
      toks.push_back(t);
      continue;
    }
    // Scientific notation `E+` / `E-` / `e+` / `e-`. Must immediately
    // follow a digit token; we still emit the token and let the classifier
    // require the sign.
    if ((c == 'E' || c == 'e') && i + 1 < fmt.size() && (fmt[i + 1] == '+' || fmt[i + 1] == '-')) {
      Token t;
      t.kind = fmt[i + 1] == '+' ? Tok::SciPlus : Tok::SciMinus;
      toks.push_back(t);
      i += 2;
      continue;
    }
    // Numeric specifiers.
    switch (c) {
      case '0': {
        Token t;
        t.kind = Tok::DigitZero;
        toks.push_back(t);
        ++i;
        continue;
      }
      case '#': {
        Token t;
        t.kind = Tok::DigitOpt;
        toks.push_back(t);
        ++i;
        continue;
      }
      case '?': {
        Token t;
        t.kind = Tok::DigitPad;
        toks.push_back(t);
        ++i;
        continue;
      }
      case '.': {
        Token t;
        t.kind = Tok::Point;
        toks.push_back(t);
        ++i;
        continue;
      }
      case ',': {
        Token t;
        t.kind = Tok::Comma;
        toks.push_back(t);
        ++i;
        continue;
      }
      case '%': {
        Token t;
        t.kind = Tok::Percent;
        toks.push_back(t);
        ++i;
        continue;
      }
      case '@': {
        Token t;
        t.kind = Tok::At;
        toks.push_back(t);
        ++i;
        continue;
      }
      default:
        break;
    }
    // Fallback: single byte literal. Note UTF-8 multi-byte characters are
    // handled one byte at a time; the output is the same as the input
    // bytes so no decoding is required.
    push_literal(i, i + 1);
    ++i;
  }
}

// ---------------------------------------------------------------------------
// Classification and `m` disambiguation.
// ---------------------------------------------------------------------------

// Returns true if `tok` is a date-family token (including elapsed brackets).
bool is_date_tok(Tok t) noexcept {
  switch (t) {
    case Tok::DateY2:
    case Tok::DateY4:
    case Tok::DateMOrMin:
    case Tok::DateMMM:
    case Tok::DateMMMM:
    case Tok::DateD:
    case Tok::DateDD:
    case Tok::DateDDD:
    case Tok::DateDDDD:
    case Tok::DateH:
    case Tok::DateHH:
    case Tok::DateS:
    case Tok::DateSS:
    case Tok::DateElapsedH:
    case Tok::DateElapsedM:
    case Tok::DateElapsedS:
    case Tok::AmPm:
    case Tok::AP:
    case Tok::DateM:
    case Tok::DateMM:
    case Tok::DateMin:
    case Tok::DateMMMin:
      return true;
    default:
      return false;
  }
}

// After tokenization, rewrite DateMOrMin tokens into either DateM/DateMM
// (month) or DateMin/DateMMMin (minute) based on surrounding context.
// Rule (matches Excel): a DateMOrMin between an hour token and a second
// token, or immediately following an hour token, or immediately preceding
// a second token, is interpreted as minute. Otherwise it's a month.
void disambiguate_minutes(std::vector<Token>& toks) noexcept {
  // Walk left->right to find hour tokens; after an hour, treat subsequent
  // DateMOrMin as minute until we see a non-date token that isn't a
  // separator (colon, space, literal). Also walk right->left: any
  // DateMOrMin immediately preceding a second token is a minute.
  auto is_minute_neighbor = [](const Token& t) {
    return t.kind == Tok::DateH || t.kind == Tok::DateHH || t.kind == Tok::DateS || t.kind == Tok::DateSS ||
           t.kind == Tok::DateElapsedH || t.kind == Tok::DateElapsedS || t.kind == Tok::DateMin ||
           t.kind == Tok::DateMMMin || t.kind == Tok::DateElapsedM;
  };
  // Pass 1 (forward): after hour, next DateMOrMin -> minute.
  for (std::size_t i = 0; i < toks.size(); ++i) {
    if (toks[i].kind == Tok::DateMOrMin) {
      // Look backward for the nearest non-literal, non-separator token.
      for (std::size_t k = i; k > 0; --k) {
        const Token& prev = toks[k - 1];
        if (prev.kind == Tok::Literal) {
          continue;  // Skip `:` or space.
        }
        if (is_minute_neighbor(prev)) {
          toks[i].kind = (toks[i].width >= 2) ? Tok::DateMMMin : Tok::DateMin;
        }
        break;
      }
    }
  }
  // Pass 2 (backward): DateMOrMin preceding a second token -> minute.
  for (std::size_t i = 0; i < toks.size(); ++i) {
    if (toks[i].kind != Tok::DateMOrMin) {
      continue;
    }
    for (std::size_t k = i + 1; k < toks.size(); ++k) {
      const Token& nxt = toks[k];
      if (nxt.kind == Tok::Literal) {
        continue;
      }
      if (nxt.kind == Tok::DateS || nxt.kind == Tok::DateSS || nxt.kind == Tok::DateElapsedS) {
        toks[i].kind = (toks[i].width >= 2) ? Tok::DateMMMin : Tok::DateMM;
        // Note: when only the following token is a second, the `m` is a
        // minute too. Use DateMin rather than DateMM if width == 1.
        if (toks[i].width < 2) {
          toks[i].kind = Tok::DateMin;
        } else {
          toks[i].kind = Tok::DateMMMin;
        }
      }
      break;
    }
  }
  // Remaining DateMOrMin tokens default to month.
  for (auto& t : toks) {
    if (t.kind == Tok::DateMOrMin) {
      t.kind = (t.width >= 2) ? Tok::DateMM : Tok::DateM;
    }
  }
}

// Populate the numeric/date summary on `section`. Also detect fractional
// seconds `.0...` that immediately follow a second token (used to format
// sub-second precision).
void classify(Section& section) noexcept {
  disambiguate_minutes(section.tokens);
  bool any_date = false;
  bool any_at = false;
  int int_zero = 0;
  int int_opt = 0;
  int int_pad = 0;
  int frac_zero = 0;
  int frac_opt = 0;
  int frac_pad = 0;
  bool has_percent = false;
  bool saw_sci = false;
  bool sci_plus = false;
  int sci_digits = 0;
  bool has_digit_before_comma = false;
  bool has_digit_after_comma = false;
  int trailing_commas = 0;
  int last_digit_index = -1;

  auto is_digit_tok = [](Tok k) { return k == Tok::DigitZero || k == Tok::DigitOpt || k == Tok::DigitPad; };

  // First locate the decimal point (if any).
  int point_index = -1;
  for (std::size_t i = 0; i < section.tokens.size(); ++i) {
    if (section.tokens[i].kind == Tok::Point) {
      point_index = static_cast<int>(i);
      break;
    }
  }
  section.has_point = point_index >= 0;

  // Count integer vs fraction digits.
  for (std::size_t i = 0; i < section.tokens.size(); ++i) {
    const Token& tk = section.tokens[i];
    if (is_date_tok(tk.kind)) {
      any_date = true;
    }
    if (tk.kind == Tok::At) {
      any_at = true;
    }
    if (tk.kind == Tok::Percent) {
      has_percent = true;
    }
    if (tk.kind == Tok::SciPlus || tk.kind == Tok::SciMinus) {
      saw_sci = true;
      sci_plus = tk.kind == Tok::SciPlus;
    }
    if (is_digit_tok(tk.kind)) {
      last_digit_index = static_cast<int>(i);
      const bool in_integer = !saw_sci && ((point_index < 0) || (static_cast<int>(i) < point_index));
      if (in_integer) {
        if (tk.kind == Tok::DigitZero) {
          ++int_zero;
        }
        if (tk.kind == Tok::DigitOpt) {
          ++int_opt;
        }
        if (tk.kind == Tok::DigitPad) {
          ++int_pad;
        }
        has_digit_before_comma = true;
      } else if (!saw_sci) {
        // Between `.` and the optional `E+/E-` marker.
        if (tk.kind == Tok::DigitZero) {
          ++frac_zero;
        }
        if (tk.kind == Tok::DigitOpt) {
          ++frac_opt;
        }
        if (tk.kind == Tok::DigitPad) {
          ++frac_pad;
        }
      } else {
        // After the scientific marker: digits describe exponent width.
        ++sci_digits;
      }
    }
  }
  // Count trailing commas between last integer digit and either `.` or
  // end-of-integer. Each such comma divides by 1000.
  if (last_digit_index >= 0) {
    std::size_t end = point_index >= 0 ? static_cast<std::size_t>(point_index) : section.tokens.size();
    for (std::size_t i = static_cast<std::size_t>(last_digit_index) + 1; i < end; ++i) {
      if (section.tokens[i].kind == Tok::Comma) {
        ++trailing_commas;
      }
    }
    // Thousands-separator test: any `,` between two digit tokens triggers it.
    bool seen_digit = false;
    for (std::size_t i = 0; i <= static_cast<std::size_t>(last_digit_index); ++i) {
      const Token& tk = section.tokens[i];
      if (is_digit_tok(tk.kind)) {
        if (seen_digit) {
          // Already had a digit; look backward for a `,` between.
        }
        seen_digit = true;
      } else if (tk.kind == Tok::Comma && seen_digit) {
        // Peek forward to confirm another digit follows.
        for (std::size_t j = i + 1; j <= static_cast<std::size_t>(last_digit_index); ++j) {
          if (is_digit_tok(section.tokens[j].kind)) {
            has_digit_after_comma = true;
            break;
          }
        }
      }
    }
  }

  // Fractional seconds: look for `.` immediately after a second token,
  // followed by one or more `0`/`#` digits. If found, remove them from the
  // fractional-digit counts (they belong to seconds, not to the numeric
  // side) and record the fractional-second digit count on the section.
  //
  // The simple rule: if the tokens contain any date token AND a
  // `DateS/DateSS/DateElapsedS` followed by `Point` then digit tokens,
  // treat that group as fractional seconds.
  if (any_date) {
    for (std::size_t i = 0; i + 1 < section.tokens.size(); ++i) {
      if ((section.tokens[i].kind == Tok::DateS || section.tokens[i].kind == Tok::DateSS ||
           section.tokens[i].kind == Tok::DateElapsedS) &&
          section.tokens[i + 1].kind == Tok::Point) {
        int digits = 0;
        std::size_t j = i + 2;
        while (j < section.tokens.size() && is_digit_tok(section.tokens[j].kind)) {
          ++digits;
          ++j;
        }
        if (digits > 0) {
          section.frac_sec_digits = digits;
          // Convert the point + digits into a FracSecDigits token so the
          // renderer skips them when emitting the regular date sequence.
          Token marker;
          marker.kind = Tok::FracSecDigits;
          marker.width = static_cast<std::uint8_t>(digits);
          section.tokens[i + 1] = marker;
          // Clear the fractional digit tokens (mark them as empty literals).
          for (std::size_t k = i + 2; k < j; ++k) {
            section.tokens[k].kind = Tok::Literal;
            section.tokens[k].lit_begin = 0;
            section.tokens[k].lit_end = 0;
          }
          // Also subtract these digits from the numeric counts (they were
          // counted as fraction_*).
          frac_zero = 0;
          frac_opt = 0;
          frac_pad = 0;
        }
        break;
      }
    }
  }

  section.is_date = any_date;
  // A section is classified as `text` whenever it contains an `@` token and
  // no numeric digit tokens. Stray date letters that slipped into literal
  // phrases (e.g. the `s` in "text is @") are tolerated: we demote them to
  // literals during text rendering by ignoring any non-`@` / non-Literal
  // token in `render_text_section`.
  section.is_text =
      any_at && int_zero == 0 && int_opt == 0 && int_pad == 0 && frac_zero == 0 && frac_opt == 0 && frac_pad == 0;
  section.integer_zero_digits = int_zero;
  section.integer_opt_digits = int_opt;
  section.integer_pad_digits = int_pad;
  section.fraction_zero_digits = frac_zero;
  section.fraction_opt_digits = frac_opt;
  section.fraction_pad_digits = frac_pad;
  section.has_percent = has_percent;
  section.trailing_comma_scale = trailing_commas;
  section.thousands_separator = has_digit_before_comma && has_digit_after_comma;
  section.has_scientific = saw_sci;
  section.sci_plus = sci_plus;
  section.sci_digits = sci_digits;
}

// ---------------------------------------------------------------------------
// Numeric rendering.
// ---------------------------------------------------------------------------

// Rounds `v` to `decimals` fractional places, half-away-from-zero. The
// result's integer and fractional digit strings are written into
// `*int_digits` and `*frac_digits` (decimal digits only, no sign, no
// leading zeros on the integer side for zero values — except the integer
// is always at least `"0"`).
//
// This helper does not handle scientific notation; see `render_scientific`.
void format_fixed_digits(double v, int decimals, bool* negative, std::string* int_digits, std::string* frac_digits) {
  *negative = std::signbit(v);
  double abs_v = std::fabs(v);
  // Use snprintf with requested precision so we pick up libc rounding.
  // `%.*f` in the C locale rounds ties away from zero on the two libcs
  // Formulon targets (glibc, musl, and Apple's libc); this matches
  // Excel's observable behaviour for TEXT's common cases.
  char buf[64];
  const int n = std::snprintf(buf, sizeof(buf), "%.*f", decimals < 0 ? 0 : decimals, abs_v);
  if (n < 0 || static_cast<std::size_t>(n) >= sizeof(buf)) {
    // Fallback: fall back to sprintf with a heap buffer (extremely rare).
    std::string out;
    out.resize(static_cast<std::size_t>(decimals < 0 ? 0 : decimals) + 32u);
    const int m = std::snprintf(&out[0], out.size(), "%.*f", decimals < 0 ? 0 : decimals, abs_v);
    if (m > 0) {
      out.resize(static_cast<std::size_t>(m));
    } else {
      out = "0";
    }
    const std::size_t dot = out.find('.');
    if (dot == std::string::npos) {
      *int_digits = out;
      frac_digits->clear();
    } else {
      *int_digits = out.substr(0, dot);
      *frac_digits = out.substr(dot + 1);
    }
    return;
  }
  std::string_view s(buf, static_cast<std::size_t>(n));
  const std::size_t dot = s.find('.');
  if (dot == std::string_view::npos) {
    int_digits->assign(s);
    frac_digits->clear();
  } else {
    int_digits->assign(s.substr(0, dot));
    frac_digits->assign(s.substr(dot + 1));
  }
}

// Insert ASCII thousands separators into `int_digits`. Walks right-to-left
// emitting a comma every three digits. `int_digits` is overwritten.
void insert_thousands(std::string& int_digits) {
  if (int_digits.size() <= 3) {
    return;
  }
  std::string out;
  out.reserve(int_digits.size() + int_digits.size() / 3);
  const std::size_t n = int_digits.size();
  const std::size_t rem = n % 3;
  std::size_t k = 0;
  if (rem > 0) {
    out.append(int_digits, 0, rem);
    k = rem;
  }
  while (k < n) {
    if (k != 0) {
      out.push_back(',');
    }
    out.append(int_digits, k, 3);
    k += 3;
  }
  int_digits = std::move(out);
}

// Render one numeric section through the walk-tokens pipeline.
//
// Steps:
//   1. Apply scaling (percent, trailing-comma divisions).
//   2. Compute fractional-digit precision = max(zero+pad, opt) ... in Excel
//      practice, the precision equals the total count of digit tokens in
//      the fractional part. Excel uses `zero + opt + pad` for the rounded
//      precision.
//   3. Round the absolute value to that precision; capture sign separately.
//   4. Walk the tokens and weave the integer/fraction digit strings into
//      the output alongside literals, commas, and percent.
void render_numeric(const Section& section, std::string_view fmt, double value, std::string& out) {
  double scaled = value;
  if (section.has_percent) {
    scaled *= 100.0;
  }
  for (int i = 0; i < section.trailing_comma_scale; ++i) {
    scaled /= 1000.0;
  }
  const int frac_digits = section.fraction_zero_digits + section.fraction_opt_digits + section.fraction_pad_digits;
  bool negative = false;
  std::string int_digits;
  std::string frac_digits_str;

  if (section.has_scientific) {
    // Scientific notation: compute mantissa/exponent, format mantissa
    // through the fixed path, then append the exponent with the configured
    // sign behaviour.
    double mantissa = scaled;
    int exponent = 0;
    if (mantissa != 0.0) {
      // Normalise mantissa so the integer part has `integer_digits_total`
      // digits; default is one digit.
      const int int_total = section.integer_zero_digits + section.integer_opt_digits + section.integer_pad_digits;
      const int want_int_digits = int_total > 0 ? int_total : 1;
      const double abs_m = std::fabs(mantissa);
      const double log10_m = std::log10(abs_m);
      // Target: floor(log10(|mantissa|)) == want_int_digits - 1.
      const int target = want_int_digits - 1;
      int shift = static_cast<int>(std::floor(log10_m)) - target;
      // Walk toward target in a stable way (avoid pow() rounding drift for
      // exact powers of 10).
      if (shift > 0) {
        for (int i = 0; i < shift; ++i) {
          mantissa /= 10.0;
        }
      } else if (shift < 0) {
        for (int i = 0; i < -shift; ++i) {
          mantissa *= 10.0;
        }
      }
      exponent = shift;
    }
    format_fixed_digits(mantissa, frac_digits, &negative, &int_digits, &frac_digits_str);
    // Fall through into the standard walker below, but also remember the
    // exponent to emit when we hit the scientific marker.
    // We lay out a specialised walker here rather than reusing the fixed
    // path's literal copier because the scientific token needs to consume
    // both the SciPlus/SciMinus token and the subsequent digit tokens that
    // describe the exponent width.
    std::string result;
    if (negative) {
      result.push_back('-');
    }
    // int_digits is already normalised by the scaling + format_fixed_digits
    // call above; append it verbatim.
    result.append(int_digits);
    if (section.has_point && frac_digits > 0) {
      result.push_back('.');
      // Pad `frac_digits_str` to `frac_digits` characters.
      if (static_cast<int>(frac_digits_str.size()) < frac_digits) {
        frac_digits_str.append(static_cast<std::size_t>(frac_digits) - frac_digits_str.size(), '0');
      } else if (static_cast<int>(frac_digits_str.size()) > frac_digits) {
        frac_digits_str.resize(static_cast<std::size_t>(frac_digits));
      }
      result.append(frac_digits_str);
    }
    // Exponent marker + digits. Sign emission: SciPlus always emits '+' or
    // '-'; SciMinus emits only '-'.
    // Use capital or lowercase 'E' matching the format string? In this
    // simplified implementation we always emit 'E'. (Mac Excel on a ja-JP
    // keyboard renders a capital E by default.)
    result.push_back('E');
    if (exponent >= 0) {
      if (section.sci_plus) {
        result.push_back('+');
      }
    } else {
      result.push_back('-');
    }
    const int abs_exp = exponent >= 0 ? exponent : -exponent;
    std::string exp_str = std::to_string(abs_exp);
    // Pad to at least `sci_digits` digits.
    if (static_cast<int>(exp_str.size()) < section.sci_digits) {
      exp_str.insert(exp_str.begin(), static_cast<std::size_t>(section.sci_digits) - exp_str.size(), '0');
    }
    result.append(exp_str);
    out.append(result);
    (void)fmt;
    return;
  }

  format_fixed_digits(scaled, frac_digits, &negative, &int_digits, &frac_digits_str);

  // Pad integer digits to the required minimum (zero + pad digits). Leading
  // `#` tokens above the actual digit count drop silently.
  const int int_min = section.integer_zero_digits + section.integer_pad_digits;
  if (static_cast<int>(int_digits.size()) < int_min) {
    int_digits.insert(0, static_cast<std::size_t>(int_min) - int_digits.size(), '0');
  }

  // Remove a redundant leading "0" for values with an integer part of zero
  // when the format has only `#` digits in the integer part.
  if (section.integer_zero_digits == 0 && section.integer_pad_digits == 0 && int_digits == "0") {
    int_digits.clear();
  }

  if (section.thousands_separator) {
    insert_thousands(int_digits);
  }

  // Fraction adjustment: the snprintf produced exactly `frac_digits` chars.
  // Trim trailing zeros matching the `#` tokens (scan right-to-left).
  int trim = section.fraction_opt_digits;
  while (trim > 0 && !frac_digits_str.empty() && frac_digits_str.back() == '0') {
    frac_digits_str.pop_back();
    --trim;
  }

  // Now walk tokens and emit output.
  std::string result;
  if (negative) {
    result.push_back('-');
  }
  bool emitted_integer = false;
  std::size_t frac_cursor = 0;

  auto emit_integer_block = [&]() {
    if (!emitted_integer) {
      result.append(int_digits);
      emitted_integer = true;
    }
  };

  // We handle the integer block as a unit: once any digit token in the
  // integer part is reached, we splat `int_digits` into the output and
  // ignore subsequent integer digit tokens (they were accounted for during
  // classification). Literals intermixed with integer digits are emitted
  // in-place; the single-block strategy is accurate for the common cases
  // (`#,##0` / `0.00` / `0%` / `0E+00`) and does not model format-strings
  // that want to interleave literal text between digit positions.
  bool past_point = false;
  for (std::size_t i = 0; i < section.tokens.size(); ++i) {
    const Token& tk = section.tokens[i];
    switch (tk.kind) {
      case Tok::DigitZero:
      case Tok::DigitOpt:
      case Tok::DigitPad:
        if (past_point) {
          if (frac_cursor < frac_digits_str.size()) {
            result.push_back(frac_digits_str[frac_cursor]);
            ++frac_cursor;
          } else {
            // Exceeds precision; emit padding based on token kind.
            if (tk.kind == Tok::DigitZero) {
              result.push_back('0');
            } else if (tk.kind == Tok::DigitPad) {
              result.push_back(' ');
            }
          }
        } else {
          emit_integer_block();
        }
        break;
      case Tok::Point:
        if (section.fraction_zero_digits + section.fraction_opt_digits + section.fraction_pad_digits > 0 ||
            !frac_digits_str.empty()) {
          result.push_back('.');
        }
        past_point = true;
        break;
      case Tok::Comma:
        // Only emit as a literal `,` when it is neither a trailing-scale
        // comma nor the thousands-separator marker (both of which were
        // handled during classification / integer block emission).
        break;
      case Tok::Percent:
        result.push_back('%');
        break;
      case Tok::Literal:
        if (tk.lit_end > tk.lit_begin) {
          result.append(fmt.data() + tk.lit_begin, tk.lit_end - tk.lit_begin);
        }
        break;
      case Tok::SciPlus:
      case Tok::SciMinus:
        // Scientific path handled above; should not reach here.
        break;
      case Tok::At:
        // `@` in a numeric context is ignored (Excel emits nothing).
        break;
      default:
        break;
    }
  }
  out.append(result);
}

// ---------------------------------------------------------------------------
// Date/time rendering.
// ---------------------------------------------------------------------------

const char* month_short(unsigned m) noexcept {
  // Mac Excel ja-JP surprisingly renders `mmm` in English (Jan/Feb/...).
  // The Japanese `N月` form is reserved for `[DBNum2]` and friends, which
  // are out of scope here.
  static const char* kTable[12] = {"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"};
  if (m < 1u || m > 12u) {
    return "";
  }
  return kTable[m - 1u];
}

const char* month_long(unsigned m) noexcept {
  // Matches Mac Excel ja-JP: `mmmm` renders as the English full name.
  static const char* kTable[12] = {"January", "February", "March",     "April",   "May",      "June",
                                   "July",    "August",   "September", "October", "November", "December"};
  if (m < 1u || m > 12u) {
    return "";
  }
  return kTable[m - 1u];
}

const char* weekday_short(int sun0) noexcept {
  // Mac Excel ja-JP `ddd` returns English 3-letter weekday abbreviations.
  static const char* kTable[7] = {"Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"};
  if (sun0 < 0 || sun0 > 6) {
    return "";
  }
  return kTable[sun0];
}

const char* weekday_long(int sun0) noexcept {
  // Mac Excel ja-JP `dddd` renders the English full weekday name.
  static const char* kTable[7] = {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"};
  if (sun0 < 0 || sun0 > 6) {
    return "";
  }
  return kTable[sun0];
}

void append_pad2(std::string& out, unsigned n) {
  if (n < 10u) {
    out.push_back('0');
  }
  out.append(std::to_string(n));
}

void render_date(const Section& section, std::string_view fmt, double serial, std::string& out) {
  if (serial < 0.0 || serial > 2958465.0) {
    // Excel rejects out-of-range serials from TEXT.
    return;
  }
  const date_time::YMD ymd = date_time::ymd_from_serial(serial);
  // Weekday Sunday=0..Saturday=6 computed from the civil day count.
  const std::int64_t days = date_time::days_from_civil(ymd.y, ymd.m, ymd.d);
  const int sun0 = static_cast<int>(((days + 4) % 7 + 7) % 7);

  // Decompose the time portion with optional fractional seconds.
  const double frac_day = serial - std::floor(serial);
  // Total seconds (float-precision) so fractional seconds survive.
  double total_seconds_f = frac_day * 86400.0;
  // Round to `frac_sec_digits` if requested, otherwise to whole seconds.
  double rounded = total_seconds_f;
  if (section.frac_sec_digits == 0) {
    rounded = std::floor(total_seconds_f + 0.5);
  } else {
    const double scale = std::pow(10.0, section.frac_sec_digits);
    rounded = std::floor(total_seconds_f * scale + 0.5) / scale;
  }
  // Extract integer h/m/s and fractional remainder.
  long long total_int_seconds = static_cast<long long>(std::floor(rounded));
  const double sub_sec_float = rounded - static_cast<double>(total_int_seconds);
  // If AM/PM is in use, we need to know it before formatting hours.
  bool use_am_pm = false;
  for (const Token& tk : section.tokens) {
    if (tk.kind == Tok::AmPm || tk.kind == Tok::AP) {
      use_am_pm = true;
      break;
    }
  }

  const long long seconds_of_day = ((total_int_seconds % 86400) + 86400) % 86400;
  unsigned hour_24 = static_cast<unsigned>(seconds_of_day / 3600);
  unsigned minute = static_cast<unsigned>((seconds_of_day / 60) % 60);
  unsigned second = static_cast<unsigned>(seconds_of_day % 60);
  bool pm = hour_24 >= 12u;
  unsigned hour_for_render = hour_24;
  if (use_am_pm) {
    hour_for_render = hour_24 % 12u;
    if (hour_for_render == 0u) {
      hour_for_render = 12u;
    }
  }

  for (std::size_t i = 0; i < section.tokens.size(); ++i) {
    const Token& tk = section.tokens[i];
    switch (tk.kind) {
      case Tok::DateY2: {
        unsigned y2 = static_cast<unsigned>(((ymd.y % 100) + 100) % 100);
        append_pad2(out, y2);
        break;
      }
      case Tok::DateY4: {
        char buf[16];
        const int n = std::snprintf(buf, sizeof(buf), "%04d", ymd.y);
        if (n > 0) {
          out.append(buf, static_cast<std::size_t>(n));
        }
        break;
      }
      case Tok::DateM:
        out.append(std::to_string(ymd.m));
        break;
      case Tok::DateMM:
        append_pad2(out, ymd.m);
        break;
      case Tok::DateMMM:
        out.append(month_short(ymd.m));
        break;
      case Tok::DateMMMM:
        out.append(month_long(ymd.m));
        break;
      case Tok::DateD:
        out.append(std::to_string(ymd.d));
        break;
      case Tok::DateDD:
        append_pad2(out, ymd.d);
        break;
      case Tok::DateDDD:
        out.append(weekday_short(sun0));
        break;
      case Tok::DateDDDD:
        out.append(weekday_long(sun0));
        break;
      case Tok::DateH:
        out.append(std::to_string(hour_for_render));
        break;
      case Tok::DateHH:
        append_pad2(out, hour_for_render);
        break;
      case Tok::DateMin:
        out.append(std::to_string(minute));
        break;
      case Tok::DateMMMin:
        append_pad2(out, minute);
        break;
      case Tok::DateS:
        out.append(std::to_string(second));
        break;
      case Tok::DateSS:
        append_pad2(out, second);
        break;
      case Tok::DateElapsedH: {
        // Total hours since serial 0 (integer floor).
        const long long total_hours = static_cast<long long>(std::floor(serial * 24.0));
        out.append(std::to_string(total_hours));
        break;
      }
      case Tok::DateElapsedM: {
        const long long total_minutes = static_cast<long long>(std::floor(serial * 1440.0));
        out.append(std::to_string(total_minutes));
        break;
      }
      case Tok::DateElapsedS: {
        const long long total_sec = static_cast<long long>(std::floor(serial * 86400.0));
        out.append(std::to_string(total_sec));
        break;
      }
      case Tok::AmPm:
        out.append(pm ? "PM" : "AM");
        break;
      case Tok::AP:
        out.append(pm ? "P" : "A");
        break;
      case Tok::FracSecDigits: {
        // Render fractional seconds at the requested precision.
        const int digits = static_cast<int>(tk.width);
        if (digits > 0) {
          out.push_back('.');
          double f = sub_sec_float;
          if (f < 0.0) {
            f = 0.0;
          }
          for (int k = 0; k < digits; ++k) {
            f *= 10.0;
            int d = static_cast<int>(std::floor(f));
            if (d > 9) {
              d = 9;
            } else if (d < 0) {
              d = 0;
            }
            out.push_back(static_cast<char>('0' + d));
            f -= static_cast<double>(d);
          }
        }
        break;
      }
      case Tok::Literal:
        if (tk.lit_end > tk.lit_begin) {
          out.append(fmt.data() + tk.lit_begin, tk.lit_end - tk.lit_begin);
        }
        break;
      default:
        break;
    }
  }
}

// ---------------------------------------------------------------------------
// Section selection and top-level apply.
// ---------------------------------------------------------------------------

// Renders the text section by walking the raw format bytes. We do not
// reuse `section.tokens` here because date letters that snuck into literal
// phrases (e.g. the `s` in "text is @") would have been promoted to
// DateS tokens during tokenisation and lost their positional info. Walking
// the raw format bytes avoids that pitfall while still honouring `"..."`
// quoted literals, `\x` / `!x` escapes, and `[...]` bracketed discards
// (e.g. colour markers).
void render_text_section(const Section& /*section*/, std::string_view fmt, std::string_view original,
                         std::string& out) {
  std::size_t i = 0;
  while (i < fmt.size()) {
    const char c = fmt[i];
    if (c == '"') {
      std::size_t j = i + 1;
      while (j < fmt.size() && fmt[j] != '"') {
        out.push_back(fmt[j]);
        ++j;
      }
      i = j < fmt.size() ? j + 1 : j;
      continue;
    }
    if ((c == '\\' || c == '!') && i + 1 < fmt.size()) {
      out.push_back(fmt[i + 1]);
      i += 2;
      continue;
    }
    if (c == '[') {
      // Skip to matching `]` (colour / locale markers discarded).
      std::size_t j = i + 1;
      while (j < fmt.size() && fmt[j] != ']') {
        ++j;
      }
      i = j < fmt.size() ? j + 1 : j;
      continue;
    }
    if (c == '@') {
      out.append(original);
      ++i;
      continue;
    }
    out.push_back(c);
    ++i;
  }
}

}  // namespace

FormatStatus apply_format(double value, std::string_view format, std::string_view original_text, std::string& out) {
  if (format.empty()) {
    return FormatStatus::kOk;
  }
  const auto sections_raw = split_sections(format);
  if (sections_raw.empty()) {
    return FormatStatus::kOk;
  }
  std::vector<Section> sections;
  sections.reserve(sections_raw.size());
  for (const auto& raw : sections_raw) {
    Section s;
    tokenize_section(raw, s);
    classify(s);
    sections.push_back(std::move(s));
  }

  // Caller is passing `original_text`: if the value is text (non-numeric
  // source) and we have a text section (index 3 for 4-section formats; any
  // `@` token in a single-section format also applies), route there.
  const bool has_original_text = !original_text.empty();
  // Decide the section to use based on Excel's rules:
  //   1 section : apply to everything; text passes unformatted unless `@`
  //               is present.
  //   2 sections: section 0 = positive/zero; section 1 = negative.
  //   3 sections: section 0 = positive; section 1 = negative; section 2 = zero.
  //   4 sections: section 0 = positive; section 1 = negative; section 2 = zero;
  //               section 3 = text.
  int chosen = 0;
  if (has_original_text) {
    if (sections.size() >= 4) {
      chosen = 3;
    } else {
      // Single-section with an `@`: route through the numeric walker but
      // `@` substitutes the text.
      chosen = 0;
    }
  } else if (value > 0.0) {
    chosen = 0;
  } else if (value < 0.0) {
    if (sections.size() >= 2) {
      chosen = 1;
    } else {
      chosen = 0;
    }
  } else {
    // Zero.
    if (sections.size() >= 3) {
      chosen = 2;
    } else {
      chosen = 0;
    }
  }

  const Section& section = sections[static_cast<std::size_t>(chosen)];
  const std::string_view raw_fmt = sections_raw[static_cast<std::size_t>(chosen)];
  if (section.has_invalid_bracket) {
    return FormatStatus::kValueError;
  }

  // For section 1 (negative) Excel emits the value's absolute representation
  // unless the format itself includes an explicit minus sign. The numeric
  // walker currently prefixes the minus from `signbit(scaled)`, so pass the
  // absolute value when we've chosen the dedicated negative section.
  double render_value = value;
  if (chosen == 1 && sections.size() >= 2) {
    render_value = std::fabs(value);
  } else if (chosen == 2 && sections.size() >= 3) {
    render_value = std::fabs(value);
  }

  if (section.is_text ||
      (has_original_text && !section.is_date && section.integer_zero_digits == 0 && section.integer_opt_digits == 0 &&
       section.integer_pad_digits == 0 && section.fraction_zero_digits == 0 && section.fraction_opt_digits == 0 &&
       section.fraction_pad_digits == 0)) {
    render_text_section(section, raw_fmt, original_text, out);
    return FormatStatus::kOk;
  }
  if (section.is_date) {
    render_date(section, raw_fmt, render_value, out);
    return FormatStatus::kOk;
  }
  render_numeric(section, raw_fmt, render_value, out);
  return FormatStatus::kOk;
}

}  // namespace text_format
}  // namespace eval
}  // namespace formulon
