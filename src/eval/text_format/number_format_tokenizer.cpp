// Copyright 2026 libraz. Licensed under the MIT License.
//
// Tokenizer and classifier for the Excel TEXT() format-string engine.
// The types consumed and produced here are declared in
// `number_format_types.h`; the renderer lives in
// `number_format_render.cpp`; the public entry point `apply_format` lives
// in `number_format.cpp`.

#include <cstddef>
#include <cstdint>
#include <string_view>
#include <vector>

#include "eval/text_format/number_format_types.h"

namespace formulon {
namespace eval {
namespace number_format_detail {
namespace {

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

}  // namespace

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

}  // namespace number_format_detail
}  // namespace eval
}  // namespace formulon
