//
// Per-token scanner for the Excel TEXT() format-string engine. Walks the
// input format byte-by-byte and produces the `Section::tokens` stream
// consumed by `classify()` (in `number_format_section.cpp`) and the
// renderers (`number_format_render.cpp`, `render_numeric.cpp`,
// `render_date.cpp`, `render_fraction.cpp`). The public entry point
// `apply_format` lives in `number_format.cpp`.
//
// Stateless scanning helpers (color/condition/DBNum specifiers, run scan,
// date-letter detection) are factored out into
// `number_format_scanner.{h,cpp}` so this TU only carries the main
// dispatch loop.

#include <cstddef>
#include <cstdint>
#include <string_view>
#include <vector>

#include "eval/eval_profile_scope.h"
#include "eval/text_format/number_format_scanner.h"
#include "eval/text_format/number_format_types.h"
#include "excel_locale.h"
#include "utils/strings.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {
namespace {

constexpr std::string_view kChineseAmPm = "上午/下午";

enum class DateLetter : std::uint8_t {
  kNone,
  kYear,
  kMonth,
  kMinute,
  kMonthOrMinute,
  kDay,
  kHour,
  kSecond,
};

DateLetter classify_date_letter(char c, const FormatLetters& letters) noexcept {
  const char lc = strings::ascii_to_lower(c);
  if (letters.case_sensitive) {
    if (c == letters.month) {
      return letters.month_contextual ? DateLetter::kMonthOrMinute : DateLetter::kMonth;
    }
    if (c == letters.minute) {
      return letters.minute_unconditional ? DateLetter::kMinute : DateLetter::kMonthOrMinute;
    }
  } else if (lc == strings::ascii_to_lower(letters.month)) {
    return DateLetter::kMonthOrMinute;
  }
  if (lc == strings::ascii_to_lower(letters.year)) {
    return DateLetter::kYear;
  }
  if (lc == strings::ascii_to_lower(letters.day)) {
    return DateLetter::kDay;
  }
  if (lc == strings::ascii_to_lower(letters.hour)) {
    return DateLetter::kHour;
  }
  if (lc == strings::ascii_to_lower(letters.second)) {
    return DateLetter::kSecond;
  }
  return DateLetter::kNone;
}

bool same_letters(const FormatLetters& a, const FormatLetters& b) noexcept {
  return a.year == b.year && a.month == b.month && a.day == b.day && a.hour == b.hour && a.minute == b.minute &&
         a.second == b.second && a.case_sensitive == b.case_sensitive &&
         a.minute_unconditional == b.minute_unconditional && a.month_contextual == b.month_contextual &&
         a.weekday == b.weekday;
}

}  // namespace

void tokenize_section(std::string_view fmt, Section& out, const FormatLetters& letters, const FormatTag* format_tag) {
  std::vector<Token>& toks = out.tokens;
  const LocaleFacts& facts = locale_facts(eval::current_eval_profile());
  const bool invariant_letters = same_letters(letters, kInvariantFormatLetters);
  auto push_literal = [&](std::size_t b, std::size_t e, bool protected_literal = false) {
    if (b == e) {
      return;
    }
    Token t;
    t.kind = Tok::Literal;
    t.lit_begin = b;
    t.lit_end = e;
    t.protected_literal = protected_literal;
    toks.push_back(t);
  };

  bool saw_color = false;
  bool saw_tag = false;
  FormatTag own_tag;
  // An unquoted digit other than 0 makes a text section invalid (locale_tokens.text_section_bare_digit).
  bool bare_digit = false;
  // Index of the first token after an unquoted colon; a digit placeholder there is #VALUE!.
  std::size_t after_colon = SIZE_MAX;
  // 0, a `[DBNumN]` index, or th-TH's `t`; resolved once the section's tag is known.
  constexpr int kThaiDigitChoice = 5;
  constexpr std::uint8_t kThaiNumeralSystem = 0x0D;
  int digit_choice = 0;
  std::size_t i = 0;
  while (i < fmt.size()) {
    const char c = fmt[i];
    // Quoted literal `"..."`.
    if (c == '"') {
      std::size_t j = i + 1;
      while (j < fmt.size() && fmt[j] != '"') {
        ++j;
      }
      if (j == fmt.size()) {
        out.has_invalid_bracket = true;
        return;
      }
      push_literal(i + 1, j, true);
      i = j + 1;
      continue;
    }
    // Escape, spacing, and fill operators each require a following scalar.
    if ((c == '\\' || (c == '!' && facts.bang_escape) || c == '_' || c == '*') && i + 1 == fmt.size()) {
      out.has_invalid_bracket = true;
      return;
    }
    // Escape `\x` or `!x` -> next UTF-8 scalar is a literal.
    if ((c == '\\' || (c == '!' && facts.bang_escape)) && i + 1 < fmt.size()) {
      const std::size_t payload_width = utf8_scalar_width(fmt, i + 1);
      push_literal(i + 1, i + 1 + payload_width, true);
      i += 1 + payload_width;
      continue;
    }
    // Bracketed specifier. Recognised kinds:
    //   `[h]` / `[m]` / `[s]`   -> elapsed-time tokens (any run length).
    //   `[$...]`                -> currency symbol and locale tag.
    //   `[Red]` / ...           -> named color qualifier; silently dropped
    //                               (no color in text).
    //   `[ColorN]`              -> indexed color qualifier; silently dropped.
    //   `[DBNum1]` ...          -> digit style on ja-JP, inert on en-US.
    // Anything else (`[>100]`, unknown qualifiers) still trips the
    // invalid-bracket flag and surfaces as #VALUE!.
    if (c == '[') {
      const std::size_t body_begin = i + 1;
      std::size_t j = i + 1;
      while (j < fmt.size() && fmt[j] != ']') {
        ++j;
      }
      if (j == fmt.size()) {
        out.has_invalid_bracket = true;
        return;
      }
      const std::string_view body = fmt.substr(body_begin, j - body_begin);
      i = j + 1;
      // Elapsed time markers: any run of `h`, `m`, or `s` (case-insensitive).
      bool all_h = !body.empty();
      bool all_m = !body.empty();
      bool all_s = !body.empty();
      for (char ch : body) {
        const char lo = strings::ascii_to_lower(ch);
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
        t.width = body.size();
        toks.push_back(t);
      } else if (all_m) {
        Token t;
        t.kind = Tok::DateElapsedM;
        t.width = body.size();
        toks.push_back(t);
      } else if (!body.empty() && body.front() == '$') {
        // Locale-currency marker. Form: `[$<symbol>-<lcid>]` or `[$<symbol>]`
        // or `[$-<lcid>]`. The `<symbol>` portion (everything after `$` up
        // to the optional `-LCID` suffix) is emitted as a literal prefix;
        // the spec after `-` switches names, calendar and digits for the
        // section (`FormatTag`). Mac Excel 365 emits e.g. `[$JPY-411]#,##0`
        // -> `JPY1,234,567`.
        //
        // `body_begin` points at the `$` in `fmt`; the symbol bytes start
        // at `body_begin + 1` and run until either `-` or the end of body.
        const std::size_t sym_begin = body_begin + 1;
        std::size_t sym_end = sym_begin + (body.size() - 1);
        for (std::size_t k = 1; k < body.size(); ++k) {
          if (body[k] == '-') {
            sym_end = body_begin + k;
            break;
          }
        }
        push_literal(sym_begin, sym_end);
        // One tag per section (locale_tokens.lcid_parse_variants).
        if (saw_tag) {
          out.has_invalid_bracket = true;
        }
        saw_tag = true;
        if (sym_end < body_begin + body.size()) {
          own_tag.present = true;
          if (!parse_format_tag(fmt.substr(sym_end + 1, body_begin + body.size() - sym_end - 1), &own_tag)) {
            out.has_invalid_bracket = true;
          }
        }
      } else if (all_s) {
        Token t;
        t.kind = Tok::DateElapsedS;
        t.width = body.size();
        toks.push_back(t);
      } else if (is_color_specifier(body)) {
        // Named colour (`[Red]`) or indexed colour (`[Color12]`). Excel discards
        // the colour inside TEXT, so the rest of the section still formats
        // the value. A section may carry at most one colour, though: Excel
        // rejects `[Red][Blue]0.00` with #VALUE! even though either bracket
        // alone is inert.
        if (saw_color) {
          out.has_invalid_bracket = true;
        }
        saw_color = true;
        out.has_color = true;
      } else if (const int dbnum = parse_dbnum_directive(body); dbnum > 0) {
        // `[DBNum1]`-`[DBNum4]`: digit style for the section, last write wins.
        digit_choice = dbnum;
      } else {
        // Conditional-section directive `[>1000]`, `[<=0]`, ...
        // Only one predicate per section; if a second one appears,
        // last-write-wins (matches Mac Excel: the rightmost directive
        // shadows earlier ones in the same section).
        CondOp cop = CondOp::kNone;
        double cval = 0.0;
        const int rc = parse_cond_directive(body, &cop, &cval);
        if (rc > 0) {
          out.cond_op = cop;
          out.cond_value = cval;
        } else {
          // Either the body doesn't look like a predicate (rc == 0) or it
          // does but the numeric tail failed to parse (rc < 0). Both surface
          // as #VALUE! through the existing invalid-bracket channel.
          out.has_invalid_bracket = true;
        }
      }
      continue;
    }
    // Underscore-skip `_X`: Excel reserves the width of character `X` and
    // emits a matching amount of whitespace. TEXT's output uses a single
    // space regardless of `X`.
    if (c == '_' && i + 1 < fmt.size()) {
      Token t;
      t.kind = Tok::Space;
      toks.push_back(t);
      i += 1 + utf8_scalar_width(fmt, i + 1);
      continue;
    }
    // Asterisk-fill `*X`: in cell formats this pads the cell with `X` to
    // fill the column width. TEXT() has no column width, so Mac Excel 365
    // emits this as a no-op (both `*` and the fill char are skipped).
    if (c == '*' && i + 1 < fmt.size()) {
      i += 1 + utf8_scalar_width(fmt, i + 1);
      continue;
    }
    // `General` keyword (case-insensitive). Must be a standalone "word":
    // we only accept it when the following byte (if any) is not an ASCII
    // letter, so `Generally` and similar words pass through as literals.
    if (c == 'G' || c == 'g') {
      auto match_general = [&](std::size_t start) -> bool {
        static const char kWord[] = "general";
        if (start + 7 > fmt.size()) {
          return false;
        }
        if (!strings::case_insensitive_starts_with(fmt.substr(start), std::string_view(kWord, 7))) {
          return false;
        }
        // Boundary check: next byte must not be a letter.
        if (start + 7 < fmt.size()) {
          const char nx = fmt[start + 7];
          if ((nx >= 'A' && nx <= 'Z') || (nx >= 'a' && nx <= 'z')) {
            return false;
          }
        }
        return true;
      };
      if (match_general(i)) {
        Token t;
        t.kind = Tok::GeneralNumber;
        toks.push_back(t);
        i += 7;
        continue;
      }
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
        return strings::case_insensitive_starts_with(fmt.substr(start), std::string_view(a, n));
      };
      if (match_ci(i, "AM/PM")) {
        Token t;
        t.kind = Tok::AmPm;
        toks.push_back(t);
        i += 5;
        continue;
      }
      if ((invariant_letters || facts.short_meridiem) && match_ci(i, "A/P")) {
        Token t;
        t.kind = Tok::AP;
        toks.push_back(t);
        i += 3;
        continue;
      }
    }
    // The Chinese designator pair is an AM/PM marker in every locale.
    if (fmt.compare(i, kChineseAmPm.size(), kChineseAmPm) == 0) {
      Token t;
      t.kind = Tok::AmPmChinese;
      toks.push_back(t);
      i += kChineseAmPm.size();
      continue;
    }
    // Weekday tokens `aaa` / `aaaa` (case-insensitive) on the locale's
    // weekday letter. `aaa` does NOT collide with `AM/PM` or `A/P` because
    // those are matched first above. A run shorter than 3 is not a weekday
    // token in Excel and is emitted as a literal.
    if (letters.weekday != '\0' && strings::ascii_to_lower(c) == letters.weekday) {
      const std::size_t run = scan_run(fmt, i, letters.weekday);
      if (run >= 3) {
        Token t;
        t.kind = (run >= 4) ? Tok::DateAaaa : Tok::DateAaa;
        t.width = static_cast<std::uint8_t>(run);
        toks.push_back(t);
        continue;
      }
      // Run of 1 or 2 weekday letters: emit as literal.
      push_literal(i - run, i);
      continue;
    }
    // ja-JP era name tokens: `g` (Roman 1-letter), `gg` (1-char kanji),
    // `ggg` or longer (full kanji name). Case-insensitive. The `General`
    // keyword check above already handled the literal "General" word. A
    // locale whose own date letter is `g` (it-IT's day) reads it as that.
    if ((c == 'g' || c == 'G') && classify_date_letter(c, letters) == DateLetter::kNone) {
      const std::size_t run = scan_run(fmt, i, 'g');
      Token t;
      t.width = static_cast<std::uint8_t>(run);
      if (run >= 3) {
        t.kind = Tok::EraGGG;
      } else if (run == 2) {
        t.kind = Tok::EraGG;
      } else {
        t.kind = Tok::EraG;
      }
      toks.push_back(t);
      continue;
    }
    // Buddhist-era year `bb` / `bbbb`, the same letter in every locale.
    if (c == 'b') {
      const std::size_t run = scan_run_exact(fmt, i, 'b');
      Token t;
      t.width = run;
      t.kind = run <= 2 ? Tok::DateB2 : Tok::DateB4;
      toks.push_back(t);
      continue;
    }
    // Date letters, read through the locale's letter table.
    if (const DateLetter letter = classify_date_letter(c, letters); letter != DateLetter::kNone) {
      const bool exact = letters.case_sensitive && (letter == DateLetter::kMonth || letter == DateLetter::kMinute ||
                                                    letter == DateLetter::kMonthOrMinute);
      const std::size_t run = exact ? scan_run_exact(fmt, i, c) : scan_run(fmt, i, strings::ascii_to_lower(c));
      Token t;
      t.width = static_cast<std::uint8_t>(run);
      switch (letter) {
        case DateLetter::kYear:
          t.kind = (run <= 2) ? Tok::DateY2 : Tok::DateY4;
          break;
        case DateLetter::kMonthOrMinute:
        case DateLetter::kMonth:
          // A run of 5 or more means "first letter of the month name"
          // (Excel's `mmmmm` convention). A contextual `m`/`mm` is resolved
          // between month and minute in pass 2.
          if (run >= 5) {
            t.kind = Tok::DateMMMMM;
          } else if (run == 4) {
            t.kind = Tok::DateMMMM;
          } else if (run == 3) {
            t.kind = Tok::DateMMM;
          } else if (letter == DateLetter::kMonth) {
            t.kind = run == 2 ? Tok::DateMM : Tok::DateM;
          } else {
            t.kind = Tok::DateMOrMin;
          }
          break;
        case DateLetter::kMinute:
          t.kind = (run >= 2) ? Tok::DateMMMin : Tok::DateMin;
          break;
        case DateLetter::kDay:
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
        case DateLetter::kHour:
          t.kind = (run >= 2) ? Tok::DateHH : Tok::DateH;
          break;
        case DateLetter::kSecond:
          t.kind = (run >= 2) ? Tok::DateSS : Tok::DateS;
          break;
        case DateLetter::kNone:
          break;
      }
      toks.push_back(t);
      continue;
    }
    // An invariant date letter the locale spells differently prints as text.
    if (is_date_letter(c)) {
      const std::size_t run = scan_run(fmt, i, strings::ascii_to_lower(c));
      push_literal(i - run, i);
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
    if (facts.blank_date_letter != '\0' && strings::ascii_to_lower(c) == facts.blank_date_letter) {
      scan_run(fmt, i, facts.blank_date_letter);
      Token t;
      t.kind = Tok::DateBlank;
      toks.push_back(t);
      continue;
    }
    if (c == 't' && facts.thai_digit_letter) {
      digit_choice = kThaiDigitChoice;
      ++i;
      continue;
    }
    // `r` spells the era (`ee`, `gggee`) in ja-JP and the year in zh-TW.
    if (c == 'r' && facts.r_letter != RLetter::kLiteral) {
      const std::size_t run = scan_run_exact(fmt, i, 'r');
      Token t;
      // zh-TW's `r` and `rr` write the year as `e` does (locale_tokens.lcid_r_letter).
      if (facts.r_letter == RLetter::kYear) {
        t.kind = Tok::EraE;
        t.width = 1;
        toks.push_back(t);
        continue;
      }
      if (run >= 2) {
        t.kind = Tok::EraGGG;
        t.width = 3;
        toks.push_back(t);
      }
      t.kind = Tok::EraEE;
      t.width = 2;
      toks.push_back(t);
      continue;
    }
    // ja-JP era year token `e` / `ee`. A bare `e` (or run of `e`) not
    // followed by `+`/`-` is the era-year placeholder when the section is
    // a date section. The renderer falls back to a literal `e` when the
    // section turns out to be numeric.
    if (c == 'e' || c == 'E') {
      const std::size_t run = scan_run(fmt, i, 'e');
      Token t;
      t.width = static_cast<std::uint8_t>(run);
      t.kind = (run >= 2) ? Tok::EraEE : Tok::EraE;
      toks.push_back(t);
      continue;
    }
    // A kept full-width percent sign scales and writes itself (locale_tokens.text_fullwidth_percent_syntax).
    constexpr std::string_view kFullwidthPercent = "\xEF\xBC\x85";
    if (fmt.substr(i, kFullwidthPercent.size()) == kFullwidthPercent) {
      Token t;
      t.kind = Tok::Percent;
      t.lit_begin = i;
      t.lit_end = i + kFullwidthPercent.size();
      toks.push_back(t);
      i = t.lit_end;
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
    // Keep the slash's source range so the classifier can tell it from a quoted slash; a kept
    // full-width slash parses as `/` and writes itself (locale_tokens.text_fullwidth_date_syntax).
    constexpr std::string_view kFullwidthSlash = "\xEF\xBC\x8F";
    constexpr std::string_view kFullwidthColon = "\xEF\xBC\x9A";
    if (c == '/' || fmt.substr(i, kFullwidthSlash.size()) == kFullwidthSlash) {
      Token t;
      t.kind = Tok::Literal;
      t.lit_begin = i;
      t.lit_end = i + (c == '/' ? 1U : kFullwidthSlash.size());
      t.fraction_slash_candidate = true;
      toks.push_back(t);
      i = t.lit_end;
      continue;
    }
    // An unescaped `n` is no format code and no text (locale_tokens.text_letter_n_time, text_letter_n_date).
    if (c == 'n') {
      out.has_invalid_bracket = true;
    }
    if (c >= '1' && c <= '9') {
      bare_digit = true;
    }
    if ((c == ':' || fmt.substr(i, kFullwidthColon.size()) == kFullwidthColon) && after_colon == SIZE_MAX) {
      after_colon = toks.size() + 1U;
    }
    // Fallback: preserve one complete UTF-8 scalar as a literal. Malformed
    // input is consumed one byte at a time by `utf8_scalar_width`, so it is
    // still copied without crashing or swallowing following syntax.
    const std::size_t literal_width = utf8_scalar_width(fmt, i);
    push_literal(i, i + literal_width);
    i += literal_width;
  }
  // A colon before a digit placeholder outside a date section is #VALUE! (locale_tokens.text_colon_before_digit).
  bool colon_before_digit = false;
  bool date_section = false;
  for (std::size_t k = 0; k < toks.size(); ++k) {
    const Tok kind = toks[k].kind;
    date_section = date_section || is_date_tok(kind);
    colon_before_digit =
        colon_before_digit ||
        (k >= after_colon && (kind == Tok::DigitZero || kind == Tok::DigitOpt || kind == Tok::DigitPad));
  }
  if (colon_before_digit && !date_section) {
    out.has_invalid_bracket = true;
  }
  if (bare_digit) {
    for (const Token& tk : toks) {
      if (tk.kind == Tok::At) {
        out.has_invalid_bracket = true;
        break;
      }
    }
  }
  // The first section's tag governs every section; a later section's own tag
  // only contributes a system date or time (locale_tokens.lcid_section_scope).
  out.tag = format_tag != nullptr ? *format_tag : own_tag;
  out.tag.system = own_tag.system;
  const bool tag_in_section = format_tag == nullptr && saw_tag;
  // A numeral-system tag decides the digits and silences `[DBNumN]`
  // (locale_tokens.lcid_numeral_systems); a tag language brings its own DBNum styles.
  if (out.tag.numeral != 0) {
    out.digit_style = numeral_system_style(out.tag.numeral);
    out.digit_style_from_tag = out.digit_style != nullptr;
  } else if (digit_choice == kThaiDigitChoice) {
    out.digit_style = numeral_system_style(kThaiNumeralSystem);
  } else if (digit_choice > 0) {
    const DbnumStyles* styles = facts.dbnum;
    // `[DBNum4]` keeps the profile's style in the tagged section (locale_tokens.lcid_dbnum4_sections).
    constexpr int kDbnum4 = 4;
    const bool profile_dbnum4 = digit_choice == kDbnum4 && (tag_in_section || facts.dbnum4_keeps_profile_style);
    if (out.tag.language_set && !profile_dbnum4) {
      const TagLanguage* language = out.tag.language;
      styles = language == nullptr ? nullptr : language->dbnum;
    }
    if (styles != nullptr) {
      out.digit_style = &(*styles)[static_cast<std::size_t>(digit_choice - 1)];
    }
  }
}

}  // namespace number_format_detail
}  // namespace text_format
}  // namespace formulon
