//
// Section splitter and classifier for the Excel TEXT() format-string
// engine. The per-token scanner is in `number_format_tokenizer.cpp`; the
// rendering pipeline is in `number_format_render.cpp` and friends. The
// public entry point `apply_format` lives in `number_format.cpp`.
//
// This translation unit owns:
//   * `split_sections` -- chops the user format on unquoted `;`.
//   * `classify`       -- second-pass token analysis (numeric/date/text
//                         classification, fraction detection, etc.).
//   * `disambiguate_minutes` (anonymous-namespace helper for `classify`)
//                       -- resolves `m`/`mm` between month and minute based
//                          on adjacent date tokens.

#include "eval/text_format/number_format_section.h"

#include <cstddef>
#include <cstdint>
#include <string_view>
#include <vector>

#include "eval/text_format/number_format_scanner.h"
#include "eval/text_format/number_format_types.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {
namespace {

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
      i += i + 1 < fmt.size() ? 1 + utf8_scalar_width(fmt, i + 1) : 1;
      // Skip escape + one complete UTF-8 scalar payload.
      continue;
    }
    if (c == '_' && i + 1 < fmt.size()) {
      i += 1 + utf8_scalar_width(fmt, i + 1);  // Skip the scalar payload.
      continue;
    }
    if (c == '*' && i + 1 < fmt.size()) {
      i += 1 + utf8_scalar_width(fmt, i + 1);  // Skip the scalar payload.
      continue;
    }
    if (c == '[') {
      // Skip to matching `]` so e.g. `[赤]` does not interact with `;`.
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

void classify(Section& section, std::string_view fmt) noexcept {
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

  bool has_general = false;
  for (const Token& tk : section.tokens) {
    if (tk.kind == Tok::GeneralNumber) {
      has_general = true;
      break;
    }
  }
  section.has_general = has_general;

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

  // Excel rejects formats that mix date tokens with number-digit tokens in
  // the same section (e.g. `mm###`). Mac Excel 365 and IronCalc both
  // surface `#VALUE!` here. Reuse the `has_invalid_bracket` channel so
  // `apply_format` already converts the section to the value-error path.
  if (section.is_date && (int_zero + int_opt + int_pad + frac_zero + frac_opt + frac_pad) > 0) {
    section.has_invalid_bracket = true;
  }

  // --- Fraction format detection (`# ?/?`, `0/0`, etc.) ----------------
  //
  // A fraction format has these features:
  //   * Exactly one literal '/' byte token in the section.
  //   * Immediately before the slash: a contiguous run of digit-placeholder
  //     tokens (`#`/`?`/`0`) -- the numerator group.
  //   * Immediately after the slash: a contiguous run of digit-placeholder
  //     tokens -- the denominator group.
  //   * Optionally before the numerator: an integer-placeholder run and
  //     exactly one literal-space byte that visually separates the integer
  //     from the numerator.
  //   * No `.` (Point) token (fractions and decimal points are mutually
  //     exclusive).
  //
  // When detected, we set `is_fraction` and the digit-group bounds; the
  // renderer's `render_numeric` branches into `render_fraction` based on
  // the flag. Date sections never qualify (the slash inside a date is a
  // date separator).
  auto is_digit_tok2 = [](Tok k) { return k == Tok::DigitZero || k == Tok::DigitOpt || k == Tok::DigitPad; };
  auto is_single_byte_literal = [&fmt](const Token& tk, char want) {
    if (tk.kind != Tok::Literal) {
      return false;
    }
    if (tk.lit_end != tk.lit_begin + 1) {
      return false;
    }
    if (tk.lit_begin >= fmt.size()) {
      return false;
    }
    return fmt[tk.lit_begin] == want;
  };
  if (!any_date && point_index < 0) {
    // Find slash candidate.
    int slash_idx = -1;
    int slash_count = 0;
    for (std::size_t i = 0; i < section.tokens.size(); ++i) {
      if (is_single_byte_literal(section.tokens[i], '/')) {
        slash_idx = static_cast<int>(i);
        ++slash_count;
      }
    }
    if (slash_count == 1 && slash_idx > 0 && static_cast<std::size_t>(slash_idx) + 1 < section.tokens.size()) {
      // Walk left from slash to find numerator group.
      int num_end = slash_idx - 1;
      int num_begin = num_end;
      while (num_begin > 0 && is_digit_tok2(section.tokens[static_cast<std::size_t>(num_begin) - 1].kind)) {
        --num_begin;
      }
      // Walk right from slash to find denominator group.
      int den_begin = slash_idx + 1;
      int den_end = den_begin;
      while (static_cast<std::size_t>(den_end) + 1 <= section.tokens.size() &&
             is_digit_tok2(section.tokens[static_cast<std::size_t>(den_end)].kind)) {
        ++den_end;
      }
      // Numerator and denominator both need to be at least one digit.
      const bool has_num =
          num_end >= num_begin && is_digit_tok2(section.tokens[static_cast<std::size_t>(num_begin)].kind);
      const bool has_den = den_end > den_begin;
      if (has_num && has_den) {
        // Optional integer group: a literal-space immediately precedes the
        // numerator group, and a digit-placeholder run precedes the space.
        int int_begin = num_begin;
        int int_end = num_begin;
        if (num_begin >= 2 && is_single_byte_literal(section.tokens[static_cast<std::size_t>(num_begin) - 1], ' ') &&
            is_digit_tok2(section.tokens[static_cast<std::size_t>(num_begin) - 2].kind)) {
          int_end = num_begin - 1;  // exclusive: stops before the space.
          int_begin = int_end;
          while (int_begin > 0 && is_digit_tok2(section.tokens[static_cast<std::size_t>(int_begin) - 1].kind)) {
            --int_begin;
          }
        }
        section.is_fraction = true;
        section.fraction_int_begin = int_begin;
        section.fraction_int_end = int_end;
        section.fraction_num_begin = num_begin;
        section.fraction_num_end = num_end + 1;  // exclusive
        section.fraction_den_begin = den_begin;
        section.fraction_den_end = den_end;
        section.fraction_slash_index = slash_idx;
        section.fraction_int_max_digits = int_end - int_begin;
        section.fraction_num_max_digits = (num_end + 1) - num_begin;
        section.fraction_den_max_digits = den_end - den_begin;
      }
    }
  }
}

}  // namespace number_format_detail
}  // namespace text_format
}  // namespace formulon
