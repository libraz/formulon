//
// Renderer orchestrator for the Excel TEXT() format-string engine.
//
// `apply_format` (in `number_format.cpp`) selects a section and dispatches
// to one of the three renderer translation units:
//   * Standard numeric / `General` -> `render_numeric.cpp`
//   * Date / time tokens          -> `render_date.cpp`
//   * Fraction tokens (`# ?/?`)   -> `render_fraction.cpp`
//
// This translation unit owns:
//   * The shared digit-substitution helpers declared in `render_common.h`
//     (DBNum tables and the `append_*` / `decimal_digits_all_zero`
//     utilities). Hosting the bodies here gives every renderer TU a
//     single linkable definition rather than duplicating tables.
//   * `render_text_section`, the walker for Excel's text-section format.
//     This walks the raw format bytes rather than the token stream so
//     date letters (`s`, `m`, ...) inside literal phrases like
//     "text is @" are not promoted to date tokens by the tokenizer.

#include <cstddef>
#include <string>
#include <string_view>
#include <utility>

#include "eval/eval_profile_scope.h"
#include "eval/text_format/number_format_types.h"
#include "eval/text_format/render_common.h"
#include "excel_locale.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {
namespace {

// Appends one group of up to four digits (not all zero) with place units;
// `*pending_zero` tracks a skipped place for the zero filler.
void append_place_group(std::string& out, const DbnumStyle& style, bool place_one, bool zero_filler,
                        std::string_view group, bool* pending_zero) {
  for (std::size_t i = 0; i < group.size(); ++i) {
    const std::size_t digit = static_cast<std::size_t>(group[i] - '0');
    const std::size_t place = group.size() - 1U - i;
    if (digit == 0U) {
      *pending_zero = true;
      continue;
    }
    if (*pending_zero && zero_filler) {
      out.append(style.digits[0]);
    }
    *pending_zero = false;
    if (place == 0U || digit != 1U || place_one) {
      out.append(style.digits[digit]);
    }
    if (place > 0U) {
      out.append(style.place_units[place - 1U]);
    }
  }
}

// Appends `digits` (no leading zeros, not empty) in groups of four under the
// group units. A part above the largest unit longer than one group is written
// digit by digit.
void append_place_number(std::string& out, const DbnumStyle& style, bool place_one, bool zero_filler,
                         std::string_view digits) {
  constexpr std::size_t kGroupDigits = 4U;
  constexpr std::size_t kTopGroup = 3U;
  bool pending_zero = false;
  const std::size_t lower_digits = kTopGroup * kGroupDigits;
  if (digits.size() > lower_digits) {
    const std::string_view top = digits.substr(0, digits.size() - lower_digits);
    if (top.size() <= kGroupDigits) {
      append_place_group(out, style, place_one, zero_filler, top, &pending_zero);
    } else {
      for (const char c : top) {
        out.append(style.digits[static_cast<std::size_t>(c - '0')]);
      }
    }
    out.append(style.group_units[kTopGroup - 1U]);
    digits.remove_prefix(top.size());
  }
  std::size_t unit = (digits.size() + kGroupDigits - 1U) / kGroupDigits;
  std::size_t begin = 0;
  while (unit > 0U) {
    --unit;
    const std::size_t end = digits.size() - unit * kGroupDigits;
    const std::string_view group = digits.substr(begin, end - begin);
    if (group.find_first_not_of('0') != std::string_view::npos) {
      // Zeros closing the previous group take no filler (一百四十万二千, dbnum_place_one_filler).
      pending_zero = false;
      append_place_group(out, style, place_one, zero_filler, group, &pending_zero);
      if (unit > 0U) {
        out.append(style.group_units[unit - 1U]);
      }
    }
    begin = end;
  }
}

}  // namespace

// --- Shared helpers declared in `render_common.h` ---------------------

std::string_view dbnum_digit_subst(const DbnumStyle* style, char c) noexcept {
  if (style == nullptr || c < '0' || c > '9') {
    return {};
  }
  return style->digits[static_cast<std::size_t>(c - '0')];
}

void append_digit_dbnum(std::string& out, const DbnumStyle* style, char c) {
  const std::string_view sub = dbnum_digit_subst(style, c);
  if (!sub.empty()) {
    out.append(sub);
  } else {
    out.push_back(c);
  }
}

void append_chars_dbnum(std::string& out, const DbnumStyle* style, std::string_view chars) {
  for (char c : chars) {
    append_digit_dbnum(out, style, c);
  }
}

void append_int_dbnum(std::string& out, long long value, const DbnumStyle* style) {
  append_chars_dbnum(out, style, std::to_string(value));
}

void append_dbnum_positional(std::string& out, const DbnumStyle* style, std::string_view digits) {
  if (style == nullptr || style->place_units[0].empty()) {
    append_chars_dbnum(out, style, digits);
    return;
  }
  const std::size_t first = digits.find_first_not_of('0');
  if (first == std::string_view::npos) {
    out.append(style->digits[0]);
    return;
  }
  append_place_number(out, *style, style->place_one, style->zero_filler, digits.substr(first));
}

void append_dbnum_date_field(std::string& out, const DbnumStyle* style, unsigned value, std::size_t width) {
  const std::string digits = std::to_string(value);
  for (std::size_t i = digits.size(); i < width; ++i) {
    append_digit_dbnum(out, style, '0');
  }
  if (style == nullptr || !style->date_positional || value == 0U) {
    append_chars_dbnum(out, style, digits);
    return;
  }
  append_place_number(out, *style, style->date_place_one, /*zero_filler=*/false, digits);
}

bool dbnum_writes_numerals(const DbnumStyle* style) noexcept {
  return style != nullptr && !style->place_units[0].empty();
}

void append_pad2(std::string& out, unsigned n) {
  if (n < 10u) {
    out.push_back('0');
  }
  out.append(std::to_string(n));
}

void append_pad2_dbnum(std::string& out, unsigned value, const DbnumStyle* style) {
  if (value < 10u) {
    append_digit_dbnum(out, style, '0');
  }
  append_int_dbnum(out, static_cast<long long>(value), style);
}

void append_percent(std::string& out, std::string_view fmt, const Token& token) {
  if (token.lit_end > token.lit_begin) {
    out.append(fmt.substr(token.lit_begin, token.lit_end - token.lit_begin));
  } else {
    out.push_back('%');
  }
}

bool decimal_digits_all_zero(std::string_view digits) noexcept {
  for (char ch : digits) {
    if (ch != '0') {
      return false;
    }
  }
  return true;
}

void cap_integer_significant_digits(std::string* digits) {
  constexpr std::size_t kSignificantDigits = 15u;
  if (digits->size() <= kSignificantDigits) {
    return;
  }
  const std::size_t original_length = digits->size();
  std::string prefix = digits->substr(0, kSignificantDigits);
  if ((*digits)[kSignificantDigits] >= '5') {
    // Half-away-from-zero round-up, with carry propagation; an all-9s
    // prefix carries out and grows by one digit (e.g. "999" -> "1000").
    std::size_t i = prefix.size();
    while (i > 0) {
      --i;
      if (prefix[i] != '9') {
        ++prefix[i];
        break;
      }
      prefix[i] = '0';
      if (i == 0) {
        prefix.insert(prefix.begin(), '1');
      }
    }
  }
  prefix.append(original_length - prefix.size(), '0');
  *digits = std::move(prefix);
}

// --- Text-section walker ----------------------------------------------

// Renders the text section by walking the raw format bytes. We do not
// reuse `section.tokens` here because date letters that snuck into literal
// phrases (e.g. the `s` in "text is @") would have been promoted to
// DateS tokens during tokenisation and lost their positional info. Walking
// the raw format bytes avoids that pitfall while still honouring `"..."`
// quoted literals, `\x` (and the ja-JP `!x`) escapes, and `[...]` bracketed discards
// (e.g. colour markers).
void render_text_section(const Section& /*section*/, std::string_view fmt, std::string_view original,
                         std::string& out) {
  const bool bang_escape = locale_facts(eval::current_eval_profile()).bang_escape;
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
    if ((c == '\\' || (c == '!' && bang_escape)) && i + 1 < fmt.size()) {
      out.push_back(fmt[i + 1]);
      i += 2;
      continue;
    }
    if (c == '[') {
      // Skip to matching `]`; a `[$JPY-411]` marker writes its symbol (locale_tokens.lcid_symbol_text_section).
      std::size_t j = i + 1;
      while (j < fmt.size() && fmt[j] != ']') {
        ++j;
      }
      if (j > i + 1 && fmt[i + 1] == '$') {
        const std::string_view body = fmt.substr(i + 2, j - i - 2);
        out.append(body.substr(0, body.find('-')));
      }
      i = j < fmt.size() ? j + 1 : j;
      continue;
    }
    if (c == '@') {
      out.append(original);
      ++i;
      continue;
    }
    if (c == '_' && i + 1 < fmt.size()) {
      // `_X` underscore-skip: emit a single space and consume both bytes.
      out.push_back(' ');
      i += 2;
      continue;
    }
    out.push_back(c);
    ++i;
  }
}

}  // namespace number_format_detail
}  // namespace text_format
}  // namespace formulon
