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

#include "eval/text_format/number_format_types.h"
#include "eval/text_format/render_common.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {
namespace {

// --- DBNum digit tables -----------------------------------------------
//
// `[DBNum1]` / `[DBNum2]` / `[DBNum3]` are *per-digit* substitutions in
// Mac Excel 365 ja-JP. Despite their popular description as "positional
// kanji" formats, the oracle corpus shows that Excel does NOT decompose
// integers into 千 / 百 / 十 groups -- it simply rewrites each ASCII
// digit through a fixed table. Concretely:
//
//   * `=TEXT(1234, "[DBNum1]0")` -> `一二三四` (NOT `一千二百三十四`).
//   * `=TEXT(1234, "[DBNum2]0")` -> `壱弐参四` (NOT `壱阡弐百参拾四`).
//   * `=TEXT(1234, "[DBNum3]0")` -> `１２３４` (full-width Arabic).
//
// Tables:
//   * DBNum1: 〇一二三四五六七八九 (weak-form everyday kanji digits).
//   * DBNum2: 零壱弐参四伍六七捌玖. Digit 4 is the everyday form 四 (not
//     大字 肆) because Mac Excel's `1234` golden ends in 四. The 5/6/7/8/9
//     entries are not exercised by the oracle corpus; we follow common
//     convention (5=伍, 8=捌, 9=玖 are 大字; 6=六, 7=七 left as everyday
//     since 4=四 sets that precedent).
//   * DBNum3: U+FF10..U+FF19 (full-width Arabic digits).

const char* dbnum1_digit(char ascii_digit) noexcept {
  // 〇一二三四五六七八九 (each is a 3-byte UTF-8 code point).
  static const char* kDigits[10] = {
      "\xE3\x80\x87",  // 〇
      "\xE4\xB8\x80",  // 一
      "\xE4\xBA\x8C",  // 二
      "\xE4\xB8\x89",  // 三
      "\xE5\x9B\x9B",  // 四
      "\xE4\xBA\x94",  // 五
      "\xE5\x85\xAD",  // 六
      "\xE4\xB8\x83",  // 七
      "\xE5\x85\xAB",  // 八
      "\xE4\xB9\x9D",  // 九
  };
  if (ascii_digit < '0' || ascii_digit > '9') {
    return "";
  }
  return kDigits[static_cast<std::size_t>(ascii_digit - '0')];
}

const char* dbnum2_digit(char ascii_digit) noexcept {
  // Mac Excel-observed 大字 mapping: ones digits = 零壱弐参四伍六七捌玖.
  // (Note 4=四 not 肆, 7=七 not 漆, per Mac Excel's empirical output.)
  static const char* kDigits[10] = {
      "\xE9\x9B\xB6",  // 零
      "\xE5\xA3\xB1",  // 壱
      "\xE5\xBC\x90",  // 弐
      "\xE5\x8F\x82",  // 参
      "\xE5\x9B\x9B",  // 四 (everyday, matches Mac Excel oracle)
      "\xE4\xBC\x8D",  // 伍
      "\xE5\x85\xAD",  // 六 (everyday, matches Mac Excel oracle)
      "\xE4\xB8\x83",  // 七 (everyday, matches Mac Excel oracle)
      "\xE6\x8D\x8C",  // 捌
      "\xE7\x8E\x96",  // 玖
  };
  if (ascii_digit < '0' || ascii_digit > '9') {
    return "";
  }
  return kDigits[static_cast<std::size_t>(ascii_digit - '0')];
}

// Full-width Arabic digits U+FF10..U+FF19 (each is a 3-byte UTF-8 sequence
// `EF BC 9X` for X in 0..9).
const char* dbnum3_digit(char ascii_digit) noexcept {
  static const char* kDigits[10] = {
      "\xEF\xBC\x90", "\xEF\xBC\x91", "\xEF\xBC\x92", "\xEF\xBC\x93", "\xEF\xBC\x94",
      "\xEF\xBC\x95", "\xEF\xBC\x96", "\xEF\xBC\x97", "\xEF\xBC\x98", "\xEF\xBC\x99",
  };
  if (ascii_digit < '0' || ascii_digit > '9') {
    return "";
  }
  return kDigits[static_cast<std::size_t>(ascii_digit - '0')];
}

}  // namespace

// --- Shared helpers declared in `render_common.h` ---------------------

std::string_view dbnum_digit_subst(DbNumMode mode, char c) noexcept {
  if (c < '0' || c > '9') {
    return {};
  }
  switch (mode) {
    case DbNumMode::kDBNum1:
      return dbnum1_digit(c);
    case DbNumMode::kDBNum2:
      return dbnum2_digit(c);
    case DbNumMode::kDBNum3:
      return dbnum3_digit(c);
    case DbNumMode::kNone:
    default:
      return {};
  }
}

void append_digit_dbnum(std::string& out, DbNumMode mode, char c) {
  const std::string_view sub = dbnum_digit_subst(mode, c);
  if (!sub.empty()) {
    out.append(sub);
  } else {
    out.push_back(c);
  }
}

void append_chars_dbnum(std::string& out, DbNumMode mode, std::string_view chars) {
  for (char c : chars) {
    append_digit_dbnum(out, mode, c);
  }
}

void append_int_dbnum(std::string& out, long long value, DbNumMode mode) {
  std::string buf = std::to_string(value);
  if (mode == DbNumMode::kNone) {
    out.append(buf);
    return;
  }
  append_chars_dbnum(out, mode, buf);
}

void append_pad2(std::string& out, unsigned n) {
  if (n < 10u) {
    out.push_back('0');
  }
  out.append(std::to_string(n));
}

void append_pad2_dbnum(std::string& out, unsigned value, DbNumMode mode) {
  if (mode == DbNumMode::kNone) {
    append_pad2(out, value);
    return;
  }
  if (value < 10u) {
    append_digit_dbnum(out, mode, '0');
  }
  append_int_dbnum(out, static_cast<long long>(value), mode);
}

bool decimal_digits_all_zero(std::string_view digits) noexcept {
  for (char ch : digits) {
    if (ch != '0') {
      return false;
    }
  }
  return true;
}

// --- Text-section walker ----------------------------------------------

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
