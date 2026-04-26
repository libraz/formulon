// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the shared date / time text parser declared in
// `date_text_parse.h`. The logic matches the grammar previously embedded in
// `src/eval/builtins/datetime.cpp`; it is extracted here so DATEVALUE,
// TIMEVALUE, and VALUE all call into the same code path.

#include "eval/date_text_parse.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "eval/date_time.h"

namespace formulon {
namespace eval {
namespace date_parse {
namespace {

// UTF-8 byte sequences for the kanji the date / time parser recognises.
// Declared as 4-byte arrays (3 UTF-8 bytes + NUL) so they forward cleanly
// into `starts_with_utf8` without array-decay pitfalls.
constexpr char kKanjiNen[4] = {'\xE5', '\xB9', '\xB4', '\0'};    // 年 year
constexpr char kKanjiGatsu[4] = {'\xE6', '\x9C', '\x88', '\0'};  // 月 month
constexpr char kKanjiNichi[4] = {'\xE6', '\x97', '\xA5', '\0'};  // 日 day
constexpr char kKanjiJi[4] = {'\xE6', '\x99', '\x82', '\0'};     // 時 hour
constexpr char kKanjiFun[4] = {'\xE5', '\x88', '\x86', '\0'};    // 分 minute
constexpr char kKanjiByou[4] = {'\xE7', '\xA7', '\x92', '\0'};   // 秒 second

// 6-byte UTF-8 sequences for the five Japanese era names.
constexpr char kEraReiwa[7] = {'\xE4', '\xBB', '\xA4', '\xE5', '\x92', '\x8C', '\0'};   // 令和
constexpr char kEraHeisei[7] = {'\xE5', '\xB9', '\xB3', '\xE6', '\x88', '\x90', '\0'};  // 平成
constexpr char kEraShowa[7] = {'\xE6', '\x98', '\xAD', '\xE5', '\x92', '\x8C', '\0'};   // 昭和
constexpr char kEraTaisho[7] = {'\xE5', '\xA4', '\xA7', '\xE6', '\xAD', '\xA3', '\0'};  // 大正
constexpr char kEraMeiji[7] = {'\xE6', '\x98', '\x8E', '\xE6', '\xB2', '\xBB', '\0'};   // 明治

// Returns true iff `s` begins with the given 3-byte UTF-8 sequence.
bool starts_with_utf8(std::string_view s, const char (&expected)[4]) noexcept {
  return s.size() >= 3 && s[0] == expected[0] && s[1] == expected[1] && s[2] == expected[2];
}

// Returns true iff `s` begins with the given 6-byte UTF-8 sequence.
bool starts_with_utf8_6(std::string_view s, const char (&expected)[7]) noexcept {
  return s.size() >= 6 && s[0] == expected[0] && s[1] == expected[1] && s[2] == expected[2] && s[3] == expected[3] &&
         s[4] == expected[4] && s[5] == expected[5];
}

// Folds full-width Arabic digits (U+FF10..U+FF19, encoded as `EF BC 90`..
// `EF BC 99` in UTF-8) into ASCII `0`..`9`. All other bytes — including
// the multi-byte kanji terminators `年/月/日/時/分/秒` and era characters
// — are passed through unchanged. Returns the folded string (or `s` itself
// when no full-width digit is present, to avoid an unnecessary copy).
std::string fold_fullwidth_digits(std::string_view s) {
  // Quick scan: if no full-width digit is present, return the input as-is.
  bool needs_fold = false;
  for (std::size_t i = 0; i + 2 < s.size(); ++i) {
    if (static_cast<unsigned char>(s[i]) == 0xEF && static_cast<unsigned char>(s[i + 1]) == 0xBC) {
      const unsigned char b2 = static_cast<unsigned char>(s[i + 2]);
      if (b2 >= 0x90 && b2 <= 0x99) {
        needs_fold = true;
        break;
      }
    }
  }
  if (!needs_fold) {
    return std::string(s);
  }
  std::string out;
  out.reserve(s.size());
  for (std::size_t i = 0; i < s.size();) {
    if (i + 2 < s.size() && static_cast<unsigned char>(s[i]) == 0xEF && static_cast<unsigned char>(s[i + 1]) == 0xBC) {
      const unsigned char b2 = static_cast<unsigned char>(s[i + 2]);
      if (b2 >= 0x90 && b2 <= 0x99) {
        out.push_back(static_cast<char>('0' + (b2 - 0x90)));
        i += 3;
        continue;
      }
    }
    out.push_back(s[i]);
    ++i;
  }
  return out;
}

// Returns the last valid day-of-month under the Gregorian leap rule.
unsigned days_in_month(int y, unsigned m) noexcept {
  static constexpr unsigned kTable[] = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};
  if (m < 1u || m > 12u) {
    return 31u;
  }
  if (m == 2u) {
    const bool leap = (y % 4 == 0 && y % 100 != 0) || (y % 400 == 0);
    return leap ? 29u : 28u;
  }
  return kTable[m - 1u];
}

// Scans 1..`max_digits` ASCII digits from the head of `s`, writes the parsed
// integer into `*out`, and advances `s` past them. Returns the number of
// digits consumed (0 if `s` did not start with a digit).
std::size_t scan_digits(std::string_view& s, int max_digits, int* out) noexcept {
  std::size_t i = 0;
  int value = 0;
  while (i < s.size() && i < static_cast<std::size_t>(max_digits)) {
    const char c = s[i];
    if (c < '0' || c > '9') {
      break;
    }
    value = value * 10 + (c - '0');
    ++i;
  }
  if (i == 0) {
    return 0;
  }
  *out = value;
  s.remove_prefix(i);
  return i;
}

// Excel's two-digit-year pivot for DATEVALUE: 00..29 -> 2000..2029,
// 30..99 -> 1930..1999. 3- and 4-digit years are returned as-is.
int expand_two_digit_year(int y, std::size_t digits) noexcept {
  if (digits <= 2) {
    return y < 30 ? (2000 + y) : (1900 + y);
  }
  return y;
}

// Case-insensitive ASCII equality for a single character.
bool ci_equal_ascii(char a, char b) noexcept {
  const char la = (a >= 'A' && a <= 'Z') ? static_cast<char>(a + ('a' - 'A')) : a;
  const char lb = (b >= 'A' && b <= 'Z') ? static_cast<char>(b + ('a' - 'A')) : b;
  return la == lb;
}

// Returns true iff `s[0..expected.size())` matches `expected` case-insensitively
// (ASCII only).
bool starts_with_ci(std::string_view s, std::string_view expected) noexcept {
  if (s.size() < expected.size()) {
    return false;
  }
  for (std::size_t i = 0; i < expected.size(); ++i) {
    if (!ci_equal_ascii(s[i], expected[i])) {
      return false;
    }
  }
  return true;
}

// Consumes a case-insensitive English month name from the head of `s` and
// advances `s` past it. On success writes the 1..12 month index into
// `*out_month` and returns true. Accepts both the three-letter abbreviation
// (Jan..Dec) and the full name (January..December); prefers the longest match
// so "June" is not cut to "Jun" + "e".
bool parse_mmm_month(std::string_view& s, int* out_month) noexcept {
  // Order matches Excel's short-form abbreviation. We try the full name first
  // to honour the longest-match rule before falling back to the 3-letter form.
  static constexpr std::string_view kFullNames[12] = {
      "January", "February", "March",     "April",   "May",      "June",
      "July",    "August",   "September", "October", "November", "December",
  };
  static constexpr std::string_view kShortNames[12] = {
      "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
  };
  for (int i = 0; i < 12; ++i) {
    if (starts_with_ci(s, kFullNames[i])) {
      s.remove_prefix(kFullNames[i].size());
      *out_month = i + 1;
      return true;
    }
  }
  for (int i = 0; i < 12; ++i) {
    if (starts_with_ci(s, kShortNames[i])) {
      s.remove_prefix(kShortNames[i].size());
      *out_month = i + 1;
      return true;
    }
  }
  return false;
}

// Parses the yyyy-first variants: `YYYY-MM-DD`, `YYYY/MM/DD`, and the kanji
// form `YYYY年MM月DD日`. Returns false if the text does not match this shape;
// in that case the caller should try the d-mmm-yyyy fall-back.
bool parse_ymd_text(std::string_view s, double* out_serial, std::string_view* rest) noexcept;

// Parses the alternate `d-mmm-yyyy` / `d mmm yyyy` / `d/mmm/yyyy` shapes,
// where the month is an English word (3-letter abbreviation or full name).
// The separator must be one of `-`, `/`, ` `, and must match on both sides of
// the month token (no mixing). Purely-numeric d-m-yyyy is intentionally not
// accepted here.
bool parse_dmy_mmm_text(std::string_view s, double* out_serial, std::string_view* rest) noexcept;

// Era ID; values matter only as enumerators within this translation unit.
enum class Era { Reiwa, Heisei, Showa, Taisho, Meiji };

// Returns the Gregorian year of era year 1 for the given era. Mac Excel
// linearly extrapolates outside the era's actual historical window — for
// example `平成32年4月30日` is accepted as 1989 + (32-1) = 2020-04-30, even
// though Heisei ended at year 31. We mirror that lenient behaviour here.
int era_year_anchor(Era e) noexcept {
  switch (e) {
    case Era::Reiwa:
      return 2019;  // 令和 1 = 2019
    case Era::Heisei:
      return 1989;  // 平成 1 = 1989
    case Era::Showa:
      return 1926;  // 昭和 1 = 1926
    case Era::Taisho:
      return 1912;  // 大正 1 = 1912
    case Era::Meiji:
      return 1868;  // 明治 1 = 1868
  }
  return 1900;  // unreachable
}

// After an era prefix has been consumed, parses the year/month/day tail in
// the kanji form `<digits>年<digits>月<digits>日`. Mac Excel rejects 元
// (gannen) and dot/slash separators when a *full-name* era prefix is used,
// so we accept only the strict kanji form here.
bool parse_era_kanji_ymd_tail(Era era, std::string_view s, double* out_serial, std::string_view* rest) noexcept {
  int era_year = 0;
  if (scan_digits(s, 4, &era_year) == 0) {
    return false;
  }
  if (era_year <= 0) {
    return false;
  }
  if (!starts_with_utf8(s, kKanjiNen)) {
    return false;
  }
  s.remove_prefix(3);
  int month = 0;
  if (scan_digits(s, 2, &month) == 0) {
    return false;
  }
  if (!starts_with_utf8(s, kKanjiGatsu)) {
    return false;
  }
  s.remove_prefix(3);
  int day = 0;
  if (scan_digits(s, 2, &day) == 0) {
    return false;
  }
  if (!starts_with_utf8(s, kKanjiNichi)) {
    return false;
  }
  s.remove_prefix(3);
  if (month < 1 || month > 12) {
    return false;
  }
  const int gy = era_year - 1 + era_year_anchor(era);
  const unsigned dim = days_in_month(gy, static_cast<unsigned>(month));
  if (day < 1 || static_cast<unsigned>(day) > dim) {
    return false;
  }
  *out_serial = date_time::serial_from_ymd(gy, static_cast<unsigned>(month), static_cast<unsigned>(day));
  *rest = s;
  return true;
}

// After an abbreviation letter has been consumed, parses the dot-separated
// `<digits>.<digits>.<digits>` tail used by `R6.4.1` etc.
bool parse_era_dot_ymd_tail(Era era, std::string_view s, double* out_serial, std::string_view* rest) noexcept {
  int era_year = 0;
  if (scan_digits(s, 4, &era_year) == 0) {
    return false;
  }
  if (era_year <= 0) {
    return false;
  }
  if (s.empty() || s[0] != '.') {
    return false;
  }
  s.remove_prefix(1);
  int month = 0;
  if (scan_digits(s, 2, &month) == 0) {
    return false;
  }
  if (s.empty() || s[0] != '.') {
    return false;
  }
  s.remove_prefix(1);
  int day = 0;
  if (scan_digits(s, 2, &day) == 0) {
    return false;
  }
  if (month < 1 || month > 12) {
    return false;
  }
  const int gy = era_year - 1 + era_year_anchor(era);
  const unsigned dim = days_in_month(gy, static_cast<unsigned>(month));
  if (day < 1 || static_cast<unsigned>(day) > dim) {
    return false;
  }
  *out_serial = date_time::serial_from_ymd(gy, static_cast<unsigned>(month), static_cast<unsigned>(day));
  *rest = s;
  return true;
}

// Recognises `<full-era-name><digits>年<digits>月<digits>日` and
// `<single-letter-era>.<digits>.<digits>.<digits>`. Returns false if the
// input does not start with one of the five recognised eras; in that case
// the caller falls back to the regular Gregorian date forms.
bool parse_era_text(std::string_view s, double* out_serial, std::string_view* rest) noexcept {
  // Full-name eras require the strict 年/月/日 grammar.
  if (starts_with_utf8_6(s, kEraReiwa)) {
    return parse_era_kanji_ymd_tail(Era::Reiwa, s.substr(6), out_serial, rest);
  }
  if (starts_with_utf8_6(s, kEraHeisei)) {
    return parse_era_kanji_ymd_tail(Era::Heisei, s.substr(6), out_serial, rest);
  }
  if (starts_with_utf8_6(s, kEraShowa)) {
    return parse_era_kanji_ymd_tail(Era::Showa, s.substr(6), out_serial, rest);
  }
  if (starts_with_utf8_6(s, kEraTaisho)) {
    return parse_era_kanji_ymd_tail(Era::Taisho, s.substr(6), out_serial, rest);
  }
  if (starts_with_utf8_6(s, kEraMeiji)) {
    return parse_era_kanji_ymd_tail(Era::Meiji, s.substr(6), out_serial, rest);
  }
  // Single-letter abbreviations require the ASCII dot-separated grammar and
  // a digit immediately after the letter (so `Mar` etc. don't get hijacked).
  if (s.size() >= 2 && s[1] >= '0' && s[1] <= '9') {
    switch (s[0]) {
      case 'R':
      case 'r':
        return parse_era_dot_ymd_tail(Era::Reiwa, s.substr(1), out_serial, rest);
      case 'H':
      case 'h':
        return parse_era_dot_ymd_tail(Era::Heisei, s.substr(1), out_serial, rest);
      case 'S':
      case 's':
        return parse_era_dot_ymd_tail(Era::Showa, s.substr(1), out_serial, rest);
      case 'T':
      case 't':
        return parse_era_dot_ymd_tail(Era::Taisho, s.substr(1), out_serial, rest);
      case 'M':
      case 'm':
        return parse_era_dot_ymd_tail(Era::Meiji, s.substr(1), out_serial, rest);
      default:
        break;
    }
  }
  return false;
}

// Parses a leading date token. See `parse_date_time_text` for the grammar.
bool parse_date_text(std::string_view s, double* out_serial, std::string_view* rest) noexcept {
  if (parse_era_text(s, out_serial, rest)) {
    return true;
  }
  if (parse_ymd_text(s, out_serial, rest)) {
    return true;
  }
  return parse_dmy_mmm_text(s, out_serial, rest);
}

bool parse_ymd_text(std::string_view s, double* out_serial, std::string_view* rest) noexcept {
  int year = 0;
  const std::size_t year_digits = scan_digits(s, 4, &year);
  if (year_digits == 0) {
    return false;
  }
  // Separator after year: '-', '/', or 年.
  bool kanji_form = false;
  if (!s.empty() && (s[0] == '-' || s[0] == '/')) {
    s.remove_prefix(1);
  } else if (starts_with_utf8(s, kKanjiNen)) {
    s.remove_prefix(3);
    kanji_form = true;
  } else {
    return false;
  }
  int month = 0;
  if (scan_digits(s, 2, &month) == 0) {
    return false;
  }
  if (kanji_form) {
    if (!starts_with_utf8(s, kKanjiGatsu)) {
      return false;
    }
    s.remove_prefix(3);
  } else {
    if (s.empty() || (s[0] != '-' && s[0] != '/')) {
      return false;
    }
    s.remove_prefix(1);
  }
  int day = 0;
  if (scan_digits(s, 2, &day) == 0) {
    return false;
  }
  if (kanji_form) {
    if (!starts_with_utf8(s, kKanjiNichi)) {
      return false;
    }
    s.remove_prefix(3);
  }
  const int expanded_year = expand_two_digit_year(year, year_digits);
  if (expanded_year < 1900 || expanded_year > 9999) {
    return false;
  }
  if (month < 1 || month > 12) {
    return false;
  }
  const unsigned dim = days_in_month(expanded_year, static_cast<unsigned>(month));
  // Excel preserves Lotus 1-2-3's fictitious 1900-02-29 (serial 60). The
  // Gregorian calendar says that day does not exist (1900 is divisible by
  // 100 but not 400), but DATEVALUE must still accept it for parity.
  const bool is_excel_ghost_day = (expanded_year == 1900 && month == 2 && day == 29);
  if (!is_excel_ghost_day && (day < 1 || static_cast<unsigned>(day) > dim)) {
    return false;
  }
  *out_serial = date_time::serial_from_ymd(expanded_year, static_cast<unsigned>(month), static_cast<unsigned>(day));
  *rest = s;
  return true;
}

bool parse_dmy_mmm_text(std::string_view s, double* out_serial, std::string_view* rest) noexcept {
  // Leading day token: exactly one or two ASCII digits.
  int day = 0;
  const std::size_t day_digits = scan_digits(s, 2, &day);
  if (day_digits == 0) {
    return false;
  }
  // Separator before the month word: '-', '/', or a single ASCII space. The
  // chosen separator is remembered and required again after the month word,
  // so inputs like "29-Feb/1900" are rejected.
  if (s.empty()) {
    return false;
  }
  const char sep = s[0];
  if (sep != '-' && sep != '/' && sep != ' ') {
    return false;
  }
  s.remove_prefix(1);
  // Month word: case-insensitive 3-letter abbreviation or full English name.
  int month = 0;
  if (!parse_mmm_month(s, &month)) {
    return false;
  }
  // Trailing separator must match the leading one exactly.
  if (s.empty() || s[0] != sep) {
    return false;
  }
  s.remove_prefix(1);
  // Year token: 1..4 ASCII digits. Two-digit values go through the same
  // Excel 1900/2000 pivot as the yyyy-first path.
  int year = 0;
  const std::size_t year_digits = scan_digits(s, 4, &year);
  if (year_digits == 0) {
    return false;
  }
  const int expanded_year = expand_two_digit_year(year, year_digits);
  if (expanded_year < 1900 || expanded_year > 9999) {
    return false;
  }
  if (month < 1 || month > 12) {
    return false;
  }
  const unsigned dim = days_in_month(expanded_year, static_cast<unsigned>(month));
  // Reuse the same 1900-02-29 ghost-day escape as the yyyy-first path so
  // `DATEVALUE("29-Feb-1900")` still resolves to serial 60.
  const bool is_excel_ghost_day = (expanded_year == 1900 && month == 2 && day == 29);
  if (!is_excel_ghost_day && (day < 1 || static_cast<unsigned>(day) > dim)) {
    return false;
  }
  *out_serial = date_time::serial_from_ymd(expanded_year, static_cast<unsigned>(month), static_cast<unsigned>(day));
  *rest = s;
  return true;
}

// Parses the kanji time form `<digits>時<digits>分[<digits>秒]`. The `分`
// segment is required (Mac Excel rejects bare `8時` as #VALUE!), and `秒`,
// when present, must follow `分`. Returns false if the input does not
// match this exact shape; on success advances `*rest` past the `分` or
// `秒` terminator.
bool parse_kanji_time_text(std::string_view s, double* out_frac, std::string_view* rest) noexcept {
  int hour = 0;
  if (scan_digits(s, 3, &hour) == 0) {
    return false;
  }
  if (!starts_with_utf8(s, kKanjiJi)) {
    return false;
  }
  s.remove_prefix(3);
  int minute = 0;
  if (scan_digits(s, 3, &minute) == 0) {
    return false;
  }
  if (!starts_with_utf8(s, kKanjiFun)) {
    return false;
  }
  s.remove_prefix(3);
  int second = 0;
  if (!s.empty() && s[0] >= '0' && s[0] <= '9') {
    if (scan_digits(s, 3, &second) == 0) {
      return false;
    }
    if (!starts_with_utf8(s, kKanjiByou)) {
      return false;
    }
    s.remove_prefix(3);
  }
  if (hour < 0 || minute < 0 || second < 0) {
    return false;
  }
  const double total_seconds =
      static_cast<double>(hour) * 3600.0 + static_cast<double>(minute) * 60.0 + static_cast<double>(second);
  *out_frac = total_seconds / 86400.0;
  *rest = s;
  return true;
}

// Parses a leading time token. See `parse_date_time_text` for the grammar.
bool parse_time_text(std::string_view s, double* out_frac, std::string_view* rest) noexcept {
  // Probe for the kanji form `H時M分[S秒]` first: scan past the leading
  // digit run and check for 時. If that matches, the entire token must
  // parse via the kanji branch — there is no ambiguity with `H:M:S`.
  {
    std::string_view probe = s;
    int dummy = 0;
    const std::size_t leading_digits = scan_digits(probe, 3, &dummy);
    if (leading_digits > 0 && starts_with_utf8(probe, kKanjiJi)) {
      return parse_kanji_time_text(s, out_frac, rest);
    }
  }
  int hour = 0;
  if (scan_digits(s, 3, &hour) == 0) {
    return false;
  }
  if (s.empty() || s[0] != ':') {
    return false;
  }
  s.remove_prefix(1);
  int minute = 0;
  if (scan_digits(s, 3, &minute) == 0 || minute < 0) {
    return false;
  }
  int second = 0;
  bool has_seconds = false;
  if (!s.empty() && s[0] == ':') {
    s.remove_prefix(1);
    if (scan_digits(s, 3, &second) == 0 || second < 0) {
      return false;
    }
    has_seconds = true;
  }
  double sub_seconds = 0.0;
  if (has_seconds && !s.empty() && s[0] == '.') {
    s.remove_prefix(1);
    double scale = 0.1;
    bool any = false;
    while (!s.empty() && s[0] >= '0' && s[0] <= '9') {
      sub_seconds += (s[0] - '0') * scale;
      scale *= 0.1;
      s.remove_prefix(1);
      any = true;
    }
    if (!any) {
      return false;
    }
  }
  bool pm = false;
  bool have_ampm = false;
  {
    std::string_view tail = s;
    std::size_t space_count = 0;
    while (!tail.empty() && (tail[0] == ' ' || tail[0] == '\t')) {
      tail.remove_prefix(1);
      ++space_count;
    }
    if (space_count >= 1 && tail.size() >= 2) {
      const char c0 = static_cast<char>(tail[0] >= 'a' && tail[0] <= 'z' ? tail[0] - ('a' - 'A') : tail[0]);
      const char c1 = static_cast<char>(tail[1] >= 'a' && tail[1] <= 'z' ? tail[1] - ('a' - 'A') : tail[1]);
      if ((c0 == 'A' || c0 == 'P') && c1 == 'M') {
        pm = (c0 == 'P');
        have_ampm = true;
        tail.remove_prefix(2);
        s = tail;
      }
    }
  }
  if (have_ampm) {
    if (hour > 12) {
      return false;
    }
    if (hour == 12 && !pm) {
      hour = 0;
    } else if (pm && hour >= 1 && hour < 12) {
      hour += 12;
    }
  }
  const double total_seconds = static_cast<double>(hour) * 3600.0 + static_cast<double>(minute) * 60.0 +
                               static_cast<double>(second) + sub_seconds;
  if (total_seconds < 0.0) {
    return false;
  }
  *out_frac = total_seconds / 86400.0;
  *rest = s;
  return true;
}

}  // namespace

std::string_view trim_date_text(std::string_view s) noexcept {
  auto is_ascii_ws = [](char c) {
    const unsigned char uc = static_cast<unsigned char>(c);
    return uc == ' ' || uc == '\t' || uc == '\n' || uc == '\r' || uc == '\v' || uc == '\f';
  };
  while (!s.empty() && is_ascii_ws(s.front())) {
    s.remove_prefix(1);
  }
  while (!s.empty() && is_ascii_ws(s.back())) {
    s.remove_suffix(1);
  }
  return s;
}

bool parse_date_time_text(std::string_view s, double* out_date_serial, double* out_time_frac, bool* out_has_date,
                          bool* out_has_time) noexcept {
  // Fold `０..９` (U+FF10..U+FF19) to ASCII before tokenisation. Mac Excel
  // accepts full-width digits anywhere ASCII digits are expected; the
  // surrounding kanji terminators / era characters / punctuation pass
  // through unchanged. `folded` owns the storage when a copy was needed.
  const std::string folded = fold_fullwidth_digits(s);
  std::string_view rest(folded);
  double serial = 0.0;
  double frac = 0.0;
  bool has_date = false;
  bool has_time = false;
  if (parse_date_text(rest, &serial, &rest)) {
    has_date = true;
    std::size_t space_count = 0;
    while (!rest.empty() && (rest[0] == ' ' || rest[0] == '\t')) {
      rest.remove_prefix(1);
      ++space_count;
    }
    if (space_count >= 1 && !rest.empty()) {
      std::string_view after = rest;
      double f = 0.0;
      if (parse_time_text(rest, &f, &after)) {
        frac = f;
        has_time = true;
        rest = after;
      }
    }
  } else {
    std::string_view after = rest;
    double f = 0.0;
    if (parse_time_text(rest, &f, &after)) {
      frac = f;
      has_time = true;
      rest = after;
    }
  }
  if (!has_date && !has_time) {
    return false;
  }
  // Tail cleanup: Mac Excel is lenient about trailing whitespace after a
  // successfully tokenised date/time, including U+3000.
  while (!rest.empty()) {
    const unsigned char uc = static_cast<unsigned char>(rest.front());
    if (uc == ' ' || uc == '\t' || uc == '\n' || uc == '\r' || uc == '\v' || uc == '\f') {
      rest.remove_prefix(1);
      continue;
    }
    if (rest.size() >= 3 && static_cast<unsigned char>(rest[0]) == 0xE3 &&
        static_cast<unsigned char>(rest[1]) == 0x80 && static_cast<unsigned char>(rest[2]) == 0x80) {
      rest.remove_prefix(3);
      continue;
    }
    break;
  }
  if (!rest.empty()) {
    return false;
  }
  *out_date_serial = serial;
  *out_time_frac = frac;
  *out_has_date = has_date;
  *out_has_time = has_time;
  return true;
}

}  // namespace date_parse
}  // namespace eval
}  // namespace formulon
