// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the shared date / time text parser declared in
// `date_text_parse.h`. The logic matches the grammar previously embedded in
// `src/eval/builtins/datetime.cpp`; it is extracted here so DATEVALUE,
// TIMEVALUE, and VALUE all call into the same code path.

#include "eval/date_text_parse.h"

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "eval/date_time.h"

namespace formulon {
namespace eval {
namespace date_parse {
namespace {

// UTF-8 byte sequences for the three kanji the date parser recognises:
// `kKanjiNen` (年, year), `kKanjiGatsu` (月, month), `kKanjiNichi` (日, day).
// Declared as 4-byte arrays (3 UTF-8 bytes + NUL) so they forward cleanly
// into `starts_with_utf8` without array-decay pitfalls.
constexpr char kKanjiNen[4] = {'\xE5', '\xB9', '\xB4', '\0'};    // 年
constexpr char kKanjiGatsu[4] = {'\xE6', '\x9C', '\x88', '\0'};  // 月
constexpr char kKanjiNichi[4] = {'\xE6', '\x97', '\xA5', '\0'};  // 日

// Returns true iff `s` begins with the given 3-byte UTF-8 sequence.
bool starts_with_utf8(std::string_view s, const char (&expected)[4]) noexcept {
  return s.size() >= 3 && s[0] == expected[0] && s[1] == expected[1] && s[2] == expected[2];
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

// Parses a leading date token. See `parse_date_time_text` for the grammar.
bool parse_date_text(std::string_view s, double* out_serial, std::string_view* rest) noexcept {
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
  if (day < 1 || static_cast<unsigned>(day) > dim) {
    return false;
  }
  *out_serial = date_time::serial_from_ymd(expanded_year, static_cast<unsigned>(month), static_cast<unsigned>(day));
  *rest = s;
  return true;
}

// Parses a leading time token. See `parse_date_time_text` for the grammar.
bool parse_time_text(std::string_view s, double* out_frac, std::string_view* rest) noexcept {
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
  std::string_view rest = s;
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
