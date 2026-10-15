//
// Implementation of the date / time text recogniser declared in
// `date_text_parse.h`.
//
// The text is folded once (digit scripts, full-width ASCII, the ideographic
// space), then read as a date, a date followed by a time and an ignored
// numeric tail, a time, or a time followed by a date. Each date reader
// returns every reading of the head of the text, longest first, so the
// composition can back off from `1 Jan 1` to `1 Jan` when a time follows.
// Kanji, era and kanji-time forms stand alone and never compose.

#include "eval/date_text_parse.h"

#include <array>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "eval/eval_profile_scope.h"
#include "eval/number_parse.h"
#include "excel_locale.h"
#include "utils/date_time.h"
#include "utils/strings.h"

namespace formulon {
namespace eval {
namespace date_parse {
namespace {

// Day-of-month clamping shares the calendar helper with the date builtins.
using date_time::days_in_month;

// UTF-8 byte sequences for the kanji the date / time parser recognises.
// Declared as 4-byte arrays (3 UTF-8 bytes + NUL) so they forward cleanly
// into `starts_with_utf8` without array-decay pitfalls.
constexpr char kKanjiNen[4] = {'\xE5', '\xB9', '\xB4', '\0'};     // 年 year
constexpr char kKanjiGatsu[4] = {'\xE6', '\x9C', '\x88', '\0'};   // 月 month
constexpr char kKanjiNichi[4] = {'\xE6', '\x97', '\xA5', '\0'};   // 日 day
constexpr char kKanjiJi[4] = {'\xE6', '\x99', '\x82', '\0'};      // 時 hour
constexpr char kKanjiFun[4] = {'\xE5', '\x88', '\x86', '\0'};     // 分 minute
constexpr char kKanjiByou[4] = {'\xE7', '\xA7', '\x92', '\0'};    // 秒 second
constexpr char kHangulNyeon[4] = {'\xEB', '\x85', '\x84', '\0'};  // 년 year
constexpr char kHangulWol[4] = {'\xEC', '\x9B', '\x94', '\0'};    // 월 month
constexpr char kHangulIl[4] = {'\xEC', '\x9D', '\xBC', '\0'};     // 일 day

// 6-byte UTF-8 sequences for the five Japanese era names.
constexpr char kEraReiwa[7] = {'\xE4', '\xBB', '\xA4', '\xE5', '\x92', '\x8C', '\0'};   // 令和
constexpr char kEraHeisei[7] = {'\xE5', '\xB9', '\xB3', '\xE6', '\x88', '\x90', '\0'};  // 平成
constexpr char kEraShowa[7] = {'\xE6', '\x98', '\xAD', '\xE5', '\x92', '\x8C', '\0'};   // 昭和
constexpr char kEraTaisho[7] = {'\xE5', '\xA4', '\xA7', '\xE6', '\xAD', '\xA3', '\0'};  // 大正
constexpr char kEraMeiji[7] = {'\xE6', '\x98', '\x8E', '\xE6', '\xB2', '\xBB', '\0'};   // 明治

constexpr std::string_view kEnglishMonths[12] = {
    "January", "February", "March",     "April",   "May",      "June",
    "July",    "August",   "September", "October", "November", "December",
};

bool is_digit(char c) noexcept {
  return c >= '0' && c <= '9';
}

bool is_ascii_alpha(char c) noexcept {
  return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
}

char ascii_upper(char c) noexcept {
  return c >= 'a' && c <= 'z' ? static_cast<char>(c - ('a' - 'A')) : c;
}

// Skips U+0020 blanks; true when at least one was skipped.
bool skip_blanks(std::string_view* s) noexcept {
  std::size_t n = 0;
  while (n < s->size() && (*s)[n] == ' ') {
    ++n;
  }
  s->remove_prefix(n);
  return n > 0;
}

// Consumes a date/time separator, accepting blanks on either side of it.
// Leaves `s` untouched and returns false when the separator is absent.
template <typename IsSep>
bool consume_spaced_sep(std::string_view* s, IsSep is_sep) noexcept {
  std::string_view t = *s;
  skip_blanks(&t);
  if (t.empty() || !is_sep(t[0])) {
    return false;
  }
  t.remove_prefix(1);
  skip_blanks(&t);
  *s = t;
  return true;
}

bool is_date_sep(char c) noexcept {
  return c == '-' || c == '/';
}

bool is_time_sep(char c) noexcept {
  return c == ':';
}

// Returns true iff `s` begins with the given 3-byte UTF-8 sequence.
bool starts_with_utf8(std::string_view s, const char (&expected)[4]) noexcept {
  return s.size() >= 3 && s[0] == expected[0] && s[1] == expected[1] && s[2] == expected[2];
}

// Folds what the recogniser reads as ASCII: the decimal digit scripts VALUE
// reads, the ideographic space (a blank) and, where the profile takes
// full-width numeric text, the full-width ASCII forms U+FF01..U+FF5E.
std::string fold_date_text(std::string_view s) {
  const std::string scripts = fold_script_digits(s);
  const bool fold_fullwidth = locale_facts(current_eval_profile()).fullwidth_numeric_text;
  std::string out;
  out.reserve(scripts.size());
  for (std::size_t i = 0; i < scripts.size();) {
    if (i + 2 < scripts.size()) {
      const auto b0 = static_cast<unsigned char>(scripts[i]);
      const auto b1 = static_cast<unsigned char>(scripts[i + 1]);
      const auto b2 = static_cast<unsigned char>(scripts[i + 2]);
      if (b0 == 0xE3 && b1 == 0x80 && b2 == 0x80) {
        out.push_back(' ');
        i += 3;
        continue;
      }
      if (fold_fullwidth && b0 == 0xEF && ((b1 == 0xBC && b2 >= 0x81) || (b1 == 0xBD && b2 <= 0x9E))) {
        const unsigned cp = 0xF000u | ((b1 & 0x3Fu) << 6) | (b2 & 0x3Fu);
        out.push_back(static_cast<char>(cp - 0xFEE0u));
        i += 3;
        continue;
      }
    }
    out.push_back(scripts[i]);
    ++i;
  }
  return out;
}

// Scans 1..`max_digits` ASCII digits from the head of `s`, writes the parsed
// integer into `*out`, and advances `s` past them. Returns the number of
// digits consumed (0 if `s` did not start with a digit).
std::size_t scan_digits(std::string_view& s, int max_digits, int* out) noexcept {
  std::size_t i = 0;
  int value = 0;
  while (i < s.size() && i < static_cast<std::size_t>(max_digits)) {
    const char c = s[i];
    if (!is_digit(c)) {
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

// A whole run of digits; the value saturates so any length stays representable.
struct DigitRun {
  int value;
  std::size_t digits;
};

constexpr int kRunCap = 1000000;

bool scan_run(std::string_view* s, DigitRun* out) noexcept {
  std::size_t i = 0;
  int value = 0;
  while (i < s->size() && is_digit((*s)[i])) {
    value = value < kRunCap ? value * 10 + ((*s)[i] - '0') : kRunCap;
    ++i;
  }
  if (i == 0) {
    return false;
  }
  *out = DigitRun{value, i};
  s->remove_prefix(i);
  return true;
}

// Excel's two-digit-year pivot for DATEVALUE: 00..29 -> 2000..2029,
// 30..99 -> 1930..1999. 3- and 4-digit years are returned as-is.
int expand_two_digit_year(int y, std::size_t digits) noexcept {
  if (digits <= 2) {
    return y < 30 ? (2000 + y) : (1900 + y);
  }
  return y;
}

// Turns a parsed (year, month, day) triple into a serial, applying the
// pivot and the calendar bounds every text date form shares. `year` is
// still in its written form; `year_digits` decides whether the pivot
// applies. Excel keeps Lotus 1-2-3's fictitious 1900-02-29 (serial 60),
// which the Gregorian calendar does not have, so that one date is
// exempted from the day-of-month bound.
//
// Returns `false` for anything outside the range DATEVALUE accepts.
bool serial_from_parsed_ymd(int year, std::size_t year_digits, int month, int day, double* out_serial) noexcept {
  const int expanded_year = expand_two_digit_year(year, year_digits);
  if (expanded_year < 1900 || expanded_year > 9999) {
    return false;
  }
  if (month < 1 || month > 12) {
    return false;
  }
  const unsigned dim = days_in_month(expanded_year, static_cast<unsigned>(month));
  const bool is_excel_ghost_day = (expanded_year == 1900 && month == 2 && day == 29);
  if (!is_excel_ghost_day && (day < 1 || static_cast<unsigned>(day) > dim)) {
    return false;
  }
  *out_serial = date_time::serial_from_ymd(expanded_year, static_cast<unsigned>(month), static_cast<unsigned>(day));
  return true;
}

// A written year: one or two digits take the two-digit window, four digits
// are the year itself; three or five and more are no year.
bool written_year_serial(DigitRun year, int month, int day, double* out) noexcept {
  return year.digits <= 4 && serial_from_parsed_ymd(year.value, year.digits, month, day, out);
}

bool current_year_serial(int current_year, int month, int day, double* out) noexcept {
  return current_year > 0 && serial_from_parsed_ymd(current_year, 4, month, day, out);
}

// Every reading of the head of a text as a date, longest first.
struct DateCandidate {
  double serial;
  std::string_view rest;
  // Kanji and era dates never take a time.
  bool standalone;
};

struct DateCandidates {
  std::array<DateCandidate, 4> items{};
  std::size_t count = 0;

  void add(double serial, std::string_view rest, bool standalone = false) noexcept {
    if (count < items.size()) {
      items[count++] = DateCandidate{serial, rest, standalone};
    }
  }
};

// ---------------------------------------------------------------------------
// Year-Month-Day profiles
// ---------------------------------------------------------------------------

enum class ShapeMatch : std::uint8_t { kNoMatch, kInvalid, kSuccess };

// Era ID; values matter only as enumerators within this translation unit.
enum class Era { Reiwa, Heisei, Showa, Taisho, Meiji };

// Gregorian year of era year 1. Excel extrapolates past an era's end
// (`平成32年4月30日` is 2020-04-30).
int era_year_anchor(Era e) noexcept {
  switch (e) {
    case Era::Reiwa:
      return 2019;
    case Era::Heisei:
      return 1989;
    case Era::Showa:
      return 1926;
    case Era::Taisho:
      return 1912;
    case Era::Meiji:
      return 1868;
  }
  return 1900;  // unreachable
}

struct EraName {
  std::string_view kanji;
  char letter;  // upper-case abbreviation; the lower-case form is accepted too
  Era era;
};

constexpr EraName kEraNames[] = {
    {std::string_view(kEraReiwa, 6), 'R', Era::Reiwa}, {std::string_view(kEraHeisei, 6), 'H', Era::Heisei},
    {std::string_view(kEraShowa, 6), 'S', Era::Showa}, {std::string_view(kEraTaisho, 6), 'T', Era::Taisho},
    {std::string_view(kEraMeiji, 6), 'M', Era::Meiji},
};

// Scans a digit run (any number of leading zeros) closed by the kanji `unit`.
bool scan_run_with_unit(std::string_view* s, const char (&unit)[4], DigitRun* out) noexcept {
  if (!scan_run(s, out) || !starts_with_utf8(*s, unit)) {
    return false;
  }
  s->remove_prefix(3);
  return true;
}

// Era years run 1..99, and a date before 1900 is no date (most of Meiji).
bool era_serial(Era era, DigitRun year, DigitRun month, DigitRun day, double* out) noexcept {
  if (year.value < 1 || year.value > 99) {
    return false;
  }
  return serial_from_parsed_ymd(era_year_anchor(era) + year.value - 1, 4, month.value, day.value, out);
}

// `<era name>Y年M月D日` and `<letter>Y.M.D` (`.`, `/` or `-`). Every field
// takes any number of leading zeros; no blanks inside.
ShapeMatch parse_era_text(std::string_view s, double* out_serial, std::string_view* rest) noexcept {
  DigitRun year{};
  DigitRun month{};
  DigitRun day{};
  for (const EraName& name : kEraNames) {
    if (s.substr(0, name.kanji.size()) == name.kanji) {
      s.remove_prefix(name.kanji.size());
      if (!scan_run_with_unit(&s, kKanjiNen, &year) || !scan_run_with_unit(&s, kKanjiGatsu, &month) ||
          !scan_run_with_unit(&s, kKanjiNichi, &day) || !era_serial(name.era, year, month, day, out_serial)) {
        return ShapeMatch::kInvalid;
      }
      *rest = s;
      return ShapeMatch::kSuccess;
    }
  }
  // A digit must follow the letter so `Mar` etc. don't get hijacked.
  if (s.size() < 2 || !is_digit(s[1])) {
    return ShapeMatch::kNoMatch;
  }
  for (const EraName& name : kEraNames) {
    if (ascii_upper(s[0]) != name.letter) {
      continue;
    }
    s.remove_prefix(1);
    const auto take_sep = [&s] {
      if (s.empty() || (s[0] != '.' && !is_date_sep(s[0]))) {
        return false;
      }
      s.remove_prefix(1);
      return true;
    };
    if (!scan_run(&s, &year) || !take_sep() || !scan_run(&s, &month) || !take_sep() || !scan_run(&s, &day) ||
        !era_serial(name.era, year, month, day, out_serial)) {
      return ShapeMatch::kInvalid;
    }
    *rest = s;
    return ShapeMatch::kSuccess;
  }
  return ShapeMatch::kNoMatch;
}

// `Y年M月D日`, `Y年M月` (the 1st) and `M月D日` (the current year), `first`
// already scanned. Fields take any number of leading zeros; the year is read
// by value, so `026年` is 2026.
bool parse_kanji_date(DigitRun first, std::string_view s, int current_year, double* out_serial,
                      std::string_view* rest) noexcept {
  DigitRun month{};
  DigitRun day{};
  if (starts_with_utf8(s, kKanjiNen)) {
    s.remove_prefix(3);
    if (!scan_run_with_unit(&s, kKanjiGatsu, &month)) {
      return false;
    }
    const std::size_t year_digits = first.value <= 99 ? 2 : 4;
    if (s.empty() || !is_digit(s[0])) {
      *rest = s;
      return serial_from_parsed_ymd(first.value, year_digits, month.value, 1, out_serial);
    }
    if (!scan_run_with_unit(&s, kKanjiNichi, &day)) {
      return false;
    }
    *rest = s;
    return serial_from_parsed_ymd(first.value, year_digits, month.value, day.value, out_serial);
  }
  s.remove_prefix(3);
  if (!scan_run_with_unit(&s, kKanjiNichi, &day)) {
    return false;
  }
  *rest = s;
  return current_year_serial(current_year, first.value, day.value, out_serial);
}

// Finishes the `<y>년 <m>월 <d>일` form after the year unit; blanks may follow
// the year and month units.
bool finish_hangul_ymd(std::string_view s, DigitRun year, double* out_serial, std::string_view* rest) noexcept {
  skip_blanks(&s);
  int month = 0;
  if (scan_digits(s, 2, &month) == 0 || !starts_with_utf8(s, kHangulWol)) {
    return false;
  }
  s.remove_prefix(3);
  skip_blanks(&s);
  int day = 0;
  if (scan_digits(s, 2, &day) == 0 || !starts_with_utf8(s, kHangulIl)) {
    return false;
  }
  s.remove_prefix(3);
  if (!written_year_serial(year, month, day, out_serial)) {
    return false;
  }
  *rest = s;
  return true;
}

// A month pair resolves m/d against the current year when the first number
// is at most 12, d/m when it is at most 31 (falling back to y/m when that is
// no day), and y/m above that or for a four-digit first number.
bool pair_serial(DigitRun a, DigitRun b, int current_year, double* out) noexcept {
  if (a.digits == 4) {
    return serial_from_parsed_ymd(a.value, 4, b.value, 1, out);
  }
  if (a.digits > 2) {
    return false;
  }
  if (a.value <= 12) {
    return current_year_serial(current_year, a.value, b.value, out);
  }
  if (a.value <= 31 && current_year_serial(current_year, b.value, a.value, out)) {
    return true;
  }
  return serial_from_parsed_ymd(a.value, 2, b.value, 1, out);
}

// `Y/M/D` and the pairs above, `first` already scanned. `/` and `-` mix
// freely (and `.` where the profile writes dotted dates), with blanks around
// each; month and day take one or two digits.
void numeric_date_candidates(DigitRun first, std::string_view s, int current_year, DateCandidates* out) noexcept {
  const bool dotted = locale_facts(current_eval_profile()).dotted_date;
  const auto is_sep = [dotted](char c) { return is_date_sep(c) || (dotted && c == '.'); };
  DigitRun second{};
  if (!consume_spaced_sep(&s, is_sep) || !scan_run(&s, &second) || second.digits > 2) {
    return;
  }
  double serial = 0.0;
  std::string_view t = s;
  DigitRun third{};
  if (consume_spaced_sep(&t, is_sep) && scan_run(&t, &third)) {
    if (third.digits <= 2 && written_year_serial(first, second.value, third.value, &serial)) {
      out->add(serial, t);
    }
    return;
  }
  if (pair_serial(first, second, current_year, &serial)) {
    out->add(serial, s);
  }
}

// Reads an English month name: the letters up to the next non-letter, at
// least three of them, beginning the full name in any case (`Sept`, `Janu`).
bool scan_month_word(std::string_view* s, int* month) noexcept {
  std::size_t n = 0;
  while (n < s->size() && is_ascii_alpha((*s)[n])) {
    ++n;
  }
  if (n < 3) {
    return false;
  }
  for (std::size_t i = 0; i < 12; ++i) {
    if (n <= kEnglishMonths[i].size() &&
        strings::case_insensitive_eq(s->substr(0, n), kEnglishMonths[i].substr(0, n))) {
      *month = static_cast<int>(i) + 1;
      s->remove_prefix(n);
      return true;
    }
  }
  return false;
}

// Between a month name and a number: nothing, blanks, or one `-` / `/` with
// blanks around it.
void skip_name_gap(std::string_view* s) noexcept {
  skip_blanks(s);
  if (!s->empty() && is_date_sep((*s)[0])) {
    s->remove_prefix(1);
    skip_blanks(s);
  }
}

// `D MMM [Y]`, `first` already scanned. The day must exist; there is no
// fallback to a year.
void day_first_name_candidates(DigitRun day, std::string_view s, int current_year, DateCandidates* out) noexcept {
  int month = 0;
  if (day.digits > 2) {
    return;
  }
  skip_name_gap(&s);
  if (!scan_month_word(&s, &month)) {
    return;
  }
  double serial = 0.0;
  std::string_view t = s;
  skip_name_gap(&t);
  DigitRun year{};
  if (scan_run(&t, &year) && written_year_serial(year, month, day.value, &serial)) {
    out->add(serial, t);
  }
  if (current_year_serial(current_year, month, day.value, &serial)) {
    out->add(serial, s);
  }
}

// `MMM D, Y`, `MMM D` and `MMM Y`. A one- or two-digit number that is no day
// of the month in the current year is a two-digit year (`Jan 32` is
// 1932-01-01); four digits are a year.
void month_first_candidates(std::string_view s, int current_year, DateCandidates* out) noexcept {
  int month = 0;
  if (!scan_month_word(&s, &month)) {
    return;
  }
  skip_name_gap(&s);
  DigitRun number{};
  if (!scan_run(&s, &number)) {
    return;
  }
  double serial = 0.0;
  std::string_view t = s;
  skip_blanks(&t);
  DigitRun year{};
  if (number.digits <= 2 && !t.empty() && t[0] == ',') {
    t.remove_prefix(1);
    if (skip_blanks(&t) && scan_run(&t, &year) && written_year_serial(year, month, number.value, &serial)) {
      out->add(serial, t);
    }
  }
  if (number.digits <= 2) {
    if (current_year_serial(current_year, month, number.value, &serial) ||
        serial_from_parsed_ymd(number.value, 2, month, 1, &serial)) {
      out->add(serial, s);
    }
  } else if (number.digits == 4 && serial_from_parsed_ymd(number.value, 4, month, 1, &serial)) {
    out->add(serial, s);
  }
}

void ymd_date_candidates(std::string_view s, int current_year, DateCandidates* out) noexcept {
  const LocaleFacts& facts = locale_facts(current_eval_profile());
  double serial = 0.0;
  std::string_view rest;
  if (facts.japanese_era) {
    const ShapeMatch era = parse_era_text(s, &serial, &rest);
    if (era == ShapeMatch::kSuccess) {
      out->add(serial, rest, true);
    }
    if (era != ShapeMatch::kNoMatch) {
      return;
    }
  }
  if (s.empty()) {
    return;
  }
  if (is_ascii_alpha(s[0])) {
    if (facts.english_month_names) {
      month_first_candidates(s, current_year, out);
    }
    return;
  }
  DigitRun first{};
  if (!scan_run(&s, &first)) {
    return;
  }
  if (starts_with_utf8(s, kKanjiNen) || starts_with_utf8(s, kKanjiGatsu)) {
    if (facts.kanji_ymd_text && parse_kanji_date(first, s, current_year, &serial, &rest)) {
      out->add(serial, rest, true);
    }
    return;
  }
  if (starts_with_utf8(s, kHangulNyeon)) {
    if (facts.hangul_ymd_text && finish_hangul_ymd(s.substr(3), first, &serial, &rest)) {
      out->add(serial, rest);
    }
    return;
  }
  numeric_date_candidates(first, s, current_year, out);
  if (facts.english_month_names) {
    day_first_name_candidates(first, s, current_year, out);
  }
}

// ---------------------------------------------------------------------------
// Month-Day-Year and Day-Month-Year profiles
// ---------------------------------------------------------------------------

// Consumes a case-insensitive month name from the head of `s` and advances
// `s` past it. On success writes the 1..12 month index into `*out_month` and
// returns true. Full names are tried before the abbreviations so "June" is
// not cut to "Jun" + "e"; within each tier the locale spelling comes first.
bool parse_mmm_month(std::string_view& s, int* out_month, bool hyphenated) noexcept {
  static constexpr std::string_view kShortNames[12] = {
      "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
  };
  const LocaleFacts& facts = locale_facts(current_eval_profile());
  const bool english = facts.english_month_names;
  for (int tier = 0; tier < 2; ++tier) {
    const bool english_tier = english || (hyphenated && tier == 1 && facts.hyphen_english_months);
    const auto& local_names = tier == 0 ? facts.month_long : facts.month_short;
    const auto* english_names = tier == 0 ? kEnglishMonths : kShortNames;
    for (std::size_t i = 0; i < 12; ++i) {
      const std::string_view candidates[2] = {local_names[i], english_tier ? english_names[i] : std::string_view()};
      for (const std::string_view name : candidates) {
        if (!name.empty() && strings::case_insensitive_starts_with(s, name)) {
          s.remove_prefix(name.size());
          *out_month = static_cast<int>(i) + 1;
          return true;
        }
      }
    }
  }
  return false;
}

// Parses numeric three-component dates. A 3+ digit leading component is
// always the year; otherwise the locale's date order places it. The
// tri-state result prevents an invalid, already-recognised date from being
// retried as another shape.
ShapeMatch parse_ordered_ymd_text(std::string_view s, double* out_serial, std::string_view* rest) noexcept {
  int first = 0;
  const std::size_t first_digits = scan_digits(s, 4, &first);
  if (first_digits == 0) {
    return ShapeMatch::kNoMatch;
  }
  const LocaleFacts& facts = locale_facts(current_eval_profile());
  const auto is_sep = [dotted = facts.dotted_date](char c) { return is_date_sep(c) || (dotted && c == '.'); };
  if (!consume_spaced_sep(&s, is_sep)) {
    return ShapeMatch::kNoMatch;
  }
  int second = 0;
  if (scan_digits(s, 2, &second) == 0) {
    return ShapeMatch::kNoMatch;
  }
  if (!consume_spaced_sep(&s, is_sep)) {
    // This is the year-less M/D shape; let parse_md_text handle it.
    return ShapeMatch::kNoMatch;
  }
  int third = 0;
  const std::size_t third_digits = scan_digits(s, 4, &third);
  if (third_digits == 0) {
    return ShapeMatch::kInvalid;
  }

  int year = 0;
  int month = 0;
  int day = 0;
  std::size_t year_digits = 0;
  if (first_digits >= 3) {
    year = first;
    year_digits = first_digits;
    month = second;
    day = third;
  } else if (facts.date_order == DateOrder::kDMY) {
    year = third;
    year_digits = third_digits;
    month = second;
    day = first;
  } else {
    year = third;
    year_digits = third_digits;
    month = first;
    day = second;
  }
  if (!serial_from_parsed_ymd(year, year_digits, month, day, out_serial)) {
    return ShapeMatch::kInvalid;
  }
  *rest = s;
  return ShapeMatch::kSuccess;
}

// Parses the year-less `M/D` / `M-D` shapes (`D/M` in day-first locales),
// using `current_year` as the year. Reached only after
// `parse_ordered_ymd_text` has failed, so a genuine 3-component date is never
// reinterpreted as month/day.
bool parse_md_text(std::string_view s, int current_year, double* out_serial, std::string_view* rest) noexcept {
  const LocaleFacts& facts = locale_facts(current_eval_profile());
  int first = 0;
  if (scan_digits(s, 2, &first) == 0) {
    return false;
  }
  bool dotted_form = false;
  if (!s.empty() && is_date_sep(s[0])) {
    s.remove_prefix(1);
  } else if (facts.dotted_date && !s.empty() && s[0] == '.') {
    s.remove_prefix(1);
    dotted_form = true;
  } else {
    return false;
  }
  const bool day_first = facts.date_order == DateOrder::kDMY;
  const std::string_view after_sep = s;
  int second = 0;
  if (scan_digits(s, 2, &second) > 0) {
    const int month = day_first ? second : first;
    const int day = day_first ? first : second;
    if (current_year_serial(current_year, month, day, out_serial)) {
      *rest = s;
      return true;
    }
  }
  // A dotted day-first text that is not a day.month reads as month.year.
  if (dotted_form && day_first) {
    std::string_view tail = after_sep;
    int year = 0;
    const std::size_t year_digits = scan_digits(tail, 4, &year);
    if (year_digits > 0 && serial_from_parsed_ymd(year, year_digits, first, 1, out_serial)) {
      *rest = tail;
      return true;
    }
  }
  return false;
}

// Scans the trailing 1..4 digit year token and resolves the date.
bool finish_with_year_token(std::string_view s, int month, int day, double* out_serial,
                            std::string_view* rest) noexcept {
  int year = 0;
  const std::size_t year_digits = scan_digits(s, 4, &year);
  if (year_digits == 0) {
    return false;
  }
  if (!serial_from_parsed_ymd(year, year_digits, month, day, out_serial)) {
    return false;
  }
  *rest = s;
  return true;
}

// Parses the `d-mmm-yyyy` / `d mmm yyyy` / `d/mmm/yyyy` shapes. The separator
// must match on both sides of the month word; purely numeric d-m-yyyy is not
// accepted here.
bool parse_dmy_mmm_text(std::string_view s, double* out_serial, std::string_view* rest) noexcept {
  int day = 0;
  if (scan_digits(s, 2, &day) == 0 || s.empty()) {
    return false;
  }
  // A dotted-date locale also writes `15. März 2024`: a dot after the day,
  // blanks around the month word.
  const char lead = s[0];
  const bool dot_form = lead == '.' && locale_facts(current_eval_profile()).dotted_date;
  if (lead != '-' && lead != '/' && lead != ' ' && !dot_form) {
    return false;
  }
  const char sep = dot_form ? ' ' : lead;
  s.remove_prefix(1);
  if (dot_form) {
    skip_blanks(&s);
  }
  int month = 0;
  if (!parse_mmm_month(s, &month, lead == '-')) {
    return false;
  }
  if (s.empty() || s[0] != sep) {
    return false;
  }
  s.remove_prefix(1);
  return finish_with_year_token(s, month, day, out_serial, rest);
}

// Parses `MMM D, YYYY`: blanks after the month word, a mandatory comma after
// the day, then at least one blank before the year.
bool parse_mmm_d_yyyy_text(std::string_view s, double* out_serial, std::string_view* rest) noexcept {
  // The month-first spelling is not a day-first locale's form.
  if (locale_facts(current_eval_profile()).date_order == DateOrder::kDMY) {
    return false;
  }
  int month = 0;
  if (!parse_mmm_month(s, &month, /*hyphenated=*/false) || !skip_blanks(&s)) {
    return false;
  }
  int day = 0;
  if (scan_digits(s, 2, &day) == 0 || s.empty() || s[0] != ',') {
    return false;
  }
  s.remove_prefix(1);
  if (!skip_blanks(&s)) {
    return false;
  }
  return finish_with_year_token(s, month, day, out_serial, rest);
}

void ordered_date_candidates(std::string_view s, int current_year, DateCandidates* out) noexcept {
  double serial = 0.0;
  std::string_view rest;
  const ShapeMatch numeric = parse_ordered_ymd_text(s, &serial, &rest);
  if (numeric == ShapeMatch::kSuccess) {
    out->add(serial, rest);
    return;
  }
  if (numeric == ShapeMatch::kInvalid) {
    return;
  }
  if (parse_dmy_mmm_text(s, &serial, &rest) || parse_mmm_d_yyyy_text(s, &serial, &rest) ||
      parse_md_text(s, current_year, &serial, &rest)) {
    out->add(serial, rest);
  }
}

void date_candidates(std::string_view s, int current_year, DateCandidates* out) noexcept {
  if (locale_facts(current_eval_profile()).date_order == DateOrder::kYMD) {
    ymd_date_candidates(s, current_year, out);
  } else {
    ordered_date_candidates(s, current_year, out);
  }
}

// ---------------------------------------------------------------------------
// Time
// ---------------------------------------------------------------------------

// One time field: up to two digits, or three / four digits worth at least 100.
bool scan_time_field(std::string_view* s, int* out) noexcept {
  DigitRun run{};
  if (!scan_run(s, &run) || run.digits > 4 || (run.digits > 2 && run.value < 100)) {
    return false;
  }
  *out = run.value;
  return true;
}

// Milliseconds of a fraction's digits, rounded half up on the decimal digits
// (`.9995` is 1000).
int round_millis(std::string_view digits) noexcept {
  int millis = 0;
  for (std::size_t i = 0; i < 3; ++i) {
    millis = millis * 10 + (i < digits.size() ? digits[i] - '0' : 0);
  }
  return digits.size() > 3 && digits[3] >= '5' ? millis + 1 : millis;
}

// Parses the kanji time form `H時M分[S秒]` (h <= 23, m <= 59, s <= 59). It
// takes no fraction, marker or trailing separator.
bool parse_kanji_time_text(std::string_view s, double* out_days, std::string_view* rest) noexcept {
  DigitRun hour{};
  DigitRun minute{};
  DigitRun second{0, 0};
  if (!scan_run_with_unit(&s, kKanjiJi, &hour) || !scan_run_with_unit(&s, kKanjiFun, &minute)) {
    return false;
  }
  if (!s.empty() && is_digit(s[0]) && !scan_run_with_unit(&s, kKanjiByou, &second)) {
    return false;
  }
  if (hour.value > 23 || minute.value > 59 || second.value > 59) {
    return false;
  }
  const double seconds = hour.value * 3600.0 + minute.value * 60.0 + second.value;
  *out_days = seconds / static_cast<double>(date_time::kSecondsPerDay);
  *rest = s;
  return true;
}

// Consumes a 12-hour marker after at least one blank: `A`, `P`, `AM` or `PM`
// in any case as a whole word, or the locale's own spelling.
bool scan_meridiem(std::string_view* s, bool* pm) noexcept {
  std::string_view tail = *s;
  if (!skip_blanks(&tail) || tail.empty()) {
    return false;
  }
  const LocaleFacts& facts = locale_facts(current_eval_profile());
  for (const bool is_pm : {false, true}) {
    const std::string_view name = is_pm ? facts.pm_name : facts.am_name;
    if (name == (is_pm ? "PM" : "AM") || tail.size() < name.size()) {
      continue;
    }
    if (strings::case_insensitive_eq(tail.substr(0, name.size()), name) &&
        (tail.size() == name.size() || tail[name.size()] == ' ')) {
      *pm = is_pm;
      *s = tail.substr(name.size());
      return true;
    }
  }
  std::size_t letters = 0;
  while (letters < tail.size() && is_ascii_alpha(tail[letters])) {
    ++letters;
  }
  const char c0 = ascii_upper(tail[0]);
  if ((c0 != 'A' && c0 != 'P') || (letters != 1 && (letters != 2 || ascii_upper(tail[1]) != 'M'))) {
    return false;
  }
  *pm = c0 == 'P';
  *s = tail.substr(letters);
  return true;
}

// Parses `h`, `h:m`, `h:m:s`, `h:m:s.f`, `m:s.f`, `h:` and `h:m.F:x` (F read
// as the seconds, its fraction kept, x ignored), then an optional 12-hour
// marker and one trailing `-`, `/`, `:` or `.`. Blanks may surround every
// separator. A field is at most two digits, or three or four worth at least
// 100; at most one field may pass its nominal range (23, 59, 59), none with a
// marker. `*rest` starts after the token.
bool parse_time_text(std::string_view s, double* out_days, std::string_view* rest) noexcept {
  const LocaleFacts& facts = locale_facts(current_eval_profile());
  int fields[3] = {0, 0, 0};
  std::size_t count = 0;
  if (!scan_time_field(&s, &fields[0])) {
    return false;
  }
  count = 1;
  while (count < 3) {
    std::string_view t = s;
    if (!consume_spaced_sep(&t, is_time_sep) || t.empty() || !is_digit(t[0])) {
      break;
    }
    if (!scan_time_field(&t, &fields[count])) {
      return false;
    }
    s = t;
    ++count;
  }
  int millis = 0;
  bool fraction = false;
  // A separator already closed the token.
  bool closed = false;
  if (count >= 2 && facts.fractional_seconds) {
    std::string_view t = s;
    skip_blanks(&t);
    if (!t.empty() && t[0] == facts.decimal_separator) {
      t.remove_prefix(1);
      std::string_view u = t;
      skip_blanks(&u);
      std::size_t n = 0;
      while (n < u.size() && is_digit(u[n])) {
        ++n;
      }
      if (n > 0) {
        const std::string_view digits = u.substr(0, n);
        u.remove_prefix(n);
        millis = round_millis(digits);
        fraction = true;
        if (count == 2) {
          std::string_view x = u;
          int ignored = 0;
          if (consume_spaced_sep(&x, is_time_sep) && !x.empty() && is_digit(x[0])) {
            std::string_view f = digits;
            if (!scan_time_field(&f, &fields[2]) || !scan_time_field(&x, &ignored)) {
              return false;
            }
            u = x;
          } else {
            fields[2] = fields[1];
            fields[1] = fields[0];
            fields[0] = 0;
          }
          count = 3;
        }
        s = u;
      } else if (facts.decimal_separator == '.') {
        // A bare `.` is the trailing separator.
        s = t;
        closed = true;
      }
    }
  }
  bool pm = false;
  const bool marker = !closed && scan_meridiem(&s, &pm);
  bool trailing_colon = false;
  if (!closed) {
    std::string_view t = s;
    const bool spaced = skip_blanks(&t);
    if (!t.empty()) {
      const char c = t[0];
      bool allowed = c == '-' || c == '/' || c == ':';
      if (c == '.') {
        allowed = marker ? (spaced ? facts.meridiem_dot_spaced : facts.meridiem_dot_attached) : !fraction;
      }
      if (allowed) {
        s = t.substr(1);
        trailing_colon = c == ':';
      }
    }
  }
  // A bare hour needs a marker or a colon (`6 PM`, `1:`).
  if (count == 1 && !marker && !trailing_colon) {
    return false;
  }
  int hour = fields[0];
  const int minute = fields[1];
  const int second = fields[2];
  if (marker) {
    if (hour > 12 || minute > 59 || second > 59) {
      return false;
    }
    if (hour == 12 && !pm) {
      hour = 0;
    } else if (pm && hour < 12) {
      hour += 12;
    }
  } else if ((hour > 23 ? 1 : 0) + (minute > 59 ? 1 : 0) + (second > 59 ? 1 : 0) > 1) {
    return false;
  }
  const double seconds = hour * 3600.0 + minute * 60.0 + second + millis / 1000.0;
  *out_days = seconds / static_cast<double>(date_time::kSecondsPerDay);
  *rest = s;
  return true;
}

// The numeric tail Excel ignores after a date-first date+time: one to four
// runs of at most four digits, separated (and optionally led and closed) by
// blanks or a single `/` / `-` with blanks around it.
bool numeric_tail_ok(std::string_view s) noexcept {
  const auto take_gap = [&s] {
    const std::size_t before = s.size();
    skip_blanks(&s);
    if (!s.empty() && is_date_sep(s[0])) {
      s.remove_prefix(1);
      skip_blanks(&s);
    }
    return s.size() != before;
  };
  take_gap();
  for (int runs = 1; runs <= 4; ++runs) {
    DigitRun run{};
    if (!scan_run(&s, &run) || run.digits > 4) {
      return false;
    }
    if (s.empty() || (take_gap() && s.empty())) {
      return true;
    }
  }
  return false;
}

bool recognise(std::string_view s, int current_year, DateTimeText* out) noexcept {
  const LocaleFacts& facts = locale_facts(current_eval_profile());
  {
    std::string_view probe = s;
    DigitRun hour{};
    if (facts.kanji_time_text && scan_run(&probe, &hour) && starts_with_utf8(probe, kKanjiJi)) {
      double days = 0.0;
      std::string_view rest;
      if (!parse_kanji_time_text(s, &days, &rest) || !rest.empty()) {
        return false;
      }
      *out = DateTimeText{days, days, false};
      return true;
    }
  }
  DateCandidates dates;
  date_candidates(s, current_year, &dates);
  for (std::size_t i = 0; i < dates.count; ++i) {
    const DateCandidate& date = dates.items[i];
    if (date.rest.empty()) {
      *out = DateTimeText{date.serial, 0.0, true};
      return true;
    }
    std::string_view rest = date.rest;
    double days = 0.0;
    std::string_view tail;
    if (date.standalone || !skip_blanks(&rest) || !parse_time_text(rest, &days, &tail) ||
        (!tail.empty() && !numeric_tail_ok(tail))) {
      continue;
    }
    *out = DateTimeText{date.serial + days, days, true};
    return true;
  }
  double days = 0.0;
  std::string_view rest;
  if (!parse_time_text(s, &days, &rest)) {
    return false;
  }
  if (rest.empty()) {
    *out = DateTimeText{days, days, false};
    return true;
  }
  if (!skip_blanks(&rest)) {
    return false;
  }
  DateCandidates after;
  date_candidates(rest, current_year, &after);
  for (std::size_t i = 0; i < after.count; ++i) {
    if (after.items[i].rest.empty() && !after.items[i].standalone) {
      *out = DateTimeText{after.items[i].serial + days, days, true};
      return true;
    }
  }
  return false;
}

}  // namespace

bool parse_date_time_text(std::string_view text, TextEdges edges, int current_year, DateTimeText* out) noexcept {
  if (numeric_text_too_long(text)) {
    return false;
  }
  if (edges == TextEdges::kTrim) {
    while (!text.empty() && text.front() == ' ') {
      text.remove_prefix(1);
    }
  }
  const std::string folded = fold_date_text(text);
  std::string_view s(folded);
  while (!s.empty() && s.back() == ' ') {
    s.remove_suffix(1);
  }
  if (s.empty() || s.front() == ' ') {
    return false;
  }
  return recognise(s, current_year, out);
}

}  // namespace date_parse
}  // namespace eval
}  // namespace formulon
