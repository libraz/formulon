//
// Implementation of the date / time text recogniser declared in
// `date_text_parse.h`.
//
// The text is folded once (decimal digit scripts and full-width letters; in
// CJK-width profiles also full-width ASCII and the ideographic space), then
// read as a date, a date followed by a time and an ignored numeric tail, a
// time, or a time followed by a date. Each date reader returns every reading
// of the head of the text, longest first, so the composition can back off
// from `1 Jan 1` to `1 Jan` when a time follows. One grammar serves every
// profile; the locale facts pick the field order, the separator set, the
// month names and the kanji forms. Kanji, era and kanji-time forms stand
// alone and never compose.

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
#include "utils/text_ops.h"

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
constexpr char kHangulSi[4] = {'\xEC', '\x8B', '\x9C', '\0'};     // 시 hour
constexpr char kHangulBun[4] = {'\xEB', '\xB6', '\x84', '\0'};    // 분 minute
constexpr char kHangulCho[4] = {'\xEC', '\xB4', '\x88', '\0'};    // 초 second

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

const LocaleFacts& profile_facts() noexcept {
  return locale_facts(current_eval_profile());
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

constexpr std::string_view kIdeographicSpace = "\xE3\x80\x80";
constexpr std::string_view kFullwidthComma = "\xEF\xBC\x8C";

// Skips U+0020 blanks and U+3000 (left unfolded outside CJK-width profiles);
// true when at least one was skipped.
bool skip_wide_blanks(std::string_view* s) noexcept {
  const std::size_t before = s->size();
  for (;;) {
    if (!s->empty() && (*s)[0] == ' ') {
      s->remove_prefix(1);
    } else if (s->substr(0, kIdeographicSpace.size()) == kIdeographicSpace) {
      s->remove_prefix(kIdeographicSpace.size());
    } else {
      return s->size() != before;
    }
  }
}

bool all_wide_blanks(std::string_view s) noexcept {
  skip_wide_blanks(&s);
  return s.empty();
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

// `/` and `-`, plus `.` where the profile writes dotted dates.
bool is_date_sep(char c) noexcept {
  return c == '-' || c == '/' || (c == '.' && profile_facts().dotted_date);
}

bool is_time_sep(char c) noexcept {
  return c == ':';
}

// Returns true iff `s` begins with the given 3-byte UTF-8 sequence.
bool starts_with_utf8(std::string_view s, const char (&expected)[4]) noexcept {
  return s.size() >= 3 && s[0] == expected[0] && s[1] == expected[1] && s[2] == expected[2];
}

// Folds what the recogniser reads as ASCII: the decimal digit scripts and
// the full-width letters everywhere; where the profile reads CJK widths, also
// the other full-width ASCII forms U+FF01..U+FF5E and the ideographic space
// (a blank).
std::string fold_date_text(std::string_view s) {
  const std::string scripts = fold_script_digits(s);
  const bool cjk_width = profile_facts().fullwidth_numeric_text;
  std::string out;
  out.reserve(scripts.size());
  for (std::size_t i = 0; i < scripts.size();) {
    if (i + 2 < scripts.size()) {
      const auto b0 = static_cast<unsigned char>(scripts[i]);
      const auto b1 = static_cast<unsigned char>(scripts[i + 1]);
      const auto b2 = static_cast<unsigned char>(scripts[i + 2]);
      if (cjk_width && b0 == 0xE3 && b1 == 0x80 && b2 == 0x80) {
        out.push_back(' ');
        i += 3;
        continue;
      }
      if (b0 == 0xEF && ((b1 == 0xBC && b2 >= 0x81) || (b1 == 0xBD && b2 <= 0x9E))) {
        const unsigned cp = 0xF000u | ((b1 & 0x3Fu) << 6) | (b2 & 0x3Fu);
        const char ascii = static_cast<char>(cp - 0xFEE0u);
        if (cjk_width || is_ascii_alpha(ascii)) {
          out.push_back(ascii);
          i += 3;
          continue;
        }
      }
    }
    out.push_back(scripts[i]);
    ++i;
  }
  return out;
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
// already scanned; `forms` limits which are dates. Fields take any number of
// leading zeros; the year is read by value, so `026年` is 2026.
bool parse_kanji_date(DigitRun first, std::string_view s, int current_year, KanjiDateText forms, double* out_serial,
                      std::string_view* rest) noexcept {
  DigitRun month{};
  DigitRun day{};
  const bool all_forms = forms == KanjiDateText::kAll;
  if (starts_with_utf8(s, kKanjiNen)) {
    s.remove_prefix(3);
    if (!scan_run_with_unit(&s, kKanjiGatsu, &month)) {
      return false;
    }
    const std::size_t year_digits = first.value <= 99 ? 2 : 4;
    if (s.empty() || !is_digit(s[0])) {
      *rest = s;
      return all_forms && serial_from_parsed_ymd(first.value, year_digits, month.value, 1, out_serial);
    }
    if (!scan_run_with_unit(&s, kKanjiNichi, &day)) {
      return false;
    }
    *rest = s;
    return serial_from_parsed_ymd(first.value, year_digits, month.value, day.value, out_serial);
  }
  s.remove_prefix(3);
  if (!all_forms || !scan_run_with_unit(&s, kKanjiNichi, &day)) {
    return false;
  }
  *rest = s;
  return current_year_serial(current_year, first.value, day.value, out_serial);
}

// One U+0020, then a digit run closed by the hangul `unit`.
bool scan_spaced_hangul_field(std::string_view* s, const char (&unit)[4], DigitRun* out) noexcept {
  if (s->size() < 2 || (*s)[0] != ' ' || !is_digit((*s)[1])) {
    return false;
  }
  s->remove_prefix(1);
  return scan_run_with_unit(s, unit, out);
}

// Finishes `<y>년 <m>월 <d>일` after the year unit: exactly one blank between
// the parts, fields with any leading zeros. A year below 100 is the year in
// 4263..4362 ending in those digits; 1900..9999 is the year itself.
bool finish_hangul_ymd(std::string_view s, DigitRun year, double* out_serial, std::string_view* rest) noexcept {
  DigitRun month{};
  DigitRun day{};
  if (!scan_spaced_hangul_field(&s, kHangulWol, &month) || !scan_spaced_hangul_field(&s, kHangulIl, &day)) {
    return false;
  }
  int full_year = year.value;
  if (year.value < 100) {
    full_year = year.value >= 63 ? 4200 + year.value : 4300 + year.value;
  } else if (year.value < 1900) {
    return false;
  }
  if (!serial_from_parsed_ymd(full_year, 4, month.value, day.value, out_serial)) {
    return false;
  }
  *rest = s;
  return true;
}

// A pair `a sep b`. Year-Month-Day profiles read m/d against the current year
// when `a` is at most 12, d/m when it is at most 31 (falling back to y/m when
// that is no day), and y/m above that or for a four-digit `a`. The others
// read their own day/month order against the current year, else m/y, where
// `b` is a one-, two- or four-digit year (day-first profiles take no year 0).
bool pair_serial(DigitRun a, DigitRun b, int current_year, double* out) noexcept {
  const DateOrder order = profile_facts().date_order;
  if (order == DateOrder::kYMD) {
    if (b.digits > 2) {
      return false;
    }
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
  if (a.digits > 2) {
    return false;
  }
  const bool day_first = order == DateOrder::kDMY;
  if (b.digits <= 2 &&
      current_year_serial(current_year, day_first ? b.value : a.value, day_first ? a.value : b.value, out)) {
    return true;
  }
  if (day_first && b.value == 0) {
    return false;
  }
  return written_year_serial(b, a.value, 1, out);
}

// A triple. A four-digit lead is the year in every profile; otherwise the
// profile's order places the fields, with no other reading when that is no
// date. Month and day take one or two digits.
bool triple_serial(DigitRun a, DigitRun b, DigitRun c, double* out) noexcept {
  const DateOrder order = profile_facts().date_order;
  if (b.digits > 2) {
    return false;
  }
  if (a.digits == 4 || order == DateOrder::kYMD) {
    return c.digits <= 2 && written_year_serial(a, b.value, c.value, out);
  }
  if (a.digits > 2) {
    return false;
  }
  const bool day_first = order == DateOrder::kDMY;
  return written_year_serial(c, day_first ? b.value : a.value, day_first ? a.value : b.value, out);
}

// Numeric pairs and triples, `first` already scanned. The date separators mix
// freely, with blanks around each.
void numeric_date_candidates(DigitRun first, std::string_view s, int current_year, DateCandidates* out) noexcept {
  DigitRun second{};
  if (!consume_spaced_sep(&s, is_date_sep) || !scan_run(&s, &second)) {
    return;
  }
  double serial = 0.0;
  std::string_view t = s;
  DigitRun third{};
  if (consume_spaced_sep(&t, is_date_sep) && scan_run(&t, &third)) {
    if (triple_serial(first, second, third, &serial)) {
      out->add(serial, t);
    }
    return;
  }
  if (pair_serial(first, second, current_year, &serial)) {
    out->add(serial, s);
  }
}

// `/M/D` where an empty leading field is the current year.
void leading_separator_candidates(std::string_view s, int current_year, DateCandidates* out) noexcept {
  DigitRun month{};
  DigitRun day{};
  double serial = 0.0;
  if (scan_run(&s, &month) && month.digits <= 2 && consume_spaced_sep(&s, is_date_sep) && scan_run(&s, &day) &&
      day.digits <= 2 && current_year_serial(current_year, month.value, day.value, &serial)) {
    out->add(serial, s);
  }
}

// A letter of a month name: ASCII, Latin-1 / Latin Extended or Cyrillic.
bool is_name_letter(std::uint32_t cp) noexcept {
  return (cp < 0x80u && is_ascii_alpha(static_cast<char>(cp))) ||
         (cp >= 0xC0u && cp <= 0x24Fu && cp != 0xD7u && cp != 0xF7u) || (cp >= 0x400u && cp <= 0x4FFu);
}

std::uint32_t name_letter_lower(std::uint32_t cp) noexcept {
  if ((cp >= 'A' && cp <= 'Z') || (cp >= 0xC0u && cp <= 0xDEu && cp != 0xD7u) || (cp >= 0x410u && cp <= 0x42Fu)) {
    return cp + 0x20u;
  }
  return cp >= 0x400u && cp <= 0x40Fu ? cp + 0x50u : cp;
}

// Byte length of the run of name letters heading `s`; `*letters` gets their count.
std::size_t name_token(std::string_view s, std::size_t* letters) noexcept {
  std::size_t i = 0;
  std::size_t n = 0;
  while (i < s.size()) {
    std::size_t len = 0;
    if (!is_name_letter(decode_utf8_step(s, i, &len))) {
      break;
    }
    i += len;
    ++n;
  }
  *letters = n;
  return i;
}

// True when `token` begins `name`, letter by letter in either case.
bool token_prefixes_name(std::string_view token, std::string_view name) noexcept {
  std::size_t i = 0;
  std::size_t j = 0;
  while (i < token.size()) {
    if (j >= name.size()) {
      return false;
    }
    std::size_t a = 0;
    std::size_t b = 0;
    if (name_letter_lower(decode_utf8_step(token, i, &a)) != name_letter_lower(decode_utf8_step(name, j, &b))) {
      return false;
    }
    i += a;
    j += b;
  }
  return true;
}

// Reads a month name: the run of letters heading `s`, at least three of them,
// matched as the profile's month-name rule says (the first month in table
// order wins); else a whole token equal to one of the locale's extra
// spellings; else, outside the English-prefix rule, one of the locale's own
// full or abbreviated spellings.
bool scan_month_word(std::string_view* s, int* month) noexcept {
  const LocaleFacts& facts = profile_facts();
  std::size_t letters = 0;
  const std::size_t bytes = name_token(*s, &letters);
  const std::string_view token = s->substr(0, bytes);
  for (std::size_t i = 0; letters >= 3 && i < 12; ++i) {
    bool match = false;
    switch (facts.month_name_text) {
      case MonthNameText::kEnglishPrefix:
        match = token_prefixes_name(token, kEnglishMonths[i]);
        break;
      case MonthNameText::kEnglishExact:
        match = strings::case_insensitive_eq(token, kEnglishMonths[i]) ||
                strings::case_insensitive_eq(token, kEnglishMonths[i].substr(0, 3));
        break;
      case MonthNameText::kLocalPrefix:
        match = token_prefixes_name(token, facts.month_long[i]);
        break;
    }
    if (match) {
      *month = static_cast<int>(i) + 1;
      s->remove_prefix(bytes);
      return true;
    }
  }
  for (std::size_t i = 0; letters > 0 && i < 12; ++i) {
    const std::string_view extra = facts.month_extra_names[i];
    if (!extra.empty() && extra.size() == bytes && token_prefixes_name(token, extra)) {
      *month = static_cast<int>(i) + 1;
      s->remove_prefix(bytes);
      return true;
    }
  }
  if (facts.month_name_text == MonthNameText::kEnglishPrefix) {
    return false;
  }
  for (const auto* names : {&facts.month_long, &facts.month_short}) {
    for (std::size_t i = 0; i < 12; ++i) {
      const std::string_view name = (*names)[i];
      std::size_t more = 0;
      if (!name.empty() && strings::case_insensitive_starts_with(*s, name) &&
          name_token(s->substr(name.size()), &more) == 0) {
        *month = static_cast<int>(i) + 1;
        s->remove_prefix(name.size());
        return true;
      }
    }
  }
  return false;
}

// Between a month name and a number: nothing, blanks, or one date separator
// with blanks around it. `wide` also takes U+3000 as a blank.
void skip_name_gap(std::string_view* s, bool wide) noexcept {
  const auto blanks = [wide](std::string_view* t) { return wide ? skip_wide_blanks(t) : skip_blanks(t); };
  blanks(s);
  if (!s->empty() && is_date_sep((*s)[0])) {
    s->remove_prefix(1);
    blanks(s);
  }
}

// Between the day and the year of `MMM D, Y`: blanks, a comma (full-width
// too) and at least one blank; or a gap holding a U+3000.
bool skip_year_comma(std::string_view* s) noexcept {
  std::string_view t = *s;
  skip_wide_blanks(&t);
  const bool ideographic = s->substr(0, s->size() - t.size()).find(kIdeographicSpace) != std::string_view::npos;
  if (!t.empty() && t[0] == ',') {
    t.remove_prefix(1);
  } else if (t.substr(0, kFullwidthComma.size()) == kFullwidthComma) {
    t.remove_prefix(kFullwidthComma.size());
  } else if (!ideographic) {
    return false;
  } else {
    *s = t;
    return true;
  }
  if (!skip_wide_blanks(&t)) {
    return false;
  }
  *s = t;
  return true;
}

// `D MMM [Y]`, `first` already scanned. The day must exist; there is no
// fallback to a year. Day-first profiles also take `D MMM, Y`, and U+3000 in
// the gaps of a form with a year.
void day_first_name_candidates(DigitRun day, std::string_view s, int current_year, DateCandidates* out) noexcept {
  if (day.digits > 2) {
    return;
  }
  const bool day_first = profile_facts().date_order == DateOrder::kDMY;
  double serial = 0.0;
  int month = 0;
  std::string_view t = s;
  skip_name_gap(&t, day_first);
  if (scan_month_word(&t, &month)) {
    bool gap = true;
    if (day_first && !t.empty() && t[0] == ',') {
      t.remove_prefix(1);
      gap = skip_wide_blanks(&t);
    } else {
      skip_name_gap(&t, day_first);
    }
    DigitRun year{};
    if (gap && scan_run(&t, &year) && written_year_serial(year, month, day.value, &serial)) {
      out->add(serial, t);
    }
  }
  t = s;
  skip_name_gap(&t, false);
  if (scan_month_word(&t, &month) && current_year_serial(current_year, month, day.value, &serial)) {
    out->add(serial, t);
  }
}

// `MMM D, Y`, `MMM D` and `MMM Y`. A one- or two-digit number that is no day
// of the month in the current year is a two-digit year (`Jan 32` is
// 1932-01-01); four digits are a year. Day-first profiles read `MMM N` as
// the year alone and have no `MMM D, Y`.
void month_first_candidates(std::string_view s, int current_year, DateCandidates* out) noexcept {
  int month = 0;
  if (!scan_month_word(&s, &month)) {
    return;
  }
  const bool day_first = profile_facts().date_order == DateOrder::kDMY;
  double serial = 0.0;
  if (!day_first) {
    std::string_view t = s;
    skip_name_gap(&t, true);
    DigitRun day{};
    DigitRun year{};
    if (scan_run(&t, &day) && day.digits <= 2 && skip_year_comma(&t) && scan_run(&t, &year) &&
        written_year_serial(year, month, day.value, &serial)) {
      out->add(serial, t);
    }
  }
  skip_name_gap(&s, false);
  DigitRun number{};
  if (!scan_run(&s, &number)) {
    return;
  }
  if (day_first || number.digits > 2) {
    if (written_year_serial(number, month, 1, &serial)) {
      out->add(serial, s);
    }
    return;
  }
  if (current_year_serial(current_year, month, number.value, &serial) ||
      serial_from_parsed_ymd(number.value, 2, month, 1, &serial)) {
    out->add(serial, s);
  }
}

void date_candidates(std::string_view s, int current_year, DateCandidates* out) noexcept {
  const LocaleFacts& facts = profile_facts();
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
  if (facts.leading_separator_current_year && s[0] == '/') {
    leading_separator_candidates(s.substr(1), current_year, out);
    return;
  }
  DigitRun first{};
  if (!scan_run(&s, &first)) {
    month_first_candidates(s, current_year, out);
    return;
  }
  if (starts_with_utf8(s, kKanjiNen) || starts_with_utf8(s, kKanjiGatsu)) {
    if (facts.kanji_date_text != KanjiDateText::kNone &&
        parse_kanji_date(first, s, current_year, facts.kanji_date_text, &serial, &rest)) {
      out->add(serial, rest, true);
    }
    return;
  }
  if (starts_with_utf8(s, kHangulNyeon)) {
    if (facts.hangul_date_time_text && finish_hangul_ymd(s.substr(3), first, &serial, &rest)) {
      out->add(serial, rest, true);
    }
    return;
  }
  numeric_date_candidates(first, s, current_year, out);
  day_first_name_candidates(first, s, current_year, out);
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

// Byte length of the marker spelling `name` heading `s` in any case, where a
// U+202F in the spelling also matches a U+0020 blank; 0 when absent.
std::size_t match_marker_name(std::string_view s, std::string_view name) noexcept {
  constexpr std::string_view kNarrowNoBreakSpace = "\xE2\x80\xAF";
  std::size_t i = 0;
  std::size_t j = 0;
  while (j < name.size()) {
    if (name.substr(j, kNarrowNoBreakSpace.size()) == kNarrowNoBreakSpace && i < s.size() && s[i] == ' ') {
      ++i;
      j += kNarrowNoBreakSpace.size();
    } else if (i < s.size() && ascii_upper(s[i]) == ascii_upper(name[j])) {
      ++i;
      ++j;
    } else {
      return 0;
    }
  }
  return i;
}

// Parses the hangul time form `H시 M분[ S초]` (h <= 23, m <= 59, s <= 59) with
// exactly one blank between the parts. It takes no fraction, marker or
// trailing separator.
bool parse_hangul_time_text(std::string_view s, double* out_days, std::string_view* rest) noexcept {
  DigitRun hour{};
  DigitRun minute{};
  DigitRun second{0, 0};
  if (!scan_run_with_unit(&s, kHangulSi, &hour) || !scan_spaced_hangul_field(&s, kHangulBun, &minute)) {
    return false;
  }
  if (!s.empty() && !scan_spaced_hangul_field(&s, kHangulCho, &second)) {
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
// in any case as a whole word, the locale's own spelling, or one of its further
// spellings.
bool scan_meridiem(std::string_view* s, bool* pm) noexcept {
  std::string_view tail = *s;
  if (!skip_blanks(&tail) || tail.empty()) {
    return false;
  }
  const LocaleFacts& facts = profile_facts();
  for (const bool is_pm : {false, true}) {
    const std::string_view name = is_pm ? facts.pm_name : facts.am_name;
    const std::size_t len = name == (is_pm ? "PM" : "AM") ? 0 : match_marker_name(tail, name);
    if (len > 0 && (tail.size() == len || tail[len] == ' ')) {
      *pm = is_pm;
      *s = tail.substr(len);
      return true;
    }
    // The further spellings end at anything but an ASCII letter or digit (`下午.`).
    const std::string_view text_name = facts.meridiem_text_names[is_pm ? 1 : 0];
    if (!text_name.empty() && tail.substr(0, text_name.size()) == text_name &&
        (tail.size() == text_name.size() ||
         (!is_ascii_alpha(tail[text_name.size()]) && !is_digit(tail[text_name.size()])))) {
      *pm = is_pm;
      *s = tail.substr(text_name.size());
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

// Whether `c` closes a time: `-`, `/` and `:` everywhere; `.` where it is the
// decimal separator or a date separator, and `,` where it is the decimal
// separator, neither after a fraction unless a marker follows it. A `.` after
// a marker follows the profile's marker-dot facts.
bool closes_time(char c, bool marker, bool spaced, bool fraction) noexcept {
  const LocaleFacts& facts = profile_facts();
  switch (c) {
    case '-':
    case '/':
    case ':':
      return true;
    case '.':
      if (facts.decimal_separator != '.' && !facts.dotted_date) {
        return false;
      }
      return marker ? (spaced ? facts.meridiem_dot_spaced : facts.meridiem_dot_attached) : !fraction;
    case ',':
      return facts.decimal_separator == ',' && (marker || !fraction);
    default:
      return false;
  }
}

// Parses `h`, `h:m`, `h:m:s`, `h:m:s.f`, `m:s.f`, `h:` and `h:m.F:x` (F read
// as the seconds, its fraction kept, x ignored; `.` stands for the decimal
// separator), then an optional 12-hour marker and one trailing separator
// (`closes_time`). Blanks may surround every separator. A field is at most two
// digits, or three or four worth at least 100; at most one field may pass its
// nominal range (23, 59, 59), none with a marker. `*rest` starts after the
// token; `*closed` tells whether a trailing separator ended it.
bool parse_time_text(std::string_view s, double* out_days, std::string_view* rest, bool* closed_out) noexcept {
  const LocaleFacts& facts = profile_facts();
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
      } else {
        // A bare decimal separator is the trailing separator.
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
    if (!t.empty() && closes_time(t[0], marker, spaced, fraction)) {
      trailing_colon = t[0] == ':';
      s = t.substr(1);
      closed = true;
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
  *closed_out = closed;
  return true;
}

// The numeric tail Excel ignores after a date-first date+time: one to four
// runs, each shaped like a time field, separated (and optionally led and
// closed) by blanks (U+3000 included) or a single date separator with blanks
// around it.
bool numeric_tail_ok(std::string_view s) noexcept {
  const auto take_gap = [&s] {
    const std::size_t before = s.size();
    skip_wide_blanks(&s);
    if (!s.empty() && is_date_sep(s[0])) {
      s.remove_prefix(1);
      skip_wide_blanks(&s);
    }
    return s.size() != before;
  };
  take_gap();
  for (int runs = 1; runs <= 4; ++runs) {
    int ignored = 0;
    if (!scan_time_field(&s, &ignored)) {
      return false;
    }
    if (s.empty() || (take_gap() && s.empty())) {
      return true;
    }
  }
  return false;
}

bool recognise(std::string_view s, int current_year, DateTimeText* out) noexcept {
  const LocaleFacts& facts = profile_facts();
  {
    std::string_view probe = s;
    DigitRun hour{};
    const bool digits = scan_run(&probe, &hour);
    const bool kanji = digits && facts.kanji_time_text && starts_with_utf8(probe, kKanjiJi);
    const bool hangul = digits && facts.hangul_date_time_text && starts_with_utf8(probe, kHangulSi);
    if (kanji || hangul) {
      double days = 0.0;
      std::string_view rest;
      const bool parsed = kanji ? parse_kanji_time_text(s, &days, &rest) : parse_hangul_time_text(s, &days, &rest);
      if (!parsed || !rest.empty()) {
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
    bool closed = false;
    if (date.standalone || !skip_blanks(&rest) || !parse_time_text(rest, &days, &tail, &closed) ||
        (!all_wide_blanks(tail) && !numeric_tail_ok(tail))) {
      continue;
    }
    *out = DateTimeText{date.serial + days, days, true};
    return true;
  }
  double days = 0.0;
  std::string_view rest;
  bool closed = false;
  if (!parse_time_text(s, &days, &rest, &closed)) {
    return false;
  }
  if (all_wide_blanks(rest)) {
    *out = DateTimeText{days, days, false};
    return true;
  }
  // A date follows a blank gap, or directly after the time's trailing separator.
  if (!skip_wide_blanks(&rest) && !closed) {
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
