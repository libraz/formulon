//
// Date / time rendering for the Excel TEXT() engine. Converts the serial
// via the shared `date_time` helpers and substitutes each `y/m/d/h/s`
// token by its textual form, with ja-JP weekday / era tables for the
// localised tokens (`aaa`, `aaaa`, `g`, `e`).

#include "eval/text_format/render_date.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <string_view>

#include "eval/date_time.h"
#include "eval/japanese_era.h"
#include "eval/text_format/number_format_types.h"
#include "eval/text_format/render_common.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {
namespace {

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

void append_elapsed_int_dbnum(std::string& out, long long value, std::uint8_t width, DbNumMode mode) {
  const std::string digits = std::to_string(value);
  const std::size_t min_width = static_cast<std::size_t>(width);
  for (std::size_t i = digits.size(); i < min_width; ++i) {
    append_digit_dbnum(out, mode, '0');
  }
  append_chars_dbnum(out, mode, digits);
}

// ja-JP weekday tokens (`aaa` / `aaaa`). Index 0 = Sunday to match the
// `sun0` index used elsewhere in this file.
const char* weekday_ja_short(int sun0) noexcept {
  // 日, 月, 火, 水, 木, 金, 土 (each is a 3-byte UTF-8 code point).
  static const char* kTable[7] = {"\xE6\x97\xA5", "\xE6\x9C\x88", "\xE7\x81\xAB", "\xE6\xB0\xB4",
                                  "\xE6\x9C\xA8", "\xE9\x87\x91", "\xE5\x9C\x9F"};
  if (sun0 < 0 || sun0 > 6) {
    return "";
  }
  return kTable[sun0];
}

const char* weekday_ja_long(int sun0) noexcept {
  // <weekday>曜日 — the suffix is `\xE6\x9B\x9C\xE6\x97\xA5` (曜日).
  static const char* kTable[7] = {
      "\xE6\x97\xA5\xE6\x9B\x9C\xE6\x97\xA5",  // 日曜日
      "\xE6\x9C\x88\xE6\x9B\x9C\xE6\x97\xA5",  // 月曜日
      "\xE7\x81\xAB\xE6\x9B\x9C\xE6\x97\xA5",  // 火曜日
      "\xE6\xB0\xB4\xE6\x9B\x9C\xE6\x97\xA5",  // 水曜日
      "\xE6\x9C\xA8\xE6\x9B\x9C\xE6\x97\xA5",  // 木曜日
      "\xE9\x87\x91\xE6\x9B\x9C\xE6\x97\xA5",  // 金曜日
      "\xE5\x9C\x9F\xE6\x9B\x9C\xE6\x97\xA5",  // 土曜日
  };
  if (sun0 < 0 || sun0 > 6) {
    return "";
  }
  return kTable[sun0];
}

// Japanese era classification. The boundary table and classifier live in
// `eval/japanese_era.{h,cpp}` so the pivot date-grouping path can share
// the same anchors. We re-export the type as a local alias so the rest
// of this TU keeps its idiomatic name.
using EraInfo = formulon::eval::japanese_era::EraInfo;

const EraInfo& classify_era(int year, unsigned month, unsigned day) noexcept {
  return formulon::eval::japanese_era::classify_era(year, month, day);
}

}  // namespace

void render_date(const Section& section, std::string_view fmt, double serial, std::string& out, bool date1904) {
  if (serial < 0.0 || serial > 2958465.0) {
    // Excel rejects out-of-range serials from TEXT.
    return;
  }
  const ::formulon::eval::date_time::YMD ymd = ::formulon::eval::date_time::ymd_from_serial(serial, date1904);
  // Weekday Sunday=0..Saturday=6 computed from the civil day count.
  const std::int64_t days = ::formulon::eval::date_time::days_from_civil(ymd.y, ymd.m, ymd.d);
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

  const DbNumMode dbnum = section.dbnum_mode;
  for (std::size_t i = 0; i < section.tokens.size(); ++i) {
    const Token& tk = section.tokens[i];
    switch (tk.kind) {
      case Tok::DateY2: {
        unsigned y2 = static_cast<unsigned>(((ymd.y % 100) + 100) % 100);
        append_pad2_dbnum(out, y2, dbnum);
        break;
      }
      case Tok::DateY4: {
        char buf[16];
        const int n = std::snprintf(buf, sizeof(buf), "%04d", ymd.y);
        if (n > 0) {
          if (dbnum == DbNumMode::kNone) {
            out.append(buf, static_cast<std::size_t>(n));
          } else {
            append_chars_dbnum(out, dbnum, std::string_view(buf, static_cast<std::size_t>(n)));
          }
        }
        break;
      }
      case Tok::DateM:
        append_int_dbnum(out, static_cast<long long>(ymd.m), dbnum);
        break;
      case Tok::DateMM:
        append_pad2_dbnum(out, ymd.m, dbnum);
        break;
      case Tok::DateMMM:
        out.append(month_short(ymd.m));
        break;
      case Tok::DateMMMM:
        out.append(month_long(ymd.m));
        break;
      case Tok::DateMMMMM: {
        // `mmmmm` (run length >= 5) emits the first letter of the English
        // month name. The month-name table only contains ASCII letters, so
        // the first byte is a complete UTF-8 code point.
        const char* name = month_long(ymd.m);
        if (name[0] != '\0') {
          out.push_back(name[0]);
        }
        break;
      }
      case Tok::DateD:
        append_int_dbnum(out, static_cast<long long>(ymd.d), dbnum);
        break;
      case Tok::DateDD:
        append_pad2_dbnum(out, ymd.d, dbnum);
        break;
      case Tok::DateDDD:
        out.append(weekday_short(sun0));
        break;
      case Tok::DateDDDD:
        out.append(weekday_long(sun0));
        break;
      case Tok::DateAaa:
        out.append(weekday_ja_short(sun0));
        break;
      case Tok::DateAaaa:
        out.append(weekday_ja_long(sun0));
        break;
      case Tok::EraG: {
        const EraInfo& era = classify_era(ymd.y, ymd.m, ymd.d);
        out.append(era.roman);
        break;
      }
      case Tok::EraGG: {
        const EraInfo& era = classify_era(ymd.y, ymd.m, ymd.d);
        out.append(era.kanji1);
        break;
      }
      case Tok::EraGGG: {
        const EraInfo& era = classify_era(ymd.y, ymd.m, ymd.d);
        out.append(era.kanji2);
        break;
      }
      case Tok::EraE: {
        const EraInfo& era = classify_era(ymd.y, ymd.m, ymd.d);
        const int era_year = ymd.y - era.year_anchor + 1;
        append_int_dbnum(out, static_cast<long long>(era_year), dbnum);
        break;
      }
      case Tok::EraEE: {
        const EraInfo& era = classify_era(ymd.y, ymd.m, ymd.d);
        const int era_year = ymd.y - era.year_anchor + 1;
        if (era_year >= 0 && era_year < 100) {
          append_pad2_dbnum(out, static_cast<unsigned>(era_year), dbnum);
        } else {
          append_int_dbnum(out, static_cast<long long>(era_year), dbnum);
        }
        break;
      }
      case Tok::DateH:
        append_int_dbnum(out, static_cast<long long>(hour_for_render), dbnum);
        break;
      case Tok::DateHH:
        append_pad2_dbnum(out, hour_for_render, dbnum);
        break;
      case Tok::DateMin:
        append_int_dbnum(out, static_cast<long long>(minute), dbnum);
        break;
      case Tok::DateMMMin:
        append_pad2_dbnum(out, minute, dbnum);
        break;
      case Tok::DateS:
        append_int_dbnum(out, static_cast<long long>(second), dbnum);
        break;
      case Tok::DateSS:
        append_pad2_dbnum(out, second, dbnum);
        break;
      case Tok::DateElapsedH: {
        // Total hours since serial 0 (integer floor).
        const long long total_hours = static_cast<long long>(std::floor(serial * 24.0));
        append_elapsed_int_dbnum(out, total_hours, tk.width, dbnum);
        break;
      }
      case Tok::DateElapsedM: {
        const long long total_minutes = static_cast<long long>(std::floor(serial * 1440.0));
        append_elapsed_int_dbnum(out, total_minutes, tk.width, dbnum);
        break;
      }
      case Tok::DateElapsedS: {
        const long long total_sec = static_cast<long long>(std::floor(serial * 86400.0));
        append_elapsed_int_dbnum(out, total_sec, tk.width, dbnum);
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
            const char ch = static_cast<char>('0' + d);
            append_digit_dbnum(out, dbnum, ch);
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
      case Tok::Space:
        // `_X` underscore-skip: emit a single space placeholder.
        out.push_back(' ');
        break;
      default:
        break;
    }
  }
}

}  // namespace number_format_detail
}  // namespace text_format
}  // namespace formulon
