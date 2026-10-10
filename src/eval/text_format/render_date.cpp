//
// Date / time rendering for the Excel TEXT() engine. Converts the serial
// via the shared `date_time` helpers and substitutes each `y/m/d/h/s`
// token by its textual form, with ja-JP weekday / era tables for the
// localised tokens (`aaa`, `aaaa`, `g`, `e`).

#include "eval/text_format/render_date.h"

#include <array>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "eval/eval_profile_scope.h"
#include "eval/text_format/number_format_scanner.h"
#include "eval/text_format/number_format_types.h"
#include "eval/text_format/render_common.h"
#include "excel_locale.h"
#include "utils/date_time.h"
#include "utils/japanese_era.h"
#include "utils/number_text.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {
namespace {

// Entry `index` of a locale name table, or "" when `index` is out of range.
template <std::size_t N>
std::string_view name_entry(const std::array<std::string_view, N>& table, long long index) noexcept {
  return (index < 0 || index >= static_cast<long long>(N)) ? std::string_view()
                                                           : table[static_cast<std::size_t>(index)];
}

// Buddhist-era year (2024 -> 2567, text_buddhist_year_bbbb).
constexpr int kBuddhistEraOffset = 543;

// Elapsed time is zero-padded to the token width and written like General.
void append_elapsed_int_dbnum(std::string& out, long long value, std::size_t width, DbNumMode mode) {
  const std::string digits = std::to_string(value);
  for (std::size_t i = digits.size(); i < width; ++i) {
    append_digit_dbnum(out, mode, '0');
  }
  append_dbnum_positional(out, mode, digits);
}

constexpr std::size_t kMaxMeaningfulFractionDigits = 15;

std::uint64_t power10(std::size_t exponent) noexcept {
  std::uint64_t result = 1;
  for (std::size_t i = 0; i < exponent; ++i) {
    result *= 10;
  }
  return result;
}

// Japanese era classification. The boundary table and classifier live in
// `eval/japanese_era.{h,cpp}` so the pivot date-grouping path can share
// the same anchors. We re-export the type as a local alias so the rest
// of this TU keeps its idiomatic name.
using EraInfo = formulon::japanese_era::EraInfo;

const EraInfo& classify_era(int year, unsigned month, unsigned day) noexcept {
  return formulon::japanese_era::classify_era(year, month, day);
}

}  // namespace

FormatStatus render_date(const Section& section, std::string_view fmt, double serial, std::string& out, bool date1904) {
  const double max_serial =
      date1904 ? kMaxDateSerial1900 - ::formulon::date_time::kDate1904EpochGap : kMaxDateSerial1900;
  if (!std::isfinite(serial) || serial < 0.0 || serial >= max_serial + 1.0) {
    // Excel rejects out-of-range serials from TEXT.
    return FormatStatus::kOverflow;
  }

  // Decompose the time portion into an integral second and a sub-second
  // fraction. Only the fraction is scaled: scaling the complete serial by
  // 10^15 would overflow before elapsed-hour/minute/second formatting sees
  // the value.
  const double day_floor_f = std::floor(serial);
  const long long day_floor = static_cast<long long>(day_floor_f);
  const double frac_day = serial - day_floor_f;
  const double total_seconds_f = frac_day * static_cast<double>(::formulon::date_time::kSecondsPerDay);
  const double whole_seconds_f = std::floor(total_seconds_f);
  const long long whole_seconds = static_cast<long long>(whole_seconds_f);
  const double subsecond = total_seconds_f - whole_seconds_f;
  const std::size_t requested_fraction_digits =
      section.frac_sec_digits > 0 ? static_cast<std::size_t>(section.frac_sec_digits) : 0U;
  const std::size_t meaningful_fraction_digits = requested_fraction_digits < kMaxMeaningfulFractionDigits
                                                     ? requested_fraction_digits
                                                     : kMaxMeaningfulFractionDigits;
  const std::uint64_t fraction_scale = power10(meaningful_fraction_digits);
  std::uint64_t fraction_ticks =
      static_cast<std::uint64_t>(std::floor(subsecond * static_cast<double>(fraction_scale) + 0.5));
  long long rounded_seconds_in_day = whole_seconds;
  if (fraction_ticks >= fraction_scale) {
    ++rounded_seconds_in_day;
    fraction_ticks = 0;
  }

  // Calendar and ordinary time tokens share the rounded day. A carry beyond
  // 9999-12-31 is a date overflow even when the source serial itself still
  // lies within the final day's fractional range.
  const long long day_carry = rounded_seconds_in_day / ::formulon::date_time::kSecondsPerDay;
  const long long day_for_calendar = day_floor + day_carry;
  if (static_cast<double>(day_for_calendar) > max_serial) {
    return FormatStatus::kOverflow;
  }
  const long long seconds_of_day = rounded_seconds_in_day % ::formulon::date_time::kSecondsPerDay;
  const double calendar_serial = static_cast<double>(day_for_calendar);
  const ::formulon::date_time::YMD ymd = date1904 ? ::formulon::date_time::ymd_from_serial(calendar_serial, true)
                                                  : ::formulon::date_time::legacy_1900_ymd(calendar_serial);
  const int sun0 = ::formulon::date_time::weekday_sun0(calendar_serial, date1904);
  const LocaleFacts& facts = locale_facts(eval::current_eval_profile());

  // Elapsed fields use the same rounded second as ordinary h/m/s fields.
  // `day_floor` is only a few million in the supported date range, so this
  // product remains well within int64_t while retaining all day carries.
  const long long total_rounded_seconds = day_floor * ::formulon::date_time::kSecondsPerDay + rounded_seconds_in_day;

  // If AM/PM is in use, we need to know it before formatting hours.
  bool use_am_pm = false;
  for (const Token& tk : section.tokens) {
    if (tk.kind == Tok::AmPm || tk.kind == Tok::AP || tk.kind == Tok::AmPmChinese) {
      use_am_pm = true;
      break;
    }
  }

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
        append_dbnum_date_field(out, dbnum, y2, 2U);
        break;
      }
      case Tok::DateY4: {
        char buf[16];
        const int n = format_signed(buf, sizeof(buf), ymd.y, 4);
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
        append_dbnum_date_field(out, dbnum, ymd.m, 1U);
        break;
      case Tok::DateMM:
        append_dbnum_date_field(out, dbnum, ymd.m, 2U);
        break;
      case Tok::DateMMM:
        out.append(name_entry(facts.month_short, static_cast<long long>(ymd.m) - 1));
        break;
      case Tok::DateMMMM:
        out.append(name_entry(facts.month_long, static_cast<long long>(ymd.m) - 1));
        break;
      case Tok::DateMMMMM: {
        // `mmmmm` (run length >= 5) emits the first scalar of the month name.
        const std::string_view name = name_entry(facts.month_long, static_cast<long long>(ymd.m) - 1);
        if (!name.empty()) {
          out.append(name.substr(0, utf8_scalar_width(name, 0)));
        }
        break;
      }
      case Tok::DateD:
        append_dbnum_date_field(out, dbnum, ymd.d, 1U);
        break;
      case Tok::DateDD:
        append_dbnum_date_field(out, dbnum, ymd.d, 2U);
        break;
      case Tok::DateDDD:
        out.append(name_entry(facts.day_short, sun0));
        break;
      case Tok::DateDDDD:
        out.append(name_entry(facts.day_long, sun0));
        break;
      case Tok::DateAaa:
        out.append(facts.weekday_short[static_cast<std::size_t>(sun0)]);
        break;
      case Tok::DateAaaa:
        out.append(facts.weekday_long[static_cast<std::size_t>(sun0)]);
        break;
      case Tok::EraG: {
        if (!facts.japanese_era) {
          break;
        }
        const EraInfo& era = classify_era(ymd.y, ymd.m, ymd.d);
        out.append(era.roman);
        break;
      }
      case Tok::EraGG: {
        if (!facts.japanese_era) {
          break;
        }
        const EraInfo& era = classify_era(ymd.y, ymd.m, ymd.d);
        out.append(era.kanji1);
        break;
      }
      case Tok::EraGGG: {
        if (!facts.japanese_era) {
          break;
        }
        const EraInfo& era = classify_era(ymd.y, ymd.m, ymd.d);
        out.append(era.kanji2);
        break;
      }
      case Tok::DateB2: {
        const auto b2 = static_cast<unsigned>(((ymd.y + kBuddhistEraOffset) % 100 + 100) % 100);
        append_pad2_dbnum(out, b2, dbnum);
        // Under a `[DBNumN]` style the two digits are written twice (locale_tokens.dbnum_date_fields).
        if (dbnum_writes_numerals(dbnum)) {
          append_pad2_dbnum(out, b2, dbnum);
        }
        break;
      }
      case Tok::DateBlank:
        break;
      case Tok::DateB4:
        append_int_dbnum(out, static_cast<long long>(ymd.y + kBuddhistEraOffset), dbnum);
        break;
      case Tok::EraE: {
        if (!facts.japanese_era) {
          append_int_dbnum(out, static_cast<long long>(ymd.y), dbnum);
          break;
        }
        const EraInfo& era = classify_era(ymd.y, ymd.m, ymd.d);
        const int era_year = ymd.y - era.year_anchor + 1;
        if (era_year >= 0) {
          append_dbnum_date_field(out, dbnum, static_cast<unsigned>(era_year), 1U);
        } else {
          append_int_dbnum(out, static_cast<long long>(era_year), dbnum);
        }
        break;
      }
      case Tok::EraEE: {
        if (!facts.japanese_era) {
          append_int_dbnum(out, static_cast<long long>(ymd.y), dbnum);
          break;
        }
        const EraInfo& era = classify_era(ymd.y, ymd.m, ymd.d);
        const int era_year = ymd.y - era.year_anchor + 1;
        if (era_year >= 0 && era_year < 100) {
          append_dbnum_date_field(out, dbnum, static_cast<unsigned>(era_year), 2U);
        } else {
          append_int_dbnum(out, static_cast<long long>(era_year), dbnum);
        }
        break;
      }
      case Tok::DateH:
        append_dbnum_date_field(out, dbnum, hour_for_render, 1U);
        break;
      case Tok::DateHH:
        append_dbnum_date_field(out, dbnum, hour_for_render, 2U);
        break;
      case Tok::DateMin:
        append_dbnum_date_field(out, dbnum, minute, 1U);
        break;
      case Tok::DateMMMin:
        append_dbnum_date_field(out, dbnum, minute, 2U);
        break;
      case Tok::DateS:
        append_dbnum_date_field(out, dbnum, second, 1U);
        break;
      case Tok::DateSS:
        append_dbnum_date_field(out, dbnum, second, 2U);
        break;
      case Tok::DateElapsedH: {
        const long long total_hours = total_rounded_seconds / 3600;
        append_elapsed_int_dbnum(out, total_hours, tk.width, dbnum);
        break;
      }
      case Tok::DateElapsedM: {
        const long long total_minutes = total_rounded_seconds / 60;
        append_elapsed_int_dbnum(out, total_minutes, tk.width, dbnum);
        break;
      }
      case Tok::DateElapsedS: {
        append_elapsed_int_dbnum(out, total_rounded_seconds, tk.width, dbnum);
        break;
      }
      case Tok::AmPm:
        out.append(pm ? facts.pm_name : facts.am_name);
        break;
      case Tok::AP:
        out.append(pm ? "P" : "A");
        break;
      case Tok::AmPmChinese:
        out.append(pm ? "下午" : "上午");
        break;
      case Tok::FracSecDigits: {
        // Render fractional seconds at the requested precision.
        const std::size_t digits = tk.width;
        if (digits > 0) {
          out.push_back(facts.decimal_separator);
          std::string fraction = std::to_string(fraction_ticks);
          if (fraction.size() < meaningful_fraction_digits) {
            fraction.insert(0, meaningful_fraction_digits - fraction.size(), '0');
          }
          append_chars_dbnum(out, dbnum, fraction);
          for (std::size_t k = meaningful_fraction_digits; k < digits; ++k) {
            append_digit_dbnum(out, dbnum, '0');
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
      case Tok::Comma:
        out.push_back(facts.group_separator);
        break;
      case Tok::Point:
        out.push_back(facts.decimal_separator);
        break;
      default:
        break;
    }
  }
  return FormatStatus::kOk;
}

}  // namespace number_format_detail
}  // namespace text_format
}  // namespace formulon
