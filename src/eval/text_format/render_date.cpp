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
// Dangi year of the Korean calendar (2024 -> 4357, locale_tokens.lcid_calendars).
constexpr int kDangiOffset = 2333;

// `CC` calendar bytes of a `[$-CCLLLL]` tag that change the output (locale_tokens.lcid_calendars).
constexpr std::uint8_t kCalendarDefault = 0x00;
constexpr std::uint8_t kCalendarEnglishNames = 0x02;
constexpr std::uint8_t kCalendarJapanese = 0x03;
constexpr std::uint8_t kCalendarTaiwan = 0x04;
constexpr std::uint8_t kCalendarKorean = 0x05;
constexpr std::uint8_t kCalendarThai = 0x07;
// Julian and Taiwan lunisolar: Gregorian dates under the profile's own names.
constexpr std::uint8_t kCalendarJulian = 0x0D;
constexpr std::uint8_t kCalendarTaiwanLunisolar = 0x15;
constexpr std::uint16_t kEnglishLcid = 0x0409;
constexpr std::uint16_t kChineseLcid = 0x0804;
constexpr std::uint16_t kTraditionalChineseLcid = 0x0404;
constexpr std::uint16_t kThaiLcid = 0x041E;
constexpr std::string_view kGannen = "元";

using Names12 = std::array<std::string_view, 12>;
using Names7 = std::array<std::string_view, 7>;

// The name tables a section writes: the profile's, or those its tag selects.
struct DateNames {
  const Names12* month_long;
  const Names12* month_short;
  const Names7* day_long;
  const Names7* day_short;
  const Names7* weekday_long;
  const Names7* weekday_short;
  std::string_view am;
  std::string_view pm;
  // A tag language writes its AM/PM for every meridiem marker, A/P included.
  bool tag_meridiem;
};

void use_tag_names(const TagLanguage& language, DateNames* names) noexcept {
  names->month_long = &language.month_long;
  names->month_short = &language.month_short;
  names->day_long = &language.day_long;
  names->day_short = &language.day_short;
  names->weekday_long = &language.day_long;
  names->weekday_short = &language.day_short;
}

DateNames date_names(const FormatTag& tag, const LocaleFacts& facts, ExcelLocale locale) noexcept {
  DateNames names{&facts.month_long, &facts.month_short,  &facts.day_long,
                  &facts.day_short,  &facts.weekday_long, &facts.weekday_short,
                  facts.am_name,     facts.pm_name,       false};
  if (tag.language != nullptr) {
    use_tag_names(*tag.language, &names);
    names.am = tag.language->am_name;
    names.pm = tag.language->pm_name;
    names.tag_meridiem = true;
  }
  // A calendar brings its own month and day names; AM/PM stay the language's.
  switch (tag.calendar) {
    case kCalendarEnglishNames:
      use_tag_names(*tag_language_for_lcid(kEnglishLcid), &names);
      break;
    case kCalendarTaiwan: {
      const TagLanguage& months = *tag_language_for_lcid(kChineseLcid);
      use_tag_names(*tag_language_for_lcid(kTraditionalChineseLcid), &names);
      names.month_long = &months.month_long;
      names.month_short = &months.month_long;
      break;
    }
    case kCalendarKorean:
    case kCalendarJulian:
    case kCalendarTaiwanLunisolar:
      use_tag_names(native_tag_language(locale), &names);
      break;
    case kCalendarThai:
      use_tag_names(*tag_language_for_lcid(kThaiLcid), &names);
      break;
    default:
      break;
  }
  return names;
}

// Elapsed time is zero-padded to the token width and written like General.
void append_elapsed_int_dbnum(std::string& out, long long value, std::size_t width, const DbnumStyle* style) {
  const std::string digits = std::to_string(value);
  for (std::size_t i = digits.size(); i < width; ++i) {
    append_digit_dbnum(out, style, '0');
  }
  append_dbnum_positional(out, style, digits);
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
  const ExcelProfile profile = eval::current_eval_profile();
  const LocaleFacts& facts = locale_facts(profile);
  const FormatTag& tag = section.tag;
  const DateNames names = date_names(tag, facts, profile.locale);
  // `e`, `g` and `r` write the era under the Japanese calendar or a Japanese
  // language; any other tag language or calendar turns them into the year.
  const bool tag_era = tag.language_set ? tag.language != nullptr && tag.language->japanese_era : facts.japanese_era;
  const bool era_on = tag.calendar == kCalendarJapanese || (tag.calendar == kCalendarDefault && tag_era);
  const EraInfo& era = classify_era(ymd.y, ymd.m, ymd.d);
  const int era_year = ymd.y - era.year_anchor + 1;
  int calendar_year = ymd.y;
  if (tag.calendar == kCalendarKorean) {
    calendar_year += kDangiOffset;
  } else if (tag.calendar == kCalendarThai) {
    calendar_year += kBuddhistEraOffset;
  }
  // `b` is the Buddhist year unless a tag other than the Thai calendar is present.
  const int b_year = tag.calendar == kCalendarThai || !tag.present ? ymd.y + kBuddhistEraOffset : ymd.y;
  bool era_name_written = false;

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

  const DbnumStyle* const dbnum = section.digit_style;
  for (std::size_t i = 0; i < section.tokens.size(); ++i) {
    const Token& tk = section.tokens[i];
    switch (tk.kind) {
      case Tok::DateY2:
      case Tok::DateY4: {
        // The Japanese calendar writes the two-digit era year for both widths;
        // the Taiwan calendar writes the full year for both.
        if (tag.calendar == kCalendarJapanese) {
          append_dbnum_date_field(out, dbnum, static_cast<unsigned>(era_year), 2U);
          break;
        }
        if (tk.kind == Tok::DateY2 && tag.calendar != kCalendarTaiwan) {
          const unsigned y2 = static_cast<unsigned>(((calendar_year % 100) + 100) % 100);
          append_dbnum_date_field(out, dbnum, y2, 2U);
          break;
        }
        char buf[16];
        const int n = format_signed(buf, sizeof(buf), calendar_year, 4);
        if (n > 0) {
          append_chars_dbnum(out, dbnum, std::string_view(buf, static_cast<std::size_t>(n)));
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
        out.append(name_entry(*names.month_short, static_cast<long long>(ymd.m) - 1));
        break;
      case Tok::DateMMMM:
        out.append(name_entry(*names.month_long, static_cast<long long>(ymd.m) - 1));
        break;
      case Tok::DateMMMMM: {
        // `mmmmm` (run length >= 5) emits the first scalar of the month name.
        const std::string_view name = name_entry(*names.month_long, static_cast<long long>(ymd.m) - 1);
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
        out.append(name_entry(*names.day_short, sun0));
        break;
      case Tok::DateDDDD:
        out.append(name_entry(*names.day_long, sun0));
        break;
      case Tok::DateAaa:
        out.append(name_entry(*names.weekday_short, sun0));
        break;
      case Tok::DateAaaa:
        out.append(name_entry(*names.weekday_long, sun0));
        break;
      case Tok::EraG:
      case Tok::EraGG:
      case Tok::EraGGG:
        if (era_on) {
          out.append(tk.kind == Tok::EraG ? era.roman : tk.kind == Tok::EraGG ? era.kanji1 : era.kanji2);
          era_name_written = true;
        }
        break;
      case Tok::DateB2: {
        const auto b2 = static_cast<unsigned>((b_year % 100 + 100) % 100);
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
        append_int_dbnum(out, static_cast<long long>(b_year), dbnum);
        break;
      case Tok::EraE:
      case Tok::EraEE: {
        if (!era_on) {
          append_int_dbnum(out, static_cast<long long>(ymd.y), dbnum);
          break;
        }
        // `-x-gannen` writes the first era year as 元 after an era name (locale_tokens.lcid_gannen).
        if (tag.gannen && era_name_written && era_year == 1 && dbnum == nullptr) {
          out.append(kGannen);
          break;
        }
        const std::size_t width = tk.kind == Tok::EraE ? 1U : 2U;
        if (era_year >= 0 && (width == 1U || era_year < 100)) {
          append_dbnum_date_field(out, dbnum, static_cast<unsigned>(era_year), width);
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
        out.append(pm ? names.pm : names.am);
        break;
      case Tok::AP:
        if (names.tag_meridiem) {
          out.append(pm ? names.pm : names.am);
        } else {
          out.append(pm ? "P" : "A");
        }
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
