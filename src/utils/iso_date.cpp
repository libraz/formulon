//
// Implementation of the strict ISO 8601 date parser. See iso_date.h.

#include "utils/iso_date.h"

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "utils/date_time.h"
#include "utils/strings.h"

namespace formulon::date_time {
namespace {

// Fractional-second digits kept: below 2^53, so the integer and its power of
// ten are exact doubles. Later digits sit far below a serial's resolution.
// @limit: internal
constexpr int kMaxFractionDigits = 15;

/// Consumes exactly `count` ASCII digits from `s` starting at `*pos`,
/// accumulating their decimal value into `*out`. Returns false (leaving
/// `*pos` / `*out` partially advanced) if fewer than `count` digits are
/// present; callers treat any false as a whole-parse failure.
bool TakeDigits(std::string_view s, std::size_t* pos, std::size_t count, int* out) noexcept {
  int value = 0;
  for (std::size_t i = 0; i < count; ++i) {
    if (*pos >= s.size()) {
      return false;
    }
    const char c = s[*pos];
    if (c < '0' || c > '9') {
      return false;
    }
    value = value * 10 + (c - '0');
    ++(*pos);
  }
  *out = value;
  return true;
}

/// Consumes a single literal character `lit` at `*pos`, advancing on match.
bool TakeChar(std::string_view s, std::size_t* pos, char lit) noexcept {
  if (*pos >= s.size() || s[*pos] != lit) {
    return false;
  }
  ++(*pos);
  return true;
}

// ISO 8601 permits a UTC marker or a signed hour/minute offset. Excel keeps
// the wall-clock value, so the validated offset is intentionally ignored.
bool TakeTimezoneSuffix(std::string_view s, std::size_t* pos) noexcept {
  if (*pos >= s.size()) {
    return true;
  }
  if (s[*pos] == 'Z' || s[*pos] == 'z') {
    ++(*pos);
    return true;
  }
  if (s[*pos] != '+' && s[*pos] != '-') {
    return false;
  }
  ++(*pos);
  int tz_h = 0;
  int tz_m = 0;
  if (!TakeDigits(s, pos, 2, &tz_h) || !TakeChar(s, pos, ':') || !TakeDigits(s, pos, 2, &tz_m)) {
    return false;
  }
  // OOXML's XSD timezone range ends at fourteen hours; 14:00 is the only valid
  // offset at that boundary. Minutes are always below 60.
  if (tz_m >= 60) {
    return false;
  }
  return tz_h < 14 || (tz_h == 14 && tz_m == 0);
}

}  // namespace

bool parse_iso_date_serial(std::string_view text, double* out_serial) noexcept {
  std::size_t pos = 0;
  int year = 0;
  int month = 0;
  int day = 0;
  if (!TakeDigits(text, &pos, 4, &year) || !TakeChar(text, &pos, '-') || !TakeDigits(text, &pos, 2, &month) ||
      !TakeChar(text, &pos, '-') || !TakeDigits(text, &pos, 2, &day)) {
    return false;
  }
  // Reject any calendar date that does not actually exist (e.g.
  // `2024-02-30`, `2023-02-29`): `serial_from_ymd` silently normalizes an
  // out-of-range day into the following month, which would violate this
  // parser's "strict" lexical contract by accepting non-canonical input.
  if (month < 1 || month > 12 || day < 1 ||
      day > static_cast<int>(date_time::days_in_month(year, static_cast<unsigned>(month)))) {
    return false;
  }

  double serial = date_time::serial_from_ymd(year, static_cast<unsigned>(month), static_cast<unsigned>(day));

  if (pos < text.size() && (text[pos] == 'T' || text[pos] == 't' || text[pos] == ' ')) {
    // A time component follows 'T' (strict OOXML) or a space (lenient producers).
    ++pos;
    int hour = 0;
    int minute = 0;
    int second = 0;
    if (!TakeDigits(text, &pos, 2, &hour) || !TakeChar(text, &pos, ':') || !TakeDigits(text, &pos, 2, &minute) ||
        !TakeChar(text, &pos, ':') || !TakeDigits(text, &pos, 2, &second)) {
      return false;
    }
    if (hour > 23 || minute > 59 || second > 59) {
      return false;
    }
    double seconds =
        static_cast<double>(hour) * 3600.0 + static_cast<double>(minute) * 60.0 + static_cast<double>(second);

    // Optional fractional seconds: '.' followed by one or more digits, read as
    // one exact integer over an exact power of ten.
    if (pos < text.size() && text[pos] == '.') {
      ++pos;
      std::uint64_t digits = 0;
      double scale = 1.0;
      int kept = 0;
      bool any = false;
      while (pos < text.size() && strings::is_ascii_digit(text[pos])) {
        if (kept < kMaxFractionDigits) {
          digits = digits * 10U + static_cast<std::uint64_t>(text[pos] - '0');
          scale *= 10.0;
          ++kept;
        }
        ++pos;
        any = true;
      }
      if (!any) {
        return false;
      }
      seconds += static_cast<double>(digits) / scale;
    }
    serial += seconds / static_cast<double>(date_time::kSecondsPerDay);
  }

  // Optional time-zone designator: 'Z' or '(+|-)hh:mm'. Excel stores a
  // wall-clock serial, so the offset is accepted but not applied. This is
  // shared by date-only and date-time forms.
  if (!TakeTimezoneSuffix(text, &pos)) {
    return false;
  }

  if (pos != text.size()) {
    return false;
  }
  *out_serial = serial;
  return true;
}

}  // namespace formulon::date_time
