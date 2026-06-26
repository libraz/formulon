// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the strict ISO 8601 date parser. See iso_date.h.

#include "io/iso_date.h"

#include <cstddef>
#include <string_view>

#include "eval/date_time.h"

namespace formulon::io {
namespace {

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
  if (month < 1 || month > 12 || day < 1 || day > 31) {
    return false;
  }

  double serial = eval::date_time::serial_from_ymd(year, static_cast<unsigned>(month), static_cast<unsigned>(day));

  if (pos < text.size()) {
    // A time component must follow, introduced by 'T' (strict OOXML) or a
    // space (lenient producers). Reject anything else.
    if (text[pos] != 'T' && text[pos] != 't' && text[pos] != ' ') {
      return false;
    }
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
    double frac =
        (static_cast<double>(hour) * 3600.0 + static_cast<double>(minute) * 60.0 + static_cast<double>(second)) /
        86400.0;

    // Optional fractional seconds: '.' followed by one or more digits.
    if (pos < text.size() && text[pos] == '.') {
      ++pos;
      double scale = 0.1;
      bool any = false;
      while (pos < text.size() && text[pos] >= '0' && text[pos] <= '9') {
        frac += (text[pos] - '0') * scale / 86400.0;
        scale *= 0.1;
        ++pos;
        any = true;
      }
      if (!any) {
        return false;
      }
    }
    serial += frac;

    // Optional time-zone designator: 'Z' or '(+|-)hh:mm'. Excel stores a
    // wall-clock serial, so the offset is accepted but not applied.
    if (pos < text.size()) {
      if (text[pos] == 'Z' || text[pos] == 'z') {
        ++pos;
      } else if (text[pos] == '+' || text[pos] == '-') {
        ++pos;
        int tz_h = 0;
        int tz_m = 0;
        if (!TakeDigits(text, &pos, 2, &tz_h) || !TakeChar(text, &pos, ':') || !TakeDigits(text, &pos, 2, &tz_m)) {
          return false;
        }
      } else {
        return false;
      }
    }
  }

  if (pos != text.size()) {
    return false;
  }
  *out_serial = serial;
  return true;
}

}  // namespace formulon::io
