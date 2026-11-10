// @size-budget: 4 KB
//
// `<sharedItems>` type flags and bounds. See the header for the rules.

#include "pivot/shared_items_hints.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <initializer_list>
#include <string>
#include <string_view>
#include <vector>

#include "pivot/date_serial.h"
#include "pivot/field_grouping.h"
#include "pivot/pivot_cache.h"
#include "utils/date_time.h"
#include "utils/number_text.h"
#include "utils/utf8_length.h"
#include "value.h"

namespace formulon::pivot {
namespace {

// Longest string `longText` stays off for, in UTF-16 units.
constexpr std::uint32_t kLongTextUnits = 255U;

void set_hint(bool& flag, bool& has, bool value) {
  flag = value;
  has = true;
}

bool flag_at(const std::vector<bool>& flags, std::size_t i) {
  return i < flags.size() && flags[i];
}

bool is_date_value(const std::vector<Value>& values, const std::vector<bool>& is_date, std::size_t i, bool date1904) {
  return values[i].is_number() && flag_at(is_date, i) && is_valid_date_serial(values[i].as_number(), date1904);
}

std::string number_text(double v) {
  std::string text;
  append_xml_number(text, v);
  return text;
}

// Calendar fields of a slot, compared in this order.
struct SlotParts {
  int year = 0;
  unsigned month = 0;
  unsigned day = 0;
  unsigned hour = 0;
  unsigned minute = 0;
  unsigned second = 0;
};

std::uint64_t bits_of(double v) {
  std::uint64_t bits = 0;
  std::memcpy(&bits, &v, sizeof bits);
  return bits;
}

double double_of(std::uint64_t bits) {
  double v = 0.0;
  std::memcpy(&v, &bits, sizeof v);
  return v;
}

SlotParts parts_of_bits(std::uint64_t bits) {
  SlotParts p;
  p.year = static_cast<std::int16_t>(static_cast<std::uint16_t>(bits & 0xFFFFU));
  p.month = static_cast<unsigned>((bits >> 16U) & 0xFFFFU);
  p.day = static_cast<unsigned>((bits >> 32U) & 0xFFU);
  p.hour = static_cast<unsigned>((bits >> 40U) & 0xFFU);
  p.minute = static_cast<unsigned>((bits >> 48U) & 0xFFU);
  p.second = static_cast<unsigned>((bits >> 56U) & 0xFFU);
  return p;
}

std::uint64_t bits_of_parts(const SlotParts& p) {
  return static_cast<std::uint64_t>(static_cast<std::uint16_t>(p.year)) |
         (static_cast<std::uint64_t>(p.month & 0xFFFFU) << 16U) | (static_cast<std::uint64_t>(p.day & 0xFFU) << 32U) |
         (static_cast<std::uint64_t>(p.hour & 0xFFU) << 40U) | (static_cast<std::uint64_t>(p.minute & 0xFFU) << 48U) |
         (static_cast<std::uint64_t>(p.second & 0xFFU) << 56U);
}

unsigned digits_at(std::string_view text, std::size_t pos, std::size_t count) {
  unsigned v = 0;
  for (std::size_t i = pos; i < pos + count && i < text.size(); ++i) {
    v = v * 10U + static_cast<unsigned>(text[i] - '0');
  }
  return v;
}

// The calendar fields of date serial `serial`, as its ISO body spells them.
SlotParts parts_of_date(double serial, bool date1904) {
  const std::string iso = group_date_iso(serial, date1904);
  SlotParts p;
  p.year = static_cast<int>(digits_at(iso, 0U, 4U));
  p.month = digits_at(iso, 5U, 2U);
  p.day = digits_at(iso, 8U, 2U);
  p.hour = digits_at(iso, 11U, 2U);
  p.minute = digits_at(iso, 14U, 2U);
  p.second = digits_at(iso, 17U, 2U);
  return p;
}

int compare_parts(const SlotParts& a, const SlotParts& b) {
  const auto cmp = [](auto x, auto y) { return x < y ? -1 : (y < x ? 1 : 0); };
  for (const int c : {cmp(a.year, b.year), cmp(a.month, b.month), cmp(a.day, b.day), cmp(a.hour, b.hour),
                      cmp(a.minute, b.minute), cmp(a.second, b.second)}) {
    if (c != 0) {
      return c;
    }
  }
  return 0;
}

// One bound of a date + number field: the date or number it last took.
struct MixedSlot {
  bool set = false;
  bool holds_date = false;
  double value = 0.0;
  SlotParts date;
};

// Sign of `value` (a date when `is_date`) against `slot`.
int compare_with_slot(double value, bool is_date, const MixedSlot& slot, bool date1904) {
  if (is_date) {
    if (slot.holds_date) {
      return value < slot.value ? -1 : (slot.value < value ? 1 : 0);
    }
    return compare_parts(parts_of_date(value, date1904), parts_of_bits(bits_of(slot.value)));
  }
  const double other = slot.holds_date ? double_of(bits_of_parts(slot.date)) : slot.value;
  return value < other ? -1 : (other < value ? 1 : 0);
}

void take(MixedSlot& slot, double value, bool is_date, bool date1904) {
  slot.set = true;
  slot.holds_date = is_date;
  slot.value = value;
  if (is_date) {
    slot.date = parts_of_date(value, date1904);
  }
}

void append_iso(std::string& out, const date_time::YMD& ymd, std::int64_t seconds) {
  append_padded(out, ymd.y, 4);
  out.push_back('-');
  append_padded(out, ymd.m, 2);
  out.push_back('-');
  append_padded(out, ymd.d, 2);
  out.push_back('T');
  append_padded(out, seconds / 3600, 2);
  out.push_back(':');
  append_padded(out, seconds / 60 % 60, 2);
  out.push_back(':');
  append_padded(out, seconds % 60, 2);
}

}  // namespace

SharedItemsHints derive_shared_items_hints(const std::vector<Value>& values, const std::vector<bool>& is_date,
                                           bool date1904) {
  std::size_t dates = 0;
  std::size_t numbers = 0;
  std::size_t strings = 0;
  std::size_t blanks = 0;
  bool integral = true;
  bool long_text = false;
  double num_lo = 0.0;
  double num_hi = 0.0;
  double date_lo = 0.0;
  double date_hi = 0.0;
  for (std::size_t i = 0; i < values.size(); ++i) {
    const Value& v = values[i];
    if (is_date_value(values, is_date, i, date1904)) {
      const double d = v.as_number();
      date_lo = dates == 0 ? d : std::min(date_lo, d);
      date_hi = dates == 0 ? d : std::max(date_hi, d);
      ++dates;
    } else if (v.is_number()) {
      const double d = v.as_number();
      num_lo = numbers == 0 ? d : std::min(num_lo, d);
      num_hi = numbers == 0 ? d : std::max(num_hi, d);
      integral = integral && std::isfinite(d) && d == std::floor(d);
      ++numbers;
    } else if (v.is_blank()) {
      ++blanks;
    } else {
      long_text = long_text || (v.is_text() && utf16_units_in(v.as_text()) > kLongTextUnits);
      ++strings;
    }
  }
  SharedItemsHints h;
  if (values.empty()) {
    return h;
  }
  if (strings == 0 && blanks == 0) {
    set_hint(h.contains_semi_mixed, h.has_contains_semi_mixed, false);
  }
  if (strings == 0 && numbers == 0) {
    set_hint(h.contains_non_date, h.has_contains_non_date, false);
  }
  if (dates > 0) {
    set_hint(h.contains_date, h.has_contains_date, true);
  }
  if (strings == 0) {
    set_hint(h.contains_string, h.has_contains_string, false);
  }
  if (blanks > 0) {
    set_hint(h.contains_blank, h.has_contains_blank, true);
  }
  if ((dates > 0 ? 1 : 0) + (numbers > 0 ? 1 : 0) + (strings > 0 ? 1 : 0) >= 2) {
    set_hint(h.contains_mixed_types, h.has_contains_mixed_types, true);
  }
  if (numbers > 0 && dates == 0) {
    set_hint(h.contains_number, h.has_contains_number, true);
    if (integral) {
      set_hint(h.contains_integer, h.has_contains_integer, true);
    }
    h.has_min_value = true;
    h.min_value = number_text(num_lo);
    h.has_max_value = true;
    h.max_value = number_text(num_hi);
  }
  if (dates > 0 && numbers == 0) {
    const GroupBounds bounds = auto_date_group_bounds(date_lo, date_hi, date1904);
    h.has_min_date = true;
    h.min_date = group_date_iso(bounds.start, date1904);
    h.has_max_date = true;
    h.max_date = group_date_iso(bounds.end, date1904);
  }
  if (long_text) {
    h.extra_attrs.emplace_back("longText", "1");
  }
  h.present = true;
  return h;
}

std::string excel_mixed_date_slot_text(double number) {
  const SlotParts p = parts_of_bits(bits_of(number));
  std::string out;
  if (p.year >= 10000) {
    out = "1899-12-30T00:00:00";
    return out;
  }
  std::int64_t days = 0;
  if (p.year <= 1900) {
    days = date_time::days_from_civil(1899, 12U, 31U);
  } else {
    // DateSerial: months past December (or month 0) carry into the year, and
    // the day counts from the first of the month.
    int year = p.year;
    unsigned month = p.month;
    if (month == 0U) {
      year -= 1;
      month = 12U;
    } else {
      year += static_cast<int>((month - 1U) / 12U);
      month = (month - 1U) % 12U + 1U;
    }
    days = date_time::days_from_civil(year, month, 1U) + static_cast<std::int64_t>(p.day) - 1;
  }
  const std::int64_t seconds = static_cast<std::int64_t>(p.hour) * 3600 + static_cast<std::int64_t>(p.minute) * 60 +
                               static_cast<std::int64_t>(p.second);
  days += seconds / date_time::kSecondsPerDay;
  append_iso(out, date_time::civil_from_days(days), seconds % date_time::kSecondsPerDay);
  return out;
}

void apply_mixed_date_slots(const std::vector<Value>& values, const std::vector<bool>& is_date, bool date1904,
                            SharedItemsHints& hints) {
  MixedSlot lo;
  MixedSlot hi;
  bool any_date = false;
  bool any_number = false;
  for (std::size_t i = 0; i < values.size(); ++i) {
    if (!values[i].is_number()) {
      continue;
    }
    const bool date = is_date_value(values, is_date, i, date1904);
    any_date = any_date || date;
    any_number = any_number || !date;
    const double v = values[i].as_number();
    if (!lo.set || compare_with_slot(v, date, lo, date1904) < 0) {
      take(lo, v, date, date1904);
    }
    if (!hi.set || compare_with_slot(v, date, hi, date1904) > 0) {
      take(hi, v, date, date1904);
    }
  }
  if (!any_date || !any_number) {
    return;
  }
  hints.has_min_date = true;
  hints.min_date = lo.holds_date ? group_date_iso(lo.value, date1904) : excel_mixed_date_slot_text(lo.value);
  hints.has_max_date = true;
  hints.max_date = hi.holds_date ? group_date_iso(auto_date_group_bounds(hi.value, hi.value, date1904).end, date1904)
                                 : excel_mixed_date_slot_text(hi.value);
}

}  // namespace formulon::pivot
