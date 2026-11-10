//
// The `<sharedItems>` type flags and bounds of one cache field, derived from
// the values it holds. The cache writer, the refresh and the read API share
// this one derivation.
//
// Values fall in four classes: S (text, boolean or error), N (number), D (a
// number marked as a date) and B (blank). A flag equal to its schema default
// is left absent, so the result spells only what Excel writes.

#ifndef FORMULON_PIVOT_SHARED_ITEMS_HINTS_H_
#define FORMULON_PIVOT_SHARED_ITEMS_HINTS_H_

#include <string>
#include <vector>

#include "pivot/pivot_cache.h"
#include "value.h"

namespace formulon::pivot {

/// The hints Excel writes for `values`; `is_date[i]` marks value `i` as a date
/// serial (a shorter vector reads as unset).
///
/// containsString is S, containsSemiMixedTypes S or B, containsNonDate S or
/// N, containsDate D, containsBlank B, containsMixedTypes two or more of S /
/// N / D. With N and no D: containsNumber, containsInteger when every number
/// is integral, and minValue / maxValue. With D and no N: minDate, and maxDate
/// one day past the latest date unless it carries a time of day. A field
/// holding both D and N carries no bounds here (`apply_mixed_date_slots`).
/// `longText` marks a string over 255 UTF-16 units. Not `present` when
/// `values` is empty.
SharedItemsHints derive_shared_items_hints(const std::vector<Value>& values, const std::vector<bool>& is_date,
                                           bool date1904);

/// The `minDate` / `maxDate` text of a slot holding `number` in a field that
/// mixes dates and numbers. The number's IEEE-754 bytes, little-endian, read
/// as `{s16 year, u16 month, u8 day, u8 hour, u8 minute, u8 second}`: a year
/// in 1901..9999 is normalised as DateSerial does (month and day carry); a
/// year up to 1900 gives 1899-12-31 plus the time; a later year gives
/// 1899-12-30T00:00:00. Hours, minutes and seconds carry into the date.
std::string excel_mixed_date_slot_text(double number);

/// Sets `minDate` / `maxDate` on `hints` for `values` (record order) when they
/// hold both dates and numbers. One min and one max slot are updated in order:
/// a date compares as a date against the slot, a number as a double against
/// the slot's 8 bytes. A slot left holding a date is written as that date
/// (`maxDate` one day later when it has no time of day); one holding a number
/// is written by `excel_mixed_date_slot_text`.
void apply_mixed_date_slots(const std::vector<Value>& values, const std::vector<bool>& is_date, bool date1904,
                            SharedItemsHints& hints);

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_SHARED_ITEMS_HINTS_H_
