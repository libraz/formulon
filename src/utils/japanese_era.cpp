//
// Japanese era classification. Anchors come from the Mac Excel 365 ja-JP
// oracle; see header for the boundary table.

#include "utils/japanese_era.h"

#include <iterator>

namespace formulon {
namespace japanese_era {

namespace {

// Newest first; the last entry (Meiji) is the fallback. UTF-8 byte literals
// for the era kanji: 3-byte single-kanji and 6-byte full-name forms, both used
// by the TEXT renderer.
constexpr EraInfo kEras[] = {
    {2019, 5u, 1u, 2019, "R", "\xE4\xBB\xA4", "\xE4\xBB\xA4\xE5\x92\x8C"},
    {1989, 1u, 8u, 1989, "H", "\xE5\xB9\xB3", "\xE5\xB9\xB3\xE6\x88\x90"},
    {1926, 12u, 25u, 1926, "S", "\xE6\x98\xAD", "\xE6\x98\xAD\xE5\x92\x8C"},
    {1912, 7u, 30u, 1912, "T", "\xE5\xA4\xA7", "\xE5\xA4\xA7\xE6\xAD\xA3"},
    {1868, 1u, 25u, 1868, "M", "\xE6\x98\x8E", "\xE6\x98\x8E\xE6\xB2\xBB"},
};

}  // namespace

const EraInfo& classify_era(int year, unsigned month, unsigned day) noexcept {
  // The first era whose start is not after (year, month, day), compared
  // lexicographically.
  for (const EraInfo& era : kEras) {
    const bool before_start = year != era.start_year     ? year < era.start_year
                              : month != era.start_month ? month < era.start_month
                                                         : day < era.start_day;
    if (!before_start) {
      return era;
    }
  }
  // Pre-Meiji: still classified as Meiji (anchor 1868). Mac Excel does
  // not validate, so an Edo-period serial would emit a non-positive era
  // year; this mirrors the parser's lenient behaviour in
  // `date_text_parse.cpp`.
  return kEras[std::size(kEras) - 1];
}

}  // namespace japanese_era
}  // namespace formulon
