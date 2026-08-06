//
// Japanese era classification. Anchors come from the Mac Excel 365 ja-JP
// oracle; see header for the boundary table.

#include "eval/japanese_era.h"

namespace formulon {
namespace eval {
namespace japanese_era {

const EraInfo& classify_era(int year, unsigned month, unsigned day) noexcept {
  // UTF-8 byte literals for the era kanji. 3-byte single-kanji and
  // 6-byte full-name forms; both are used by the TEXT renderer.
  static const EraInfo kReiwa{2019, 5u, 1u, 2019, "R", "\xE4\xBB\xA4", "\xE4\xBB\xA4\xE5\x92\x8C"};
  static const EraInfo kHeisei{1989, 1u, 8u, 1989, "H", "\xE5\xB9\xB3", "\xE5\xB9\xB3\xE6\x88\x90"};
  static const EraInfo kShowa{1926, 12u, 25u, 1926, "S", "\xE6\x98\xAD", "\xE6\x98\xAD\xE5\x92\x8C"};
  static const EraInfo kTaisho{1912, 7u, 30u, 1912, "T", "\xE5\xA4\xA7", "\xE5\xA4\xA7\xE6\xAD\xA3"};
  static const EraInfo kMeiji{1868, 1u, 25u, 1868, "M", "\xE6\x98\x8E", "\xE6\x98\x8E\xE6\xB2\xBB"};
  // Order (year, month, day) lexicographically.
  auto cmp = [](int y1, unsigned m1, unsigned d1, int y2, unsigned m2, unsigned d2) {
    if (y1 != y2) {
      return y1 < y2;
    }
    if (m1 != m2) {
      return m1 < m2;
    }
    return d1 < d2;
  };
  if (!cmp(year, month, day, kReiwa.start_year, kReiwa.start_month, kReiwa.start_day)) {
    return kReiwa;
  }
  if (!cmp(year, month, day, kHeisei.start_year, kHeisei.start_month, kHeisei.start_day)) {
    return kHeisei;
  }
  if (!cmp(year, month, day, kShowa.start_year, kShowa.start_month, kShowa.start_day)) {
    return kShowa;
  }
  if (!cmp(year, month, day, kTaisho.start_year, kTaisho.start_month, kTaisho.start_day)) {
    return kTaisho;
  }
  // Pre-Meiji: still classified as Meiji (anchor 1868). Mac Excel does
  // not validate, so an Edo-period serial would emit a non-positive era
  // year; this mirrors the parser's lenient behaviour in
  // `date_text_parse.cpp`.
  return kMeiji;
}

}  // namespace japanese_era
}  // namespace eval
}  // namespace formulon
