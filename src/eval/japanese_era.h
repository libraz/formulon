// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Japanese era (和暦) classification helpers shared by:
//
//   * `eval/text_format/number_format_render.cpp` — TEXT() format codes
//     `g` / `gg` / `ggg` and the era-year computation embedded in those
//     pipelines.
//   * `pivot/pivot_evaluator.cpp` (planned) — date grouping with
//     `CalendarSystem::Japanese`.
//
// The boundary anchors are the Mac Excel 365 ja-JP oracle values; do not
// adjust without re-verifying against the oracle. The complementary
// parser-side anchors (era + era-year -> Gregorian year) live in
// `date_text_parse.cpp`'s anonymous namespace because they are only
// consumed there.

#ifndef FORMULON_EVAL_JAPANESE_ERA_H_
#define FORMULON_EVAL_JAPANESE_ERA_H_

namespace formulon {
namespace eval {
namespace japanese_era {

/// Per-era metadata. `roman` / `kanji1` / `kanji2` are byte literals owned
/// by the implementation (static storage); callers may take and copy
/// pointers freely. The string contents:
///
/// * `roman`: 1-byte ASCII abbreviation used by the format code `g`
///   (`R` / `H` / `S` / `T` / `M`).
/// * `kanji1`: 3-byte UTF-8 single-kanji abbreviation used by `gg`
///   (`令` / `平` / `昭` / `大` / `明`).
/// * `kanji2`: 6-byte UTF-8 full era name used by `ggg`
///   (`令和` / `平成` / `昭和` / `大正` / `明治`).
struct EraInfo {
  int start_year;
  unsigned start_month;
  unsigned start_day;
  int year_anchor;     ///< Gregorian year corresponding to era year 1.
  const char* roman;
  const char* kanji1;
  const char* kanji2;
};

/// Returns the era spanning the given Gregorian (year, month, day).
///
/// Boundaries (Mac Excel 365 ja-JP):
///
///   * 1868-01-25 .. 1912-07-29 -> Meiji
///   * 1912-07-30 .. 1926-12-24 -> Taisho
///   * 1926-12-25 .. 1989-01-07 -> Showa
///   * 1989-01-08 .. 2019-04-30 -> Heisei
///   * 2019-05-01 onward         -> Reiwa
///
/// Pre-Meiji dates fall through to Meiji's anchor; the era-year math then
/// produces a non-positive result, mirroring how Mac Excel renders such
/// inputs (no validation).
const EraInfo& classify_era(int year, unsigned month, unsigned day) noexcept;

}  // namespace japanese_era
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_JAPANESE_ERA_H_
