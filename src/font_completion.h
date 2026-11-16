#ifndef FORMULON_FONT_COMPLETION_H_
#define FORMULON_FONT_COMPLETION_H_

// The `<family>` and `<charset>` Excel holds for a font that names a face and states no family, charset or scheme (it
// completes them when it loads the font), and the canonical name it gives an alias. A face missing here is held as
// stored, as are Aptos and Aptos Narrow, which Excel completes with neither. A font renamed by a cell-format patch
// is given the same values, except that a face equal to the Normal font's copies the Normal font's.

#include <cassert>
#include <cstdint>
#include <string>
#include <string_view>

#include "excel_profile.h"
#include "styles.h"

namespace formulon {

/// `FontCompletion::ja_family` value meaning the row's `family` holds in a ja profile too.
inline constexpr std::uint8_t kSameFamilyInJa = 0xFFU;

struct FontCompletion {
  std::string_view name;
  std::uint8_t family;      // 0 leaves `<family>` absent
  std::uint8_t charset;     // 0 leaves `<charset>` absent
  std::string_view rename;  // empty keeps the name
  std::uint8_t ja_family =
      kSameFamilyInJa;  // `family` under a ja profile; only ja and en are measured, the rest follow en
};

inline constexpr FontCompletion kFontCompletions[] = {
    {"Arial", 2, 0, {}},
    {"Helvetica", 2, 0, {}},
    {"Calibri", 2, 0, {}},
    {"Tahoma", 2, 0, {}},
    {"Verdana", 2, 0, {}},
    {"Helvetica Neue", 2, 0, {}},
    {"Cambria", 1, 0, {}},
    {"Times New Roman", 1, 0, {}},
    {"Courier New", 1, 0, {}},
    {"Comic Sans MS", 4, 0, {}},
    {"Wingdings", 0, 2, {}},
    {"Symbol", 0, 2, {}},
    {"SimSun", 0, 134, {}},
    {"\xE3\x83\xA1\xE3\x82\xA4\xE3\x83\xAA\xE3\x82\xAA", 2, 128, {}},  // メイリオ
    {"\xE3\x83\x92\xE3\x83\xA9\xE3\x82\xAE\xE3\x83\x8E\xE8\xA7\x92\xE3\x82\xB4\xE3\x82\xB7\xE3\x83\x83\xE3\x82\xAF",
     2,
     128,
     {}},  // ヒラギノ角ゴシック
    {"Malgun Gothic", 2, 129, {}},
    // 游ゴシック is family 3 under a ja profile and has none elsewhere; Excel's stock theme-linked record states 2.
    {"\xE6\xB8\xB8\xE3\x82\xB4\xE3\x82\xB7\xE3\x83\x83\xE3\x82\xAF", 0, 128, {}, 3},
    {"\xE6\xB8\xB8\xE6\x98\x8E\xE6\x9C\x9D", 1, 128, {}},               // 游明朝
    {"\xE7\xAD\x89\xE7\xBA\xBF", 4, 134, {}},                           // 等线
    {"\xEB\xA7\x91\xEC\x9D\x80 \xEA\xB3\xA0\xEB\x94\x95", 2, 129, {}},  // 맑은 고딕
    {"\xE6\x96\xB0\xE7\xB4\xB0\xE6\x98\x8E\xE9\xAB\x94", 1, 136, {}},   // 新細明體
    {"\xEF\xBC\xAD\xEF\xBC\xB3 \xEF\xBC\xB0\xE3\x82\xB4\xE3\x82\xB7\xE3\x83\x83\xE3\x82\xAF",
     2,
     128,
     {}},  // ＭＳ Ｐゴシック
    {"\xEF\xBC\xAD\xEF\xBC\xB3 \xE3\x82\xB4\xE3\x82\xB7\xE3\x83\x83\xE3\x82\xAF", 2, 128, {}},  // ＭＳ ゴシック
    {"\xEF\xBC\xAD\xEF\xBC\xB3 \xE6\x98\x8E\xE6\x9C\x9D", 1, 128, {}},                          // ＭＳ 明朝
    {"Meiryo", 2, 128, {}},
    {"Meiryo UI", 2, 128, {}},
    // MS Pゴシック is written as ＭＳ Ｐゴシック.
    {"MS P\xE3\x82\xB4\xE3\x82\xB7\xE3\x83\x83\xE3\x82\xAF", 2, 128,
     "\xEF\xBC\xAD\xEF\xBC\xB3 \xEF\xBC\xB0\xE3\x82\xB4\xE3\x82\xB7\xE3\x83\x83\xE3\x82\xAF"},
};

/// The row of the face named exactly `name`, or nullptr.
inline const FontCompletion* find_font_completion(std::string_view name) noexcept {
  for (const FontCompletion& row : kFontCompletions) {
    if (row.name == name) {
      return &row;
    }
  }
  return nullptr;
}

/// `<family>` the row's face gets under `locale`.
inline std::uint8_t completion_family(const FontCompletion& row, ExcelLocale locale) noexcept {
  return locale == ExcelLocale::kJaJP && row.ja_family != kSameFamilyInJa ? row.ja_family : row.family;
}

/// Completes a font that states no family, no charset and no scheme from its face's row, and renames an alias; any
/// other font is left alone. A face equal to the Normal font's gets no more than any other.
inline void complete_font_from_face(FontRecord& font, ExcelLocale locale) {
  if (font.has_family || font.has_charset || font.scheme != 0U) {
    return;
  }
  const FontCompletion* row = find_font_completion(font.name);
  if (row == nullptr) {
    return;
  }
  // A font linked to the theme takes its name from the theme (`resolve_scheme_font_names`), never from this table.
  assert(font.scheme == 0U);
  font.family = completion_family(*row, locale);
  font.has_family = font.family != 0U;
  font.charset = row->charset;
  font.has_charset = row->charset != 0U;
  if (!row->rename.empty()) {
    font.name = std::string(row->rename);
  }
}

}  // namespace formulon

#endif  // FORMULON_FONT_COMPLETION_H_
