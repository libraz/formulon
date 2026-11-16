#ifndef FORMULON_FONT_COMPLETION_H_
#define FORMULON_FONT_COMPLETION_H_

// The `<family>` and `<charset>` Excel's save fills in on a literal font that names a face and carries neither, and
// the canonical name it writes for an alias. A face missing here is written as stored, as are Aptos and Aptos Narrow,
// which Excel completes with neither.

#include <cstdint>
#include <string_view>

namespace formulon {

struct FontCompletion {
  std::string_view name;
  std::uint8_t family;      // 0 leaves `<family>` absent
  std::uint8_t charset;     // 0 leaves `<charset>` absent
  std::string_view rename;  // empty keeps the name
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
    // A literal 游ゴシック without either is family 3; Excel's stock theme-linked record states 2 itself.
    {"\xE6\xB8\xB8\xE3\x82\xB4\xE3\x82\xB7\xE3\x83\x83\xE3\x82\xAF", 3, 128, {}},
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

/// The completion for a literal face named exactly `name`, or nullptr.
inline const FontCompletion* find_font_completion(std::string_view name) noexcept {
  for (const FontCompletion& row : kFontCompletions) {
    if (row.name == name) {
      return &row;
    }
  }
  return nullptr;
}

}  // namespace formulon

#endif  // FORMULON_FONT_COMPLETION_H_
