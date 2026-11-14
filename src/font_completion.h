#ifndef FORMULON_FONT_COMPLETION_H_
#define FORMULON_FONT_COMPLETION_H_

// The `<family>` and `<charset>` Excel's save fills in on a literal font that names a face and carries neither, and
// the canonical name it writes for an alias. A face missing here is written as stored.

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
