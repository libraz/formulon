#ifndef FORMULON_EXCEL_LOCALE_H_
#define FORMULON_EXCEL_LOCALE_H_

#include <array>
#include <cstdint>
#include <string_view>

#include "excel_profile.h"

namespace formulon {

enum class DateOrder : std::uint8_t {
  kMDY = 0,
  kYMD = 1,
};

enum class SbcsCodepage : std::uint8_t {
  kMacRoman = 0,
  kWindows1252 = 1,
  kCp932 = 2,
};

// Criteria, lookup, and D-function text matching use the host's width and
// kana folding rules. The Windows path retains the engine's existing
// non-Mac behavior.
enum class WidthFolding : std::uint8_t {
  kWin = 0,
  kMac = 1,
};

struct LocaleFacts {
  std::string_view grand_total;
  std::string_view hierarchy_grand_total;
  std::string_view row_field_prefix;
  std::string_view column_field_prefix;
  std::string_view value_field_prefix;
  std::string_view dollar_format;
  std::uint8_t dollar_default_decimals;
  std::string_view currency_symbol;
  DateOrder date_order;
  bool dbcs;
  bool fullwidth_numeric_text;
  bool kanji_date_text;
  bool japanese_era;
  bool ja_format_syntax;
  std::array<std::string_view, 7> weekday_long;
  std::array<std::string_view, 7> weekday_short;
  bool char_snaps_near_integer;
  bool phonetic;
};

const LocaleFacts& locale_facts(ExcelProfile profile) noexcept;
SbcsCodepage sbcs_codepage(ExcelProfile profile) noexcept;
WidthFolding width_folding(ExcelProfile profile) noexcept;

}  // namespace formulon

#endif  // FORMULON_EXCEL_LOCALE_H_
