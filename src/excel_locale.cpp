#include "excel_locale.h"

namespace formulon {
namespace {

// @size-budget: 2 KB
constexpr LocaleFacts kJapaneseFacts{
    "合計",
    "総計",
    "行フィールド ",
    "列フィールド ",
    "値 ",
    "¥#,##0;¥-#,##0",
    0U,
    "¥",
    DateOrder::kYMD,
    true,
    true,
    true,
    true,
    true,
    {"日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"},
    {"日", "月", "火", "水", "木", "金", "土"},
    false,
    true,
};

constexpr LocaleFacts kEnglishFacts{
    "Total",
    "Grand Total",
    "Row Field ",
    "Column Field ",
    "Value ",
    "$#,##0.00_);($#,##0.00)",
    2U,
    "",
    DateOrder::kMDY,
    false,
    false,
    false,
    false,
    false,
    {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"},
    {"Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"},
    true,
    false,
};

}  // namespace

const LocaleFacts& locale_facts(ExcelProfile profile) noexcept {
  return profile.locale == ExcelLocale::kJaJP ? kJapaneseFacts : kEnglishFacts;
}

SbcsCodepage sbcs_codepage(ExcelProfile profile) noexcept {
  if (profile.locale == ExcelLocale::kJaJP) {
    return SbcsCodepage::kCp932;
  }
  return profile.host == ExcelHost::kMac365 ? SbcsCodepage::kMacRoman : SbcsCodepage::kWindows1252;
}

WidthFolding width_folding(ExcelProfile profile) noexcept {
  return profile.host == ExcelHost::kMac365 ? WidthFolding::kMac : WidthFolding::kWin;
}

}  // namespace formulon
