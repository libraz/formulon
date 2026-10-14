#include "num_fmt_locale.h"

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>

#include "eval/text_format/display_text.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "styles.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

// 2024-03-04 13:30:18.
constexpr double kStamp = 45355.56270833333;

struct MeasuredRow {
  ExcelLocale locale;
  std::uint16_t id;
  double value;
  const char* text;
};

// Mac Excel 16.113.4 `Range.Text` of a cell holding built-in `id` (no <numFmts>
// entry), in each UI locale, with the accounting repeat-fill run removed.
constexpr MeasuredRow kMeasured[] = {
    {ExcelLocale::kJaJP, 5, -1234.567, "¥-1,235"},
    {ExcelLocale::kJaJP, 14, kStamp, "2024/3/4"},
    {ExcelLocale::kJaJP, 15, kStamp, "4-Mar-24"},
    {ExcelLocale::kJaJP, 22, kStamp, "2024/3/4 13:30"},
    {ExcelLocale::kJaJP, 27, kStamp, "R6.3.4"},
    {ExcelLocale::kJaJP, 34, kStamp, "2024年3月"},
    {ExcelLocale::kJaJP, 42, -1234.567, " ¥-1,235 "},
    {ExcelLocale::kJaJP, 42, 0.0, " ¥- "},
    {ExcelLocale::kJaJP, 43, 0.0, " -   "},
    {ExcelLocale::kJaJP, 47, 1234.567, "36:28.8"},
    {ExcelLocale::kJaJP, 59, 1234.567, "1235"},
    {ExcelLocale::kJaJP, 63, -1234.567, "¥-1,235"},
    {ExcelLocale::kJaJP, 71, kStamp, "2024/3/4"},
    {ExcelLocale::kEnUS, 5, -1234.567, "($1,235)"},
    {ExcelLocale::kEnUS, 14, kStamp, "3/4/24"},
    {ExcelLocale::kEnUS, 15, kStamp, "4-Mar-24"},
    {ExcelLocale::kEnUS, 22, kStamp, "3/4/24 13:30"},
    {ExcelLocale::kEnUS, 27, kStamp, "3/4/24"},
    {ExcelLocale::kEnUS, 34, kStamp, "13:30:18"},
    {ExcelLocale::kEnUS, 42, -1234.567, " $(1,235)"},
    {ExcelLocale::kEnUS, 42, 0.0, " $- "},
    {ExcelLocale::kEnUS, 43, 0.0, " -   "},
    {ExcelLocale::kEnUS, 47, 1234.567, "36:28.8"},
    {ExcelLocale::kEnUS, 59, 1234.567, "1235"},
    {ExcelLocale::kEnUS, 63, -1234.567, "($1,235)"},
    {ExcelLocale::kEnUS, 71, kStamp, "3/4/24"},
    {ExcelLocale::kDeDE, 5, -1234.567, "-1.235 €"},
    {ExcelLocale::kDeDE, 14, kStamp, "04.03.24"},
    {ExcelLocale::kDeDE, 15, kStamp, "04. Mär 24"},
    {ExcelLocale::kDeDE, 22, kStamp, "04.03.24 13:30"},
    {ExcelLocale::kDeDE, 27, kStamp, "04.03.24"},
    {ExcelLocale::kDeDE, 34, kStamp, "13:30:18"},
    {ExcelLocale::kDeDE, 42, -1234.567, "-1.235 € "},
    {ExcelLocale::kDeDE, 42, 0.0, " - € "},
    {ExcelLocale::kDeDE, 43, 0.0, " -   "},
    {ExcelLocale::kDeDE, 47, 1234.567, "36:28,8"},
    {ExcelLocale::kDeDE, 59, 1234.567, "1235"},
    {ExcelLocale::kDeDE, 63, -1234.567, "-1.235 €"},
    {ExcelLocale::kDeDE, 71, kStamp, "04.03.24"},
    {ExcelLocale::kFrFR, 5, -1234.567, "(1 235 €)"},
    {ExcelLocale::kFrFR, 14, kStamp, "04/03/2024"},
    {ExcelLocale::kFrFR, 15, kStamp, "04-mars-24"},
    {ExcelLocale::kFrFR, 22, kStamp, "04/03/2024 13:30"},
    {ExcelLocale::kFrFR, 27, kStamp, "04/03/2024"},
    {ExcelLocale::kFrFR, 34, kStamp, "13:30:18"},
    {ExcelLocale::kFrFR, 42, -1234.567, " (1 235) € "},
    {ExcelLocale::kFrFR, 42, 0.0, " -  € "},
    {ExcelLocale::kFrFR, 43, 0.0, " -    "},
    {ExcelLocale::kFrFR, 47, 1234.567, "36:28,8"},
    {ExcelLocale::kFrFR, 59, 1234.567, "1235"},
    {ExcelLocale::kFrFR, 63, -1234.567, "(1 235 €)"},
    {ExcelLocale::kFrFR, 71, kStamp, "04/03/2024"},
    {ExcelLocale::kZhCN, 5, -1234.567, "(¥1,235)"},
    {ExcelLocale::kZhCN, 14, kStamp, "2024/3/4"},
    {ExcelLocale::kZhCN, 15, kStamp, "4-Mar-24"},
    {ExcelLocale::kZhCN, 22, kStamp, "2024/3/4 13:30"},
    {ExcelLocale::kZhCN, 27, kStamp, "2024年3月"},
    {ExcelLocale::kZhCN, 34, kStamp, "PMtw1时30分"},
    {ExcelLocale::kZhCN, 42, -1234.567, " ¥(1,235)"},
    {ExcelLocale::kZhCN, 42, 0.0, " ¥- "},
    {ExcelLocale::kZhCN, 43, 0.0, " -   "},
    {ExcelLocale::kZhCN, 47, 1234.567, "36:28.8"},
    {ExcelLocale::kZhCN, 59, 1234.567, "1235"},
    {ExcelLocale::kZhCN, 63, -1234.567, "(¥1,235)"},
    {ExcelLocale::kZhCN, 71, kStamp, "2024/3/4"},
    {ExcelLocale::kKoKR, 5, -1234.567, "(₩1,235)"},
    {ExcelLocale::kKoKR, 14, kStamp, "2024.3.4"},
    {ExcelLocale::kKoKR, 15, kStamp, "4.Mar.24"},
    {ExcelLocale::kKoKR, 22, kStamp, "2024.3.4 13:30"},
    {ExcelLocale::kKoKR, 27, kStamp, "2024年 3月 4日"},
    {ExcelLocale::kKoKR, 34, kStamp, "2024/3/4"},
    {ExcelLocale::kKoKR, 42, -1234.567, " ₩(1,235)"},
    {ExcelLocale::kKoKR, 42, 0.0, " ₩- "},
    {ExcelLocale::kKoKR, 43, 0.0, " -   "},
    {ExcelLocale::kKoKR, 47, 1234.567, "36:28.8"},
    {ExcelLocale::kKoKR, 59, 1234.567, "1235"},
    {ExcelLocale::kKoKR, 63, -1234.567, "(₩1,235)"},
    {ExcelLocale::kKoKR, 71, kStamp, "2024.3.4"},
    {ExcelLocale::kThTH, 5, -1234.567, "(฿1,235)"},
    {ExcelLocale::kThTH, 14, kStamp, "4/3/24"},
    {ExcelLocale::kThTH, 15, kStamp, "4-มี.ค.-24"},
    {ExcelLocale::kThTH, 22, kStamp, "4/3/24 13:30"},
    {ExcelLocale::kThTH, 27, kStamp, "4/3/24"},
    {ExcelLocale::kThTH, 34, kStamp, "13:30:18"},
    {ExcelLocale::kThTH, 42, -1234.567, " ฿(1,235)"},
    {ExcelLocale::kThTH, 42, 0.0, " ฿- "},
    {ExcelLocale::kThTH, 43, 0.0, " -   "},
    {ExcelLocale::kThTH, 47, 1234.567, "36:28.8"},
    {ExcelLocale::kThTH, 59, 1234.567, "๑๒๓๕"},
    {ExcelLocale::kThTH, 63, -1234.567, "(฿๑,๒๓๕)"},
    {ExcelLocale::kThTH, 71, kStamp, "๔/๓/๖๗"},
    {ExcelLocale::kEsES, 5, -1234.567, "-1.235 €"},
    {ExcelLocale::kEsES, 14, kStamp, "4/3/24"},
    {ExcelLocale::kEsES, 15, kStamp, "4-mar-24"},
    {ExcelLocale::kEsES, 22, kStamp, "4/3/24 13:30"},
    {ExcelLocale::kEsES, 27, kStamp, "4/3/24"},
    {ExcelLocale::kEsES, 34, kStamp, "13:30:18"},
    {ExcelLocale::kEsES, 42, -1234.567, "-1.235 € "},
    {ExcelLocale::kEsES, 42, 0.0, " - € "},
    {ExcelLocale::kEsES, 43, 0.0, " -   "},
    {ExcelLocale::kEsES, 47, 1234.567, "36:28,8"},
    {ExcelLocale::kEsES, 59, 1234.567, "1235"},
    {ExcelLocale::kEsES, 63, -1234.567, "-1.235 €"},
    {ExcelLocale::kEsES, 71, kStamp, "4/3/24"},
    {ExcelLocale::kEsMX, 5, -1234.567, "-$1,235"},
    {ExcelLocale::kEsMX, 14, kStamp, "04/03/24"},
    {ExcelLocale::kEsMX, 15, kStamp, "04-mar-24"},
    {ExcelLocale::kEsMX, 22, kStamp, "04/03/24 13:30"},
    {ExcelLocale::kEsMX, 27, kStamp, "04/03/24"},
    {ExcelLocale::kEsMX, 34, kStamp, "13:30:18"},
    {ExcelLocale::kEsMX, 42, -1234.567, "-$1,235 "},
    {ExcelLocale::kEsMX, 42, 0.0, " $- "},
    {ExcelLocale::kEsMX, 43, 0.0, " -   "},
    {ExcelLocale::kEsMX, 47, 1234.567, "36:28.8"},
    {ExcelLocale::kEsMX, 59, 1234.567, "1235"},
    {ExcelLocale::kEsMX, 63, -1234.567, "-$1,235"},
    {ExcelLocale::kEsMX, 71, kStamp, "04/03/24"},
    {ExcelLocale::kPtBR, 5, -1234.567, "-R$ 1.235"},
    {ExcelLocale::kPtBR, 14, kStamp, "04/03/2024"},
    {ExcelLocale::kPtBR, 15, kStamp, "04/mar/24"},
    {ExcelLocale::kPtBR, 22, kStamp, "04/03/2024 13:30"},
    {ExcelLocale::kPtBR, 27, kStamp, "04/03/2024"},
    {ExcelLocale::kPtBR, 34, kStamp, "13:30:18"},
    {ExcelLocale::kPtBR, 42, -1234.567, "-R$ 1.235 "},
    {ExcelLocale::kPtBR, 42, 0.0, " R$ - "},
    {ExcelLocale::kPtBR, 43, 0.0, " -   "},
    {ExcelLocale::kPtBR, 47, 1234.567, "36:28,8"},
    {ExcelLocale::kPtBR, 59, 1234.567, "1235"},
    {ExcelLocale::kPtBR, 63, -1234.567, "-R$ 1.235"},
    {ExcelLocale::kPtBR, 71, kStamp, "04/03/2024"},
    {ExcelLocale::kRuRU, 5, -1234.567, "-1 235 ₽"},
    {ExcelLocale::kRuRU, 14, kStamp, "04.03.2024"},
    {ExcelLocale::kRuRU, 15, kStamp, "04.март.24"},
    {ExcelLocale::kRuRU, 22, kStamp, "04.03.2024 13:30"},
    {ExcelLocale::kRuRU, 27, kStamp, "04.03.2024"},
    {ExcelLocale::kRuRU, 34, kStamp, "13:30:18"},
    {ExcelLocale::kRuRU, 42, -1234.567, "-1 235 ₽ "},
    {ExcelLocale::kRuRU, 42, 0.0, " - ₽ "},
    {ExcelLocale::kRuRU, 43, 0.0, " -   "},
    {ExcelLocale::kRuRU, 47, 1234.567, "36:28,8"},
    {ExcelLocale::kRuRU, 59, 1234.567, "1235"},
    {ExcelLocale::kRuRU, 63, -1234.567, "-1 235 ₽"},
    {ExcelLocale::kRuRU, 71, kStamp, "04.03.2024"},
    {ExcelLocale::kZhTW, 5, -1234.567, "($1,235)"},
    {ExcelLocale::kZhTW, 14, kStamp, "2024/3/4"},
    {ExcelLocale::kZhTW, 15, kStamp, "4-Mar-24"},
    {ExcelLocale::kZhTW, 22, kStamp, "2024/3/4 13:30"},
    {ExcelLocale::kZhTW, 27, kStamp, "2024/3/4"},
    {ExcelLocale::kZhTW, 34, kStamp, "下午1時30分"},
    {ExcelLocale::kZhTW, 42, -1234.567, " $(1,235)"},
    {ExcelLocale::kZhTW, 42, 0.0, " $- "},
    {ExcelLocale::kZhTW, 43, 0.0, " -   "},
    {ExcelLocale::kZhTW, 47, 1234.567, "36:28.8"},
    {ExcelLocale::kZhTW, 59, 1234.567, "1235"},
    {ExcelLocale::kZhTW, 63, -1234.567, "($1,235)"},
    {ExcelLocale::kZhTW, 71, kStamp, "2024/3/4"},
    {ExcelLocale::kItIT, 5, -1234.567, "-1.235 €"},
    {ExcelLocale::kItIT, 14, kStamp, "04/03/24"},
    {ExcelLocale::kItIT, 15, kStamp, "04-mar-24"},
    {ExcelLocale::kItIT, 22, kStamp, "04/03/24 13:30"},
    {ExcelLocale::kItIT, 27, kStamp, "04/03/24"},
    {ExcelLocale::kItIT, 34, kStamp, "13:30:18"},
    {ExcelLocale::kItIT, 42, -1234.567, "-1.235 € "},
    {ExcelLocale::kItIT, 42, 0.0, " - € "},
    {ExcelLocale::kItIT, 43, 0.0, " -   "},
    {ExcelLocale::kItIT, 47, 1234.567, "36:28,8"},
    {ExcelLocale::kItIT, 59, 1234.567, "1235"},
    {ExcelLocale::kItIT, 63, -1234.567, "-1.235 €"},
    {ExcelLocale::kItIT, 71, kStamp, "04/03/24"},
    {ExcelLocale::kNlNL, 5, -1234.567, "(€ 1.235)"},
    {ExcelLocale::kNlNL, 14, kStamp, "04-03-2024"},
    {ExcelLocale::kNlNL, 15, kStamp, "04-mrt-24"},
    {ExcelLocale::kNlNL, 22, kStamp, "04-03-2024 13:30"},
    {ExcelLocale::kNlNL, 27, kStamp, "04-03-2024"},
    {ExcelLocale::kNlNL, 34, kStamp, "13:30:18"},
    {ExcelLocale::kNlNL, 42, -1234.567, " € (1.235)"},
    {ExcelLocale::kNlNL, 42, 0.0, " € - "},
    {ExcelLocale::kNlNL, 43, 0.0, " -   "},
    {ExcelLocale::kNlNL, 47, 1234.567, "36:28,8"},
    {ExcelLocale::kNlNL, 59, 1234.567, "1235"},
    {ExcelLocale::kNlNL, 63, -1234.567, "(€ 1.235)"},
    {ExcelLocale::kNlNL, 71, kStamp, "04-03-2024"},
};

text_format::DisplayText RenderBuiltin(std::uint16_t id, double value, ExcelProfile profile) {
  const StylesTable styles;
  const std::optional<std::string_view> code = effective_num_fmt(styles, id, profile);
  return text_format::format_value_for_display(Value::number(value), code ? *code : std::string_view{}, false, profile);
}

TEST(NumFmtLocale, MatchesMeasuredDisplayOnBothHosts) {
  for (const MeasuredRow& row : kMeasured) {
    for (const ExcelHost host : {ExcelHost::kMac365, ExcelHost::kWin365}) {
      const ExcelProfile profile{host, row.locale};
      SCOPED_TRACE(excel_profile_id(profile));
      SCOPED_TRACE(row.id);
      EXPECT_EQ(RenderBuiltin(row.id, row.value, profile).text, row.text);
    }
  }
}

TEST(NumFmtLocale, IdsWithoutCodeRenderAsGeneral) {
  for (const std::uint16_t id : {std::uint16_t{23}, std::uint16_t{100}, std::uint16_t{163}, std::uint16_t{164}}) {
    EXPECT_TRUE(localized_builtin_num_fmt(id, mac_365_en_us_profile()).empty()) << id;
  }
  // ja-JP gives 23-26 the dollar forms.
  EXPECT_EQ(RenderBuiltin(23, -1234.567, mac_365_ja_jp_profile()).text, "($1,235)");
}

TEST(NumFmtLocale, BuiltinIdKnownInAnyLocale) {
  for (const std::uint16_t id : {std::uint16_t{0}, std::uint16_t{5}, std::uint16_t{27}, std::uint16_t{44},
                                 std::uint16_t{59}, std::uint16_t{81}}) {
    EXPECT_TRUE(is_builtin_num_fmt_id(id)) << id;
  }
  for (const std::uint16_t id : {std::uint16_t{82}, std::uint16_t{163}, std::uint16_t{164}}) {
    EXPECT_FALSE(is_builtin_num_fmt_id(id)) << id;
  }
}

TEST(NumFmtLocale, ThaiDigitIdsAreThaiOnlyInThTh) {
  const ExcelProfile th{ExcelHost::kMac365, ExcelLocale::kThTH};
  EXPECT_EQ(RenderBuiltin(59, 1234.567, th).text, "๑๒๓๕");
  EXPECT_EQ(RenderBuiltin(81, kStamp, th).text, "4/3/67");
  EXPECT_EQ(RenderBuiltin(81, kStamp, mac_365_en_us_profile()).text, "30:18.0");
}

TEST(NumFmtLocale, StoredResolutionKeepsEcmaTable) {
  const StylesTable styles;
  EXPECT_EQ(effective_num_fmt(styles, 14U), std::optional<std::string_view>("mm-dd-yy"));
  EXPECT_EQ(effective_num_fmt(styles, 14U, mac_365_ja_jp_profile()), std::optional<std::string_view>("yyyy/m/d"));
  EXPECT_EQ(effective_num_fmt(styles, 5U), std::nullopt);
}

TEST(NumFmtLocale, CustomRecordOverridesLocalizedBuiltin) {
  StylesTable styles;
  styles.num_fmt_strings.emplace_back("d mmm");
  styles.num_fmts.push_back(NumFmtRecord{14U, 0U});
  EXPECT_EQ(effective_num_fmt(styles, 14U, mac_365_ja_jp_profile()), std::optional<std::string_view>("d mmm"));
}

TEST(NumFmtLocale, CellDisplayUsesWorkbookLocale) {
  struct Case {
    ExcelProfile profile;
    const char* text;
  };
  const Case cases[] = {
      {mac_365_ja_jp_profile(), "2024/3/4"},
      {win_365_ja_jp_profile(), "2024/3/4"},
      {mac_365_en_us_profile(), "3/4/24"},
      {ExcelProfile{ExcelHost::kMac365, ExcelLocale::kDeDE}, "04.03.24"},
      {ExcelProfile{ExcelHost::kMac365, ExcelLocale::kKoKR}, "2024.3.4"},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(excel_profile_id(c.profile));
    Workbook workbook = Workbook::create();
    workbook.set_excel_profile(c.profile);
    StylesTable& styles = workbook.mutable_styles();
    CellXf xf;
    xf.num_fmt_id = 14U;
    xf.apply_number_format = true;
    styles.cell_xfs.push_back(xf);
    ASSERT_TRUE(static_cast<bool>(workbook.set_cell_value(0U, 0U, 0U, Value::number(kStamp))));
    ASSERT_TRUE(static_cast<bool>(
        workbook.set_cell_xf_index(0U, 0U, 0U, static_cast<std::uint32_t>(styles.cell_xfs.size() - 1U))));
    EXPECT_EQ(text_format::format_cell_for_display(workbook, workbook.sheet(0), 0U, 0U).text, c.text);
    EXPECT_EQ(workbook.styles().cell_xfs.back().num_fmt_id, 14U);
  }
}

}  // namespace
}  // namespace formulon
