// Save-time cell style plan, checked against Mac Excel 365 ja-JP re-saves
// (case ids from the cellstyle measurement: cs1..cs5).

#include "cell_style_save_plan.h"

#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <string>
#include <utility>
#include <vector>

#include "eval/collation.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "sheet_layout.h"
#include "styles.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

constexpr std::uint32_t kNo = CellStyleRecord::kBuiltinIdNone;

CellStyleRecord Style(const char* name, std::uint32_t xf, std::uint32_t builtin = kNo, std::uint32_t level = 0U,
                      bool hidden = false, bool custom_builtin = false) {
  CellStyleRecord record;
  record.name = name;
  record.xf_id = xf;
  record.builtin_id = builtin;
  record.i_level = level;
  record.hidden = hidden;
  record.custom_builtin = custom_builtin;
  return record;
}

std::vector<std::uint32_t> Iota(std::uint32_t count) {
  std::vector<std::uint32_t> out(count);
  for (std::uint32_t i = 0; i < count; ++i) {
    out[i] = i;
  }
  return out;
}

// `cell_xf_parents[c]` is cell xf c's style xf.
StylesTable Table(std::vector<CellStyleRecord> records, std::uint32_t style_xf_count,
                  const std::vector<std::uint32_t>& cell_xf_parents) {
  StylesTable table;
  table.cell_styles = std::move(records);
  table.cell_style_xfs.assign(style_xf_count, CellXf{});
  for (const std::uint32_t parent : cell_xf_parents) {
    CellXf xf;
    xf.xf_id = parent;
    table.cell_xfs.push_back(xf);
  }
  return table;
}

UsedCellXfs Used(std::size_t count, std::initializer_list<std::uint32_t> on) {
  UsedCellXfs used;
  used.used.assign(count, false);
  for (const std::uint32_t c : on) {
    used.used[c] = true;
  }
  return used;
}

UsedCellXfs Every(std::size_t count) {
  UsedCellXfs used;
  used.used.assign(count, true);
  return used;
}

struct WantStyle {
  const char* name;
  std::uint32_t xf;
  std::uint32_t builtin;
  std::uint32_t level;
  bool hidden;
  bool custom_builtin;
};

WantStyle Want(const char* name, std::uint32_t xf, std::uint32_t builtin = kNo, std::uint32_t level = 0U,
               bool hidden = false, bool custom_builtin = false) {
  return WantStyle{name, xf, builtin, level, hidden, custom_builtin};
}

std::string Describe(const CellStyleSavePlan& plan) {
  std::string out;
  for (const CellStyleRecord& style : plan.cell_styles()) {
    out += "[" + style.name + " xf=" + std::to_string(style.xf_id) + "]";
  }
  return out;
}

void ExpectStyles(const CellStyleSavePlan& plan, const std::vector<WantStyle>& want) {
  const std::vector<CellStyleRecord>& got = plan.cell_styles();
  ASSERT_EQ(got.size(), want.size()) << Describe(plan);
  for (std::size_t i = 0; i < want.size(); ++i) {
    EXPECT_EQ(got[i].name, want[i].name) << i << " " << Describe(plan);
    EXPECT_EQ(got[i].xf_id, want[i].xf) << i << " " << Describe(plan);
    EXPECT_EQ(got[i].builtin_id, want[i].builtin) << i << " " << want[i].name;
    EXPECT_EQ(got[i].i_level, want[i].level) << i << " " << want[i].name;
    EXPECT_EQ(got[i].hidden, want[i].hidden) << i << " " << want[i].name;
    EXPECT_EQ(got[i].custom_builtin, want[i].custom_builtin) << i << " " << want[i].name;
  }
}

void ExpectParents(const CellStyleSavePlan& plan, const std::vector<std::uint32_t>& old_to_new) {
  for (std::uint32_t x = 0; x < old_to_new.size(); ++x) {
    EXPECT_EQ(plan.cell_xf_parent(x), old_to_new[x]) << "old style xf " << x;
  }
}

// The authored table shared by the cs1 cases: style xf i and cell xf i carry number format i.
StylesTable Cs1Authored() {
  return Table({Style("Heading 2", 13, 17), Style("Total", 1, 25), Style("aaa", 2), Style("標準", 0, 0),
                Style("Comma", 3, 3), Style("zzz", 4), Style("見出し 2", 12), Style("Heading 1", 5, 16),
                Style("Good", 6, 26), Style("Hyperlink", 7, 8), Style("RowLevel_1", 8, 1, 0),
                Style("Percent", 9, 5, 0, false, true), Style("Bad", 10, 27, 0, true), Style("Foo", 11, 4)},
               14, Iota(14));
}

TEST(CellStyleSavePlan, Cs1C1EveryCellXfUsed) {
  const CellStyleSavePlan plan = build_cell_style_save_plan(Cs1Authored(), Every(14), mac_365_ja_jp_profile());
  ExpectStyles(plan, {Want("aaa", 3), Want("zzz", 5), Want("パーセント", 9, 5, 0, false, true),
                      Want("ハイパーリンク", 8, 8), Want("悪い", 10, 27, 0, true), Want("桁区切り [0.00]", 4, 3),
                      Want("見出し 1", 6, 16), Want("見出し 2", 12, 17), Want("行レベル_1", 1, 1, 0),
                      Want("集計", 2, 25), Want("通貨 [0.00]", 11, 4), Want("標準", 0, 0), Want("良い", 7, 26)});
  EXPECT_EQ(plan.style_xf_order(), (std::vector<std::uint32_t>{0, 8, 1, 2, 3, 4, 5, 6, 7, 9, 10, 11, 12}));
  // Cell xf 13 named the merged custom's partner; both now name the survivor.
  ExpectParents(plan, {0, 2, 3, 4, 5, 6, 7, 8, 1, 9, 10, 11, 12, 12});
  EXPECT_FALSE(plan.is_identity());
}

TEST(CellStyleSavePlan, Cs1C0UnusedBuiltinsDrop) {
  const CellStyleSavePlan plan = build_cell_style_save_plan(Cs1Authored(), Used(14, {0}), mac_365_ja_jp_profile());
  ExpectStyles(plan, {Want("aaa", 1), Want("zzz", 2), Want("パーセント", 3, 5, 0, false, true),
                      Want("悪い", 4, 27, 0, true), Want("標準", 0, 0)});
  EXPECT_EQ(plan.style_xf_order(), (std::vector<std::uint32_t>{0, 2, 4, 9, 10}));
  // A dropped destination maps to 0 (only unused cell xfs reach one).
  EXPECT_EQ(plan.cell_xf_parent(13), 0U);
  EXPECT_EQ(plan.cell_xf_parent(5), 0U);
  EXPECT_EQ(plan.cell_xf_parent(9), 3U);
}

TEST(CellStyleSavePlan, Cs1C3PartlyUsedKeepsUsedBuiltins) {
  const CellStyleSavePlan plan =
      build_cell_style_save_plan(Cs1Authored(), Used(14, {0, 2, 3, 5}), mac_365_ja_jp_profile());
  ExpectStyles(plan,
               {Want("aaa", 1), Want("zzz", 3), Want("パーセント", 5, 5, 0, false, true), Want("悪い", 6, 27, 0, true),
                Want("桁区切り [0.00]", 2, 3), Want("見出し 1", 4, 16), Want("標準", 0, 0)});
  EXPECT_EQ(plan.style_xf_order(), (std::vector<std::uint32_t>{0, 2, 3, 4, 5, 9, 10}));
  EXPECT_EQ(plan.cell_xf_parent(2), 1U);
  EXPECT_EQ(plan.cell_xf_parent(3), 2U);
  EXPECT_EQ(plan.cell_xf_parent(5), 4U);
}

TEST(CellStyleSavePlan, Cs1C1ResavedIsIdentity) {
  const StylesTable resaved =
      Table({Style("aaa", 3), Style("zzz", 5), Style("パーセント", 9, 5, 0, false, true), Style("ハイパーリンク", 8, 8),
             Style("悪い", 10, 27, 0, true), Style("桁区切り [0.00]", 4, 3), Style("見出し 1", 6, 16),
             Style("見出し 2", 12, 17), Style("行レベル_1", 1, 1, 0), Style("集計", 2, 25), Style("通貨 [0.00]", 11, 4),
             Style("標準", 0, 0), Style("良い", 7, 26)},
            13, {0, 2, 3, 4, 5, 6, 7, 8, 1, 9, 10, 11, 12, 12});
  const CellStyleSavePlan plan = build_cell_style_save_plan(resaved, Every(14), mac_365_ja_jp_profile());
  EXPECT_TRUE(plan.is_identity()) << Describe(plan);
  ExpectParents(plan, Iota(13));
}

TEST(CellStyleSavePlan, IdentityBrokenByNameOrderMergeAndDrop) {
  const StylesTable base = Table({Style("aaa", 1), Style("標準", 0, 0)}, 2, {0, 1});
  EXPECT_TRUE(build_cell_style_save_plan(base, Used(2, {0}), mac_365_ja_jp_profile()).is_identity());

  StylesTable english = base;
  english.cell_styles[1].name = "Normal";
  EXPECT_FALSE(build_cell_style_save_plan(english, Used(2, {0}), mac_365_ja_jp_profile()).is_identity());

  const StylesTable unsorted = Table({Style("標準", 0, 0), Style("aaa", 1)}, 2, {0, 1});
  EXPECT_FALSE(build_cell_style_save_plan(unsorted, Used(2, {0}), mac_365_ja_jp_profile()).is_identity());

  const StylesTable xf_order = Table({Style("aaa", 0), Style("標準", 1, 0)}, 2, {1, 0});
  EXPECT_FALSE(build_cell_style_save_plan(xf_order, Used(2, {0}), mac_365_ja_jp_profile()).is_identity());

  const StylesTable merged = Table({Style("ｶ", 1), Style("か", 2), Style("標準", 0, 0)}, 3, {0, 1, 2});
  EXPECT_FALSE(build_cell_style_save_plan(merged, Every(3), mac_365_ja_jp_profile()).is_identity());

  const StylesTable dropped = Table({Style("良い", 1, 26), Style("標準", 0, 0)}, 2, {0, 1});
  EXPECT_FALSE(build_cell_style_save_plan(dropped, Used(2, {0}), mac_365_ja_jp_profile()).is_identity());
}

TEST(CellStyleSavePlan, Cs2X1OutlineXfsFollowNormalByLevelRowFirst) {
  const StylesTable table =
      Table({Style("標準", 0, 0), Style("Total", 1, 25), Style("aaa", 2), Style("ColLevel_1", 3, 2, 0), Style("zzz", 4),
             Style("RowLevel_2", 5, 1, 1), Style("Comma", 6, 3), Style("RowLevel_1", 7, 1, 0), Style("Good", 8, 26)},
            9, Iota(9));
  const CellStyleSavePlan plan = build_cell_style_save_plan(table, Every(9), mac_365_ja_jp_profile());
  ExpectStyles(plan, {Want("aaa", 5), Want("zzz", 6), Want("桁区切り [0.00]", 7, 3), Want("行レベル_1", 1, 1, 0),
                      Want("行レベル_2", 3, 1, 1), Want("集計", 4, 25), Want("標準", 0, 0), Want("良い", 8, 26),
                      Want("列レベル_1", 2, 2, 0)});
  EXPECT_EQ(plan.style_xf_order(), (std::vector<std::uint32_t>{0, 7, 3, 5, 1, 2, 4, 6, 8}));
  ExpectParents(plan, {0, 4, 5, 2, 6, 3, 7, 1, 8});
}

TEST(CellStyleSavePlan, Cs2X3SortAndKanaWidthMerge) {
  const StylesTable table =
      Table({Style("標準", 0, 0), Style("b", 1), Style("B2", 2), Style("a10", 3), Style("a9", 4), Style("Ａx", 5),
             Style("_x", 6), Style("1x", 7), Style("ä", 8), Style("ab", 9), Style("Ac", 10), Style("ｶ", 11),
             Style("カ", 12), Style("か", 13), Style("漢", 14), Style("Z", 15), Style("-y", 16), Style(" sp", 17)},
            18, Iota(18));
  const CellStyleSavePlan plan = build_cell_style_save_plan(table, Used(18, {0}), mac_365_ja_jp_profile());
  // ｶ / カ / か are one style: the first's name, the last's xf.
  ExpectStyles(plan, {Want(" sp", 15), Want("_x", 6), Want("-y", 14), Want("1x", 7), Want("ä", 8), Want("a10", 3),
                      Want("a9", 4), Want("ab", 9), Want("Ac", 10), Want("Ａx", 5), Want("b", 1), Want("B2", 2),
                      Want("Z", 13), Want("ｶ", 11), Want("漢", 12), Want("標準", 0, 0)});
  EXPECT_EQ(plan.style_xf_order(), (std::vector<std::uint32_t>{0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 13, 14, 15, 16, 17}));
  EXPECT_EQ(plan.cell_xf_parent(11), 11U);
  EXPECT_EQ(plan.cell_xf_parent(12), 11U);
  EXPECT_EQ(plan.cell_xf_parent(13), 11U);
}

TEST(CellStyleSavePlan, Cs2X4RepresentativeFlagsKeepUnusedStyles) {
  const StylesTable table = Table({Style("標準", 0, 0), Style("Good", 1, 26, 0, true), Style("Neutral", 2, 28),
                                   Style("c", 3, kNo, 0, true), Style("Bad", 4, 27)},
                                  5, Iota(5));
  const CellStyleSavePlan plan = build_cell_style_save_plan(table, Used(5, {0}), mac_365_ja_jp_profile());
  ExpectStyles(plan, {Want("c", 2, kNo, 0, true), Want("標準", 0, 0), Want("良い", 1, 26, 0, true)});
  EXPECT_EQ(plan.style_xf_order(), (std::vector<std::uint32_t>{0, 1, 3}));
}

TEST(CellStyleSavePlan, FlagsOfLaterMergedRecordDoNotKeepGroup) {
  // The group takes the first record's flags; a hidden or edited later twin does not keep it.
  const StylesTable table =
      Table({Style("標準", 0, 0), Style("Good", 1, 26), Style("良い", 2, kNo, 0, true, true)}, 3, Iota(3));
  const CellStyleSavePlan plan = build_cell_style_save_plan(table, Used(3, {0}), mac_365_ja_jp_profile());
  ExpectStyles(plan, {Want("標準", 0, 0)});
  EXPECT_EQ(plan.style_xf_order(), (std::vector<std::uint32_t>{0}));
}

TEST(CellStyleSavePlan, Cs4R6BuiltinThenCustomStaysBuiltin) {
  const StylesTable table = Table({Style("標準", 0, 0), Style("Heading 2", 1, 17), Style("見出し 2", 2)}, 3, Iota(3));
  const CellStyleSavePlan plan = build_cell_style_save_plan(table, Every(3), mac_365_ja_jp_profile());
  ExpectStyles(plan, {Want("見出し 2", 1, 17), Want("標準", 0, 0)});
  EXPECT_EQ(plan.style_xf_order(), (std::vector<std::uint32_t>{0, 2}));
  ExpectParents(plan, {0, 1, 1});
}

TEST(CellStyleSavePlan, Cs4R7CustomThenBuiltinStaysCustom) {
  const StylesTable table = Table({Style("標準", 0, 0), Style("見出し 2", 1), Style("Heading 2", 2, 17)}, 3, Iota(3));
  const CellStyleSavePlan plan = build_cell_style_save_plan(table, Every(3), mac_365_ja_jp_profile());
  ExpectStyles(plan, {Want("見出し 2", 1), Want("標準", 0, 0)});
  EXPECT_EQ(plan.style_xf_order(), (std::vector<std::uint32_t>{0, 2}));
  ExpectParents(plan, {0, 1, 1});
}

TEST(CellStyleSavePlan, Cs4R8CustomThenBuiltinKeptUnused) {
  const StylesTable table = Table({Style("標準", 0, 0), Style("見出し 2", 1), Style("Heading 2", 2, 17)}, 3, Iota(3));
  const CellStyleSavePlan plan = build_cell_style_save_plan(table, Used(3, {0}), mac_365_ja_jp_profile());
  ExpectStyles(plan, {Want("見出し 2", 1), Want("標準", 0, 0)});
  EXPECT_EQ(plan.style_xf_order(), (std::vector<std::uint32_t>{0, 2}));
}

TEST(CellStyleSavePlan, Cs4R4SameBuiltinIdTwiceMerges) {
  const StylesTable table =
      Table({Style("標準", 0, 0), Style("見出し 2", 1, 17), Style("Heading 2", 2, 17)}, 3, Iota(3));
  const CellStyleSavePlan plan = build_cell_style_save_plan(table, Every(3), mac_365_ja_jp_profile());
  ExpectStyles(plan, {Want("見出し 2", 1, 17), Want("標準", 0, 0)});
  EXPECT_EQ(plan.style_xf_order(), (std::vector<std::uint32_t>{0, 2}));
  ExpectParents(plan, {0, 1, 1});
}

TEST(CellStyleSavePlan, Cs4R3HalfWidthAndHiraganaMerge) {
  const StylesTable table = Table({Style("標準", 0, 0), Style("ｶ", 1), Style("か", 2)}, 3, Iota(3));
  const CellStyleSavePlan plan = build_cell_style_save_plan(table, Every(3), mac_365_ja_jp_profile());
  ExpectStyles(plan, {Want("ｶ", 1), Want("標準", 0, 0)});
  ExpectParents(plan, {0, 1, 1});
}

TEST(CellStyleSavePlan, Cs4R1R2LocaleSpellingWithoutBuiltinIdIsCustom) {
  const StylesTable table = Table({Style("標準", 0, 0), Style("見出し 3", 1), Style("Heading 3", 2)}, 3, {0});
  const CellStyleSavePlan plan = build_cell_style_save_plan(table, Used(1, {0}), mac_365_ja_jp_profile());
  ExpectStyles(plan, {Want("Heading 3", 2), Want("見出し 3", 1), Want("標準", 0, 0)});
}

TEST(CellStyleSavePlan, SharedStyleXfGoesToFirstRecordsGroup) {
  // Style xf 1 is named by "Heading 1" (merged with a later custom on xf 2) and by "foo".
  const StylesTable table =
      Table({Style("標準", 0, 0), Style("Heading 1", 1, 16), Style("foo", 1), Style("見出し 1", 2)}, 3, {0, 1});
  const CellStyleSavePlan plan = build_cell_style_save_plan(table, Used(2, {1}), mac_365_ja_jp_profile());
  ExpectStyles(plan, {Want("foo", 1), Want("見出し 1", 2, 16), Want("標準", 0, 0)});
  EXPECT_EQ(plan.style_xf_order(), (std::vector<std::uint32_t>{0, 1, 2}));
  ExpectParents(plan, {0, 2, 2});
}

TEST(CellStyleSavePlan, OrphanStyleXfKeptOnlyWhenUsed) {
  // Style xfs 2 and 3 have no cell style; cell xf 2 uses 2, cell xf 3 (unused) names 3.
  const StylesTable table = Table({Style("標準", 0, 0), Style("a", 1)}, 4, {0, 1, 2, 3});
  const CellStyleSavePlan plan = build_cell_style_save_plan(table, Used(4, {0, 2}), mac_365_ja_jp_profile());
  ExpectStyles(plan, {Want("a", 1), Want("標準", 0, 0)});
  EXPECT_EQ(plan.style_xf_order(), (std::vector<std::uint32_t>{0, 1, 2}));
  ExpectParents(plan, {0, 1, 2, 0});
  EXPECT_FALSE(plan.is_identity());
}

TEST(CellStyleSavePlan, EmptyTablesSynthesizeNormal) {
  const StylesTable table;
  const CellStyleSavePlan plan = build_cell_style_save_plan(table, UsedCellXfs{}, mac_365_ja_jp_profile());
  ExpectStyles(plan, {Want("標準", 0, 0)});
  EXPECT_EQ(plan.style_xf_order(), (std::vector<std::uint32_t>{0}));
  EXPECT_EQ(plan.cell_xf_parent(0), 0U);
  EXPECT_EQ(plan.cell_xf_parent(7), 0U);
  EXPECT_TRUE(plan.is_identity());
}

TEST(CellStyleSavePlan, MissingNormalIsSynthesizedFirst) {
  const StylesTable table = Table({Style("aaa", 1)}, 2, {0, 1});
  const CellStyleSavePlan plan = build_cell_style_save_plan(table, Used(2, {0}), mac_365_ja_jp_profile());
  ExpectStyles(plan, {Want("aaa", 1), Want("標準", 0, 0)});
  EXPECT_EQ(plan.style_xf_order(), (std::vector<std::uint32_t>{0, 1}));
  EXPECT_FALSE(plan.is_identity());
}

TEST(CellStyleSavePlan, Cs5N1CustomHyojunBeforeNormalMergesIntoNormal) {
  const StylesTable table = Table({Style("標準", 1), Style("Normal", 0, 0)}, 2, {0, 1});
  const CellStyleSavePlan plan = build_cell_style_save_plan(table, Every(2), mac_365_ja_jp_profile());
  ExpectStyles(plan, {Want("標準", 0, 0)});
  EXPECT_EQ(plan.style_xf_order(), (std::vector<std::uint32_t>{0}));
  ExpectParents(plan, {0, 0});
}

TEST(CellStyleSavePlan, AllUsedKeepsEveryGroup) {
  const StylesTable table = Table({Style("標準", 0, 0), Style("Good", 1, 26)}, 2, {0});
  UsedCellXfs used;
  used.all_used = true;
  const CellStyleSavePlan plan = build_cell_style_save_plan(table, used, mac_365_ja_jp_profile());
  ExpectStyles(plan, {Want("標準", 0, 0), Want("良い", 1, 26)});
  EXPECT_EQ(plan.style_xf_order(), (std::vector<std::uint32_t>{0, 1}));
}

TEST(CellStyleSavePlan, SortTiesKeepSourceOrder) {
  // Under zh, A and its full-width form are distinct styles but sort equal.
  const ExcelProfile zh{ExcelHost::kMac365, ExcelLocale::kZhCN};
  ASSERT_NE(eval::collate("A", "Ａ", zh, eval::CollationStrength::kCompare), 0);
  ASSERT_EQ(eval::collate("A", "Ａ", zh, eval::CollationStrength::kDynamic), 0);
  const auto order_of = [&zh](const char* first, const char* second) {
    const StylesTable table = Table({Style("Normal", 0, 0), Style(first, 1), Style(second, 2)}, 3, {0});
    const CellStyleSavePlan plan = build_cell_style_save_plan(table, Used(1, {0}), zh);
    std::vector<std::string> customs;
    for (const CellStyleRecord& style : plan.cell_styles()) {
      if (style.builtin_id == kNo) {
        customs.push_back(style.name);
      }
    }
    return customs;
  };
  EXPECT_EQ(order_of("Ａ", "A"), (std::vector<std::string>{"Ａ", "A"}));
  EXPECT_EQ(order_of("A", "Ａ"), (std::vector<std::string>{"A", "Ａ"}));
}

TEST(CellStyleSavePlan, SavedFileNamesCellStyle) {
  // Only the custom "mine" and Normal survive; Neutral and Good are dropped as unused.
  const StylesTable table =
      Table({Style("標準", 0, 0), Style("Neutral", 1, 28), Style("Good", 2, 26), Style("mine", 3)}, 4, Iota(4));
  const CellStyleSavePlan plan = build_cell_style_save_plan(table, Used(4, {0}), mac_365_ja_jp_profile());
  // A dropped built-in's locale name still names a style.
  EXPECT_TRUE(saved_file_names_cell_style(plan, "どちらでもない"));
  EXPECT_TRUE(saved_file_names_cell_style(plan, "良い"));
  EXPECT_TRUE(saved_file_names_cell_style(plan, "mine"));
  EXPECT_TRUE(saved_file_names_cell_style(plan, "ＭＩＮＥ"));
  EXPECT_TRUE(saved_file_names_cell_style(plan, "標準"));
  // A built-in the book does not hold names nothing, even under its locale name.
  EXPECT_FALSE(saved_file_names_cell_style(plan, "悪い"));
  EXPECT_FALSE(saved_file_names_cell_style(plan, "行レベル_1"));
  // English spellings and unknown names do not.
  EXPECT_FALSE(saved_file_names_cell_style(plan, "Neutral"));
  EXPECT_FALSE(saved_file_names_cell_style(plan, "Good"));
  EXPECT_FALSE(saved_file_names_cell_style(plan, "NoSuch"));
}

TEST(CellStyleSavePlan, CollectUsedCellXfsReadsCellsRunsRowsAndColumns) {
  Workbook wb = Workbook::create();
  wb.mutable_styles().cell_xfs.resize(8);
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_xf_index(0, 0, 2U);
  sheet.set_cell_xf_index(1, 1, 3U);
  sheet.set_cell_xf_index(2, 2, 99U);
  RowLayout styled_row;
  styled_row.row = 4;
  styled_row.has_style = true;
  styled_row.style_xf = 4U;
  RowLayout plain_row;
  plain_row.row = 5;
  plain_row.style_xf = 5U;
  sheet.mutable_layout().row_overrides = {styled_row, plain_row};
  ColumnLayout styled_col;
  styled_col.first = 3;
  styled_col.last = 3;
  styled_col.has_style = true;
  styled_col.style_xf = 6U;
  sheet.mutable_layout().columns.push_back(styled_col);

  const UsedCellXfs used = collect_used_cell_xfs(wb);
  EXPECT_FALSE(used.all_used);
  ASSERT_EQ(used.used.size(), 8U);
  EXPECT_EQ(used.used, (std::vector<bool>{false, false, true, true, true, false, true, false}));
}

TEST(CellStyleSavePlan, CollectUsedCellXfsOpaqueSheetUsesAll) {
  Workbook wb = Workbook::create();
  const std::size_t index = wb.add_sheet("Macro1");
  wb.sheet(index).set_opaque_ooxml_sheet("xl/macrosheets/sheet1.xml",
                                         "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet");
  EXPECT_TRUE(collect_used_cell_xfs(wb).all_used);
}

// Cell xf compaction. `num_fmts[c]` is cell xf c's number format, `parents[c]` its style xf.
StylesTable CellXfTable(const std::vector<std::uint32_t>& num_fmts, const std::vector<std::uint32_t>& parents = {}) {
  StylesTable table;
  table.cell_style_xfs.assign(1, CellXf{});
  for (std::size_t c = 0; c < num_fmts.size(); ++c) {
    CellXf xf;
    xf.num_fmt_id = static_cast<std::uint16_t>(num_fmts[c]);
    xf.apply_number_format = num_fmts[c] != 0U;
    xf.xf_id = c < parents.size() ? parents[c] : 0U;
    table.cell_xfs.push_back(xf);
  }
  return table;
}

// The cx1 authored table: cell xf k carries number format k unless `overrides` names it.
std::vector<std::uint32_t> Cx1NumFmts(std::initializer_list<std::pair<std::uint32_t, std::uint32_t>> overrides) {
  std::vector<std::uint32_t> out = Iota(13);
  for (const auto& o : overrides) {
    out[o.first] = o.second;
  }
  return out;
}

CellStyleSavePlan Plan(const StylesTable& table, const UsedCellXfs& used) {
  return build_cell_style_save_plan(table, used, mac_365_ja_jp_profile());
}

std::vector<std::uint32_t> Indices(const CellStyleSavePlan& plan, std::uint32_t count) {
  std::vector<std::uint32_t> out(count);
  for (std::uint32_t c = 0; c < count; ++c) {
    out[c] = plan.cell_xf_index(c);
  }
  return out;
}

TEST(CellStyleSavePlan, Cx1D1EqualCellXfsCollapseIntoTheFirst) {
  const CellStyleSavePlan plan = Plan(CellXfTable(Cx1NumFmts({{1, 3}, {2, 3}})), Used(13, {2, 1, 3}));
  // Cell xf 3 carries number format 3 as well, so all three are one group.
  EXPECT_EQ(plan.cell_xf_order(), (std::vector<std::uint32_t>{0, 1}));
  EXPECT_EQ(plan.cell_xf_index(1), 1U);
  EXPECT_EQ(plan.cell_xf_index(2), 1U);
  EXPECT_EQ(plan.cell_xf_index(3), 1U);
  EXPECT_FALSE(plan.cell_xfs_identity());
}

TEST(CellStyleSavePlan, Cx1D3SurvivorsKeepSourceOrder) {
  const CellStyleSavePlan plan = Plan(CellXfTable(Cx1NumFmts({})), Used(13, {9, 4, 7}));
  EXPECT_EQ(plan.cell_xf_order(), (std::vector<std::uint32_t>{0, 4, 7, 9}));
  EXPECT_EQ(plan.cell_xf_index(9), 3U);
  EXPECT_EQ(plan.cell_xf_index(4), 1U);
}

TEST(CellStyleSavePlan, Cx1D7CellXfEqualToXf0CollapsesIntoXf0) {
  const CellStyleSavePlan plan = Plan(CellXfTable(Cx1NumFmts({{5, 0}})), Used(13, {5, 7}));
  EXPECT_EQ(plan.cell_xf_order(), (std::vector<std::uint32_t>{0, 7}));
  EXPECT_EQ(plan.cell_xf_index(5), 0U);
  EXPECT_EQ(plan.cell_xf_index(7), 1U);
}

TEST(CellStyleSavePlan, Cx1D13GroupSitsAtItsUnusedFirstMember) {
  const CellStyleSavePlan plan = Plan(CellXfTable(Cx1NumFmts({{6, 2}})), Used(13, {6, 4}));
  EXPECT_EQ(plan.cell_xf_order(), (std::vector<std::uint32_t>{0, 2, 4}));
  EXPECT_EQ(plan.cell_xf_index(6), 1U);
  EXPECT_EQ(plan.cell_xf_index(4), 2U);
}

TEST(CellStyleSavePlan, Cx1D2ApplyFlagsKeepCellXfsApart) {
  StylesTable table = CellXfTable(Cx1NumFmts({{1, 3}, {2, 3}}));
  table.cell_xfs[2].apply_number_format = false;
  const CellStyleSavePlan plan = Plan(table, Used(13, {1, 2}));
  EXPECT_EQ(plan.cell_xf_order(), (std::vector<std::uint32_t>{0, 1, 2}));
}

TEST(CellStyleSavePlan, CellXfsAllUsedIsIdentity) {
  const CellStyleSavePlan plan = Plan(CellXfTable(Cx1NumFmts({{1, 3}, {2, 3}})), Every(13));
  UsedCellXfs all;
  all.all_used = true;
  const CellStyleSavePlan opaque = Plan(CellXfTable(Cx1NumFmts({{1, 3}, {2, 3}})), all);
  EXPECT_FALSE(plan.cell_xfs_identity());
  EXPECT_TRUE(opaque.cell_xfs_identity());
  EXPECT_EQ(opaque.cell_xf_order(), Iota(13));
  EXPECT_EQ(Indices(opaque, 13), Iota(13));
}

TEST(CellStyleSavePlan, CellXfsEveryXfUsedAndDistinctIsIdentity) {
  const CellStyleSavePlan plan = Plan(CellXfTable(Cx1NumFmts({})), Every(13));
  EXPECT_TRUE(plan.cell_xfs_identity());
  EXPECT_EQ(Indices(plan, 13), Iota(13));
}

TEST(CellStyleSavePlan, CellXfsEmptyTableWritesNoOrder) {
  const CellStyleSavePlan plan = Plan(StylesTable{}, UsedCellXfs{});
  EXPECT_TRUE(plan.cell_xf_order().empty());
  EXPECT_TRUE(plan.cell_xfs_identity());
  EXPECT_EQ(plan.cell_xf_index(5), 5U);
}

TEST(CellStyleSavePlan, CellXfIndexOutsideTheTableIsKept) {
  const CellStyleSavePlan plan = Plan(CellXfTable(Cx1NumFmts({})), Used(13, {4}));
  EXPECT_EQ(plan.cell_xf_index(13), 13U);
  EXPECT_EQ(plan.cell_xf_index(99), 99U);
}

TEST(CellStyleSavePlan, CellXfsCompareTheWrittenFontFillAndBorder) {
  StylesTable table = CellXfTable({0, 3, 3});
  table.fonts.resize(1);
  table.cell_xfs[1].font_index = 5U;
  table.cell_xfs[1].fill_index = 7U;
  table.cell_xfs[1].border_index = 9U;
  const CellStyleSavePlan plan = Plan(table, Used(3, {1, 2}));
  EXPECT_EQ(plan.cell_xf_order(), (std::vector<std::uint32_t>{0, 1}));
  EXPECT_EQ(plan.cell_xf_index(2), 1U);
}

TEST(CellStyleSavePlan, CellXfsWhoseParentsMergeCollapse) {
  StylesTable table = Table({Style("標準", 0, 0), Style("Good", 1, 26), Style("Good", 2, 26)}, 3, {0, 1, 2});
  table.cell_xfs[1].num_fmt_id = 3U;
  table.cell_xfs[2].num_fmt_id = 3U;
  const CellStyleSavePlan plan = Plan(table, Used(3, {1, 2}));
  ASSERT_EQ(plan.cell_xf_parent(1), plan.cell_xf_parent(2));
  EXPECT_EQ(plan.cell_xf_order(), (std::vector<std::uint32_t>{0, 1}));
  EXPECT_EQ(plan.cell_xf_index(2), 1U);
}

TEST(CellStyleSavePlan, CellXfUnderADroppedStyleDoesNotCollapseIntoOneUnderNormal) {
  // Cell xf 1 is unused under Good, which drops; cell xf 3 is used under Normal with the same content.
  StylesTable table = Table({Style("標準", 0, 0), Style("Good", 1, 26)}, 2, {0, 1, 0, 0});
  table.cell_xfs[1].num_fmt_id = 5U;
  table.cell_xfs[2].num_fmt_id = 7U;
  table.cell_xfs[3].num_fmt_id = 5U;
  const CellStyleSavePlan plan = Plan(table, Used(4, {2, 3}));
  ASSERT_EQ(plan.cell_xf_parent(1), 0U);
  EXPECT_EQ(plan.cell_xf_order(), (std::vector<std::uint32_t>{0, 2, 3}));
  EXPECT_EQ(plan.cell_xf_index(3), 2U);
}

TEST(CellStyleSavePlan, CellXfsDifferingInExtLstStayApart) {
  StylesTable table = CellXfTable({0, 3, 3});
  table.cell_xfs[2].trailing_children_xml = "<extLst><ext uri=\"{X}\"/></extLst>";
  const CellStyleSavePlan plan = Plan(table, Used(3, {1, 2}));
  EXPECT_EQ(plan.cell_xf_order(), (std::vector<std::uint32_t>{0, 1, 2}));
}

TEST(CellStyleSavePlan, CollectUsedCellXfsChartsheetDoesNotUseAll) {
  Workbook wb = Workbook::create();
  const std::size_t index = wb.add_sheet("Chart1");
  wb.sheet(index).set_opaque_ooxml_sheet(
      "xl/chartsheets/sheet1.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet");
  EXPECT_FALSE(collect_used_cell_xfs(wb).all_used);
}

TEST(CellStyleSavePlan, CollectUsedCellXfsRetainedRevisionLogUsesAll) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.add_passthrough_part(PassthroughPart("xl/revisions/revisionLog1.xml",
                                                      "application/vnd.openxmlformats-officedocument."
                                                      "spreadsheetml.revisionLog+xml",
                                                      {}))
                  .has_value());
  EXPECT_TRUE(collect_used_cell_xfs(wb).all_used);
}

TEST(CellStyleSavePlan, CollectUsedCellXfsRetainedMacroSheetPartUsesAll) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.add_passthrough_part(PassthroughPart("xl/macrosheets/sheet1.bin", "", {})).has_value());
  EXPECT_TRUE(collect_used_cell_xfs(wb).all_used);
}

// Merged-cell entries: one per group, after the other survivors, in first-appearance order.
UsedCellXfs UsedMerged(std::uint32_t count, std::initializer_list<std::uint32_t> plain,
                       std::vector<std::uint32_t> merged) {
  UsedCellXfs out = Used(count, plain);
  out.merged = std::move(merged);
  return out;
}

std::vector<std::uint32_t> MergedIndices(const CellStyleSavePlan& plan, std::uint32_t count) {
  std::vector<std::uint32_t> out(count);
  for (std::uint32_t c = 0; c < count; ++c) {
    out[c] = plan.merged_cell_xf_index(c);
  }
  return out;
}

TEST(CellStyleSavePlan, MergedCellsTakeAnEntryOfTheirOwnAfterTheOthers) {
  // A merged cell and a plain cell on the same xf 3 split into two entries.
  const CellStyleSavePlan plan = Plan(CellXfTable(Cx1NumFmts({})), UsedMerged(13, {3}, {3}));
  EXPECT_EQ(plan.cell_xf_order(), (std::vector<std::uint32_t>{0, 3, 3}));
  EXPECT_EQ(plan.cell_xf_index(3), 1U);
  EXPECT_EQ(plan.merged_cell_xf_index(3), 2U);
  EXPECT_FALSE(plan.cell_xfs_identity());
}

TEST(CellStyleSavePlan, MergedEntriesFollowFirstAppearanceNotSourceOrder) {
  const CellStyleSavePlan plan = Plan(CellXfTable(Cx1NumFmts({})), UsedMerged(13, {4}, {5, 3}));
  EXPECT_EQ(plan.cell_xf_order(), (std::vector<std::uint32_t>{0, 4, 5, 3}));
  EXPECT_EQ(plan.merged_cell_xf_index(5), 2U);
  EXPECT_EQ(plan.merged_cell_xf_index(3), 3U);
}

TEST(CellStyleSavePlan, MergedCellsShareOneEntryPerGroup) {
  // Cell xfs 1 and 2 are equal; both only inside merges.
  const CellStyleSavePlan plan = Plan(CellXfTable(Cx1NumFmts({{1, 3}, {2, 3}})), UsedMerged(13, {}, {1, 2}));
  EXPECT_EQ(plan.cell_xf_order(), (std::vector<std::uint32_t>{0, 1}));
  EXPECT_EQ(plan.merged_cell_xf_index(1), 1U);
  EXPECT_EQ(plan.merged_cell_xf_index(2), 1U);
}

TEST(CellStyleSavePlan, MergedCellsOnXf0TakeTheirOwnEntry) {
  const CellStyleSavePlan plan = Plan(CellXfTable(Cx1NumFmts({})), UsedMerged(13, {3}, {0}));
  EXPECT_EQ(plan.cell_xf_order(), (std::vector<std::uint32_t>{0, 3, 0}));
  EXPECT_EQ(plan.cell_xf_index(0), 0U);
  EXPECT_EQ(plan.merged_cell_xf_index(0), 2U);
}

TEST(CellStyleSavePlan, ExcelsOwnMergedEntryIsIdentity) {
  // As Excel writes it: xf 1 equals xf 0 and only merged cells use it.
  const CellStyleSavePlan plan = Plan(CellXfTable({0, 0}), UsedMerged(2, {}, {1}));
  EXPECT_EQ(plan.cell_xf_order(), (std::vector<std::uint32_t>{0, 1}));
  EXPECT_EQ(plan.merged_cell_xf_index(1), 1U);
  EXPECT_TRUE(plan.cell_xfs_identity());
}

TEST(CellStyleSavePlan, MergedIndicesWithoutMergesAreThePlainIndices) {
  const CellStyleSavePlan plan = Plan(CellXfTable(Cx1NumFmts({{6, 2}})), Used(13, {6, 4}));
  EXPECT_EQ(MergedIndices(plan, 13), Indices(plan, 13));
}

TEST(CellStyleSavePlan, CollectUsedCellXfsSendsMergedCellsAndRunColumnsApart) {
  Workbook wb = Workbook::create();
  wb.mutable_styles().cell_xfs.resize(8);
  Sheet& sheet = wb.sheet(0);
  // Row 3: C3 (xf 5, merged), D3 blank run xf 5 (merged), F3 xf 6 plain.
  sheet.set_cell_value(2, 2, Value::number(1.0));
  sheet.set_cell_xf_index(2, 2, 5U);
  sheet.set_cell_xf_index(2, 3, 5U);
  sheet.set_cell_value(2, 5, Value::number(1.0));
  sheet.set_cell_xf_index(2, 5, 6U);
  // Row 1: a run over A1:E1 with B1:C1 merged; A1, D1 and E1 stay plain.
  sheet.set_style_runs(0, 0, 5, 4U);
  // Row 2: A2 xf 3 merged, appearing before row 3 although its xf is lower.
  sheet.set_cell_value(1, 0, Value::number(1.0));
  sheet.set_cell_xf_index(1, 0, 3U);
  sheet.mutable_merges() = {MergeRange{2, 2, 2, 3}, MergeRange{0, 1, 0, 2}, MergeRange{1, 0, 1, 1}};
  const UsedCellXfs used = collect_used_cell_xfs(wb);
  EXPECT_EQ(used.used, (std::vector<bool>{false, false, false, false, true, false, true, false}));
  EXPECT_EQ(used.merged, (std::vector<std::uint32_t>{4, 3, 5}));
}

}  // namespace
}  // namespace formulon
