// Per-locale PivotTable labels against what Mac Excel 365 renders and names
// in each UI locale.

#include "pivot/pivot_locale.h"

#include <array>
#include <cstddef>
#include <string>
#include <string_view>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_types.h"
#include "pivot_layout_test_helpers.h"

namespace formulon::pivot {
namespace {

struct MeasuredLocale {
  const char* profile_id;
  // Data-field names Excel gave field `Amt`, in `Aggregation` order.
  std::array<const char*, 11> data_field_names;
  const char* grand_total;
  const char* row_labels;
  const char* column_labels;
  const char* blank_item;
  const char* all_pages;
  const char* multiple_items;
  const char* north_subtotal;  // tabular subtotal row of group `North`
};

const std::array<MeasuredLocale, 7> kMeasured{{
    {"mac-365-ja_JP",
     {"合計 / Amt", "個数 / Amt", "平均 / Amt", "最大 / Amt", "最小 / Amt", "積 / Amt", "数値の個数 / Amt",
      "標本標準偏差 / Amt", "標準偏差 / Amt", "標本分散 / Amt", "分散 / Amt"},
     "総計",
     "行ラベル",
     "列ラベル",
     "(空白)",
     "(すべて)",
     "(複数のアイテム)",
     "North 集計"},
    {"mac-365-en_US",
     {"Sum of Amt", "Count of Amt", "Average of Amt", "Max of Amt", "Min of Amt", "Product of Amt", "Count of Amt",
      "StdDev of Amt", "StdDevp of Amt", "Var of Amt", "Varp of Amt"},
     "Grand Total",
     "Row Labels",
     "Column Labels",
     "(blank)",
     "(All)",
     "(Multiple Items)",
     "North Total"},
    {"mac-365-de_DE",
     {"Summe von Amt", "Anzahl von Amt", "Mittelwert von Amt", "Max. von Amt", "Min. von Amt", "Produkt von Amt",
      "Anzahl von Amt", "STABW von Amt", "Standardabweichung (Grundgesamtheit) von Amt", "Var von Amt",
      "Varianz (Grundgesamtheit) von Amt"},
     "Gesamtergebnis",
     "Zeilenbeschriftungen",
     "Spaltenbeschriftungen",
     "(Leer)",
     "(Alle)",
     "(Mehrere Elemente)",
     "North Ergebnis"},
    {"mac-365-fr_FR",
     {"Somme de Amt", "Nombre de Amt", "Moyenne de Amt", "Max. de Amt", "Min. de Amt", "Produit de Amt",
      "Nombre de Amt", "Écartype de Amt", "Écartypep de Amt", "Var de Amt", "Varp de Amt"},
     "Total général",
     "Étiquettes de lignes",
     "Étiquettes de colonnes",
     "(vide)",
     "(Tous)",
     "(Plusieurs éléments)",
     "Total North"},
    {"mac-365-zh_CN",
     {"求和项:Amt", "计数项:Amt", "平均值项:Amt", "最大值项:Amt", "最小值项:Amt", "乘积项:Amt", "计数项:Amt",
      "标准偏差项:Amt", "总体标准偏差项:Amt", "方差项:Amt", "总体方差项:Amt"},
     "总计",
     "行标签",
     "列标签",
     "(空白)",
     "(全部)",
     "(多项)",
     "North 汇总"},
    {"mac-365-ko_KR",
     {"합계 : Amt", "개수 : Amt", "평균 : Amt", "최대 : Amt", "최소 : Amt", "곱 : Amt", "개수 : Amt",
      "표본 표준 편차 : Amt", "표준 편차 : Amt", "표본 분산 : Amt", "분산 : Amt"},
     "총합계",
     "행 레이블",
     "열 레이블",
     "(비어 있음)",
     "(모두)",
     "(다중 항목)",
     "North 요약"},
    {"mac-365-th_TH",
     {"ผลรวม ของ Amt", "นับจำนวน ของ Amt", "ค่าเฉลี่ย ของ Amt", "สูงสุด ของ Amt", "ต่ำสุด ของ Amt", "ผลคูณ ของ Amt",
      "นับจำนวน ของ Amt", "ส่วนเบี่ยงเบนมาตรฐาน ของ Amt", "ส่วนเบี่ยงเบนมาตรฐานของประชากร ของ Amt", "ค่าความแปรปรวน ของ Amt",
      "ค่าความแปรปรวนของประชากร ของ Amt"},
     "ผลรวมทั้งหมด",
     "ป้ายชื่อแถว",
     "ป้ายชื่อคอลัมน์",
     "(ว่าง)",
     "(ทั้งหมด)",
     "(หลายรายการ)",
     "North ผลรวม"},
}};

ExcelProfile profile_of(const char* id) {
  ExcelProfile profile;
  EXPECT_TRUE(parse_excel_profile_id(id, &profile)) << id;
  return profile;
}

TEST(PivotLocale, DataFieldNamesMatchExcelPerLocale) {
  for (const MeasuredLocale& m : kMeasured) {
    const ExcelProfile profile = profile_of(m.profile_id);
    for (std::size_t agg = 0; agg < m.data_field_names.size(); ++agg) {
      EXPECT_EQ(data_field_display_name(static_cast<Aggregation>(agg), "Amt", profile), m.data_field_names[agg])
          << m.profile_id << " aggregation " << agg;
    }
  }
}

TEST(PivotLocale, LayoutLabelsMatchExcelPerLocale) {
  for (const MeasuredLocale& m : kMeasured) {
    const PivotLayoutOptions options = pivot_layout_options_for(profile_of(m.profile_id));
    EXPECT_EQ(options.grand_total_label, m.grand_total) << m.profile_id;
    EXPECT_EQ(options.row_labels_label, m.row_labels) << m.profile_id;
    EXPECT_EQ(options.column_labels_label, m.column_labels) << m.profile_id;
    EXPECT_EQ(options.blank_item_label, m.blank_item) << m.profile_id;
    EXPECT_EQ(options.all_pages_label, m.all_pages) << m.profile_id;
    EXPECT_EQ(options.multiple_items_label, m.multiple_items) << m.profile_id;
  }
}

TEST(PivotLocale, TabularSubtotalRowIsLabelledPerLocale) {
  const PivotCache cache = layout_test::build_basic_cache();
  PivotTable table = layout_test::build_table(/*row=*/{0, 1}, /*col=*/{});
  table.set_layout(PivotLayout::Tabular);
  for (const MeasuredLocale& m : kMeasured) {
    const PivotLayoutOptions options = pivot_layout_options_for(profile_of(m.profile_id));
    auto result_or = evaluate(table, cache, options);
    ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
    auto cells_or = layout(table, result_or.value(), options);
    ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;
    bool found = false;
    for (const PivotCell& cell : cells_or.value().cells) {
      if (cell.kind == PivotCellKind::RowSubtotal && cell.value.is_text()) {
        EXPECT_EQ(cell.value.as_text(), m.north_subtotal) << m.profile_id;
        found = true;
        break;
      }
    }
    EXPECT_TRUE(found) << m.profile_id << ": no subtotal label cell";
  }
}

}  // namespace
}  // namespace formulon::pivot
