#include <cstdint>
#include <string>

#include "cell.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "style_resolve.h"
#include "workbook.h"

namespace formulon {
namespace {

constexpr const char* kSparklines =
    "<ext uri=\"{05C60535-1F16-4fd2-B633-F4F36F0B64E0}\"><x14:sparklineGroups>"
    "<x14:sparklineGroup><x14:sparklines count=\"1\"><x14:sparkline><xm:f>Sheet1!E1:E6</xm:f>"
    "<xm:sqref>G1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup>"
    "</x14:sparklineGroups></ext>";

std::uint32_t XfOf(const Sheet& sheet, std::uint32_t row, std::uint32_t col) {
  return select_effective_xf(sheet, row, col).xf_index;
}

TEST(InsertFormatInheritance, InsertedRowsTakeRowAndCellStyleOfRowAbove) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_xf_index(1, 2, 7U);
  RowLayout styled;
  styled.row = 1;
  styled.has_style = true;
  styled.style_xf = 4U;
  sheet.mutable_layout().row_overrides.push_back(styled);
  sheet.set_cell_xf_index(2, 0, 9U);

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 2, 2)));

  for (const std::uint32_t row : {2U, 3U}) {
    EXPECT_EQ(XfOf(sheet, row, 2), 7U) << row;
    EXPECT_EQ(XfOf(sheet, row, 5), 4U) << row;
    EXPECT_EQ(select_effective_xf(sheet, row, 5).source, StyleSource::kRow) << row;
    EXPECT_EQ(XfOf(sheet, row, 0), 4U) << row;
  }
  EXPECT_EQ(XfOf(sheet, 1, 2), 7U);
  EXPECT_EQ(XfOf(sheet, 4, 0), 9U);
}

TEST(InsertFormatInheritance, UnformattedRowAboveLeavesInsertedRowsPlain) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(1, 0, Value::number(1.0));
  sheet.set_cell_xf_index(2, 0, 5U);

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 2, 1)));

  EXPECT_EQ(sheet.cell_at(2, 0), nullptr);
  EXPECT_EQ(XfOf(sheet, 2, 0), 0U);
  EXPECT_EQ(XfOf(sheet, 3, 0), 5U);
}

TEST(InsertFormatInheritance, InsertAtFirstRowInheritsNothing) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_xf_index(0, 0, 3U);

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 0, 1)));

  EXPECT_EQ(XfOf(sheet, 0, 0), 0U);
  EXPECT_EQ(XfOf(sheet, 1, 0), 3U);
}

TEST(InsertFormatInheritance, ExplicitDefaultStyleIsCopiedOverRowStyle) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  RowLayout styled;
  styled.row = 1;
  styled.has_style = true;
  styled.style_xf = 4U;
  sheet.mutable_layout().row_overrides.push_back(styled);
  sheet.set_cell_xf_index(1, 1, 0U);

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 2, 1)));

  EXPECT_EQ(XfOf(sheet, 2, 1), 0U);
  EXPECT_EQ(select_effective_xf(sheet, 2, 1).source, StyleSource::kCell);
  EXPECT_EQ(XfOf(sheet, 2, 3), 4U);
}

TEST(InsertFormatInheritance, InsertedColumnsTakeColumnAndCellStyleOfColumnToTheLeft) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  ColumnLayout styled;
  styled.first = 1;
  styled.last = 1;
  styled.has_style = true;
  styled.style_xf = 6U;
  sheet.mutable_layout().columns.push_back(styled);
  sheet.set_cell_xf_index(3, 1, 8U);

  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, 2, 2)));

  for (const std::uint32_t col : {2U, 3U}) {
    EXPECT_EQ(XfOf(sheet, 3, col), 8U) << col;
    EXPECT_EQ(XfOf(sheet, 9, col), 6U) << col;
    EXPECT_EQ(select_effective_xf(sheet, 9, col).source, StyleSource::kColumn) << col;
  }
  EXPECT_EQ(XfOf(sheet, 9, 4), 0U);
}

TEST(InsertFormatInheritance, InsertInsideStyledColumnSpanKeepsOneSpan) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  ColumnLayout styled;
  styled.first = 0;
  styled.last = 3;
  styled.has_style = true;
  styled.style_xf = 6U;
  sheet.mutable_layout().columns.push_back(styled);

  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, 2, 1)));

  ASSERT_EQ(sheet.layout().columns.size(), 1U);
  EXPECT_EQ(sheet.layout().columns[0].last, 4U);
}

TEST(InsertFormatInheritance, InsertAtFirstColumnInheritsNothing) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_xf_index(0, 0, 3U);

  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, 0, 1)));

  EXPECT_EQ(XfOf(sheet, 0, 0), 0U);
  EXPECT_EQ(XfOf(sheet, 0, 1), 3U);
}

TEST(InsertFormatInheritance, InsertedRowsCloneSparklineHostedInRowAbove) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_ext_lst_xml(kSparklines);

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 1, 2)));

  const std::string& xml = sheet.ext_lst_xml();
  EXPECT_NE(xml.find("<xm:f>Sheet1!E1:E8</xm:f><xm:sqref>G1</xm:sqref>"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<xm:f>Sheet1!E2:E9</xm:f><xm:sqref>G2</xm:sqref>"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<xm:f>Sheet1!E3:E10</xm:f><xm:sqref>G3</xm:sqref>"), std::string::npos) << xml;
  EXPECT_NE(xml.find("count=\"3\""), std::string::npos) << xml;
}

TEST(InsertFormatInheritance, SparklineBelowInsertionPointIsNotCloned) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_ext_lst_xml(kSparklines);

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 0, 1)));

  const std::string& xml = sheet.ext_lst_xml();
  EXPECT_NE(xml.find("<xm:sqref>G2</xm:sqref>"), std::string::npos) << xml;
  EXPECT_EQ(xml.find("G1"), std::string::npos) << xml;
  EXPECT_NE(xml.find("count=\"1\""), std::string::npos) << xml;
}

TEST(InsertFormatInheritance, InsertedColumnsCloneSparklineHostedInColumnToTheLeft) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_ext_lst_xml(kSparklines);

  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, 7, 1)));

  const std::string& xml = sheet.ext_lst_xml();
  EXPECT_NE(xml.find("<xm:f>Sheet1!E1:E6</xm:f><xm:sqref>G1</xm:sqref>"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<xm:sqref>H1</xm:sqref>"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<xm:f>Sheet1!F1:F6</xm:f>"), std::string::npos) << xml;
}

}  // namespace
}  // namespace formulon
