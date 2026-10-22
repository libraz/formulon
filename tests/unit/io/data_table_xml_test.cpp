// `<f t="dataTable">` attribute parsing and re-emission.

#include "io/data_table_xml.h"

#include <string>

#include "data_table.h"
#include "gtest/gtest.h"

namespace formulon::io {
namespace {

TEST(DataTableXml, ParsesEveryAttribute) {
  DataTable table;
  ASSERT_TRUE(parse_data_table(DataTableAttrs{"B2:D5", "A1", "A2", "1", "0", "0", "1", "1", "true"}, &table));
  EXPECT_EQ(table.ref, (MergeRange{1, 1, 4, 3}));
  EXPECT_TRUE(table.two_dimensional);
  EXPECT_FALSE(table.row_oriented);
  EXPECT_EQ(table.r1, (DataTableInput{true, false, 0, 0}));
  EXPECT_EQ(table.r2, (DataTableInput{true, true, 1, 0}));
  EXPECT_TRUE(table.always_calc);
  EXPECT_TRUE(table.always_calc_array);
}

TEST(DataTableXml, AbsentAttributesTakeTheirDefaults) {
  DataTable table;
  ASSERT_TRUE(parse_data_table(DataTableAttrs{"C3", "B1", {}, {}, {}, {}, {}, {}, {}}, &table));
  EXPECT_EQ(table.ref, (MergeRange{2, 2, 2, 2}));
  EXPECT_FALSE(table.two_dimensional);
  EXPECT_FALSE(table.row_oriented);
  EXPECT_TRUE(table.r1.present);
  EXPECT_FALSE(table.r2.present);
  EXPECT_EQ(table.explicit_attrs, 0U);
}

TEST(DataTableXml, RejectsAMissingOrMalformedRef) {
  DataTable table;
  EXPECT_FALSE(parse_data_table(DataTableAttrs{}, &table));
  EXPECT_FALSE(parse_data_table(DataTableAttrs{"nonsense", "A1", {}, {}, {}, {}, {}, {}, {}}, &table));
  EXPECT_FALSE(parse_data_table(DataTableAttrs{"A1:B2 D4", "A1", {}, {}, {}, {}, {}, {}, {}}, &table));
}

TEST(DataTableXml, AnUnparseableInputIsAbsentAndNotWritten) {
  DataTable table;
  ASSERT_TRUE(parse_data_table(DataTableAttrs{"B2:B3", "#REF!", {}, {}, {}, {}, {}, {}, {}}, &table));
  EXPECT_FALSE(table.r1.present);
  std::string xml;
  append_data_table_formula(xml, table);
  EXPECT_EQ(xml.find("r1="), std::string::npos);
}

TEST(DataTableXml, ReEmitsWhatTheSourceSpelledOut) {
  DataTable table;
  ASSERT_TRUE(parse_data_table(DataTableAttrs{"B2:D5", "A1", "A2", "1", "0", "0", "0", "1", {}}, &table));
  std::string xml;
  append_data_table_formula(xml, table);
  EXPECT_EQ(
      xml,
      "<f t=\"dataTable\" ref=\"B2:D5\" dt2D=\"1\" dtr=\"0\" r1=\"A1\" r2=\"A2\" del1=\"0\" del2=\"0\" ca=\"1\"/>");

  // Defaults the source left out stay out; set flags are written.
  DataTable bare;
  ASSERT_TRUE(parse_data_table(DataTableAttrs{"B2:B3", "F1", {}, {}, {}, {}, {}, {}, {}}, &bare));
  bare.r1.deleted = true;
  xml.clear();
  append_data_table_formula(xml, bare);
  EXPECT_EQ(xml, "<f t=\"dataTable\" ref=\"B2:B3\" r1=\"F1\" del1=\"1\"/>");
}

}  // namespace
}  // namespace formulon::io
