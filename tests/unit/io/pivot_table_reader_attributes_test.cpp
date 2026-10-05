// Pivot-table definition reader tests grouped by XML concern.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/pivot_table_reader.h"
#include "io/pivot_table_writer.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot_table_reader_test_helpers.h"
#include "utils/error.h"

namespace formulon::io {
namespace {
using namespace pivot_table_reader_test_support;
TEST(PivotTableReader, ShowDataAsAttributesAreParsed) {
  // <dataField> exercises showDataAs / baseField / baseItem decoding.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"2\">");
  xml.append("  <pivotField axis=\"axisRow\" name=\"R\"/>");
  xml.append("  <pivotField dataField=\"1\" name=\"V\"/>");
  xml.append("</pivotFields>");
  xml.append("<dataFields count=\"1\">");
  xml.append(
      "  <dataField name=\"X\" fld=\"0\" subtotal=\"sum\""
      " showDataAs=\"difference\" baseField=\"1\" baseItem=\"3\"/>");
  xml.append("</dataFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.data_fields().size(), 1U);
  const pivot::PivotDataField& df = table.data_fields()[0];
  EXPECT_EQ(df.show_as, pivot::ShowValuesAs::DifferenceFrom);
  ASSERT_TRUE(df.show_as_base_field.has_value());
  EXPECT_EQ(*df.show_as_base_field, 1U);
  ASSERT_TRUE(df.show_as_base_item.has_value());
  EXPECT_EQ(*df.show_as_base_item, 3U);
}
TEST(PivotTableReader, EmptyPassthroughWhenNoExtensionsPresent) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"0\"/>");
  xml.append("</pivotTableDefinition>");
  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or));
  EXPECT_TRUE(table_or.value().raw_passthrough_xml().empty());
}
TEST(PivotTableReader, GrandTotalsDefaultTrueWhenAbsent) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"0\"/>");
  xml.append("</pivotTableDefinition>");
  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  EXPECT_TRUE(table_or.value().grand_totals_rows());
  EXPECT_TRUE(table_or.value().grand_totals_cols());
}
TEST(PivotTableReader, GrandTotalsOffIsRead) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs);
  xml.append(" name=\"P\" cacheId=\"0\" rowGrandTotals=\"0\" colGrandTotals=\"0\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"0\"/>");
  xml.append("</pivotTableDefinition>");
  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  EXPECT_FALSE(table_or.value().grand_totals_rows());
  EXPECT_FALSE(table_or.value().grand_totals_cols());
}
TEST(PivotTableReader, LocationRequiredAttributesAreRead) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append(
      "<location ref=\"A3:D10\" firstHeaderRow=\"1\" firstDataRow=\"2\" firstDataCol=\"3\" "
      "rowPageCount=\"4\" colPageCount=\"5\"/>");
  xml.append("<pivotFields count=\"0\"/>");
  xml.append("</pivotTableDefinition>");
  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_TRUE(table.location_first_header_row().has_value());
  EXPECT_EQ(*table.location_first_header_row(), 1U);
  ASSERT_TRUE(table.location_first_data_row().has_value());
  EXPECT_EQ(*table.location_first_data_row(), 2U);
  ASSERT_TRUE(table.location_first_data_col().has_value());
  EXPECT_EQ(*table.location_first_data_col(), 3U);
  ASSERT_TRUE(table.location_row_page_count().has_value());
  EXPECT_EQ(*table.location_row_page_count(), 4U);
  ASSERT_TRUE(table.location_col_page_count().has_value());
  EXPECT_EQ(*table.location_col_page_count(), 5U);
}
TEST(PivotTableReader, LocationOptionalAttributesAbsentStayEmpty) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  // Only `ref` present; all offset attributes absent.
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"0\"/>");
  xml.append("</pivotTableDefinition>");
  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  EXPECT_FALSE(table.location_first_header_row().has_value());
  EXPECT_FALSE(table.location_first_data_row().has_value());
  EXPECT_FALSE(table.location_first_data_col().has_value());
  EXPECT_FALSE(table.location_row_page_count().has_value());
  EXPECT_FALSE(table.location_col_page_count().has_value());
}
TEST(PivotTableReader, CustomSubtotalsAndDefaultSubtotalSurviveRoundTrip) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"1\">");
  // defaultSubtotal turned OFF, with explicit Average + Max custom
  // subtotals selected.
  xml.append("  <pivotField axis=\"axisRow\" name=\"R\" defaultSubtotal=\"0\" avgSubtotal=\"1\" maxSubtotal=\"1\"/>");
  xml.append("</pivotFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.fields().size(), 1U);
  const pivot::PivotField& f = table.fields()[0];
  EXPECT_FALSE(f.default_subtotal);
  ASSERT_EQ(f.subtotal_fns.size(), 2U);
  EXPECT_EQ(f.subtotal_fns[0], pivot::SubtotalFn::Average);
  EXPECT_EQ(f.subtotal_fns[1], pivot::SubtotalFn::Max);

  // Write -> read again: the custom selection must not revert to default.
  const std::string round = write_pivot_table_definition(table);
  auto reparsed_or = read_pivot_table_definition(Bytes(round));
  ASSERT_TRUE(static_cast<bool>(reparsed_or)) << reparsed_or.error().message;
  const pivot::PivotField& f2 = reparsed_or.value().fields()[0];
  EXPECT_FALSE(f2.default_subtotal);
  ASSERT_EQ(f2.subtotal_fns.size(), 2U);
  EXPECT_EQ(f2.subtotal_fns[0], pivot::SubtotalFn::Average);
  EXPECT_EQ(f2.subtotal_fns[1], pivot::SubtotalFn::Max);
}
TEST(PivotTableReader, DefaultSubtotalDefaultsTrueWhenAbsent) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"1\"><pivotField axis=\"axisRow\" name=\"R\"/></pivotFields>");
  xml.append("</pivotTableDefinition>");
  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const pivot::PivotField& f = table_or.value().fields()[0];
  EXPECT_TRUE(f.default_subtotal);
  EXPECT_TRUE(f.subtotal_fns.empty());
}
TEST(PivotTableReader, RowItemsAndColItemsSurviveRoundTrip) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<pivotFields count=\"0\"/>");
  xml.append("<rowItems count=\"2\"><i><x/></i><i t=\"grand\"><x/></i></rowItems>");
  xml.append("<colItems count=\"1\"><i><x/></i></colItems>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  // rowItems / colItems are binned into their schema-position buffers so
  // the writer can re-emit them before <dataFields> (a single tail buffer
  // would place them after <dataFields> and trip Excel's repair).
  const pivot::PivotTable& table = table_or.value();
  EXPECT_NE(table.raw_passthrough_after_row_fields().find("<rowItems"), std::string::npos);
  EXPECT_NE(table.raw_passthrough_after_row_fields().find("t=\"grand\""), std::string::npos);
  EXPECT_NE(table.raw_passthrough_after_col_fields().find("<colItems"), std::string::npos);
  // They must NOT leak into the tail buffer.
  EXPECT_EQ(table.raw_passthrough_xml().find("<rowItems"), std::string::npos);

  // Write -> read again: the layout-item cache must still be present, and
  // rowItems must precede colItems in the emitted bytes.
  const std::string round = write_pivot_table_definition(table);
  const std::size_t p_rowitems = round.find("<rowItems");
  const std::size_t p_colitems = round.find("<colItems");
  ASSERT_NE(p_rowitems, std::string::npos) << round;
  ASSERT_NE(p_colitems, std::string::npos) << round;
  EXPECT_LT(p_rowitems, p_colitems);
  auto reparsed_or = read_pivot_table_definition(Bytes(round));
  ASSERT_TRUE(static_cast<bool>(reparsed_or)) << reparsed_or.error().message;
  const pivot::PivotTable& reparsed = reparsed_or.value();
  EXPECT_NE(reparsed.raw_passthrough_after_row_fields().find("<rowItems"), std::string::npos);
  EXPECT_NE(reparsed.raw_passthrough_after_col_fields().find("<colItems"), std::string::npos);
}

}  // namespace
}  // namespace formulon::io
