// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for `formulon::io::read_table`. The reader parses one
// `xl/tables/tableN.xml` part and returns the metadata required for
// round-trip preservation. Calculated-column formulas and structured
// references are deliberately out of scope at this layer.

#include "io/tables_reader.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "utils/error.h"

namespace formulon {
namespace io {
namespace {

std::vector<std::uint8_t> Bytes(std::string_view xml) {
  return std::vector<std::uint8_t>(xml.begin(), xml.end());
}

constexpr std::string_view kXmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

// ---------------------------------------------------------------------------
// Happy paths
// ---------------------------------------------------------------------------

TEST(TablesReader, MinimalTableTwoColumns) {
  std::string xml(kXmlDecl);
  xml.append(
      "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"1\" name=\"T1\" "
      "displayName=\"T1\" ref=\"A1:B5\">");
  xml.append("  <tableColumns count=\"2\">");
  xml.append("    <tableColumn id=\"1\" name=\"A\"/>");
  xml.append("    <tableColumn id=\"2\" name=\"B\"/>");
  xml.append("  </tableColumns>");
  xml.append("</table>");

  auto table_or = read_table(Bytes(xml), 0U);
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const TableMetadata& t = table_or.value();
  EXPECT_EQ(t.id, 1U);
  EXPECT_EQ(t.name, "T1");
  EXPECT_EQ(t.display_name, "T1");
  EXPECT_EQ(t.ref, "A1:B5");
  EXPECT_EQ(t.sheet_index, 0U);
  EXPECT_TRUE(t.header_row);
  EXPECT_FALSE(t.totals_row);
  ASSERT_EQ(t.columns.size(), 2U);
  EXPECT_EQ(t.columns[0].id, 1U);
  EXPECT_EQ(t.columns[0].name, "A");
  EXPECT_EQ(t.columns[1].id, 2U);
  EXPECT_EQ(t.columns[1].name, "B");
}

TEST(TablesReader, HeaderRowCountZeroDisablesHeader) {
  std::string xml(kXmlDecl);
  xml.append(
      "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"7\" name=\"NoHdr\" "
      "displayName=\"NoHdr\" ref=\"A1:C3\" headerRowCount=\"0\">");
  xml.append("  <tableColumns count=\"3\">");
  xml.append("    <tableColumn id=\"1\" name=\"X\"/>");
  xml.append("    <tableColumn id=\"2\" name=\"Y\"/>");
  xml.append("    <tableColumn id=\"3\" name=\"Z\"/>");
  xml.append("  </tableColumns>");
  xml.append("</table>");

  auto table_or = read_table(Bytes(xml), 1U);
  ASSERT_TRUE(static_cast<bool>(table_or));
  EXPECT_FALSE(table_or.value().header_row);
  EXPECT_FALSE(table_or.value().totals_row);
}

TEST(TablesReader, TotalsRowCountOneEnablesTotalsRow) {
  std::string xml(kXmlDecl);
  xml.append(
      "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"3\" name=\"WithTotals\" "
      "displayName=\"WithTotals\" ref=\"A1:B6\" totalsRowCount=\"1\">");
  xml.append("  <tableColumns count=\"2\">");
  xml.append("    <tableColumn id=\"1\" name=\"Item\"/>");
  xml.append("    <tableColumn id=\"2\" name=\"Qty\"/>");
  xml.append("  </tableColumns>");
  xml.append("</table>");

  auto table_or = read_table(Bytes(xml), 0U);
  ASSERT_TRUE(static_cast<bool>(table_or));
  EXPECT_TRUE(table_or.value().header_row);
  EXPECT_TRUE(table_or.value().totals_row);
}

TEST(TablesReader, ColumnTotalsLabelAndFunctionCaptured) {
  std::string xml(kXmlDecl);
  xml.append(
      "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"4\" name=\"Sales\" "
      "displayName=\"Sales\" ref=\"A1:C7\" totalsRowCount=\"1\">");
  xml.append("  <tableColumns count=\"3\">");
  xml.append("    <tableColumn id=\"1\" name=\"Region\" totalsRowLabel=\"Total\"/>");
  xml.append("    <tableColumn id=\"2\" name=\"Q1\" totalsRowFunction=\"sum\"/>");
  xml.append("    <tableColumn id=\"3\" name=\"Q2\" totalsRowFunction=\"average\"/>");
  xml.append("  </tableColumns>");
  xml.append("</table>");

  auto table_or = read_table(Bytes(xml), 0U);
  ASSERT_TRUE(static_cast<bool>(table_or));
  ASSERT_EQ(table_or.value().columns.size(), 3U);
  EXPECT_EQ(table_or.value().columns[0].totals_label, "Total");
  EXPECT_TRUE(table_or.value().columns[0].totals_function.empty());
  EXPECT_EQ(table_or.value().columns[1].totals_function, "sum");
  EXPECT_TRUE(table_or.value().columns[1].totals_label.empty());
  EXPECT_EQ(table_or.value().columns[2].totals_function, "average");
}

TEST(TablesReader, SheetIndexParameterIsPropagated) {
  // The reader does not infer the owning sheet from the part name; the
  // caller (ooxml_reader) supplies it from the rels walk. Verify the
  // raw parameter survives the parse for a few representative values.
  std::string xml(kXmlDecl);
  xml.append(
      "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"2\" name=\"T2\" "
      "displayName=\"T2\" ref=\"A1:A1\">");
  xml.append("  <tableColumns count=\"1\"><tableColumn id=\"1\" name=\"X\"/></tableColumns>");
  xml.append("</table>");

  for (std::size_t sheet_index : {std::size_t{0}, std::size_t{5}, std::size_t{42}}) {
    auto table_or = read_table(Bytes(xml), sheet_index);
    ASSERT_TRUE(static_cast<bool>(table_or)) << "sheet_index=" << sheet_index;
    EXPECT_EQ(table_or.value().sheet_index, sheet_index);
  }
}

// ---------------------------------------------------------------------------
// Error paths
// ---------------------------------------------------------------------------

TEST(TablesReader, MissingTableRootIsContentTypeInvalid) {
  std::string xml(kXmlDecl);
  xml.append("<notATable id=\"1\" name=\"T1\" ref=\"A1:B5\"/>");

  auto table_or = read_table(Bytes(xml), 0U);
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoContentTypeInvalid);
}

TEST(TablesReader, MissingRefAttributeIsCorruption) {
  std::string xml(kXmlDecl);
  xml.append(
      "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"1\" name=\"T1\" "
      "displayName=\"T1\">");
  xml.append("  <tableColumns count=\"1\"><tableColumn id=\"1\" name=\"X\"/></tableColumns>");
  xml.append("</table>");

  auto table_or = read_table(Bytes(xml), 0U);
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(TablesReader, MalformedXmlIsParseError) {
  // Unterminated element — pugixml should refuse it.
  std::string xml(kXmlDecl);
  xml.append("<table id=\"1\" ref=\"A1:B2\"><tableColumns");

  auto table_or = read_table(Bytes(xml), 0U);
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoXmlParse);
}

}  // namespace
}  // namespace io
}  // namespace formulon
