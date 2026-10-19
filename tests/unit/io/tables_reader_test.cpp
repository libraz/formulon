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
#include "io/ooxml_writer.h"
#include "io/xml_utils.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

constexpr std::string_view kXmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

std::vector<std::uint8_t> Bytes(std::string_view xml) {
  return std::vector<std::uint8_t>(xml.begin(), xml.end());
}

std::string SerializeTablePart(const TableMetadata& table) {
  Workbook workbook = Workbook::create_empty();
  workbook.add_sheet("Sheet1");
  workbook.set_tables({table});
  auto package_or = write_ooxml(workbook);
  if (!package_or) {
    ADD_FAILURE() << "write_ooxml failed: " << package_or.error().message;
    return {};
  }

  ZipReader zip;
  const std::vector<std::uint8_t>& package = package_or.value();
  auto open_or = zip.open(ByteSpan{package.data(), package.size()});
  if (!open_or) {
    ADD_FAILURE() << "opening the written package failed: " << open_or.error().message;
    return {};
  }
  auto table_xml_or = zip.read_entry("xl/tables/table1.xml");
  if (!table_xml_or) {
    ADD_FAILURE() << "reading xl/tables/table1.xml failed: " << table_xml_or.error().message;
    return {};
  }
  const std::vector<std::uint8_t>& bytes = table_xml_or.value();
  return std::string(bytes.begin(), bytes.end());
}

void ExpectTableColumnNamespaceRoundTrip(std::string_view table_columns_attrs, std::string_view column_attrs,
                                         std::string_view prefix, std::string_view expected_uri,
                                         std::string_view retained_attr_name, std::string_view retained_attr_value,
                                         std::string_view extra_prefix = {}, std::string_view extra_uri = {}) {
  std::string xml(kXmlDecl);
  xml.append("<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ");
  xml.append("xmlns:x=\"urn:root\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" ");
  xml.append("id=\"1\" name=\"Table1\" displayName=\"Table1\" ref=\"A1:A2\">");
  xml.append("<tableColumns count=\"1\"");
  if (!table_columns_attrs.empty()) {
    xml.push_back(' ');
    xml.append(table_columns_attrs);
  }
  xml.append("><tableColumn id=\"1\" name=\"Output\" totalsRowLabel=\"Total\"");
  if (!column_attrs.empty()) {
    xml.push_back(' ');
    xml.append(column_attrs);
  }
  xml.append("><calculatedColumnFormula>SUM(A1)</calculatedColumnFormula></tableColumn></tableColumns></table>");

  auto table_or = read_table(Bytes(xml), 0U);
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read_table failed: " << table_or.error().message;
  ASSERT_EQ(table_or.value().columns.size(), 1U);
  EXPECT_EQ(table_or.value().columns[0].id, 1U);
  EXPECT_EQ(table_or.value().columns[0].name, "Output");
  EXPECT_EQ(table_or.value().columns[0].totals_label, "Total");
  EXPECT_EQ(table_or.value().columns[0].calculated_column_formula, "SUM(A1)");

  const std::string output = SerializeTablePart(table_or.value());
  ASSERT_FALSE(output.empty());
  pugi::xml_document doc;
  const pugi::xml_parse_result parse = doc.load_buffer(output.data(), output.size());
  ASSERT_TRUE(parse) << parse.description();
  const pugi::xml_node table = doc.child("table");
  ASSERT_TRUE(table);
  EXPECT_EQ(find_namespace_binding(table, "x"), "urn:root");
  const pugi::xml_node column = table.child("tableColumns").child("tableColumn");
  ASSERT_TRUE(column);
  EXPECT_EQ(find_namespace_binding(column, prefix), expected_uri);
  if (!extra_prefix.empty()) {
    EXPECT_EQ(find_namespace_binding(column, extra_prefix), extra_uri);
  }
  const std::string retained_name(retained_attr_name);
  EXPECT_EQ(std::string_view(column.attribute(retained_name.c_str()).value()), retained_attr_value);
  EXPECT_EQ(std::string_view(column.attribute("id").value()), "1");
  EXPECT_EQ(std::string_view(column.attribute("name").value()), "Output");
  EXPECT_EQ(std::string_view(column.attribute("totalsRowLabel").value()), "Total");
  ASSERT_TRUE(column.child("calculatedColumnFormula"));
  EXPECT_EQ(std::string_view(column.child("calculatedColumnFormula").text().as_string()), "SUM(A1)");
}

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

TEST(TablesReader, ColumnLocalNamespaceDeclarationSurvivesSerialization) {
  ExpectTableColumnNamespaceRoundTrip("", "xmlns:y=\"urn:local\"", "y", "urn:local", "xmlns:y", "urn:local");
}

TEST(TablesReader, InheritedTableColumnsNamespaceSurvivesSerialization) {
  ExpectTableColumnNamespaceRoundTrip("xmlns:x=\"urn:columns\"", "x:tag=\"inherited\"", "x", "urn:columns", "x:tag",
                                      "inherited");
}

TEST(TablesReader, ColumnNamespaceShadowingRootSurvivesSerialization) {
  ExpectTableColumnNamespaceRoundTrip("", "xmlns:x=\"urn:column\" x:tag=\"shadowed\"", "x", "urn:column", "x:tag",
                                      "shadowed");
}

TEST(TablesReader, IgnorableOnlyPrefixBindingSurvivesSerialization) {
  ExpectTableColumnNamespaceRoundTrip("xmlns:x=\"urn:columns\"", "mc:Ignorable=\"x\"", "x", "urn:columns",
                                      "mc:Ignorable", "x");
}

TEST(TablesReader, AliasedIgnorableOnlyPrefixBindingSurvivesSerialization) {
  ExpectTableColumnNamespaceRoundTrip(
      "xmlns:x=\"urn:columns\"",
      "xmlns:m=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" m:Ignorable=\"x\"", "x", "urn:columns",
      "m:Ignorable", "x", "m", "http://schemas.openxmlformats.org/markup-compatibility/2006");
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

TEST(TablesReader, CalculatedColumnFormulaPreserved) {
  std::string xml(kXmlDecl);
  xml.append(
      "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"5\" name=\"CalcTbl\" "
      "displayName=\"CalcTbl\" ref=\"A1:C5\">");
  xml.append("  <tableColumns count=\"3\">");
  xml.append("    <tableColumn id=\"1\" name=\"A\"/>");
  xml.append("    <tableColumn id=\"2\" name=\"B\"/>");
  xml.append(
      "    <tableColumn id=\"3\" name=\"Total\">"
      "<calculatedColumnFormula>SUM(Tbl[@[A]:[B]])</calculatedColumnFormula>"
      "</tableColumn>");
  xml.append("  </tableColumns>");
  xml.append("</table>");

  auto table_or = read_table(Bytes(xml), 0U);
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const TableMetadata& t = table_or.value();
  ASSERT_EQ(t.columns.size(), 3U);
  EXPECT_EQ(t.columns[0].calculated_column_formula, "");
  EXPECT_EQ(t.columns[1].calculated_column_formula, "");
  EXPECT_EQ(t.columns[2].calculated_column_formula, "SUM(Tbl[@[A]:[B]])");
}

TEST(TablesReader, CalculatedColumnFormulaWithXmlEscapesPreserved) {
  // Verify the parser un-escapes the standard XML entities (`&lt;`,
  // `&amp;`, `&gt;`) before storing the formula string. This is what
  // pugixml's `text().as_string()` already does — the test pins the
  // contract so a future refactor cannot silently double-escape.
  std::string xml(kXmlDecl);
  xml.append(
      "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"6\" name=\"EscTbl\" "
      "displayName=\"EscTbl\" ref=\"A1:B2\">");
  xml.append("  <tableColumns count=\"1\">");
  xml.append(
      "    <tableColumn id=\"1\" name=\"E\">"
      "<calculatedColumnFormula>IF(A1&lt;5,&quot;x&quot;,&quot;y&quot;)&amp;\"!\"</calculatedColumnFormula>"
      "</tableColumn>");
  xml.append("  </tableColumns>");
  xml.append("</table>");

  auto table_or = read_table(Bytes(xml), 0U);
  ASSERT_TRUE(static_cast<bool>(table_or));
  ASSERT_EQ(table_or.value().columns.size(), 1U);
  // Un-escaped: `<` and `&` and `"` show up as literal characters.
  EXPECT_EQ(table_or.value().columns[0].calculated_column_formula, "IF(A1<5,\"x\",\"y\")&\"!\"");
}

TEST(TablesReader, CalculatedColumnFormulaDecodesOoxmlTextEscapes) {
  std::string xml(kXmlDecl);
  xml.append(
      "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"8\" name=\"EscTbl\" "
      "displayName=\"EscTbl\" ref=\"A1:B2\">");
  xml.append("  <tableColumns count=\"1\">");
  xml.append(
      "    <tableColumn id=\"1\" name=\"E\"><calculatedColumnFormula>"
      "literal _x005F_x0041_ &amp; &lt;猫&gt; _x000D_control _x0001_ 🙂"
      "</calculatedColumnFormula></tableColumn>");
  xml.append("  </tableColumns>");
  xml.append("</table>");

  auto table_or = read_table(Bytes(xml), 0U);
  ASSERT_TRUE(static_cast<bool>(table_or));
  ASSERT_EQ(table_or.value().columns.size(), 1U);
  EXPECT_EQ(table_or.value().columns[0].calculated_column_formula, "literal _x0041_ & <猫> \rcontrol \x01 🙂");
}

TEST(TablesReader, CalculatedColumnFormulaWithOoxmlControlCharacterEscapeDecoded) {
  // `_x0001_` is the OOXML notation for a control character XML 1.0
  // cannot carry literally (ooxml_writer.cpp's AppendXmlEscaped emits it
  // for such a byte inside this same element); the reader must decode it
  // back, matching every other element-text slot.
  std::string xml(kXmlDecl);
  xml.append(
      "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"6\" name=\"EscTbl\" "
      "displayName=\"EscTbl\" ref=\"A1:B2\">");
  xml.append("  <tableColumns count=\"1\">");
  xml.append(
      "    <tableColumn id=\"1\" name=\"E\">"
      "<calculatedColumnFormula>\"a_x0001_b\"</calculatedColumnFormula>"
      "</tableColumn>");
  xml.append("  </tableColumns>");
  xml.append("</table>");

  auto table_or = read_table(Bytes(xml), 0U);
  ASSERT_TRUE(static_cast<bool>(table_or));
  ASSERT_EQ(table_or.value().columns.size(), 1U);
  EXPECT_EQ(table_or.value().columns[0].calculated_column_formula, std::string("\"a\x01"
                                                                               "b\"",
                                                                               5));
}

TEST(TablesReader, MissingCalculatedColumnFormulaIsEmptyString) {
  std::string xml(kXmlDecl);
  xml.append(
      "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"8\" name=\"NoCalc\" "
      "displayName=\"NoCalc\" ref=\"A1:A2\">");
  xml.append("  <tableColumns count=\"1\"><tableColumn id=\"1\" name=\"X\"/></tableColumns>");
  xml.append("</table>");

  auto table_or = read_table(Bytes(xml), 0U);
  ASSERT_TRUE(static_cast<bool>(table_or));
  ASSERT_EQ(table_or.value().columns.size(), 1U);
  EXPECT_TRUE(table_or.value().columns[0].calculated_column_formula.empty());
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

TEST(TablesReader, OutOfRangeIdFallsBackToDefault) {
  // Regression: tableId="3000000000" used to clamp to LONG_MAX on 32-bit
  // hosts via std::strtol. ParseU32Attr now ranges-checks against
  // UINT32_MAX and falls back to the default (0) on overflow.
  std::string xml(kXmlDecl);
  xml.append(
      "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "id=\"99999999999\" name=\"T1\" displayName=\"T1\" ref=\"A1:B5\">");
  xml.append("  <tableColumns count=\"1\"><tableColumn id=\"1\" name=\"X\"/></tableColumns>");
  xml.append("</table>");

  auto table_or = read_table(Bytes(xml), 0U);
  ASSERT_TRUE(static_cast<bool>(table_or));
  // Default fallback is 0, not a wrapped value.
  EXPECT_EQ(table_or.value().id, 0U);
}

TEST(TablesReader, NegativeIdFallsBackToDefault) {
  std::string xml(kXmlDecl);
  xml.append(
      "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "id=\"-1\" name=\"T1\" displayName=\"T1\" ref=\"A1:B5\">");
  xml.append("  <tableColumns count=\"1\"><tableColumn id=\"1\" name=\"X\"/></tableColumns>");
  xml.append("</table>");

  auto table_or = read_table(Bytes(xml), 0U);
  ASSERT_TRUE(static_cast<bool>(table_or));
  // Negative integers must not be reinterpreted as huge unsigned values.
  EXPECT_EQ(table_or.value().id, 0U);
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
