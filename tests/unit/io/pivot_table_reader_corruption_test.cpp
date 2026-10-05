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
TEST(PivotTableReader, MissingDataFieldNameIsCorruption) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <dataFields count=\"1\">");
  xml.append("    <dataField fld=\"1\" subtotal=\"sum\"/>");
  xml.append("  </dataFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoSheetCorrupt);
}
TEST(PivotTableReader, MissingLocationIsCorruption) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <pivotFields count=\"0\"/>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoSheetCorrupt);
}
TEST(PivotTableReader, UnparseableLocationRefIsCorruption) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"not-a-range\"/>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoSheetCorrupt);
}
TEST(PivotTableReader, MalformedRootIsContentTypeInvalid) {
  std::string xml(kXmlDecl);
  xml.append("<foo").append(kPivotNs).append("><location ref=\"A1\"/></foo>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoContentTypeInvalid);
}
TEST(PivotTableReader, MalformedXmlIsParseError) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append("><location");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoXmlParse);
}
TEST(PivotTableReader, UnmodelledChildrenCapturedAsPassthrough) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("<location ref=\"A1:B2\"/>");
  xml.append("<calculatedItems count=\"1\"><calculatedItem name=\"Avg\" formula=\"=A1/B1\"/></calculatedItems>");
  xml.append("<pivotTableStyleInfo name=\"PivotStyleLight16\" showRowHeaders=\"1\"/>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const std::string& passthrough = table_or.value().raw_passthrough_xml();
  // Both unrecognised elements survive verbatim (order-preserving).
  EXPECT_NE(passthrough.find("<calculatedItems"), std::string::npos);
  EXPECT_NE(passthrough.find("formula=\"=A1/B1\""), std::string::npos);
  EXPECT_NE(passthrough.find("<pivotTableStyleInfo"), std::string::npos);
  EXPECT_NE(passthrough.find("PivotStyleLight16"), std::string::npos);
  // Recognised elements (location) are NOT in the passthrough.
  EXPECT_EQ(passthrough.find("<location"), std::string::npos);
}

}  // namespace
}  // namespace formulon::io
