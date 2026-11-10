// Pivot-table definition reader tests grouped by XML concern.

#include "io/pivot_table_reader.h"

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <string_view>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/pivot_table_writer.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot_table_reader_test_helpers.h"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon::io {
namespace {
using namespace pivot_table_reader_test_support;
TEST(PivotTableReader, MinimalHappyPath) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P1\" cacheId=\"3\">");
  xml.append("  <location ref=\"A3:D10\"/>");
  xml.append("  <pivotFields count=\"0\"/>");
  xml.append("  <dataFields count=\"0\"/>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  EXPECT_EQ(table.name(), "P1");
  EXPECT_EQ(table.pivot_cache_id(), 3U);
  EXPECT_EQ(table.anchor_row(), 2U);
  EXPECT_EQ(table.anchor_col(), 0U);
  EXPECT_EQ(table.span_rows(), 8U);
  EXPECT_EQ(table.span_cols(), 4U);
  EXPECT_TRUE(table.fields().empty());
  EXPECT_TRUE(table.data_fields().empty());
}
TEST(PivotTableReader, MultipleFieldsAxisDecoding) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <pivotFields count=\"3\">");
  xml.append("    <pivotField axis=\"axisRow\" name=\"R\"/>");
  xml.append("    <pivotField axis=\"axisCol\"/>");
  xml.append("    <pivotField dataField=\"1\"/>");  // unset axis -> Value
  xml.append("  </pivotFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.fields().size(), 3U);
  EXPECT_EQ(table.fields()[0].axis, pivot::PivotAxis::Row);
  EXPECT_EQ(table.fields()[0].custom_name, "R");
  EXPECT_EQ(table.fields()[1].axis, pivot::PivotAxis::Col);
  EXPECT_EQ(table.fields()[2].axis, pivot::PivotAxis::Value);
}
TEST(PivotTableReader, ItemsDecodingSkipsSubtotalMarkers) {
  // Mix of real items (with `x` indices), a hidden item (`h="1"`), and a
  // subtotal marker (`t="default"`). The marker must be skipped; the
  // hidden item must round-trip with `visible = false`.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <pivotFields count=\"1\">");
  xml.append("    <pivotField axis=\"axisRow\">");
  xml.append("      <items count=\"4\">");
  xml.append("        <item x=\"0\"/>");
  xml.append("        <item h=\"1\" x=\"1\"/>");
  xml.append("        <item x=\"2\"/>");
  xml.append("        <item t=\"default\"/>");
  xml.append("      </items>");
  xml.append("    </pivotField>");
  xml.append("  </pivotFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.fields().size(), 1U);
  const auto& items = table.fields()[0].items;
  ASSERT_EQ(items.size(), 3U);
  EXPECT_TRUE(items[0].visible);
  EXPECT_FALSE(items[1].visible);
  EXPECT_TRUE(items[2].visible);
}
TEST(PivotTableReader, ItemWithMalformedCacheIndexIsRejected) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <pivotFields count=\"1\">");
  xml.append("    <pivotField axis=\"axisRow\"><items><item x=\"not-an-index\"/></items></pivotField>");
  xml.append("  </pivotFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoSheetCorrupt);
  EXPECT_NE(table_or.error().message.find("invalid x attribute"), std::string::npos);
}
TEST(PivotTableReader, RowAndColFieldOrder) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <pivotFields count=\"4\">");
  xml.append("    <pivotField axis=\"axisRow\"/>");
  xml.append("    <pivotField axis=\"axisRow\"/>");
  xml.append("    <pivotField axis=\"axisCol\"/>");
  xml.append("    <pivotField/>");
  xml.append("  </pivotFields>");
  xml.append("  <rowFields count=\"2\">");
  xml.append("    <field x=\"0\"/>");
  xml.append("    <field x=\"1\"/>");
  xml.append("  </rowFields>");
  xml.append("  <colFields count=\"1\">");
  xml.append("    <field x=\"2\"/>");
  xml.append("  </colFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.row_field_order().size(), 2U);
  EXPECT_EQ(table.row_field_order()[0], 0U);
  EXPECT_EQ(table.row_field_order()[1], 1U);
  ASSERT_EQ(table.col_field_order().size(), 1U);
  EXPECT_EQ(table.col_field_order()[0], 2U);
}
TEST(PivotTableReader, SortTypeDescendingReversesTheField) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <pivotFields count=\"1\">");
  xml.append("    <pivotField axis=\"axisRow\" sortType=\"descending\"/>");
  xml.append("  </pivotFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.fields().size(), 1U);
  EXPECT_FALSE(table.fields()[0].sort.ascending);
  EXPECT_FALSE(table.fields()[0].sort.manual);
}
TEST(PivotTableReader, SortTypeManualDefersToItemOrder) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <pivotFields count=\"1\">");
  xml.append("    <pivotField axis=\"axisRow\" sortType=\"manual\"/>");
  xml.append("  </pivotFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.fields().size(), 1U);
  EXPECT_TRUE(table.fields()[0].sort.manual);
}
TEST(PivotTableReader, MissingSortTypeKeepsAscendingDefault) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <pivotFields count=\"1\">");
  xml.append("    <pivotField axis=\"axisRow\"/>");
  xml.append("  </pivotFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.fields().size(), 1U);
  EXPECT_TRUE(table.fields()[0].sort.manual);
}
TEST(PivotTableReader, ExplicitAscendingSortTypeIsLabelOrder) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <pivotFields count=\"1\">");
  xml.append("    <pivotField axis=\"axisRow\" sortType=\"ascending\"/>");
  xml.append("  </pivotFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  EXPECT_TRUE(table_or.value().fields()[0].sort.ascending);
  EXPECT_FALSE(table_or.value().fields()[0].sort.manual);
}
TEST(PivotTableReader, ValuesFieldMarkerIsKeptOutOfFieldOrder) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <pivotFields count=\"2\">");
  xml.append("    <pivotField axis=\"axisCol\"/>");
  xml.append("    <pivotField dataField=\"1\"/>");
  xml.append("  </pivotFields>");
  xml.append("  <colFields count=\"2\">");
  xml.append("    <field x=\"0\"/>");
  xml.append("    <field x=\"-2\"/>");
  xml.append("  </colFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.col_field_order().size(), 1U);
  EXPECT_EQ(table.col_field_order()[0], 0U);
  ASSERT_TRUE(table.col_values_position().has_value());
  EXPECT_EQ(table.col_values_position().value(), 1U);
  EXPECT_FALSE(table.row_values_position().has_value());
}
TEST(PivotTableReader, MultipleDataFieldsOnSameSourceField) {
  // GETPIVOTDATA's display-name lookup requires that two data fields on
  // the same source column (Sum + Average of "Amount") survive as two
  // distinct PivotDataField entries.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <pivotFields count=\"4\">");
  xml.append("    <pivotField/>");
  xml.append("    <pivotField/>");
  xml.append("    <pivotField/>");
  xml.append("    <pivotField dataField=\"1\"/>");
  xml.append("  </pivotFields>");
  xml.append("  <dataFields count=\"2\">");
  xml.append("    <dataField name=\"Sum of Amount\" fld=\"3\" subtotal=\"sum\"/>");
  xml.append("    <dataField name=\"Average of Amount\" fld=\"3\" subtotal=\"average\"/>");
  xml.append("  </dataFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.data_fields().size(), 2U);
  EXPECT_EQ(table.data_fields()[0].name, "Sum of Amount");
  EXPECT_EQ(table.data_fields()[0].field_index, 3U);
  EXPECT_EQ(table.data_fields()[0].aggregation, pivot::Aggregation::Sum);
  EXPECT_EQ(table.data_fields()[1].name, "Average of Amount");
  EXPECT_EQ(table.data_fields()[1].field_index, 3U);
  EXPECT_EQ(table.data_fields()[1].aggregation, pivot::Aggregation::Average);
}
TEST(PivotTableReader, DataFieldWithEmptyOrMissingNameKeepsEmptyName) {
  for (const char* attr : {" name=\"\"", ""}) {
    std::string xml(kXmlDecl);
    xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
    xml.append("  <location ref=\"A1:B2\"/>");
    xml.append("  <pivotFields count=\"1\"><pivotField dataField=\"1\"/></pivotFields>");
    xml.append("  <dataFields count=\"1\">");
    xml.append("    <dataField").append(attr).append(" fld=\"0\" subtotal=\"sum\"/>");
    xml.append("  </dataFields>");
    xml.append("</pivotTableDefinition>");

    auto table_or = read_pivot_table_definition(Bytes(xml));
    ASSERT_TRUE(static_cast<bool>(table_or)) << "attr=[" << attr << "] read failed: " << table_or.error().message;
    ASSERT_EQ(table_or.value().data_fields().size(), 1U);
    EXPECT_TRUE(table_or.value().data_fields()[0].name.empty()) << "attr=[" << attr << "]";
    const std::string round = write_pivot_table_definition(table_or.value());
    EXPECT_NE(round.find("<dataField name=\"\""), std::string::npos) << round;
  }
}
TEST(PivotTableReader, DataFieldAggregationMappingExhaustive) {
  // One <dataField> per Aggregation enum value the OOXML subtotal
  // attribute can name.
  struct Case {
    const char* attr;
    pivot::Aggregation expected;
  };
  const Case cases[] = {
      {"sum", pivot::Aggregation::Sum},
      {"count", pivot::Aggregation::Count},
      {"average", pivot::Aggregation::Average},
      {"max", pivot::Aggregation::Max},
      {"min", pivot::Aggregation::Min},
      {"product", pivot::Aggregation::Product},
      {"countNums", pivot::Aggregation::CountNumbers},
      {"stdDev", pivot::Aggregation::StdDev},
      {"var", pivot::Aggregation::Var},
  };

  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <dataFields count=\"9\">");
  for (const Case& c : cases) {
    xml.append("    <dataField name=\"agg_")
        .append(c.attr)
        .append("\" fld=\"0\" subtotal=\"")
        .append(c.attr)
        .append("\"/>");
  }
  xml.append("  </dataFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  ASSERT_EQ(table.data_fields().size(), sizeof(cases) / sizeof(cases[0]));
  for (std::size_t i = 0; i < table.data_fields().size(); ++i) {
    EXPECT_EQ(table.data_fields()[i].aggregation, cases[i].expected) << "case=" << cases[i].attr;
  }
}
TEST(PivotTableReader, UnknownSubtotalFallsBackToSum) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"0\">");
  xml.append("  <location ref=\"A1:B2\"/>");
  xml.append("  <dataFields count=\"1\">");
  xml.append("    <dataField name=\"Weird of X\" fld=\"0\" subtotal=\"weirdo\"/>");
  xml.append("  </dataFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  ASSERT_EQ(table_or.value().data_fields().size(), 1U);
  EXPECT_EQ(table_or.value().data_fields()[0].aggregation, pivot::Aggregation::Sum);
}
TEST(PivotTableReader, PageFieldsCarryTheSelectedItem) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"7\">");
  xml.append("  <location ref=\"A1:C5\"/>");
  xml.append("  <pivotFields count=\"1\"><pivotField axis=\"axisPage\"/></pivotFields>");
  xml.append("  <pageFields count=\"1\">");
  xml.append("    <pageField fld=\"0\" item=\"2\"/>");
  xml.append("  </pageFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  const pivot::PivotTable& table = table_or.value();
  EXPECT_EQ(table.name(), "P");
  EXPECT_EQ(table.pivot_cache_id(), 7U);
  ASSERT_EQ(table.fields().size(), 1U);
  EXPECT_EQ(table.fields()[0].axis, pivot::PivotAxis::Page);

  ASSERT_EQ(table.page_fields().size(), 1U);
  EXPECT_EQ(table.page_fields()[0].field_index, 0U);
  ASSERT_TRUE(table.page_fields()[0].item_index.has_value());
  EXPECT_EQ(*table.page_fields()[0].item_index, 2U);
  // The file wrote no `hier`, so none is kept.
  EXPECT_FALSE(table.page_fields()[0].hierarchy.has_value());
}
TEST(PivotTableReader, PageFieldWithoutASelectionLeavesTheItemAbsent) {
  // The unfiltered state: Excel writes no `item`, and the field is showing
  // every item rather than item 0.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("  <location ref=\"A1:C5\"/>");
  xml.append("  <pivotFields count=\"1\"><pivotField axis=\"axisPage\"/></pivotFields>");
  xml.append("  <pageFields count=\"1\"><pageField fld=\"0\" hier=\"-1\"/></pageFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  ASSERT_EQ(table_or.value().page_fields().size(), 1U);
  EXPECT_FALSE(table_or.value().page_fields()[0].item_index.has_value());
}
TEST(PivotTableReader, PageFieldWithoutAFieldIndexIsSkipped) {
  // `fld` names the field; absent, the entry designates nothing. Defaulting
  // it to 0 would draw a page header over a field the user never filtered by.
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("  <location ref=\"A1:C5\"/>");
  xml.append("  <pivotFields count=\"2\"><pivotField axis=\"axisPage\"/><pivotField axis=\"axisRow\"/></pivotFields>");
  xml.append("  <pageFields count=\"2\"><pageField item=\"1\"/><pageField fld=\"0\"/></pageFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  ASSERT_EQ(table_or.value().page_fields().size(), 1U);
  EXPECT_EQ(table_or.value().page_fields()[0].field_index, 0U);
}
TEST(PivotTableReader, PageFieldOrderOverridesDocumentOrder) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("  <location ref=\"A1:C5\"/>");
  xml.append("  <pivotFields count=\"2\"><pivotField axis=\"axisPage\"/><pivotField axis=\"axisPage\"/></pivotFields>");
  xml.append("  <pageFields count=\"2\"><pageField fld=\"1\"/><pageField fld=\"0\"/></pageFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  EXPECT_EQ(table_or.value().page_field_order(), (std::vector<std::uint32_t>{1, 0}));
}
TEST(PivotTableReader, PageAxisWithoutAPageFieldsBlockFallsBackToDocumentOrder) {
  std::string xml(kXmlDecl);
  xml.append("<pivotTableDefinition").append(kPivotNs).append(" name=\"P\" cacheId=\"1\">");
  xml.append("  <location ref=\"A1:C5\"/>");
  xml.append("  <pivotFields count=\"3\">");
  xml.append("    <pivotField axis=\"axisRow\"/><pivotField axis=\"axisPage\"/><pivotField axis=\"axisPage\"/>");
  xml.append("  </pivotFields>");
  xml.append("</pivotTableDefinition>");

  auto table_or = read_pivot_table_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(table_or)) << "read failed: " << table_or.error().message;
  EXPECT_TRUE(table_or.value().page_fields().empty());
  EXPECT_EQ(table_or.value().page_field_order(), (std::vector<std::uint32_t>{1, 2}));
}

TEST(PivotTableReader, LastSavedPartIsTheEarliestCreated) {
  Workbook wb = Workbook::create_empty();
  ASSERT_TRUE(static_cast<bool>(wb.add_sheet_entered("First")));
  ASSERT_TRUE(static_cast<bool>(wb.add_sheet_entered("Second")));
  const auto add = [&wb](std::size_t sheet, const char* name, std::uint64_t key) {
    auto table = std::make_unique<pivot::PivotTable>();
    table->set_name(name);
    table->set_creation_order(key);
    wb.sheet(sheet).add_pivot_table(std::move(table));
  };
  add(0U, "A", 0U);
  add(0U, "B", 0U);
  add(0U, "C", 0U);
  // A pivot already keyed elsewhere in the workbook sets the floor for the new keys.
  add(1U, "Z", 7U);
  assign_read_creation_orders(wb, 0U, 0U);
  const auto& tables = wb.sheet(0U).pivot_tables();
  ASSERT_EQ(tables.size(), 3U);
  // The first part read is the newest and the last is the earliest created.
  EXPECT_GT(tables[0]->creation_order(), tables[1]->creation_order());
  EXPECT_GT(tables[1]->creation_order(), tables[2]->creation_order());
  EXPECT_GT(tables[2]->creation_order(), 7U);
  EXPECT_EQ(wb.sheet(1U).pivot_tables()[0]->creation_order(), 7U);

  // Pivots read earlier in the same sheet keep their keys.
  add(0U, "D", 0U);
  const std::uint64_t first = tables[0]->creation_order();
  assign_read_creation_orders(wb, 0U, 3U);
  EXPECT_EQ(tables[0]->creation_order(), first);
  EXPECT_GT(tables[3]->creation_order(), first);
}

// Excel's pivot_blocked.xlsx: a report saved blocked, a one-cell location with
// zero offsets over the anchor A3 of Sheet2, whose rich value is {6, 8, 8, 7}.
std::vector<std::uint8_t> PivotBlockedSample() {
  return test::read_file_bytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/rich_errors/pivot_blocked.xlsx");
}

std::vector<std::uint8_t> ReplacePart(const std::vector<std::uint8_t>& package, const std::string& name,
                                      const std::string& body) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(test::span_of(package))));
  mz_zip_archive writer{};
  EXPECT_TRUE(mz_zip_writer_init_heap(&writer, 0, 4096) != MZ_FALSE);
  for (const std::string& entry : zip.list_entries()) {
    auto part_or = zip.read_entry(entry);
    EXPECT_TRUE(static_cast<bool>(part_or));
    const std::vector<std::uint8_t> part =
        entry == name ? std::vector<std::uint8_t>(body.begin(), body.end()) : part_or.value();
    EXPECT_TRUE(mz_zip_writer_add_mem(&writer, entry.c_str(), part.data(), part.size(),
                                      static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)) != MZ_FALSE);
  }
  void* archive = nullptr;
  std::size_t archive_size = 0;
  EXPECT_TRUE(mz_zip_writer_finalize_heap_archive(&writer, &archive, &archive_size) != MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<std::uint8_t*>(archive),
                                static_cast<std::uint8_t*>(archive) + archive_size);
  mz_free(archive);
  mz_zip_writer_end(&writer);
  return out;
}

TEST(PivotTableReader, BlockedAnchorRestoresTheBlockedState) {
  auto loaded = read_ooxml(test::span_of(PivotBlockedSample()));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  Workbook& wb = loaded.value().workbook;
  const Sheet* sheet = wb.sheet_by_name("Sheet2");
  ASSERT_NE(sheet, nullptr);
  ASSERT_EQ(sheet->pivot_tables().size(), 1U);
  const pivot::PivotTable& table = *sheet->pivot_tables()[0];
  ASSERT_TRUE(table.blocked().has_value());
  const pivot::BlockedReport loaded_state = *table.blocked();
  // The body rows and columns come from the rich value's offsets; the size is re-planned.
  EXPECT_EQ(loaded_state.rows - loaded_state.rows_above, 9U);
  EXPECT_EQ(loaded_state.cols, 7U);
  EXPECT_FALSE(loaded_state.size_known);
  EXPECT_FALSE(loaded_state.past_edge);
  EXPECT_EQ(loaded_state.cause, SpillCause::kOccupied);
  const Cell* anchor = sheet->cell_at(table.anchor_row(), table.anchor_col());
  ASSERT_NE(anchor, nullptr);
  EXPECT_TRUE(anchor->cached_value.is_error() && anchor->cached_value.as_error() == ErrorCode::Spill);

  // The first recalc plans the report; it is still blocked, now with a known size.
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_TRUE(table.blocked().has_value());
  EXPECT_TRUE(table.blocked()->size_known);
  EXPECT_EQ(table.blocked()->cols, 7U);
  EXPECT_EQ(table.blocked()->rows - table.blocked()->rows_above, 9U);
}

TEST(PivotTableReader, AOneCellReportWithoutARichErrorIsNotBlocked) {
  const std::vector<std::uint8_t> sample = PivotBlockedSample();
  std::string sheet_xml;
  ASSERT_TRUE(test::extract_part(test::span_of(sample), "xl/worksheets/sheet1.xml", &sheet_xml));
  const std::string anchor = "<c r=\"A3\" s=\"1\" t=\"e\" vm=\"1\"><v>#VALUE!</v></c>";
  const std::size_t at = sheet_xml.find(anchor);
  ASSERT_NE(at, std::string::npos) << sheet_xml;
  sheet_xml.replace(at, anchor.size(), "<c r=\"A3\" s=\"1\"/>");
  auto loaded = read_ooxml(test::span_of(ReplacePart(sample, "xl/worksheets/sheet1.xml", sheet_xml)));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  const Sheet* sheet = loaded.value().workbook.sheet_by_name("Sheet2");
  ASSERT_NE(sheet, nullptr);
  ASSERT_EQ(sheet->pivot_tables().size(), 1U);
  EXPECT_FALSE(sheet->pivot_tables()[0]->blocked().has_value());
}

}  // namespace
}  // namespace formulon::io
