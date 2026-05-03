// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Unit tests for `formulon::io::read_pivot_cache_definition` and
// `formulon::io::read_pivot_cache_records`. Each test feeds a hand-rolled
// OOXML byte vector into the reader and asserts the populated
// `pivot::PivotCache` shape; the workbook integration (rels resolution,
// xlsx fixture loading) is covered separately.

#include "io/pivot_cache_reader.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "pivot/pivot_cache.h"
#include "utils/error.h"
#include "value.h"

namespace formulon::io {
namespace {

std::vector<std::uint8_t> Bytes(std::string_view xml) {
  return std::vector<std::uint8_t>(xml.begin(), xml.end());
}

constexpr std::string_view kXmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
constexpr std::string_view kPivotNs = " xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"";

// ---------------------------------------------------------------------------
// Definition: happy paths
// ---------------------------------------------------------------------------

TEST(PivotCacheReader, DefinitionMinimalSingleField) {
  std::string xml(kXmlDecl);
  xml.append("<pivotCacheDefinition").append(kPivotNs).append(">");
  xml.append("  <cacheSource type=\"worksheet\"/>");
  xml.append("  <cacheFields count=\"1\">");
  xml.append("    <cacheField name=\"Region\">");
  xml.append("      <sharedItems count=\"3\">");
  xml.append("        <s v=\"East\"/>");
  xml.append("        <n v=\"42\"/>");
  xml.append("        <m/>");
  xml.append("      </sharedItems>");
  xml.append("    </cacheField>");
  xml.append("  </cacheFields>");
  xml.append("</pivotCacheDefinition>");

  auto cache_or = read_pivot_cache_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(cache_or)) << "read failed: " << cache_or.error().message;
  const pivot::PivotCache& cache = cache_or.value();
  ASSERT_EQ(cache.fields().size(), 1U);
  EXPECT_EQ(cache.fields()[0].name, "Region");
  ASSERT_EQ(cache.fields()[0].shared_items.size(), 3U);
  EXPECT_TRUE(cache.fields()[0].shared_items[0].is_text());
  EXPECT_EQ(cache.fields()[0].shared_items[0].as_text(), "East");
  EXPECT_TRUE(cache.fields()[0].shared_items[1].is_number());
  EXPECT_DOUBLE_EQ(cache.fields()[0].shared_items[1].as_number(), 42.0);
  EXPECT_TRUE(cache.fields()[0].shared_items[2].is_blank());
}

TEST(PivotCacheReader, DefinitionMultipleFieldsMixedTypes) {
  std::string xml(kXmlDecl);
  xml.append("<pivotCacheDefinition").append(kPivotNs).append(">");
  xml.append("  <cacheFields count=\"2\">");
  xml.append("    <cacheField name=\"Region\">");
  xml.append("      <sharedItems>");
  xml.append("        <s v=\"East\"/>");
  xml.append("        <s v=\"West\"/>");
  xml.append("      </sharedItems>");
  xml.append("    </cacheField>");
  xml.append("    <cacheField name=\"Qty\">");
  xml.append("      <sharedItems>");
  xml.append("        <n v=\"1.5\"/>");
  xml.append("        <n v=\"-2\"/>");
  xml.append("        <n v=\"0\"/>");
  xml.append("      </sharedItems>");
  xml.append("    </cacheField>");
  xml.append("  </cacheFields>");
  xml.append("</pivotCacheDefinition>");

  auto cache_or = read_pivot_cache_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(cache_or));
  const pivot::PivotCache& cache = cache_or.value();
  ASSERT_EQ(cache.fields().size(), 2U);

  EXPECT_EQ(cache.fields()[0].name, "Region");
  ASSERT_EQ(cache.fields()[0].shared_items.size(), 2U);
  EXPECT_EQ(cache.fields()[0].shared_items[0].as_text(), "East");
  EXPECT_EQ(cache.fields()[0].shared_items[1].as_text(), "West");

  EXPECT_EQ(cache.fields()[1].name, "Qty");
  ASSERT_EQ(cache.fields()[1].shared_items.size(), 3U);
  EXPECT_DOUBLE_EQ(cache.fields()[1].shared_items[0].as_number(), 1.5);
  EXPECT_DOUBLE_EQ(cache.fields()[1].shared_items[1].as_number(), -2.0);
  EXPECT_DOUBLE_EQ(cache.fields()[1].shared_items[2].as_number(), 0.0);
}

TEST(PivotCacheReader, DefinitionBooleanAndErrorSharedItems) {
  std::string xml(kXmlDecl);
  xml.append("<pivotCacheDefinition").append(kPivotNs).append(">");
  xml.append("  <cacheFields count=\"1\">");
  xml.append("    <cacheField name=\"Mixed\">");
  xml.append("      <sharedItems>");
  xml.append("        <b v=\"1\"/>");
  xml.append("        <b v=\"0\"/>");
  xml.append("        <e v=\"#DIV/0!\"/>");
  xml.append("        <e v=\"#N/A\"/>");
  xml.append("      </sharedItems>");
  xml.append("    </cacheField>");
  xml.append("  </cacheFields>");
  xml.append("</pivotCacheDefinition>");

  auto cache_or = read_pivot_cache_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(cache_or));
  const pivot::PivotCache& cache = cache_or.value();
  ASSERT_EQ(cache.fields().size(), 1U);
  ASSERT_EQ(cache.fields()[0].shared_items.size(), 4U);
  EXPECT_TRUE(cache.fields()[0].shared_items[0].is_boolean());
  EXPECT_TRUE(cache.fields()[0].shared_items[0].as_boolean());
  EXPECT_TRUE(cache.fields()[0].shared_items[1].is_boolean());
  EXPECT_FALSE(cache.fields()[0].shared_items[1].as_boolean());
  EXPECT_TRUE(cache.fields()[0].shared_items[2].is_error());
  EXPECT_EQ(cache.fields()[0].shared_items[2].as_error(), ErrorCode::Div0);
  EXPECT_TRUE(cache.fields()[0].shared_items[3].is_error());
  EXPECT_EQ(cache.fields()[0].shared_items[3].as_error(), ErrorCode::NA);
}

TEST(PivotCacheReader, DefinitionRangeTypedFieldHasEmptySharedItems) {
  // `<sharedItems>` with `containsNumber="1"` and range attributes but
  // no per-item children describes a numeric range. The records part is
  // expected to carry inline `<n>` cells rather than `<x>` indices.
  std::string xml(kXmlDecl);
  xml.append("<pivotCacheDefinition").append(kPivotNs).append(">");
  xml.append("  <cacheFields count=\"1\">");
  xml.append("    <cacheField name=\"Sales\">");
  xml.append(
      "      <sharedItems containsString=\"0\" containsNumber=\"1\" containsInteger=\"1\" minValue=\"0\" "
      "maxValue=\"100\"/>");
  xml.append("    </cacheField>");
  xml.append("  </cacheFields>");
  xml.append("</pivotCacheDefinition>");

  auto cache_or = read_pivot_cache_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(cache_or));
  const pivot::PivotCache& cache = cache_or.value();
  ASSERT_EQ(cache.fields().size(), 1U);
  EXPECT_EQ(cache.fields()[0].name, "Sales");
  EXPECT_TRUE(cache.fields()[0].shared_items.empty());
}

// ---------------------------------------------------------------------------
// Definition: error paths
// ---------------------------------------------------------------------------

TEST(PivotCacheReader, DefinitionMalformedRootIsContentTypeInvalid) {
  std::string xml(kXmlDecl);
  xml.append("<foo").append(kPivotNs).append("><cacheFields/></foo>");

  auto cache_or = read_pivot_cache_definition(Bytes(xml));
  ASSERT_FALSE(static_cast<bool>(cache_or));
  EXPECT_EQ(cache_or.error().code, FormulonErrorCode::kIoContentTypeInvalid);
}

TEST(PivotCacheReader, DefinitionMalformedXmlIsParseError) {
  std::string xml(kXmlDecl);
  xml.append("<pivotCacheDefinition").append(kPivotNs).append("><cacheFields");

  auto cache_or = read_pivot_cache_definition(Bytes(xml));
  ASSERT_FALSE(static_cast<bool>(cache_or));
  EXPECT_EQ(cache_or.error().code, FormulonErrorCode::kIoXmlParse);
}

TEST(PivotCacheReader, DefinitionExternalCacheSourceIsContentTypeInvalid) {
  std::string xml(kXmlDecl);
  xml.append("<pivotCacheDefinition").append(kPivotNs).append(">");
  xml.append("  <cacheSource type=\"external\"/>");
  xml.append("  <cacheFields count=\"0\"/>");
  xml.append("</pivotCacheDefinition>");

  auto cache_or = read_pivot_cache_definition(Bytes(xml));
  ASSERT_FALSE(static_cast<bool>(cache_or));
  EXPECT_EQ(cache_or.error().code, FormulonErrorCode::kIoContentTypeInvalid);
}

TEST(PivotCacheReader, DefinitionUnparseableNumberIsCorruption) {
  std::string xml(kXmlDecl);
  xml.append("<pivotCacheDefinition").append(kPivotNs).append(">");
  xml.append("  <cacheFields count=\"1\">");
  xml.append("    <cacheField name=\"X\">");
  xml.append("      <sharedItems><n v=\"not-a-number\"/></sharedItems>");
  xml.append("    </cacheField>");
  xml.append("  </cacheFields>");
  xml.append("</pivotCacheDefinition>");

  auto cache_or = read_pivot_cache_definition(Bytes(xml));
  ASSERT_FALSE(static_cast<bool>(cache_or));
  EXPECT_EQ(cache_or.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

// ---------------------------------------------------------------------------
// Records: happy paths
// ---------------------------------------------------------------------------

TEST(PivotCacheReader, RecordsIndexedLookupResolvesSharedItems) {
  // Two-field discrete cache with three rows of `<x v="N">` indices.
  std::string def_xml(kXmlDecl);
  def_xml.append("<pivotCacheDefinition").append(kPivotNs).append(">");
  def_xml.append("  <cacheFields count=\"2\">");
  def_xml.append("    <cacheField name=\"Region\">");
  def_xml.append("      <sharedItems><s v=\"East\"/><s v=\"West\"/></sharedItems>");
  def_xml.append("    </cacheField>");
  def_xml.append("    <cacheField name=\"Status\">");
  def_xml.append("      <sharedItems><s v=\"Open\"/><s v=\"Closed\"/></sharedItems>");
  def_xml.append("    </cacheField>");
  def_xml.append("  </cacheFields>");
  def_xml.append("</pivotCacheDefinition>");
  auto cache_or = read_pivot_cache_definition(Bytes(def_xml));
  ASSERT_TRUE(static_cast<bool>(cache_or));
  pivot::PivotCache cache = std::move(cache_or.value());

  std::string rec_xml(kXmlDecl);
  rec_xml.append("<pivotCacheRecords").append(kPivotNs).append(" count=\"3\">");
  rec_xml.append("  <r><x v=\"0\"/><x v=\"1\"/></r>");
  rec_xml.append("  <r><x v=\"1\"/><x v=\"0\"/></r>");
  rec_xml.append("  <r><x v=\"0\"/><x v=\"0\"/></r>");
  rec_xml.append("</pivotCacheRecords>");
  auto status = read_pivot_cache_records(Bytes(rec_xml), cache);
  ASSERT_TRUE(static_cast<bool>(status)) << "records read failed: " << status.error().message;

  ASSERT_EQ(cache.records().size(), 3U);
  ASSERT_EQ(cache.records()[0].cells.size(), 2U);
  EXPECT_EQ(cache.records()[0].cells[0].as_text(), "East");
  EXPECT_EQ(cache.records()[0].cells[1].as_text(), "Closed");
  EXPECT_EQ(cache.records()[1].cells[0].as_text(), "West");
  EXPECT_EQ(cache.records()[1].cells[1].as_text(), "Open");
  EXPECT_EQ(cache.records()[2].cells[0].as_text(), "East");
  EXPECT_EQ(cache.records()[2].cells[1].as_text(), "Open");
}

TEST(PivotCacheReader, RecordsInlineNumericValues) {
  // Range-typed numeric field: shared_items is empty, records carry
  // inline `<n v="...">` cells.
  std::string def_xml(kXmlDecl);
  def_xml.append("<pivotCacheDefinition").append(kPivotNs).append(">");
  def_xml.append("  <cacheFields count=\"1\">");
  def_xml.append("    <cacheField name=\"Amount\">");
  def_xml.append("      <sharedItems containsNumber=\"1\" minValue=\"0\" maxValue=\"500\"/>");
  def_xml.append("    </cacheField>");
  def_xml.append("  </cacheFields>");
  def_xml.append("</pivotCacheDefinition>");
  auto cache_or = read_pivot_cache_definition(Bytes(def_xml));
  ASSERT_TRUE(static_cast<bool>(cache_or));
  pivot::PivotCache cache = std::move(cache_or.value());
  ASSERT_TRUE(cache.fields()[0].shared_items.empty());

  std::string rec_xml(kXmlDecl);
  rec_xml.append("<pivotCacheRecords").append(kPivotNs).append(" count=\"3\">");
  rec_xml.append("  <r><n v=\"100\"/></r>");
  rec_xml.append("  <r><n v=\"250.5\"/></r>");
  rec_xml.append("  <r><n v=\"0\"/></r>");
  rec_xml.append("</pivotCacheRecords>");
  auto status = read_pivot_cache_records(Bytes(rec_xml), cache);
  ASSERT_TRUE(static_cast<bool>(status));

  ASSERT_EQ(cache.records().size(), 3U);
  EXPECT_DOUBLE_EQ(cache.records()[0].cells[0].as_number(), 100.0);
  EXPECT_DOUBLE_EQ(cache.records()[1].cells[0].as_number(), 250.5);
  EXPECT_DOUBLE_EQ(cache.records()[2].cells[0].as_number(), 0.0);
}

// ---------------------------------------------------------------------------
// Records: error / forgiving paths
// ---------------------------------------------------------------------------

TEST(PivotCacheReader, RecordsIndexOutOfBoundsIsCorruption) {
  std::string def_xml(kXmlDecl);
  def_xml.append("<pivotCacheDefinition").append(kPivotNs).append(">");
  def_xml.append("  <cacheFields count=\"1\">");
  def_xml.append("    <cacheField name=\"R\">");
  def_xml.append("      <sharedItems><s v=\"A\"/><s v=\"B\"/></sharedItems>");
  def_xml.append("    </cacheField>");
  def_xml.append("  </cacheFields>");
  def_xml.append("</pivotCacheDefinition>");
  auto cache_or = read_pivot_cache_definition(Bytes(def_xml));
  ASSERT_TRUE(static_cast<bool>(cache_or));
  pivot::PivotCache cache = std::move(cache_or.value());

  std::string rec_xml(kXmlDecl);
  rec_xml.append("<pivotCacheRecords").append(kPivotNs).append(" count=\"1\">");
  rec_xml.append("  <r><x v=\"999\"/></r>");
  rec_xml.append("</pivotCacheRecords>");
  auto status = read_pivot_cache_records(Bytes(rec_xml), cache);
  ASSERT_FALSE(static_cast<bool>(status));
  EXPECT_EQ(status.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(PivotCacheReader, RecordsMissingTrailingValuesPadWithBlank) {
  // Two fields, but the only `<r>` carries one child. The reader pads
  // the second cell with `Value::blank()` so consumers can index by
  // field index without bounds-checking.
  std::string def_xml(kXmlDecl);
  def_xml.append("<pivotCacheDefinition").append(kPivotNs).append(">");
  def_xml.append("  <cacheFields count=\"2\">");
  def_xml.append("    <cacheField name=\"A\">");
  def_xml.append("      <sharedItems><s v=\"x\"/><s v=\"y\"/></sharedItems>");
  def_xml.append("    </cacheField>");
  def_xml.append("    <cacheField name=\"B\">");
  def_xml.append("      <sharedItems><s v=\"u\"/></sharedItems>");
  def_xml.append("    </cacheField>");
  def_xml.append("  </cacheFields>");
  def_xml.append("</pivotCacheDefinition>");
  auto cache_or = read_pivot_cache_definition(Bytes(def_xml));
  ASSERT_TRUE(static_cast<bool>(cache_or));
  pivot::PivotCache cache = std::move(cache_or.value());

  std::string rec_xml(kXmlDecl);
  rec_xml.append("<pivotCacheRecords").append(kPivotNs).append(" count=\"1\">");
  rec_xml.append("  <r><x v=\"0\"/></r>");
  rec_xml.append("</pivotCacheRecords>");
  auto status = read_pivot_cache_records(Bytes(rec_xml), cache);
  ASSERT_TRUE(static_cast<bool>(status));

  ASSERT_EQ(cache.records().size(), 1U);
  ASSERT_EQ(cache.records()[0].cells.size(), 2U);
  EXPECT_EQ(cache.records()[0].cells[0].as_text(), "x");
  EXPECT_TRUE(cache.records()[0].cells[1].is_blank());
}

TEST(PivotCacheReader, RecordsMalformedRootIsContentTypeInvalid) {
  pivot::PivotCache cache;  // empty cache is fine; root check fires first
  std::string xml(kXmlDecl);
  xml.append("<notRecords").append(kPivotNs).append("><r/></notRecords>");

  auto status = read_pivot_cache_records(Bytes(xml), cache);
  ASSERT_FALSE(static_cast<bool>(status));
  EXPECT_EQ(status.error().code, FormulonErrorCode::kIoContentTypeInvalid);
}

}  // namespace
}  // namespace formulon::io
