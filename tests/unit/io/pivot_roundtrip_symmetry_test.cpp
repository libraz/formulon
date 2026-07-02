// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Read<->write symmetry tests for the pivot parts, built on
// tests/support/roundtrip_symmetry.h. These complement (do not replace) the
// component reader/writer unit tests: they feed real Excel-shaped
// pivotCacheDefinition / pivotCacheRecords / pivotTableDefinition XML through
// `read_* -> write_*` and assert declaratively that the attribute set survives
// -- the "reader understands an attribute the writer drops" defect class.
//
// Covered: root/field passthrough attributes, sharedItems content hints,
// <fieldGroup>, the record index-vs-inline distinction (a 100.5 inline value
// must not collapse to an `<x v="100"/>` index), pivotField item `x=`
// ordering + hidden `h="1"`, and the subtotalTop / defaultSubtotal tri-state.

#include <cstdint>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/pivot_cache_reader.h"
#include "io/pivot_cache_writer.h"
#include "io/pivot_table_reader.h"
#include "io/pivot_table_writer.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pugixml.hpp"
#include "support/roundtrip_symmetry.h"

namespace formulon {
namespace {

std::vector<std::uint8_t> Bytes(const std::string& s) {
  return std::vector<std::uint8_t>(s.begin(), s.end());
}

// Default + relationships namespaces exactly as Excel writes them.
constexpr const char* kNs =
    " xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
    " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";

// ---------------------------------------------------------------------------
// pivotCacheDefinition
// ---------------------------------------------------------------------------

TEST(PivotSymmetry, CacheDefinitionAttributesAndHintsSurvive) {
  const std::string before_str = std::string("<?xml version=\"1.0\"?><pivotCacheDefinition") + kNs +
                                 " refreshedBy=\"Alice\" refreshOnLoad=\"1\" createdVersion=\"8\" updatedVersion=\"8\">"
                                 "<cacheSource type=\"worksheet\"><worksheetSource ref=\"A1:B4\" sheet=\"Data\"/>"
                                 "</cacheSource>"
                                 "<cacheFields count=\"3\">"
                                 // Discrete string field with shared items + hints.
                                 "<cacheField name=\"Region\"><sharedItems containsBlank=\"1\" count=\"2\">"
                                 "<s v=\"East\"/><s v=\"West\"/></sharedItems></cacheField>"
                                 // Range-typed numeric field: hints only, no shared items.
                                 "<cacheField name=\"Amount\"><sharedItems containsNumber=\"1\" containsString=\"0\" "
                                 "minValue=\"1\" maxValue=\"9\"/></cacheField>"
                                 // Grouped (non-database) field carrying a <fieldGroup>.
                                 "<cacheField name=\"Years\" databaseField=\"0\"><fieldGroup par=\"0\">"
                                 "<rangePr groupBy=\"years\"/><groupItems count=\"1\"><s v=\"2024\"/></groupItems>"
                                 "</fieldGroup></cacheField>"
                                 "</cacheFields></pivotCacheDefinition>";
  auto cache = io::read_pivot_cache_definition(Bytes(before_str));
  ASSERT_TRUE(static_cast<bool>(cache)) << "read failed: " << cache.error().message;
  const std::string after_str = io::write_pivot_cache_definition(cache.value());

  pugi::xml_document before;
  pugi::xml_document after;
  ASSERT_TRUE(test::parse_xml(before_str, &before));
  ASSERT_TRUE(test::parse_xml(after_str, &after));

  // Root passthrough attributes.
  EXPECT_TRUE(test::attributes_preserved(before, after, "//pivotCacheDefinition",
                                         {"refreshedBy", "refreshOnLoad", "createdVersion", "updatedVersion"}));
  // Worksheet source ref / sheet.
  EXPECT_TRUE(test::attributes_preserved(before, after, "//worksheetSource", {"ref", "sheet"}));
  // Field names.
  EXPECT_TRUE(test::attributes_preserved(before, after, "(//cacheField)[1]", {"name"}));
  EXPECT_TRUE(test::attributes_preserved(before, after, "(//cacheField)[3]", {"name", "databaseField"}));
  // sharedItems content hints (discrete + range-typed).
  EXPECT_TRUE(test::attributes_preserved(before, after, "//cacheField[@name='Region']/sharedItems",
                                         {"containsBlank", "count"}));
  EXPECT_TRUE(test::attributes_preserved(before, after, "//cacheField[@name='Amount']/sharedItems",
                                         {"containsNumber", "containsString", "minValue", "maxValue"}));
  // Shared item values.
  EXPECT_TRUE(test::attributes_preserved(before, after, "//cacheField[@name='Region']/sharedItems/s[1]", {"v"}));
  EXPECT_TRUE(test::attributes_preserved(before, after, "//cacheField[@name='Region']/sharedItems/s[2]", {"v"}));
  // <fieldGroup> survives verbatim.
  EXPECT_TRUE(after.select_node("//cacheField[@name='Years']/fieldGroup")) << after_str;
  EXPECT_TRUE(test::attributes_preserved(before, after, "//fieldGroup", {"par"}));
}

// ---------------------------------------------------------------------------
// pivotCacheRecords: index-vs-inline distinction
// ---------------------------------------------------------------------------

TEST(PivotSymmetry, CacheRecordsIndexAndInlineDistinctionSurvive) {
  // Field 0 (Region) has shared items -> records reference them by <x v="N"/>.
  // Field 1 (Amount) is range-typed -> records carry inline <n v="..."/>.
  const std::string def_str = std::string("<?xml version=\"1.0\"?><pivotCacheDefinition") + kNs +
                              " recordCount=\"2\">"
                              "<cacheSource type=\"worksheet\"><worksheetSource ref=\"A1:B3\" sheet=\"Data\"/>"
                              "</cacheSource><cacheFields count=\"2\">"
                              "<cacheField name=\"Region\"><sharedItems count=\"2\"><s v=\"East\"/><s v=\"West\"/>"
                              "</sharedItems></cacheField>"
                              "<cacheField name=\"Amount\"><sharedItems containsNumber=\"1\"/></cacheField>"
                              "</cacheFields></pivotCacheDefinition>";
  auto cache = io::read_pivot_cache_definition(Bytes(def_str));
  ASSERT_TRUE(static_cast<bool>(cache)) << "def read failed: " << cache.error().message;

  const std::string before_str = std::string("<?xml version=\"1.0\"?><pivotCacheRecords") + kNs +
                                 " count=\"2\">"
                                 "<r><x v=\"0\"/><n v=\"100.5\"/></r>"
                                 "<r><x v=\"1\"/><n v=\"3\"/></r></pivotCacheRecords>";
  pivot::PivotCache filled = std::move(cache.value());
  const auto rc = io::read_pivot_cache_records(Bytes(before_str), filled);
  ASSERT_TRUE(static_cast<bool>(rc)) << "records read failed: " << rc.error().message;
  const std::string after_str = io::write_pivot_cache_records(filled);

  pugi::xml_document before;
  pugi::xml_document after;
  ASSERT_TRUE(test::parse_xml(before_str, &before));
  ASSERT_TRUE(test::parse_xml(after_str, &after));

  // Index cells stay `<x>`; inline cells stay `<n>` with their exact value.
  EXPECT_TRUE(test::attributes_preserved(before, after, "//r[1]/x", {"v"}));
  EXPECT_TRUE(test::attributes_preserved(before, after, "//r[1]/n", {"v"}));
  EXPECT_TRUE(test::attributes_preserved(before, after, "//r[2]/x", {"v"}));
  // The decimal inline value must survive as `<n v="100.5"/>`, NOT truncate to
  // an `<x v="100"/>` index.
  EXPECT_TRUE(after.select_node("//r[1]/n[@v='100.5']")) << after_str;
  EXPECT_FALSE(after.select_node("//r[1]/x[@v='100']")) << after_str;
}

// ---------------------------------------------------------------------------
// pivotTableDefinition
// ---------------------------------------------------------------------------

TEST(PivotSymmetry, TableDefinitionAttributesItemsAndTriStateSurvive) {
  const std::string before_str =
      std::string("<?xml version=\"1.0\"?><pivotTableDefinition") + kNs +
      " name=\"P\" cacheId=\"0\" updatedVersion=\"8\" createdVersion=\"8\" itemPrintTitles=\"1\" indent=\"0\">"
      "<location ref=\"A1:B10\" firstHeaderRow=\"1\" firstDataRow=\"2\" firstDataCol=\"1\"/>"
      "<pivotFields count=\"2\">"
      // Row field: non-default tri-state (both OFF) + unmodelled attrs + items.
      "<pivotField axis=\"axisRow\" subtotalTop=\"0\" defaultSubtotal=\"0\" compact=\"0\" showAll=\"0\">"
      "<items count=\"3\"><item x=\"2\"/><item x=\"0\" h=\"1\"/><item x=\"1\"/></items></pivotField>"
      "<pivotField dataField=\"1\" compact=\"0\"/>"
      "</pivotFields>"
      "<rowFields count=\"1\"><field x=\"0\"/></rowFields>"
      "<dataFields count=\"1\"><dataField name=\"Sum of Amount\" fld=\"1\" baseField=\"0\" "
      "baseItem=\"0\"/></dataFields>"
      "</pivotTableDefinition>";
  auto table = io::read_pivot_table_definition(Bytes(before_str));
  ASSERT_TRUE(static_cast<bool>(table)) << "read failed: " << table.error().message;
  const std::string after_str = io::write_pivot_table_definition(table.value());

  pugi::xml_document before;
  pugi::xml_document after;
  ASSERT_TRUE(test::parse_xml(before_str, &before));
  ASSERT_TRUE(test::parse_xml(after_str, &after));

  // Root passthrough attributes.
  EXPECT_TRUE(
      test::attributes_preserved(before, after, "//pivotTableDefinition",
                                 {"name", "cacheId", "updatedVersion", "createdVersion", "itemPrintTitles", "indent"}));
  // Row pivotField: modelled tri-state (both explicitly OFF) + passthrough.
  EXPECT_TRUE(test::attributes_preserved(before, after, "(//pivotField)[1]",
                                         {"axis", "subtotalTop", "defaultSubtotal", "compact", "showAll"}));
  // Item `x=` ordering and the hidden flag `h="1"` on the middle item.
  EXPECT_TRUE(test::attributes_preserved(before, after, "(//pivotField)[1]/items/item[1]", {"x"}));
  EXPECT_TRUE(test::attributes_preserved(before, after, "(//pivotField)[1]/items/item[2]", {"x", "h"}));
  EXPECT_TRUE(test::attributes_preserved(before, after, "(//pivotField)[1]/items/item[3]", {"x"}));
  // rowFields / dataFields references.
  EXPECT_TRUE(test::attributes_preserved(before, after, "//rowFields/field[1]", {"x"}));
  EXPECT_TRUE(test::attributes_preserved(before, after, "//dataField[1]", {"name", "fld"}));
}

}  // namespace
}  // namespace formulon
