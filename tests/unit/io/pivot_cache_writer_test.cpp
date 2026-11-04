//
// Unit tests for `formulon::io::write_pivot_cache_definition` and
// `formulon::io::write_pivot_cache_records`. Each test builds a hand-
// rolled `pivot::PivotCache`, writes the two parts, and (where useful)
// pipes the bytes back through the symmetric reader to assert a clean
// round-trip on the data the readers actually consume.

#include "io/pivot_cache_writer.h"

#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_writer.h"
#include "io/pivot_cache_reader.h"
#include "io/zip_reader.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pivot/record_access.h"
#include "value.h"
#include "workbook.h"

namespace formulon::io {
namespace {

std::vector<std::uint8_t> Bytes(std::string_view xml) {
  return std::vector<std::uint8_t>(xml.begin(), xml.end());
}

// Convenience: appends `s` to `cache.mutable_text_storage()` and returns
// a `Value::text` aliasing that storage. Mirrors how the reader hands
// out text payloads, so values built this way live as long as `cache`.
Value MakeText(pivot::PivotCache& cache, std::string s) {
  cache.mutable_text_storage().push_back(std::move(s));
  return Value::text(cache.mutable_text_storage().back());
}

// ---------------------------------------------------------------------------
// Empty cache
// ---------------------------------------------------------------------------

TEST(PivotCacheWriter, EmptyCacheRoundTrips) {
  pivot::PivotCache empty;
  const std::string def_xml = write_pivot_cache_definition(empty);
  const std::string rec_xml = write_pivot_cache_records(empty);

  auto def_or = read_pivot_cache_definition(Bytes(def_xml));
  ASSERT_TRUE(static_cast<bool>(def_or)) << "definition read failed: " << def_or.error().message;
  pivot::PivotCache parsed = std::move(def_or.value());
  EXPECT_TRUE(parsed.fields().empty());

  auto rec_or = read_pivot_cache_records(Bytes(rec_xml), parsed);
  ASSERT_TRUE(static_cast<bool>(rec_or)) << "records read failed: " << rec_or.error().message;
  EXPECT_TRUE(parsed.records().empty());
}

// ---------------------------------------------------------------------------
// Definition: round-trip via reader
// ---------------------------------------------------------------------------

TEST(PivotCacheWriter, DefinitionRoundTripsThroughReader) {
  pivot::PivotCache cache;
  // Field 0: shared-items text field "Region" with two values.
  pivot::PivotCacheField region;
  region.name = "Region";
  region.shared_items.push_back(MakeText(cache, "North"));
  region.shared_items.push_back(MakeText(cache, "South"));
  cache.mutable_fields().push_back(std::move(region));

  // Field 1: range-typed numeric field "Amount" with empty shared_items.
  pivot::PivotCacheField amount;
  amount.name = "Amount";
  cache.mutable_fields().push_back(std::move(amount));

  const std::string xml = write_pivot_cache_definition(cache);
  auto parsed_or = read_pivot_cache_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  const pivot::PivotCache& parsed = parsed_or.value();

  ASSERT_EQ(parsed.fields().size(), 2U);
  EXPECT_EQ(parsed.fields()[0].name, "Region");
  ASSERT_EQ(parsed.fields()[0].shared_items.size(), 2U);
  EXPECT_TRUE(parsed.fields()[0].shared_items[0].is_text());
  EXPECT_EQ(parsed.fields()[0].shared_items[0].as_text(), "North");
  EXPECT_TRUE(parsed.fields()[0].shared_items[1].is_text());
  EXPECT_EQ(parsed.fields()[0].shared_items[1].as_text(), "South");

  EXPECT_EQ(parsed.fields()[1].name, "Amount");
  EXPECT_TRUE(parsed.fields()[1].shared_items.empty());
}

// ---------------------------------------------------------------------------
// Records: round-trip via reader
// ---------------------------------------------------------------------------

TEST(PivotCacheWriter, RecordsRoundTripThroughReader) {
  // Same shape as DefinitionRoundTripsThroughReader: field 0 is shared
  // (Region), field 1 is range-typed numeric (Amount). Records carry a
  // mix of `<x>` indices (resolved to "North" / "South") and inline
  // numeric cells (100 / 200 / 300).
  pivot::PivotCache cache;
  pivot::PivotCacheField region;
  region.name = "Region";
  region.shared_items.push_back(MakeText(cache, "North"));
  region.shared_items.push_back(MakeText(cache, "South"));
  cache.mutable_fields().push_back(std::move(region));

  pivot::PivotCacheField amount;
  amount.name = "Amount";
  cache.mutable_fields().push_back(std::move(amount));

  pivot::PivotCacheRecord r0;
  r0.cells.push_back(Value::number(0.0));  // index 0 -> "North"
  r0.cells.push_back(Value::number(100.0));
  cache.mutable_records().push_back(std::move(r0));
  pivot::PivotCacheRecord r1;
  r1.cells.push_back(Value::number(1.0));  // index 1 -> "South"
  r1.cells.push_back(Value::number(200.0));
  cache.mutable_records().push_back(std::move(r1));
  pivot::PivotCacheRecord r2;
  r2.cells.push_back(Value::number(0.0));  // index 0 -> "North"
  r2.cells.push_back(Value::number(300.0));
  cache.mutable_records().push_back(std::move(r2));

  const std::string def_xml = write_pivot_cache_definition(cache);
  const std::string rec_xml = write_pivot_cache_records(cache);

  auto parsed_or = read_pivot_cache_definition(Bytes(def_xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or));
  pivot::PivotCache parsed = std::move(parsed_or.value());
  auto status = read_pivot_cache_records(Bytes(rec_xml), parsed);
  ASSERT_TRUE(static_cast<bool>(status));

  ASSERT_EQ(parsed.records().size(), 3U);
  ASSERT_EQ(parsed.records()[0].cells.size(), 2U);

  // The shared field (0) stores the raw shared_items index as a Number; the
  // resolved value comes from the `cell_value` accessor. The range-typed
  // field (1) stores its value inline.
  EXPECT_TRUE(parsed.records()[0].cells[0].is_number());
  EXPECT_DOUBLE_EQ(parsed.records()[0].cells[0].as_number(), 0.0);
  EXPECT_EQ(pivot::cell_value(parsed, parsed.records()[0], 0).as_text(), "North");
  EXPECT_TRUE(parsed.records()[0].cells[1].is_number());
  EXPECT_DOUBLE_EQ(parsed.records()[0].cells[1].as_number(), 100.0);

  EXPECT_TRUE(parsed.records()[1].cells[0].is_number());
  EXPECT_DOUBLE_EQ(parsed.records()[1].cells[0].as_number(), 1.0);
  EXPECT_EQ(pivot::cell_value(parsed, parsed.records()[1], 0).as_text(), "South");
  EXPECT_DOUBLE_EQ(parsed.records()[1].cells[1].as_number(), 200.0);

  EXPECT_TRUE(parsed.records()[2].cells[0].is_number());
  EXPECT_DOUBLE_EQ(parsed.records()[2].cells[0].as_number(), 0.0);
  EXPECT_EQ(pivot::cell_value(parsed, parsed.records()[2], 0).as_text(), "North");
  EXPECT_DOUBLE_EQ(parsed.records()[2].cells[1].as_number(), 300.0);
}

// ---------------------------------------------------------------------------
// Discrete numeric shared field: read -> resolve -> write -> re-read identity
// ---------------------------------------------------------------------------

TEST(PivotCacheWriter, DiscreteNumericFieldReadWriteRereadIdentity) {
  // A discrete numeric field is a shared field whose shared_items are
  // numbers. The record cells carry `<x>` indices into those items, so the
  // resolved value must come through the `cell_value` accessor on both the
  // initial read and after a write/re-read round-trip.
  constexpr std::string_view kDefXml =
      "<?xml version=\"1.0\"?>"
      "<pivotCacheDefinition>"
      "<cacheFields count=\"2\">"
      "<cacheField name=\"Score\">"
      "<sharedItems><n v=\"1.5\"/><n v=\"-2\"/><n v=\"0\"/></sharedItems>"
      "</cacheField>"
      "<cacheField name=\"Amount\">"
      "<sharedItems containsNumber=\"1\"/>"
      "</cacheField>"
      "</cacheFields>"
      "</pivotCacheDefinition>";
  // Field 0 indices: 2 -> 0.0, 0 -> 1.5, 1 -> -2.0. Field 1 inline numbers.
  constexpr std::string_view kRecXml =
      "<?xml version=\"1.0\"?>"
      "<pivotCacheRecords count=\"3\">"
      "<r><x v=\"2\"/><n v=\"10\"/></r>"
      "<r><x v=\"0\"/><n v=\"20\"/></r>"
      "<r><x v=\"1\"/><n v=\"30\"/></r>"
      "</pivotCacheRecords>";

  auto def_or = read_pivot_cache_definition(Bytes(kDefXml));
  ASSERT_TRUE(static_cast<bool>(def_or)) << "definition read failed: " << def_or.error().message;
  pivot::PivotCache cache1 = std::move(def_or.value());
  auto rec_status = read_pivot_cache_records(Bytes(kRecXml), cache1);
  ASSERT_TRUE(static_cast<bool>(rec_status)) << "records read failed: " << rec_status.error().message;

  ASSERT_EQ(cache1.records().size(), 3U);

  // Initial read: shared field cells hold raw indices; resolution via the
  // accessor must yield the discrete numeric values.
  EXPECT_TRUE(cache1.records()[0].cells[0].is_number());
  EXPECT_DOUBLE_EQ(cache1.records()[0].cells[0].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(pivot::cell_value(cache1, cache1.records()[0], 0).as_number(), 0.0);
  EXPECT_DOUBLE_EQ(pivot::cell_value(cache1, cache1.records()[1], 0).as_number(), 1.5);
  EXPECT_DOUBLE_EQ(pivot::cell_value(cache1, cache1.records()[2], 0).as_number(), -2.0);

  // Inline range-typed field 1 carries the values directly.
  EXPECT_DOUBLE_EQ(cache1.records()[0].cells[1].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(cache1.records()[1].cells[1].as_number(), 20.0);
  EXPECT_DOUBLE_EQ(cache1.records()[2].cells[1].as_number(), 30.0);

  // Write back out and re-read: the resolved discrete values must survive
  // unchanged (index-vs-value confusion would corrupt them here).
  const std::string def_xml2 = write_pivot_cache_definition(cache1);
  const std::string rec_xml2 = write_pivot_cache_records(cache1);

  auto def2_or = read_pivot_cache_definition(Bytes(def_xml2));
  ASSERT_TRUE(static_cast<bool>(def2_or)) << "re-read definition failed: " << def2_or.error().message;
  pivot::PivotCache cache2 = std::move(def2_or.value());
  auto rec2_status = read_pivot_cache_records(Bytes(rec_xml2), cache2);
  ASSERT_TRUE(static_cast<bool>(rec2_status)) << "re-read records failed: " << rec2_status.error().message;

  ASSERT_EQ(cache2.records().size(), 3U);
  EXPECT_DOUBLE_EQ(pivot::cell_value(cache2, cache2.records()[0], 0).as_number(), 0.0);
  EXPECT_DOUBLE_EQ(pivot::cell_value(cache2, cache2.records()[1], 0).as_number(), 1.5);
  EXPECT_DOUBLE_EQ(pivot::cell_value(cache2, cache2.records()[2], 0).as_number(), -2.0);

  EXPECT_DOUBLE_EQ(cache2.records()[0].cells[1].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(cache2.records()[1].cells[1].as_number(), 20.0);
  EXPECT_DOUBLE_EQ(cache2.records()[2].cells[1].as_number(), 30.0);
}

// ---------------------------------------------------------------------------
// All five Value kinds in shared_items
// ---------------------------------------------------------------------------

TEST(PivotCacheWriter, AllValueKindsInSharedItems) {
  pivot::PivotCache cache;
  pivot::PivotCacheField mixed;
  mixed.name = "Mixed";
  mixed.shared_items.push_back(MakeText(cache, "hello"));
  mixed.shared_items.push_back(Value::number(42.5));
  mixed.shared_items.push_back(Value::boolean(true));
  mixed.shared_items.push_back(Value::blank());
  mixed.shared_items.push_back(Value::error(ErrorCode::Ref));
  cache.mutable_fields().push_back(std::move(mixed));

  const std::string xml = write_pivot_cache_definition(cache);
  auto parsed_or = read_pivot_cache_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  const pivot::PivotCache& parsed = parsed_or.value();
  ASSERT_EQ(parsed.fields().size(), 1U);
  ASSERT_EQ(parsed.fields()[0].shared_items.size(), 5U);

  EXPECT_TRUE(parsed.fields()[0].shared_items[0].is_text());
  EXPECT_EQ(parsed.fields()[0].shared_items[0].as_text(), "hello");
  EXPECT_TRUE(parsed.fields()[0].shared_items[1].is_number());
  EXPECT_DOUBLE_EQ(parsed.fields()[0].shared_items[1].as_number(), 42.5);
  EXPECT_TRUE(parsed.fields()[0].shared_items[2].is_boolean());
  EXPECT_TRUE(parsed.fields()[0].shared_items[2].as_boolean());
  EXPECT_TRUE(parsed.fields()[0].shared_items[3].is_blank());
  EXPECT_TRUE(parsed.fields()[0].shared_items[4].is_error());
  EXPECT_EQ(parsed.fields()[0].shared_items[4].as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// Number formatting round-trips through strtod cleanly
// ---------------------------------------------------------------------------

TEST(PivotCacheWriter, NumberFormatRoundTrips) {
  pivot::PivotCache cache;
  pivot::PivotCacheField nums;
  nums.name = "N";
  nums.shared_items.push_back(Value::number(0.1));
  nums.shared_items.push_back(Value::number(1.5));
  nums.shared_items.push_back(Value::number(1e20));
  nums.shared_items.push_back(Value::number(-7.0));
  nums.shared_items.push_back(Value::number(0.0));
  cache.mutable_fields().push_back(std::move(nums));

  const std::string xml = write_pivot_cache_definition(cache);
  auto parsed_or = read_pivot_cache_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  const pivot::PivotCache& parsed = parsed_or.value();
  ASSERT_EQ(parsed.fields()[0].shared_items.size(), 5U);
  EXPECT_DOUBLE_EQ(parsed.fields()[0].shared_items[0].as_number(), 0.1);
  EXPECT_DOUBLE_EQ(parsed.fields()[0].shared_items[1].as_number(), 1.5);
  EXPECT_DOUBLE_EQ(parsed.fields()[0].shared_items[2].as_number(), 1e20);
  EXPECT_DOUBLE_EQ(parsed.fields()[0].shared_items[3].as_number(), -7.0);
  EXPECT_DOUBLE_EQ(parsed.fields()[0].shared_items[4].as_number(), 0.0);
}

// The spelling, not just the value: the cache writer shares its dtoa with
// the sheet writer, so a number reads the same wherever in the package it
// lands -- and the same way Excel writes it. `%.17g` also round-trips,
// which is why the value assertions above cannot see the difference, but
// it renders 0.1 as "0.10000000000000001".
TEST(PivotCacheWriter, NumbersUseTheExcelForm) {
  pivot::PivotCache cache;
  pivot::PivotCacheField nums;
  nums.name = "N";
  nums.shared_items.push_back(Value::number(0.1));
  nums.shared_items.push_back(Value::number(1.0 / 3.0));
  nums.shared_items.push_back(Value::number(1e-7));
  nums.shared_items.push_back(Value::number(-0.0));
  cache.mutable_fields().push_back(std::move(nums));

  const std::string xml = write_pivot_cache_definition(cache);
  EXPECT_NE(xml.find("<n v=\"0.1\"/>"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<n v=\"0.33333333333333331\"/>"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<n v=\"9.9999999999999995E-8\"/>"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<n v=\"0\"/>"), std::string::npos) << xml;
  EXPECT_EQ(xml.find("0.10000000000000001"), std::string::npos) << xml;
}

// ---------------------------------------------------------------------------
// XML escaping
// ---------------------------------------------------------------------------

TEST(PivotCacheWriter, XmlEscapingApplies) {
  pivot::PivotCache cache;
  pivot::PivotCacheField field;
  field.name = "<R&D>";
  field.shared_items.push_back(MakeText(cache, "a&b"));
  field.shared_items.push_back(MakeText(cache, "quote\"here"));
  field.shared_items.push_back(MakeText(cache, "<tag>"));
  cache.mutable_fields().push_back(std::move(field));

  const std::string xml = write_pivot_cache_definition(cache);

  // Sanity: escaped sequences are present in the bytes; the raw '<' / '&'
  // inside the field name and shared item text must not survive verbatim
  // as element/attribute characters.
  EXPECT_NE(xml.find("&lt;R&amp;D&gt;"), std::string::npos);
  EXPECT_NE(xml.find("a&amp;b"), std::string::npos);
  EXPECT_NE(xml.find("quote&quot;here"), std::string::npos);
  EXPECT_NE(xml.find("&lt;tag&gt;"), std::string::npos);

  // Round-trip: the reader must recover the original (unescaped) bytes.
  auto parsed_or = read_pivot_cache_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(parsed_or)) << "read failed: " << parsed_or.error().message;
  const pivot::PivotCache& parsed = parsed_or.value();
  ASSERT_EQ(parsed.fields().size(), 1U);
  EXPECT_EQ(parsed.fields()[0].name, "<R&D>");
  ASSERT_EQ(parsed.fields()[0].shared_items.size(), 3U);
  EXPECT_EQ(parsed.fields()[0].shared_items[0].as_text(), "a&b");
  EXPECT_EQ(parsed.fields()[0].shared_items[1].as_text(), "quote\"here");
  EXPECT_EQ(parsed.fields()[0].shared_items[2].as_text(), "<tag>");
}

// ---------------------------------------------------------------------------
// Smoke: leading bytes form a well-shaped XML declaration
// ---------------------------------------------------------------------------

TEST(PivotCacheWriter, EmittedXmlIsValidUTF8WithBOMlessHeader) {
  pivot::PivotCache cache;
  const std::string def_xml = write_pivot_cache_definition(cache);
  const std::string rec_xml = write_pivot_cache_records(cache);

  // No UTF-8 BOM, declaration leads.
  ASSERT_GE(def_xml.size(), 6U);
  EXPECT_NE(static_cast<unsigned char>(def_xml[0]), 0xEFU);
  EXPECT_EQ(def_xml.substr(0, 19), std::string("<?xml version=\"1.0\""));

  ASSERT_GE(rec_xml.size(), 6U);
  EXPECT_NE(static_cast<unsigned char>(rec_xml[0]), 0xEFU);
  EXPECT_EQ(rec_xml.substr(0, 19), std::string("<?xml version=\"1.0\""));
}

// ---------------------------------------------------------------------------
// recordCount attribute matches records().size()
// ---------------------------------------------------------------------------

TEST(PivotCacheWriter, RecordCountAttributeMatchesRecords) {
  pivot::PivotCache cache;
  pivot::PivotCacheField f;
  f.name = "X";
  f.shared_items.push_back(MakeText(cache, "a"));
  cache.mutable_fields().push_back(std::move(f));

  // Push five records; the value content does not matter for this test.
  for (int i = 0; i < 5; ++i) {
    pivot::PivotCacheRecord r;
    r.cells.push_back(Value::number(0.0));
    cache.mutable_records().push_back(std::move(r));
  }

  const std::string def_xml = write_pivot_cache_definition(cache);
  EXPECT_NE(def_xml.find("recordCount=\"5\""), std::string::npos);

  const std::string rec_xml = write_pivot_cache_records(cache);
  EXPECT_NE(rec_xml.find("count=\"5\""), std::string::npos);
}

// ---------------------------------------------------------------------------
// refreshOnLoad: set for in-memory caches, left alone for captured ones
// ---------------------------------------------------------------------------

TEST(PivotCacheWriter, AuthoredCacheAsksExcelToRefreshOnLoad) {
  pivot::PivotCache cache;
  EXPECT_NE(write_pivot_cache_definition(cache).find(" refreshOnLoad=\"1\""), std::string::npos);
}

TEST(PivotCacheWriter, CapturedRootAttributesSuppressSynthesisedRefreshOnLoad) {
  pivot::PivotCache cache;
  cache.mutable_passthrough_attrs().emplace_back("createdVersion", "8");
  const std::string xml = write_pivot_cache_definition(cache);
  EXPECT_EQ(xml.find("refreshOnLoad"), std::string::npos) << xml;
  EXPECT_NE(xml.find("createdVersion=\"8\""), std::string::npos) << xml;
}

// ---------------------------------------------------------------------------
// Record cell index / inline discrimination (item d)
// ---------------------------------------------------------------------------

TEST(PivotCacheWriter, InlineNumberInSharedFieldStaysInline) {
  // A shared field whose record carries an inline numeric value (not an
  // <x> index) must round-trip as <n>, not be mistaken for an index.
  pivot::PivotCache cache;
  pivot::PivotCacheField f;
  f.name = "Region";
  f.shared_items.push_back(MakeText(cache, "East"));
  f.shared_items.push_back(MakeText(cache, "West"));
  cache.mutable_fields().push_back(std::move(f));

  // Record 0: shared index 1 ("West"). Record 1: an inline number 42
  // that happens to look like it could be an index but is not.
  pivot::PivotCacheRecord r0;
  r0.cells.push_back(Value::number(1.0));
  r0.cell_is_index.push_back(true);
  cache.mutable_records().push_back(std::move(r0));
  pivot::PivotCacheRecord r1;
  r1.cells.push_back(Value::number(42.0));
  r1.cell_is_index.push_back(false);
  cache.mutable_records().push_back(std::move(r1));

  const std::string rec_xml = write_pivot_cache_records(cache);
  EXPECT_NE(rec_xml.find("<x v=\"1\"/>"), std::string::npos);
  EXPECT_NE(rec_xml.find("<n v=\"42\"/>"), std::string::npos);
  // The inline 42 must NOT be emitted as an index.
  EXPECT_EQ(rec_xml.find("<x v=\"42\"/>"), std::string::npos);

  // Round-trip: the reader must restore both the values and the flags.
  auto rec_or = read_pivot_cache_records(Bytes(rec_xml), cache);
  ASSERT_TRUE(static_cast<bool>(rec_or)) << "records read failed: " << rec_or.error().message;
}

TEST(PivotCacheWriter, IndexInlineFlagsSurviveReadRoundTrip) {
  pivot::PivotCache cache;
  pivot::PivotCacheField shared;
  shared.name = "Region";
  shared.shared_items.push_back(MakeText(cache, "East"));
  shared.shared_items.push_back(MakeText(cache, "West"));
  cache.mutable_fields().push_back(std::move(shared));
  pivot::PivotCacheField range;
  range.name = "Amount";  // range-typed: no shared_items
  cache.mutable_fields().push_back(std::move(range));

  pivot::PivotCacheRecord r;
  r.cells.push_back(Value::number(0.0));  // index -> "East"
  r.cell_is_index.push_back(true);
  r.cells.push_back(Value::number(99.5));  // inline number
  r.cell_is_index.push_back(false);
  cache.mutable_records().push_back(std::move(r));

  const std::string rec_xml = write_pivot_cache_records(cache);
  std::vector<std::uint8_t> bytes = Bytes(rec_xml);
  // Reparse into a fresh cache with the same field shape.
  pivot::PivotCache reparsed;
  {
    pivot::PivotCacheField s2;
    s2.name = "Region";
    s2.shared_items.push_back(MakeText(reparsed, "East"));
    s2.shared_items.push_back(MakeText(reparsed, "West"));
    reparsed.mutable_fields().push_back(std::move(s2));
    pivot::PivotCacheField a2;
    a2.name = "Amount";
    reparsed.mutable_fields().push_back(std::move(a2));
  }
  auto rec_or = read_pivot_cache_records(bytes, reparsed);
  ASSERT_TRUE(static_cast<bool>(rec_or)) << "records read failed: " << rec_or.error().message;
  ASSERT_EQ(reparsed.records().size(), 1U);
  const pivot::PivotCacheRecord& rr = reparsed.records()[0];
  ASSERT_EQ(rr.cell_is_index.size(), 2U);
  EXPECT_TRUE(rr.cell_is_index[0]);
  EXPECT_FALSE(rr.cell_is_index[1]);
  // cell_value resolves the index to the shared value and returns the
  // inline number verbatim.
  EXPECT_EQ(pivot::cell_value(reparsed, rr, 0).as_text(), "East");
  EXPECT_EQ(pivot::cell_value(reparsed, rr, 1).as_number(), 99.5);
}

// ---------------------------------------------------------------------------
// M-23: sharedItems content hints round-trip verbatim, including the
// default-true ones and hints on a discrete field.
// ---------------------------------------------------------------------------

TEST(PivotCacheWriter, DefaultTrueHintRoundTripsAsFalse) {
  // A field that explicitly turns off a default-true hint
  // (containsSemiMixedTypes="0") must re-emit "0", not drop it (which would
  // let the reader's default flip it back to true).
  const std::string_view def =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" recordCount=\"0\">"
      "<cacheSource type=\"worksheet\"/><cacheFields count=\"1\">"
      "<cacheField name=\"N\"><sharedItems containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\" "
      "minValue=\"1\" maxValue=\"9\"/></cacheField></cacheFields></pivotCacheDefinition>";
  auto def_or = read_pivot_cache_definition(Bytes(def));
  ASSERT_TRUE(static_cast<bool>(def_or)) << def_or.error().message;
  const pivot::SharedItemsHints& h = def_or.value().fields()[0].shared_items_hints;
  EXPECT_TRUE(h.has_contains_semi_mixed);
  EXPECT_FALSE(h.contains_semi_mixed);

  const std::string xml = write_pivot_cache_definition(def_or.value());
  EXPECT_NE(xml.find("containsSemiMixedTypes=\"0\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("containsString=\"0\""), std::string::npos) << xml;
  // Re-read: the "0" values survive rather than reverting to default true.
  auto reparsed = read_pivot_cache_definition(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(reparsed)) << reparsed.error().message;
  EXPECT_FALSE(reparsed.value().fields()[0].shared_items_hints.contains_semi_mixed);
}

TEST(PivotCacheWriter, DiscreteFieldHintsSurvive) {
  // A discrete (shared-items) field with a containsBlank hint must keep it
  // through the round trip (previously discrete-field hints were dropped).
  const std::string_view def =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" recordCount=\"0\">"
      "<cacheSource type=\"worksheet\"/><cacheFields count=\"1\">"
      "<cacheField name=\"R\"><sharedItems containsBlank=\"1\" count=\"2\"><s v=\"a\"/><s v=\"b\"/></sharedItems>"
      "</cacheField></cacheFields></pivotCacheDefinition>";
  auto def_or = read_pivot_cache_definition(Bytes(def));
  ASSERT_TRUE(static_cast<bool>(def_or)) << def_or.error().message;
  EXPECT_TRUE(def_or.value().fields()[0].shared_items_hints.contains_blank);

  const std::string xml = write_pivot_cache_definition(def_or.value());
  EXPECT_NE(xml.find("containsBlank=\"1\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("<s v=\"a\"/>"), std::string::npos) << xml;
}

// ---------------------------------------------------------------------------
// H-19: a grouping-derived field (databaseField="0") is excluded from record
// output, its <fieldGroup> round-trips, and record arity = database fields.
// ---------------------------------------------------------------------------

TEST(PivotCacheWriter, GroupedFieldExcludedFromRecordsAndFieldGroupSurvives) {
  const std::string_view def =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" recordCount=\"2\">"
      "<cacheSource type=\"worksheet\"/><cacheFields count=\"2\">"
      "<cacheField name=\"Date\"><sharedItems containsNumber=\"1\" minValue=\"1\" maxValue=\"9\"/></cacheField>"
      "<cacheField name=\"Years\" databaseField=\"0\"><fieldGroup par=\"0\"><rangePr groupBy=\"years\"/>"
      "<groupItems count=\"1\"><s v=\"2024\"/></groupItems></fieldGroup></cacheField>"
      "</cacheFields></pivotCacheDefinition>";
  auto def_or = read_pivot_cache_definition(Bytes(def));
  ASSERT_TRUE(static_cast<bool>(def_or)) << def_or.error().message;
  pivot::PivotCache cache = std::move(def_or.value());
  ASSERT_EQ(cache.fields().size(), 2U);
  EXPECT_TRUE(cache.fields()[0].is_database_field);
  EXPECT_FALSE(cache.fields()[1].is_database_field);
  ASSERT_TRUE(cache.fields()[1].field_group.has_value());
  EXPECT_EQ(cache.fields()[1].field_group->par, std::optional<std::uint32_t>(0U));

  // Records carry one cell per database field (the Date field only).
  const std::string_view rec =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"2\">"
      "<r><n v=\"3\"/></r><r><n v=\"7\"/></r></pivotCacheRecords>";
  auto rec_or = read_pivot_cache_records(Bytes(rec), cache);
  ASSERT_TRUE(static_cast<bool>(rec_or)) << rec_or.error().message;
  ASSERT_EQ(cache.records().size(), 2U);
  // Cells are aligned to cache.fields(): Date populated, group slot blank.
  EXPECT_DOUBLE_EQ(cache.records()[0].cells[0].as_number(), 3.0);
  EXPECT_TRUE(cache.records()[0].cells[1].is_blank());

  const std::string def_xml = write_pivot_cache_definition(cache);
  EXPECT_NE(def_xml.find("databaseField=\"0\""), std::string::npos) << def_xml;
  EXPECT_NE(def_xml.find("<fieldGroup"), std::string::npos) << def_xml;

  const std::string rec_xml = write_pivot_cache_records(cache);
  // Each <r> emits exactly one cell (the database field), not two.
  EXPECT_NE(rec_xml.find("<r><n v=\"3\"/></r>"), std::string::npos) << rec_xml;
  EXPECT_NE(rec_xml.find("<r><n v=\"7\"/></r>"), std::string::npos) << rec_xml;
}

// ---------------------------------------------------------------------------
// Unmodelled pivotCacheDefinition root attributes round-trip verbatim.
// ---------------------------------------------------------------------------

TEST(PivotCacheWriter, UnknownRootAttributesRoundTrip) {
  const std::string_view def =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\" "
      "recordCount=\"0\" refreshedBy=\"Alice\" refreshOnLoad=\"1\" createdVersion=\"8\">"
      "<cacheSource type=\"worksheet\"/><cacheFields count=\"0\"/></pivotCacheDefinition>";
  auto def_or = read_pivot_cache_definition(Bytes(def));
  ASSERT_TRUE(static_cast<bool>(def_or)) << def_or.error().message;
  const auto& attrs = def_or.value().passthrough_attrs();
  auto has_attr = [&](const std::string& n, const std::string& v) {
    for (const auto& [name, value] : attrs) {
      if (name == n) {
        return value == v;
      }
    }
    return false;
  };
  EXPECT_TRUE(has_attr("refreshedBy", "Alice"));
  EXPECT_TRUE(has_attr("createdVersion", "8"));
  // refreshOnLoad is modelled, not passed through.
  EXPECT_FALSE(has_attr("refreshOnLoad", "1"));
  EXPECT_TRUE(def_or.value().refresh_on_load());
  // r:id / recordCount / namespace decls are not captured as passthrough.
  for (const auto& [name, value] : attrs) {
    (void)value;
    EXPECT_NE(name, "r:id");
    EXPECT_NE(name, "recordCount");
    EXPECT_NE(name, "xmlns");
    EXPECT_NE(name, "xmlns:r");
  }

  const std::string xml = write_pivot_cache_definition(def_or.value());
  EXPECT_NE(xml.find("refreshedBy=\"Alice\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("refreshOnLoad=\"1\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("createdVersion=\"8\""), std::string::npos) << xml;
}

// ---------------------------------------------------------------------------
// cacheField numFmtId and <fieldGroup>
// ---------------------------------------------------------------------------

// The shape Excel saves for a date field grouped by months: the base keeps
// its shared items and points at the group field through `par`; the group
// field derives from the base and lists every month plus the two boundaries.
constexpr std::string_view kMonthGroupedCache =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\" "
    "refreshedBy=\"x\" recordCount=\"0\"><cacheSource type=\"worksheet\"><worksheetSource ref=\"A1:B3\" "
    "sheet=\"Sheet1\"/></cacheSource><cacheFields count=\"3\">"
    "<cacheField name=\"Date\" numFmtId=\"14\"><sharedItems containsSemiMixedTypes=\"0\" containsNonDate=\"0\" "
    "containsDate=\"1\" containsString=\"0\" minDate=\"2024-01-15T00:00:00\" maxDate=\"2024-07-23T00:00:00\" "
    "count=\"2\"><d v=\"2024-01-15T00:00:00\"/><d v=\"2024-07-22T00:00:00\"/></sharedItems>"
    "<fieldGroup par=\"2\"/></cacheField>"
    "<cacheField name=\"Amt\" numFmtId=\"0\"><sharedItems containsNumber=\"1\"/></cacheField>"
    "<cacheField name=\"Month (Date)\" numFmtId=\"0\" databaseField=\"0\"><fieldGroup base=\"0\">"
    "<rangePr groupBy=\"months\" startDate=\"2024-01-15T00:00:00\" endDate=\"2024-07-23T00:00:00\"/>"
    "<groupItems count=\"4\"><s v=\"&lt;2024/1/15\"/><s v=\"1\"/><s v=\"2\"/><s v=\">2024/7/23\"/></groupItems>"
    "</fieldGroup></cacheField></cacheFields></pivotCacheDefinition>";

TEST(PivotCacheWriter, NumFmtIdIsModelledAndAbsenceSurvives) {
  auto cache_or = read_pivot_cache_definition(Bytes(kMonthGroupedCache));
  ASSERT_TRUE(static_cast<bool>(cache_or)) << cache_or.error().message;
  const pivot::PivotCache& cache = cache_or.value();
  EXPECT_EQ(cache.fields()[0].num_fmt_id, std::optional<std::uint32_t>(14U));
  EXPECT_EQ(cache.fields()[2].num_fmt_id, std::optional<std::uint32_t>(0U));
  for (const pivot::PivotCacheField& field : cache.fields()) {
    for (const auto& attr : field.extra_attrs) {
      EXPECT_NE(attr.first, "numFmtId");
    }
  }
  const std::string xml = write_pivot_cache_definition(cache);
  EXPECT_NE(xml.find("<cacheField name=\"Date\" numFmtId=\"14\">"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<cacheField name=\"Month (Date)\" numFmtId=\"0\" databaseField=\"0\">"), std::string::npos)
      << xml;

  pivot::PivotCache bare;
  bare.mutable_fields().push_back(pivot::PivotCacheField{"Region", {}});
  EXPECT_EQ(write_pivot_cache_definition(bare).find("numFmtId"), std::string::npos);
}

TEST(PivotCacheWriter, DateRangeGroupIsReadStructurally) {
  auto cache_or = read_pivot_cache_definition(Bytes(kMonthGroupedCache));
  ASSERT_TRUE(static_cast<bool>(cache_or)) << cache_or.error().message;
  const pivot::PivotCache& cache = cache_or.value();
  ASSERT_TRUE(cache.fields()[0].field_group.has_value());
  const pivot::FieldGroup& base = *cache.fields()[0].field_group;
  EXPECT_EQ(base.par, std::optional<std::uint32_t>(2U));
  EXPECT_FALSE(base.base.has_value());
  EXPECT_EQ(base.kind, pivot::GroupKind::None);
  EXPECT_FALSE(cache.fields()[1].field_group.has_value());

  ASSERT_TRUE(cache.fields()[2].field_group.has_value());
  const pivot::FieldGroup& group = *cache.fields()[2].field_group;
  EXPECT_FALSE(group.par.has_value());
  EXPECT_EQ(group.base, std::optional<std::uint32_t>(0U));
  ASSERT_EQ(group.kind, pivot::GroupKind::Range);
  EXPECT_EQ(group.range.by, pivot::GroupBy::Months);
  EXPECT_TRUE(group.range.is_date);
  EXPECT_EQ(group.range.start_date, "2024-01-15T00:00:00");
  EXPECT_EQ(group.range.end_date, "2024-07-23T00:00:00");
  EXPECT_TRUE(group.range.auto_start);
  EXPECT_TRUE(group.range.auto_end);
  EXPECT_EQ(group.range.interval, 1.0);
  ASSERT_EQ(group.group_items.size(), 4U);
  EXPECT_EQ(group.group_items[0].as_text(), "<2024/1/15");
  EXPECT_EQ(group.group_items[3].as_text(), ">2024/7/23");
  EXPECT_TRUE(cache.fields()[2].shared_items.empty());
}

TEST(PivotCacheWriter, DateRangeGroupWritesAsExcelDoes) {
  auto cache_or = read_pivot_cache_definition(Bytes(kMonthGroupedCache));
  ASSERT_TRUE(static_cast<bool>(cache_or)) << cache_or.error().message;
  const std::string xml = write_pivot_cache_definition(cache_or.value());
  EXPECT_NE(xml.find("<d v=\"2024-01-15T00:00:00\"/><d v=\"2024-07-22T00:00:00\"/></sharedItems><fieldGroup par=\"2\"/>"
                     "</cacheField>"),
            std::string::npos)
      << xml;
  // A group field carries no `<sharedItems>`.
  EXPECT_NE(xml.find("databaseField=\"0\"><fieldGroup base=\"0\"><rangePr groupBy=\"months\" "
                     "startDate=\"2024-01-15T00:00:00\" endDate=\"2024-07-23T00:00:00\"/><groupItems count=\"4\">"
                     "<s v=\"&lt;2024/1/15\"/><s v=\"1\"/><s v=\"2\"/>"),
            std::string::npos)
      << xml;
}

TEST(PivotCacheWriter, FieldGroupShapesRoundTrip) {
  const std::string_view def =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" recordCount=\"0\">"
      "<cacheSource type=\"worksheet\"/><cacheFields count=\"4\">"
      "<cacheField name=\"N\"><sharedItems containsNumber=\"1\" minValue=\"1\" maxValue=\"9\"/></cacheField>"
      "<cacheField name=\"N2\" databaseField=\"0\"><fieldGroup base=\"0\" future=\"x\"><rangePr autoStart=\"0\" "
      "autoEnd=\"0\" startNum=\"0.5\" endNum=\"10\" groupInterval=\"2.5\"/><groupItems count=\"3\"><s v=\"a\"/>"
      "<n v=\"2\"/><m/></groupItems></fieldGroup></cacheField>"
      "<cacheField name=\"D\" databaseField=\"0\"><fieldGroup base=\"0\"><rangePr groupBy=\"days\" "
      "startDate=\"2024-01-15T00:00:00\" endDate=\"2024-07-23T00:00:00\" groupInterval=\"7\"/></fieldGroup>"
      "</cacheField>"
      "<cacheField name=\"G\" databaseField=\"0\"><fieldGroup base=\"0\"><discretePr count=\"3\"><x v=\"1\"/>"
      "<x v=\"0\"/><x v=\"1\"/></discretePr><groupItems count=\"2\"><s v=\"Group1\"/><s v=\"Group2\"/>"
      "</groupItems></fieldGroup></cacheField></cacheFields></pivotCacheDefinition>";
  auto first_or = read_pivot_cache_definition(Bytes(def));
  ASSERT_TRUE(static_cast<bool>(first_or)) << first_or.error().message;
  const pivot::PivotCache& first = first_or.value();

  const pivot::FieldGroup& numeric = *first.fields()[1].field_group;
  EXPECT_EQ(numeric.range.by, pivot::GroupBy::Range);
  EXPECT_FALSE(numeric.range.is_date);
  EXPECT_FALSE(numeric.range.auto_start);
  EXPECT_FALSE(numeric.range.auto_end);
  EXPECT_EQ(numeric.range.start_num, 0.5);
  EXPECT_EQ(numeric.range.end_num, 10.0);
  EXPECT_EQ(numeric.range.interval, 2.5);
  ASSERT_EQ(numeric.extra_attrs.size(), 1U);
  EXPECT_EQ(numeric.extra_attrs[0], (std::pair<std::string, std::string>("future", "x")));
  const pivot::FieldGroup& days = *first.fields()[2].field_group;
  EXPECT_EQ(days.range.by, pivot::GroupBy::Days);
  EXPECT_EQ(days.range.interval, 7.0);
  EXPECT_TRUE(days.group_items.empty());
  const pivot::FieldGroup& discrete = *first.fields()[3].field_group;
  EXPECT_EQ(discrete.kind, pivot::GroupKind::Discrete);
  EXPECT_EQ(discrete.discrete_map, (std::vector<std::uint32_t>{1U, 0U, 1U}));

  const std::string written = write_pivot_cache_definition(first);
  EXPECT_NE(written.find("<rangePr autoStart=\"0\" autoEnd=\"0\" startNum=\"0.5\" endNum=\"10\" "
                         "groupInterval=\"2.5\"/>"),
            std::string::npos)
      << written;
  EXPECT_NE(written.find("<rangePr groupBy=\"days\" startDate=\"2024-01-15T00:00:00\" "
                         "endDate=\"2024-07-23T00:00:00\" groupInterval=\"7\"/></fieldGroup>"),
            std::string::npos)
      << written;
  EXPECT_NE(written.find("<discretePr count=\"3\"><x v=\"1\"/><x v=\"0\"/><x v=\"1\"/></discretePr>"),
            std::string::npos)
      << written;

  auto second_or = read_pivot_cache_definition(Bytes(written));
  ASSERT_TRUE(static_cast<bool>(second_or)) << second_or.error().message;
  EXPECT_EQ(write_pivot_cache_definition(second_or.value()), written);
  const pivot::FieldGroup& again = *second_or.value().fields()[1].field_group;
  ASSERT_EQ(again.group_items.size(), 3U);
  EXPECT_EQ(again.group_items[0].as_text(), "a");
  EXPECT_EQ(again.group_items[1].as_number(), 2.0);
  EXPECT_TRUE(again.group_items[2].is_blank());
}

TEST(PivotCacheWriter, DatesRoundTripAsDatesInSharedGroupAndRecordItems) {
  const std::string_view def =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" recordCount=\"2\">"
      "<cacheSource type=\"worksheet\"/><cacheFields count=\"3\">"
      "<cacheField name=\"Mixed\"><sharedItems containsMixedTypes=\"1\" count=\"2\"><d v=\"2024-01-15T13:30:00\"/>"
      "<n v=\"45306\"/></sharedItems></cacheField>"
      "<cacheField name=\"When\"><sharedItems containsSemiMixedTypes=\"0\" containsNonDate=\"0\" containsDate=\"1\" "
      "containsString=\"0\" minDate=\"2024-01-15T00:00:00\" maxDate=\"2024-01-17T00:00:00\"/></cacheField>"
      "<cacheField name=\"G\" databaseField=\"0\"><fieldGroup base=\"0\"><discretePr count=\"2\"><x v=\"0\"/>"
      "<x v=\"0\"/></discretePr><groupItems count=\"2\"><d v=\"2024-01-15T00:00:00\"/><n v=\"7\"/></groupItems>"
      "</fieldGroup></cacheField></cacheFields></pivotCacheDefinition>";
  auto def_or = read_pivot_cache_definition(Bytes(def));
  ASSERT_TRUE(static_cast<bool>(def_or)) << def_or.error().message;
  pivot::PivotCache cache = std::move(def_or.value());
  EXPECT_EQ(cache.fields()[0].shared_item_is_date, (std::vector<bool>{true, false}));
  EXPECT_EQ(cache.fields()[2].field_group->group_item_is_date, (std::vector<bool>{true, false}));

  const std::string_view rec =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"2\">"
      "<r><x v=\"0\"/><d v=\"2024-01-16T06:00:00\"/></r><r><x v=\"1\"/><n v=\"5\"/></r></pivotCacheRecords>";
  ASSERT_TRUE(static_cast<bool>(read_pivot_cache_records(Bytes(rec), cache)));
  EXPECT_EQ(cache.records()[0].cell_is_date, (std::vector<bool>{false, true, false}));
  EXPECT_TRUE(cache.records()[1].cell_is_date.empty());

  const std::string def_xml = write_pivot_cache_definition(cache);
  EXPECT_NE(def_xml.find("<d v=\"2024-01-15T13:30:00\"/><n v=\"45306\"/></sharedItems>"), std::string::npos) << def_xml;
  EXPECT_NE(def_xml.find("<groupItems count=\"2\"><d v=\"2024-01-15T00:00:00\"/><n v=\"7\"/></groupItems>"),
            std::string::npos)
      << def_xml;
  const std::string rec_xml = write_pivot_cache_records(cache);
  EXPECT_NE(rec_xml.find("<r><x v=\"0\"/><d v=\"2024-01-16T06:00:00\"/></r><r><x v=\"1\"/><n v=\"5\"/></r>"),
            std::string::npos)
      << rec_xml;
}

// ---------------------------------------------------------------------------
// Calculated fields / items, cache options and saveData
// ---------------------------------------------------------------------------

constexpr std::string_view kRootNs =
    " xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
    " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";

std::string DefinitionXml(std::string_view root_attrs, std::string_view fields, std::string_view trailing) {
  std::string xml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<pivotCacheDefinition");
  xml.append(kRootNs).append(root_attrs).append(">");
  xml.append("<cacheSource type=\"worksheet\"><worksheetSource ref=\"A1:D6\" sheet=\"Sheet1\"/></cacheSource>");
  xml.append("<cacheFields count=\"2\">").append(fields).append("</cacheFields>");
  xml.append(trailing).append("</pivotCacheDefinition>");
  return xml;
}

pivot::PivotCache ReadCache(std::string_view xml) {
  auto cache_or = read_pivot_cache_definition(Bytes(xml));
  EXPECT_TRUE(static_cast<bool>(cache_or)) << (cache_or ? "" : cache_or.error().message);
  return cache_or ? std::move(cache_or.value()) : pivot::PivotCache{};
}

bool Has(const std::string& haystack, std::string_view needle) {
  return haystack.find(needle) != std::string::npos;
}

// The measured shape of a calculated field.
constexpr std::string_view kCalcFieldXml =
    "<cacheField name=\"Price\" numFmtId=\"0\" formula=\"Amt/Qty\" databaseField=\"0\"/>";

TEST(PivotCacheCalculated, CalculatedFieldRoundTripsAsMeasured) {
  const std::string input =
      DefinitionXml(" refreshedBy=\"libraz\"",
                    "<cacheField name=\"Amt\" numFmtId=\"0\"><sharedItems containsNumber=\"1\"/></cacheField>" +
                        std::string(kCalcFieldXml),
                    "");
  const pivot::PivotCache cache = ReadCache(input);
  const std::string out = write_pivot_cache_definition(cache);
  EXPECT_TRUE(Has(out, kCalcFieldXml)) << out;
  // The formula is stored without a leading '=' and is not an extra attribute.
  EXPECT_FALSE(Has(out, "formula=\"=")) << out;
  EXPECT_EQ(write_pivot_cache_definition(ReadCache(out)), out);
}

TEST(PivotCacheCalculated, AuthoredCalculatedFieldWritesNoChildren) {
  pivot::PivotCache cache;
  pivot::PivotCacheField calc;
  calc.name = "Price";
  calc.formula = "Amt/Qty";
  calc.is_database_field = false;
  calc.num_fmt_id = 0U;
  cache.mutable_fields().push_back(std::move(calc));
  EXPECT_TRUE(Has(write_pivot_cache_definition(cache), kCalcFieldXml));
}

constexpr std::string_view kCalcItemsXml =
    "<calculatedItems count=\"1\"><calculatedItem formula=\"Region[North]+Region[South]\">"
    "<pivotArea cacheIndex=\"1\" outline=\"0\" fieldPosition=\"0\"><references count=\"1\">"
    "<reference field=\"0\" count=\"1\"><x v=\"3\"/></reference></references></pivotArea>"
    "</calculatedItem></calculatedItems>";

constexpr std::string_view kCalcItemFields =
    "<cacheField name=\"Region\" numFmtId=\"0\"><sharedItems count=\"4\"><s v=\"East\"/><s v=\"North\"/>"
    "<s v=\"South\"/><s v=\"X\" f=\"1\"/></sharedItems></cacheField>"
    "<cacheField name=\"Amt\" numFmtId=\"0\"><sharedItems containsNumber=\"1\"/></cacheField>";

TEST(PivotCacheCalculated, CalculatedItemRoundTripsAsMeasured) {
  const std::string input = DefinitionXml(" refreshedBy=\"libraz\"", kCalcItemFields, kCalcItemsXml);
  const std::string out = write_pivot_cache_definition(ReadCache(input));
  EXPECT_TRUE(Has(out, "<s v=\"X\" f=\"1\"/>")) << out;
  EXPECT_FALSE(Has(out, "<s v=\"South\" f=")) << out;
  EXPECT_TRUE(Has(out, kCalcItemsXml)) << out;
  EXPECT_EQ(write_pivot_cache_definition(ReadCache(out)), out);
}

TEST(PivotCacheCalculated, CalculatedItemsSitBetweenTheSurroundingTrailingChildren) {
  const std::string trailing =
      "<cacheHierarchies count=\"0\"/>" + std::string(kCalcItemsXml) + "<extLst><ext uri=\"{u}\"/></extLst>";
  const std::string out =
      write_pivot_cache_definition(ReadCache(DefinitionXml(" refreshedBy=\"x\"", kCalcItemFields, trailing)));
  EXPECT_TRUE(Has(out, "</cacheFields>" + trailing + "</pivotCacheDefinition>")) << out;
}

TEST(PivotCacheCalculated, ExtLstWithoutCalculatedItemsStaysAfterTheCacheHierarchies) {
  const std::string trailing = "<cacheHierarchies count=\"0\"/><extLst><ext uri=\"{u}\"/></extLst>";
  pivot::PivotCache cache = ReadCache(DefinitionXml(" refreshedBy=\"x\"", kCalcItemFields, trailing));
  EXPECT_TRUE(Has(write_pivot_cache_definition(cache), "</cacheFields>" + trailing + "</pivotCacheDefinition>"));
  pivot::PivotCalculatedItem item;
  item.formula = "Region[East]";
  item.area.references.push_back({});
  item.area.references.back().field = 0U;
  item.area.references.back().items = {3U};
  cache.mutable_calculated_items().push_back(std::move(item));
  const std::string out = write_pivot_cache_definition(cache);
  const std::size_t hierarchies = out.find("<cacheHierarchies");
  const std::size_t calculated = out.find("<calculatedItems");
  const std::size_t ext = out.find("<extLst");
  ASSERT_NE(calculated, std::string::npos) << out;
  EXPECT_LT(hierarchies, calculated);
  EXPECT_LT(calculated, ext);
}

TEST(PivotCacheCalculated, UnknownAreaAttributesAndChildrenSurvive) {
  const std::string items =
      "<calculatedItems count=\"1\"><calculatedItem field=\"0\" formula=\"Region[East]\">"
      "<pivotArea cacheIndex=\"1\" dataOnly=\"0\"><references count=\"1\"><reference field=\"0\" count=\"1\" "
      "selected=\"0\"><x v=\"3\"/></reference></references><extLst><ext uri=\"{a}\"/></extLst></pivotArea>"
      "<extLst><ext uri=\"{b}\"/></extLst></calculatedItem></calculatedItems>";
  const std::string out = write_pivot_cache_definition(ReadCache(DefinitionXml("", kCalcItemFields, items)));
  EXPECT_TRUE(Has(out, items)) << out;
}

TEST(PivotCacheOptions, SaveDataLimitAndRefreshOnLoadRoundTrip) {
  const std::string input = DefinitionXml(
      " refreshedBy=\"x\" saveData=\"0\" missingItemsLimit=\"0\" refreshOnLoad=\"1\"", kCalcItemFields, "");
  const std::string out = write_pivot_cache_definition(ReadCache(input));
  EXPECT_TRUE(Has(out, " saveData=\"0\"")) << out;
  EXPECT_TRUE(Has(out, " missingItemsLimit=\"0\"")) << out;
  EXPECT_TRUE(Has(out, " refreshOnLoad=\"1\"")) << out;
  EXPECT_EQ(out.find("refreshOnLoad"), out.rfind("refreshOnLoad")) << out;
  EXPECT_FALSE(Has(out, "r:id=")) << out;
}

TEST(PivotCacheOptions, DefaultsWriteNothing) {
  const std::string out =
      write_pivot_cache_definition(ReadCache(DefinitionXml(" refreshedBy=\"x\"", kCalcItemFields, "")));
  EXPECT_FALSE(Has(out, "saveData")) << out;
  EXPECT_FALSE(Has(out, "missingItemsLimit")) << out;
  EXPECT_FALSE(Has(out, "refreshOnLoad")) << out;
  EXPECT_TRUE(Has(out, " r:id=\"rId1\"")) << out;
}

TEST(PivotCacheOptions, MissingItemsLimitMaxIsWrittenAsSet) {
  pivot::PivotCache cache = ReadCache(DefinitionXml(" refreshedBy=\"x\"", kCalcItemFields, ""));
  cache.set_missing_items_limit(32500U);
  EXPECT_TRUE(Has(write_pivot_cache_definition(cache), " missingItemsLimit=\"32500\""));
}

TEST(PivotCacheOptions, UserRefreshOnLoadIsOrredWithThePlanRule) {
  pivot::PivotCache cache = ReadCache(DefinitionXml(" refreshedBy=\"x\"", kCalcItemFields, ""));
  PivotCacheWritePlan plan;
  plan.refresh_on_load = false;
  EXPECT_FALSE(Has(write_pivot_cache_definition(cache, &plan), "refreshOnLoad"));
  cache.set_refresh_on_load(true);
  const std::string user = write_pivot_cache_definition(cache, &plan);
  EXPECT_TRUE(Has(user, " refreshOnLoad=\"1\"")) << user;
  cache.set_refresh_on_load(false);
  plan.refresh_on_load = true;
  const std::string rule = write_pivot_cache_definition(cache, &plan);
  EXPECT_TRUE(Has(rule, " refreshOnLoad=\"1\"")) << rule;
  EXPECT_EQ(rule.find("refreshOnLoad"), rule.rfind("refreshOnLoad")) << rule;
}

// Saves a one-pivot workbook whose cache has `save_data`, returning the package parts.
class PackageParts {
 public:
  explicit PackageParts(bool save_data) {
    Workbook wb = Workbook::create();
    wb.set_cell_value(0, 0, 0, Value::number(1.0));
    auto cache = std::make_unique<pivot::PivotCache>();
    cache->set_cache_id(7U);
    cache->set_save_data(save_data);
    cache->mutable_worksheet_source().present = true;
    cache->mutable_worksheet_source().ref = "A1:B2";
    cache->mutable_worksheet_source().sheet = wb.sheet(0).name();
    cache->mutable_fields().push_back(pivot::PivotCacheField{"Region", {}});
    cache->mutable_fields().push_back(pivot::PivotCacheField{"Amt", {}});
    wb.add_pivot_cache(std::move(cache));
    auto table = std::make_unique<pivot::PivotTable>();
    table->set_name("P");
    table->set_pivot_cache_id(7U);
    table->set_anchor(0U, 0U, 1U, 1U);
    wb.sheet(0).add_pivot_table(std::move(table));
    auto saved = write_ooxml(wb);
    ok_ = static_cast<bool>(saved);
    if (ok_) {
      bytes_ = std::move(saved.value());
      ok_ = static_cast<bool>(zip_.open(ByteSpan{bytes_.data(), bytes_.size()}));
    }
  }

  bool ok() const { return ok_; }
  bool Has(std::string_view part) const { return zip_.has_entry(part); }
  std::string Read(std::string_view part) const {
    auto bytes = zip_.read_entry(part);
    return bytes ? std::string(bytes.value().begin(), bytes.value().end()) : std::string();
  }

 private:
  std::vector<std::uint8_t> bytes_;
  ZipReader zip_;
  bool ok_ = false;
};

TEST(PivotCacheSaveData, SavedDataKeepsTheRecordsPart) {
  const PackageParts pkg(/*save_data=*/true);
  ASSERT_TRUE(pkg.ok());
  EXPECT_TRUE(pkg.Has("xl/pivotCache/pivotCacheRecords1.xml"));
  EXPECT_TRUE(pkg.Has("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"));
  EXPECT_TRUE(Has(pkg.Read("[Content_Types].xml"), "/xl/pivotCache/pivotCacheRecords1.xml"));
  EXPECT_TRUE(Has(pkg.Read("xl/pivotCache/pivotCacheDefinition1.xml"), " r:id=\"rId1\""));
}

TEST(PivotCacheSaveData, DroppedDataLeavesNoRecordsReference) {
  const PackageParts pkg(/*save_data=*/false);
  ASSERT_TRUE(pkg.ok());
  EXPECT_TRUE(pkg.Has("xl/pivotCache/pivotCacheDefinition1.xml"));
  EXPECT_FALSE(pkg.Has("xl/pivotCache/pivotCacheRecords1.xml"));
  EXPECT_FALSE(pkg.Has("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"));
  EXPECT_FALSE(Has(pkg.Read("[Content_Types].xml"), "pivotCacheRecords"));
  const std::string definition = pkg.Read("xl/pivotCache/pivotCacheDefinition1.xml");
  EXPECT_FALSE(Has(definition, "r:id=")) << definition;
  EXPECT_TRUE(Has(definition, " saveData=\"0\"")) << definition;
}

}  // namespace
}  // namespace formulon::io
