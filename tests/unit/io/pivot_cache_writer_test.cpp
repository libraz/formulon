// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for `formulon::io::write_pivot_cache_definition` and
// `formulon::io::write_pivot_cache_records`. Each test builds a hand-
// rolled `pivot::PivotCache`, writes the two parts, and (where useful)
// pipes the bytes back through the symmetric reader to assert a clean
// round-trip on the data the readers actually consume.

#include "io/pivot_cache_writer.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/pivot_cache_reader.h"
#include "pivot/pivot_cache.h"
#include "value.h"

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
  EXPECT_TRUE(parsed.records()[0].cells[0].is_text());
  EXPECT_EQ(parsed.records()[0].cells[0].as_text(), "North");
  EXPECT_TRUE(parsed.records()[0].cells[1].is_number());
  EXPECT_DOUBLE_EQ(parsed.records()[0].cells[1].as_number(), 100.0);

  EXPECT_EQ(parsed.records()[1].cells[0].as_text(), "South");
  EXPECT_DOUBLE_EQ(parsed.records()[1].cells[1].as_number(), 200.0);

  EXPECT_EQ(parsed.records()[2].cells[0].as_text(), "North");
  EXPECT_DOUBLE_EQ(parsed.records()[2].cells[1].as_number(), 300.0);
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

}  // namespace
}  // namespace formulon::io
