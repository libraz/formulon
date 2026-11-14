#include "io/xlsb/pivot_cache_definition_records.h"

#include <cstdint>
#include <initializer_list>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/xlsb/guid_text.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/zip_reader.h"
#include "pivot_twin_corpus.h"
#include "utils/a1_ref.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

namespace pcd = pivot_cache_definition;

constexpr const char* kPartPrefix = "xl/pivotCache/pivotCacheDefinition";

std::string Fixture(const std::string& relative) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/" + relative;
}

std::vector<std::uint8_t> PartNamed(const std::string& package, const std::string& name) {
  for (auto& part : pivot_twins::Parts(package, kPartPrefix)) {
    if (part.first == name) {
      return std::move(part.second);
    }
  }
  ADD_FAILURE() << "no part " << name << " in " << package;
  return {};
}

/// The `.xlsx` twin's XML for the `.bin` part `bin_name` of `xlsb_path`.
std::string TwinXml(const std::string& xlsb_path, const std::string& bin_name) {
  const std::string xlsx_path = xlsb_path.substr(0, xlsb_path.size() - 1U) + "x";
  const std::string xml_name = bin_name.substr(0, bin_name.size() - 3U) + "xml";
  const std::vector<std::uint8_t> package = ReadFixtureFile(xlsx_path);
  ZipReader zip;
  if (!zip.open(SpanOf(package))) {
    ADD_FAILURE() << "cannot open " << xlsx_path;
    return {};
  }
  auto bytes = zip.read_entry(xml_name);
  if (!bytes) {
    ADD_FAILURE() << "no " << xml_name << " in " << xlsx_path;
    return {};
  }
  return std::string(bytes.value().begin(), bytes.value().end());
}

/// The start tag of the `index`-th `<tag ...>` element of `xml`, or empty.
std::string StartTag(const std::string& xml, const std::string& tag, std::size_t index = 0) {
  const std::string open = "<" + tag;
  std::size_t seen = 0;
  for (std::size_t pos = xml.find(open); pos != std::string::npos; pos = xml.find(open, pos + 1U)) {
    const char next = xml[pos + open.size()];
    if (next != ' ' && next != '>' && next != '/') {
      continue;
    }
    if (seen++ == index) {
      return xml.substr(pos, xml.find('>', pos) - pos + 1U);
    }
  }
  return {};
}

std::size_t CountTags(const std::string& xml, const std::string& tag) {
  std::size_t count = 0;
  while (!StartTag(xml, tag, count).empty()) {
    ++count;
  }
  return count;
}

/// The value of attribute `name` in the start tag `element`, or empty.
std::string Attr(const std::string& element, const std::string& name) {
  const std::string key = " " + name + "=\"";
  const std::size_t pos = element.find(key);
  if (pos == std::string::npos) {
    return {};
  }
  const std::size_t begin = pos + key.size();
  return element.substr(begin, element.find('"', begin) - begin);
}

std::vector<std::uint8_t> InsertAfterFirst(const std::vector<std::uint8_t>& part, std::uint16_t after,
                                           std::uint16_t type) {
  for (const FramedRecord& rec : split_records(part)) {
    if (rec.type == after) {
      std::vector<std::uint8_t> out(part.begin(), part.begin() + static_cast<std::ptrdiff_t>(rec.end));
      emit_record(out, type, std::vector<std::uint8_t>(8U, 0U));
      out.insert(out.end(), part.begin() + static_cast<std::ptrdiff_t>(rec.end), part.end());
      return out;
    }
  }
  ADD_FAILURE() << "no record " << after;
  return part;
}

TEST(XlsbPivotCacheDefinitionRecords, RoundTripsEveryCorpusPart) {
  std::size_t parts = 0;
  for (const std::string& path : pivot_twins::CorpusPaths()) {
    const auto package_parts = pivot_twins::Parts(path, kPartPrefix);
    EXPECT_FALSE(package_parts.empty()) << path;
    for (const auto& [name, bytes] : package_parts) {
      ++parts;
      auto doc = decode_pivot_cache_definition(SpanOf(bytes));
      ASSERT_TRUE(doc) << path << " " << name << ": " << doc.error().message << " " << doc.error().context;
      EXPECT_EQ(encode_pivot_cache_definition(doc.value()), bytes) << path << " " << name;
    }
  }
  EXPECT_GE(parts, 89U);
}

TEST(XlsbPivotCacheDefinitionRecords, HeaderSourceAndFieldsMatchTwinXlsx) {
  for (const std::string& path : pivot_twins::CorpusPaths()) {
    for (const auto& [name, bytes] : pivot_twins::Parts(path, kPartPrefix)) {
      SCOPED_TRACE(path + " " + name);
      auto decoded = decode_pivot_cache_definition(SpanOf(bytes));
      ASSERT_TRUE(decoded);
      const pcd::Document& doc = decoded.value();
      const std::string xml = TwinXml(path, name);
      const std::string root = StartTag(xml, "pivotCacheDefinition");
      ASSERT_FALSE(root.empty());

      ASSERT_TRUE(doc.uid.has_value());
      EXPECT_EQ(FormatGuid(doc.uid->data()), Attr(root, "xr:uid"));
      EXPECT_EQ(doc.records_rel_id, Attr(root, "r:id"));
      EXPECT_EQ(doc.refreshed_by, Attr(root, "refreshedBy"));
      EXPECT_EQ(std::to_string(doc.record_count), Attr(root, "recordCount"));
      EXPECT_EQ(std::to_string(doc.refreshed_version), Attr(root, "refreshedVersion"));
      EXPECT_EQ(std::to_string(doc.min_refreshable_version), Attr(root, "minRefreshableVersion"));
      EXPECT_EQ(std::to_string(doc.created_version), Attr(root, "createdVersion"));
      EXPECT_EQ((doc.flags & pcd::kCacheRefreshOnLoad) != 0, Attr(root, "refreshOnLoad") == "1");
      EXPECT_TRUE(Attr(root, "missingItemsLimit").empty());
      EXPECT_EQ(doc.missing_items_limit, 0xFFFFFFFFU);

      const std::string source = StartTag(xml, "worksheetSource");
      EXPECT_EQ(doc.source.sheet, Attr(source, "sheet"));
      EXPECT_EQ(a1::encode_a1(doc.source.ref.first_row, doc.source.ref.first_col) + ":" +
                    a1::encode_a1(doc.source.ref.last_row, doc.source.ref.last_col),
                Attr(source, "ref"));

      ASSERT_EQ(doc.fields.size(), CountTags(xml, "cacheField"));
      for (std::size_t i = 0; i < doc.fields.size(); ++i) {
        const std::string field = StartTag(xml, "cacheField", i);
        EXPECT_EQ(doc.fields[i].name, Attr(field, "name"));
        EXPECT_EQ(std::to_string(doc.fields[i].num_fmt_id), Attr(field, "numFmtId"));
        EXPECT_EQ((doc.fields[i].flags & pcd::kFieldDatabase) != 0, Attr(field, "databaseField") != "0");
      }

      EXPECT_EQ(doc.has_pcd14, xml.find("<x14:pivotCacheDefinition/>") != std::string::npos);
      EXPECT_EQ(doc.data_source_clean.has_value(), xml.find("<xlpdsc:dataSourceClean>1<") != std::string::npos);
      if (doc.data_source_clean.has_value()) {
        EXPECT_EQ(*doc.data_source_clean, 1U);
      }
      EXPECT_EQ(doc.calculated_items.empty(), xml.find("<calculatedItems") == std::string::npos);
    }
  }
}

TEST(XlsbPivotCacheDefinitionRecords, A5C9SumNumberMissingAndUnusedItems) {
  const auto bytes = PartNamed(Fixture("xlsb_pivot_twins/A5_C9__sum.xlsb"), "xl/pivotCache/pivotCacheDefinition1.bin");
  auto decoded = decode_pivot_cache_definition(SpanOf(bytes));
  ASSERT_TRUE(decoded);
  const pcd::Document& doc = decoded.value();
  ASSERT_EQ(doc.fields.size(), 5U);

  // Region: <s v="N"/><s v="S"/><s v="North" u="1"/><m u="1"/><s v="South" u="1"/>, containsBlank.
  const pcd::Field& region = doc.fields[0];
  ASSERT_TRUE(region.items.has_value());
  EXPECT_NE(region.items->flags & pcd::kAtblBlank, 0);
  const auto& items = region.items->entries;
  ASSERT_EQ(items.size(), 5U);
  EXPECT_EQ(items[0].item.type, pcd::ItemType::kString);
  EXPECT_EQ(items[0].item.text, "N");
  EXPECT_EQ(items[1].item.text, "S");
  EXPECT_EQ(items[2].item.type, pcd::ItemType::kAttributedString);
  EXPECT_EQ(items[2].item.text, "North");
  EXPECT_EQ(items[2].item.attributes, pcd::kItemUnused);
  EXPECT_EQ(items[3].item.type, pcd::ItemType::kAttributedMissing);
  EXPECT_EQ(items[3].item.attributes, pcd::kItemUnused);
  EXPECT_EQ(items[4].item.type, pcd::ItemType::kAttributedString);
  EXPECT_EQ(items[4].item.text, "South");
  EXPECT_EQ(region.items->item_count(), 5U);

  // Q: <n v="1"/><n v="2"/>, minValue 1, maxValue 2, containsInteger.
  const pcd::Field& q = doc.fields[3];
  EXPECT_EQ(q.name, "Q");
  ASSERT_TRUE(q.items.has_value());
  EXPECT_NE(q.items->flags & pcd::kAtblMinMax, 0);
  EXPECT_NE(q.items->flags & pcd::kAtblInteger, 0);
  EXPECT_EQ(q.items->min, 1.0);
  EXPECT_EQ(q.items->max, 2.0);
  ASSERT_EQ(q.items->entries.size(), 2U);
  EXPECT_FALSE(q.items->entries[0].is_run);
  EXPECT_EQ(q.items->entries[0].item.type, pcd::ItemType::kNumber);
  EXPECT_EQ(q.items->entries[0].item.number, 1.0);
  EXPECT_EQ(q.items->entries[1].item.number, 2.0);

  // Amt: minValue -0.4, maxValue 9007199254740992, no items, not integer.
  const pcd::Field& amt = doc.fields[4];
  ASSERT_TRUE(amt.items.has_value());
  EXPECT_TRUE(amt.items->entries.empty());
  EXPECT_EQ(amt.items->flags & pcd::kAtblInteger, 0);
  EXPECT_NE(amt.items->flags & pcd::kAtblNumber, 0);
  EXPECT_EQ(amt.items->min, -0.4);
  EXPECT_EQ(amt.items->max, 9007199254740992.0);
}

TEST(XlsbPivotCacheDefinitionRecords, C13CalculatedItemMarkerAndBlock) {
  const auto bytes =
      PartNamed(Fixture("xlsb_pivot_twins/C13_calc_item.xlsb"), "xl/pivotCache/pivotCacheDefinition1.bin");
  auto decoded = decode_pivot_cache_definition(SpanOf(bytes));
  ASSERT_TRUE(decoded);
  const pcd::Document& doc = decoded.value();
  ASSERT_EQ(doc.fields.size(), 5U);

  // Region: East, West, North, then the calculated item <s v="EW" f="1"/>.
  ASSERT_TRUE(doc.fields[0].items.has_value());
  const auto& items = doc.fields[0].items->entries;
  ASSERT_EQ(items.size(), 4U);
  EXPECT_EQ(items[3].item.type, pcd::ItemType::kAttributedString);
  EXPECT_EQ(items[3].item.text, "EW");
  EXPECT_EQ(items[3].item.attributes, pcd::kItemFormula);

  // Date: minDate 2024-01-05 = serial 45296.
  EXPECT_EQ(doc.fields[2].num_fmt_id, 14U);
  ASSERT_TRUE(doc.fields[2].items.has_value());
  EXPECT_EQ(doc.fields[2].items->min, 45296.0);

  ASSERT_EQ(doc.calculated_items.size(), 26U);
  EXPECT_EQ(doc.calculated_items.front().type, pivot_records::kBeginPCDCalcItems);
  EXPECT_EQ(doc.calculated_items.front().payload, (std::vector<std::uint8_t>{1, 0, 0, 0}));
  EXPECT_EQ(doc.calculated_items[1].type, pivot_records::kBeginPCDCalcItem);
  EXPECT_EQ(doc.calculated_items.back().type, pivot_records::kEndPCDCalcItems);
}

TEST(XlsbPivotCacheDefinitionRecords, DaysEndDateRunAndGroupingBlocks) {
  const auto bytes =
      PartNamed(Fixture("win/pivot_days_end/pivot_days_end.xlsb"), "xl/pivotCache/pivotCacheDefinition1.bin");
  auto decoded = decode_pivot_cache_definition(SpanOf(bytes));
  ASSERT_TRUE(decoded);
  const pcd::Document& doc = decoded.value();
  EXPECT_EQ(doc.source.sheet, "data");
  EXPECT_EQ(doc.record_count, 71U);
  EXPECT_TRUE(doc.has_pcd14);
  EXPECT_FALSE(doc.data_source_clean.has_value());
  ASSERT_EQ(doc.fields.size(), 3U);

  // d: 71 <d> items 1899-12-30 .. 1900-03-10 in one date run, minDate/maxDate as serials 0 and 71.
  const pcd::Field& d = doc.fields[0];
  EXPECT_EQ(d.name, "d");
  EXPECT_EQ(d.num_fmt_id, 14U);
  ASSERT_TRUE(d.items.has_value());
  EXPECT_EQ(d.items->flags & pcd::kAtblNonDate, 0);
  EXPECT_NE(d.items->flags & pcd::kAtblDate, 0);
  EXPECT_EQ(d.items->min, 0.0);
  EXPECT_EQ(d.items->max, 71.0);
  ASSERT_EQ(d.items->entries.size(), 1U);
  ASSERT_TRUE(d.items->entries[0].is_run);
  const pcd::ItemRun& run = d.items->entries[0].run;
  EXPECT_EQ(run.type, pcd::RunType::kDates);
  ASSERT_EQ(run.dates.size(), 71U);
  EXPECT_EQ(run.dates[0].year, 1900U);
  EXPECT_EQ(run.dates[0].month, 1U);
  EXPECT_EQ(run.dates[0].day, 0U);
  EXPECT_EQ(run.dates[1].day, 1U);
  EXPECT_EQ(run.dates.back().month, 3U);
  EXPECT_EQ(run.dates.back().day, 10U);
  EXPECT_EQ(run.dates.back().hour, 0U);

  // <fieldGroup par="2"/> after the item table.
  ASSERT_EQ(d.grouping.size(), 2U);
  EXPECT_EQ(d.grouping[0].type, pivot_records::kBeginPCDFGroup);
  EXPECT_EQ(d.grouping[0].payload, (std::vector<std::uint8_t>{2, 0, 0, 0, 0xFF, 0xFF, 0xFF, 0xFF}));
  EXPECT_EQ(d.grouping[1].type, pivot_records::kEndPCDFGroup);

  // The grouped field (databaseField="0") has no item table, only its fieldGroup block.
  const pcd::Field& grouped = doc.fields[2];
  EXPECT_EQ(grouped.name, "\xE6\x97\xA5 (d)");
  EXPECT_EQ(grouped.flags, 0U);
  EXPECT_FALSE(grouped.items.has_value());
  ASSERT_FALSE(grouped.grouping.empty());
  EXPECT_EQ(grouped.grouping.front().type, pivot_records::kBeginPCDFGroup);
  EXPECT_EQ(grouped.grouping[1].type, pivot_records::kBeginPCDFGRange);
  EXPECT_EQ(grouped.grouping.back().type, pivot_records::kEndPCDFGroup);
}

TEST(XlsbPivotCacheDefinitionRecords, Week1900DateRun) {
  const auto bytes =
      PartNamed(Fixture("win/pivot_week_1900/pivot_week_1900.xlsb"), "xl/pivotCache/pivotCacheDefinition1.bin");
  auto decoded = decode_pivot_cache_definition(SpanOf(bytes));
  ASSERT_TRUE(decoded);
  const pcd::Document& doc = decoded.value();
  ASSERT_EQ(doc.fields.size(), 3U);
  ASSERT_TRUE(doc.fields[0].items.has_value());
  EXPECT_EQ(doc.fields[0].items->item_count(), 71U);
  ASSERT_TRUE(doc.fields[0].items->entries[0].is_run);
  EXPECT_EQ(doc.fields[0].items->entries[0].run.type, pcd::RunType::kDates);
  EXPECT_EQ(doc.fields[0].items->entries[0].run.dates.back().day, 10U);
}

TEST(XlsbPivotCacheDefinitionRecords, LoneItemRecordsOfTheWrongSizeAreCorrupt) {
  const auto base = PartNamed(Fixture("xlsb_pivot_twins/B01_base.xlsb"), "xl/pivotCache/pivotCacheDefinition1.bin");
  ASSERT_TRUE(decode_pivot_cache_definition(SpanOf(base)));
  // Each inserted record carries 8 zero bytes: too long for 20, 22 and 23, and one item too many for the count.
  for (const std::uint16_t lone : std::initializer_list<std::uint16_t>{20, 22, 23, 25}) {
    const auto bytes = InsertAfterFirst(base, pivot_records::kBeginPCDFAtbl, lone);
    auto decoded = decode_pivot_cache_definition(SpanOf(bytes));
    ASSERT_FALSE(decoded) << lone;
    EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt) << lone;
  }
}

TEST(XlsbPivotCacheDefinitionRecords, LoneItemRecordsRoundTrip) {
  for (const char* stem : {"item_bare_m", "item_mixed_d", "item_b_e"}) {
    for (const char* save : {".xlsb", ".2.xlsb"}) {
      const std::string package = Fixture(std::string("xlsb_parity/pivot/") + stem + save);
      const auto bytes = PartNamed(package, "xl/pivotCache/pivotCacheDefinition1.bin");
      auto decoded = decode_pivot_cache_definition(SpanOf(bytes));
      ASSERT_TRUE(decoded) << package << ": " << decoded.error().message;
      EXPECT_EQ(encode_pivot_cache_definition(decoded.value()), bytes) << package;
    }
  }
}

TEST(XlsbPivotCacheDefinitionRecords, UnknownRecordInsideAFieldIsRefused) {
  const auto base = PartNamed(Fixture("xlsb_pivot_twins/B01_base.xlsb"), "xl/pivotCache/pivotCacheDefinition1.bin");
  for (const std::uint16_t after : {pivot_records::kBeginPCDFAtbl, pivot_records::kEndPCDFAtbl}) {
    auto decoded = decode_pivot_cache_definition(SpanOf(InsertAfterFirst(base, after, 999)));
    ASSERT_FALSE(decoded) << after;
    EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbUnknownRecord) << after;
    EXPECT_NE(decoded.error().message.rfind("unmeasured", 0), 0U) << decoded.error().message;
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
