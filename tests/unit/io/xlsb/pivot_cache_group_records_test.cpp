#include "io/xlsb/pivot_cache_group_records.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/xlsb/pivot_area_records.h"
#include "io/xlsb/pivot_calc_item_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/zip_reader.h"
#include "pivot_twin_corpus.h"
#include "pugixml.hpp"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using pivot_records::RawRecord;

const char* const kCachePrefix = "xl/pivotCache/pivotCacheDefinition";
const char* const kTablePrefix = "xl/pivotTables/pivotTable";

/// A part split into records, with each record's byte range kept.
struct Stream {
  std::vector<std::uint8_t> bytes;
  std::vector<std::size_t> begin;
  std::vector<std::size_t> end;
  std::vector<RawRecord> raw;
};

Stream Split(std::vector<std::uint8_t> bytes) {
  Stream s;
  s.bytes = std::move(bytes);
  for (const FramedRecord& rec : split_records(s.bytes)) {
    s.begin.push_back(rec.begin);
    s.end.push_back(rec.end);
    s.raw.push_back(
        RawRecord{rec.type, std::vector<std::uint8_t>(rec.payload.data, rec.payload.data + rec.payload.size)});
  }
  return s;
}

std::vector<std::uint8_t> Frame(const std::vector<RawRecord>& records) {
  std::vector<std::uint8_t> out;
  for (const RawRecord& rec : records) {
    emit_record(out, rec.type, rec.payload);
  }
  return out;
}

std::vector<std::uint8_t> Slice(const Stream& s, std::size_t first, std::size_t last) {
  return std::vector<std::uint8_t>(s.bytes.begin() + static_cast<std::ptrdiff_t>(s.begin[first]),
                                   s.bytes.begin() + static_cast<std::ptrdiff_t>(s.end[last - 1U]));
}

std::string ErrorText(const Error& e) {
  return e.message + " " + e.context;
}

/// Decodes every block opening with `begin_type` in every cache-definition and
/// pivot-table part of the corpus and checks encode(decode(x)) == x. Returns
/// the number of blocks seen.
template <typename Decode, typename Encode>
int RoundTripEveryBlock(std::uint16_t begin_type, Decode decode, Encode encode) {
  int blocks = 0;
  for (const std::string& path : pivot_twins::CorpusPaths()) {
    for (const char* prefix : {kCachePrefix, kTablePrefix}) {
      for (auto& part : pivot_twins::Parts(path, prefix)) {
        const Stream s = Split(std::move(part.second));
        EXPECT_FALSE(s.raw.empty()) << path << " " << part.first;
        for (std::size_t i = 0; i < s.raw.size(); ++i) {
          if (s.raw[i].type != begin_type) {
            continue;
          }
          ++blocks;
          std::size_t pos = i;
          auto doc = decode(s.raw, &pos);
          if (!doc) {
            ADD_FAILURE() << path << " " << part.first << " record " << i << ": " << ErrorText(doc.error());
            continue;
          }
          std::vector<RawRecord> out;
          encode(doc.value(), &out);
          EXPECT_EQ(Frame(out), Slice(s, i, pos)) << path << " " << part.first << " record " << i;
        }
      }
    }
  }
  return blocks;
}

std::string XlsxTwin(const std::string& xlsb_path) {
  return xlsb_path.substr(0, xlsb_path.size() - 5U) + ".xlsx";
}

/// Loads `entry` of the package at `path` into `*doc`.
bool LoadXml(const std::string& path, const std::string& entry, pugi::xml_document* doc) {
  const std::vector<std::uint8_t> package = ReadFixtureFile(path);
  ZipReader zip;
  if (!zip.open(SpanOf(package))) {
    return false;
  }
  auto bytes = zip.read_entry(entry);
  if (!bytes) {
    return false;
  }
  return static_cast<bool>(doc->load_buffer(bytes.value().data(), bytes.value().size()));
}

std::string XmlEntryFor(const std::string& bin_entry) {
  return bin_entry.substr(0, bin_entry.size() - 4U) + ".xml";
}

std::string Fixture(const std::string& relative) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/" + relative;
}

/// The single part of `path` under `prefix`.
Stream OnlyPart(const std::string& path, const std::string& prefix, std::string* entry) {
  auto parts = pivot_twins::Parts(path, prefix);
  EXPECT_EQ(parts.size(), 1U) << path;
  if (parts.empty()) {
    return Stream{};
  }
  *entry = parts.front().first;
  return Split(std::move(parts.front().second));
}

std::size_t IndexOf(const Stream& s, std::uint16_t type) {
  for (std::size_t i = 0; i < s.raw.size(); ++i) {
    if (s.raw[i].type == type) {
      return i;
    }
  }
  return s.raw.size();
}

// ---- Round trips over the whole corpus -------------------------------------

TEST(XlsbPivotCacheGroupRecords, EveryFieldGroupRoundTrips) {
  const int blocks = RoundTripEveryBlock(pivot_records::kBeginPCDFGroup, pivot_group::decode_field_group,
                                         pivot_group::encode_field_group);
  EXPECT_GE(blocks, 14);
}

TEST(XlsbPivotCacheGroupRecords, EveryCalcItemBlockRoundTrips) {
  const int blocks = RoundTripEveryBlock(pivot_records::kBeginPCDCalcItems, pivot_calc::decode_calc_items,
                                         pivot_calc::encode_calc_items);
  EXPECT_GE(blocks, 1);
}

TEST(XlsbPivotCacheGroupRecords, EveryPivotAreaRoundTrips) {
  const int blocks =
      RoundTripEveryBlock(pivot_records::kBeginPRule, pivot_area::decode_pivot_area, pivot_area::encode_pivot_area);
  EXPECT_GE(blocks, 2);
}

// ---- Grouping ---------------------------------------------------------------

// Every group in the two Windows grouping books against the twin's fieldGroup.
TEST(XlsbPivotCacheGroupRecords, FieldGroupsMatchTwinXml) {
  int checked = 0;
  for (const char* book : {"win/pivot_days_end/pivot_days_end.xlsb", "win/pivot_week_1900/pivot_week_1900.xlsb"}) {
    const std::string path = Fixture(book);
    for (auto& part : pivot_twins::Parts(path, kCachePrefix)) {
      pugi::xml_document xml;
      ASSERT_TRUE(LoadXml(XlsxTwin(path), XmlEntryFor(part.first), &xml)) << part.first;
      std::vector<pugi::xml_node> fields;
      for (pugi::xml_node f : xml.child("pivotCacheDefinition").child("cacheFields").children("cacheField")) {
        fields.push_back(f);
      }
      const Stream s = Split(std::move(part.second));
      std::size_t field = 0;
      for (std::size_t i = 0; i < s.raw.size(); ++i) {
        if (s.raw[i].type == pivot_records::kEndPCDField) {
          ++field;
        }
        if (s.raw[i].type != pivot_records::kBeginPCDFGroup) {
          continue;
        }
        std::size_t pos = i;
        auto group = pivot_group::decode_field_group(s.raw, &pos);
        ASSERT_TRUE(group) << ErrorText(group.error());
        ASSERT_LT(field, fields.size());
        const pugi::xml_node fg = fields[field].child("fieldGroup");
        ASSERT_TRUE(fg) << part.first << " field " << field;
        const pugi::xml_attribute par = fg.attribute("par");
        const pugi::xml_attribute base = fg.attribute("base");
        EXPECT_EQ(group.value().par, par ? par.as_uint() : pivot_group::kNoField);
        EXPECT_EQ(group.value().base, base ? base.as_uint() : pivot_group::kNoField);

        const pugi::xml_node range = fg.child("rangePr");
        ASSERT_EQ(group.value().range.has_value(), static_cast<bool>(range));
        if (range) {
          const pivot_group::RangeGroup& r = *group.value().range;
          ASSERT_STREQ(range.attribute("groupBy").value(), "days");
          EXPECT_EQ(r.group_by, pivot_group::kGroupByDays);
          EXPECT_EQ((r.flags & pivot_group::kRangeAutoStart) != 0, range.attribute("autoStart").as_bool(true));
          EXPECT_EQ((r.flags & pivot_group::kRangeAutoEnd) != 0, range.attribute("autoEnd").as_bool(true));
          EXPECT_EQ((r.flags & pivot_group::kRangeDates) != 0, static_cast<bool>(range.attribute("startDate")));
          EXPECT_DOUBLE_EQ(r.interval, range.attribute("groupInterval").as_double(1.0));
        }

        const pugi::xml_node items = fg.child("groupItems");
        ASSERT_EQ(group.value().items.has_value(), static_cast<bool>(items));
        if (items) {
          std::vector<std::string> expected;
          for (pugi::xml_node v : items.children("s")) {
            expected.emplace_back(v.attribute("v").value());
          }
          EXPECT_EQ(*group.value().items, expected);
          EXPECT_EQ(group.value().items->size(), items.attribute("count").as_uint());
        }
        ++checked;
      }
    }
  }
  EXPECT_GE(checked, 14);
}

TEST(XlsbPivotCacheGroupRecords, DaysEndRangeFields) {
  std::string entry;
  const Stream s =
      OnlyPart(Fixture("win/pivot_days_end/pivot_days_end.xlsb"), "xl/pivotCache/pivotCacheDefinition1.bin", &entry);
  std::size_t pos = IndexOf(s, pivot_records::kBeginPCDFGRange);
  ASSERT_LT(pos, s.raw.size());
  --pos;  // the BrtBeginPCDFGroup that owns the range
  auto group = pivot_group::decode_field_group(s.raw, &pos);
  ASSERT_TRUE(group) << ErrorText(group.error());
  EXPECT_EQ(group.value().par, pivot_group::kNoField);
  EXPECT_EQ(group.value().base, 0U);
  ASSERT_TRUE(group.value().range.has_value());
  const pivot_group::RangeGroup& r = *group.value().range;
  EXPECT_EQ(r.group_by, pivot_group::kGroupByDays);
  EXPECT_EQ(r.flags, pivot_group::kRangeAutoStart | pivot_group::kRangeDates);
  EXPECT_DOUBLE_EQ(r.start, 0.0);
  EXPECT_DOUBLE_EQ(r.end, 60.0);
  EXPECT_DOUBLE_EQ(r.interval, 7.0);
  ASSERT_TRUE(group.value().items.has_value());
  ASSERT_EQ(group.value().items->size(), 11U);
  EXPECT_EQ(group.value().items->front(), "<1900/1/0");
  EXPECT_EQ(group.value().items->back(), ">1900/2/29");
}

TEST(XlsbPivotCacheGroupRecords, GroupRefusesUnknownChild) {
  std::vector<RawRecord> records;
  pivot_group::FieldGroup group;
  group.base = 0;
  pivot_group::encode_field_group(group, &records);
  records.insert(records.begin() + 1, RawRecord{pivot_records::kBeginPCDCalcItems, {0, 0, 0, 0}});
  std::size_t pos = 0;
  auto decoded = pivot_group::decode_field_group(records, &pos);
  ASSERT_FALSE(decoded);
  EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbUnknownRecord);
}

TEST(XlsbPivotCacheGroupRecords, GroupRefusesNonStringItemRun) {
  std::vector<RawRecord> records;
  pivot_group::FieldGroup group;
  group.items = std::vector<std::string>{"a"};
  pivot_group::encode_field_group(group, &records);
  records[2].payload[0] = 0x01;  // a number run
  std::size_t pos = 0;
  EXPECT_FALSE(pivot_group::decode_field_group(records, &pos));
}

// ---- Calculated items -------------------------------------------------------

TEST(XlsbPivotCacheGroupRecords, CalcItemFieldsMatchTwin) {
  const std::string path = Fixture("xlsb_pivot_twins/C13_calc_item.xlsb");
  std::string entry;
  const Stream s = OnlyPart(path, kCachePrefix, &entry);
  std::size_t pos = IndexOf(s, pivot_records::kBeginPCDCalcItems);
  ASSERT_LT(pos, s.raw.size());
  auto decoded = pivot_calc::decode_calc_items(s.raw, &pos);
  ASSERT_TRUE(decoded) << ErrorText(decoded.error());
  ASSERT_EQ(decoded.value().size(), 1U);
  const pivot_calc::CalcItem& item = decoded.value().front();

  pugi::xml_document xml;
  ASSERT_TRUE(LoadXml(XlsxTwin(path), XmlEntryFor(entry), &xml));
  const pugi::xml_node def = xml.child("pivotCacheDefinition");
  const pugi::xml_node calc = def.child("calculatedItems").child("calculatedItem");
  ASSERT_TRUE(calc);
  ASSERT_STREQ(calc.attribute("formula").value(), "Region[East]+Region[West]");
  EXPECT_FALSE(calc.attribute("field"));
  EXPECT_EQ(item.field, pivot_calc::kNoField);

  // Region[East]+Region[West] = PtgSxName 0, PtgSxName 1, PtgAdd.
  const std::vector<std::uint8_t> rgce = {0x18, 0x1d, 0, 0, 0, 0, 0x18, 0x1d, 1, 0, 0, 0, 0x03};
  EXPECT_EQ(item.rgce, rgce);
  EXPECT_TRUE(item.rgcb.empty());
  auto tokens = pivot_calc::walk_calc_item_formula(item.rgce);
  ASSERT_TRUE(tokens) << ErrorText(tokens.error());
  ASSERT_EQ(tokens.value().size(), 3U);
  EXPECT_EQ(tokens.value()[0].ptg, pivot_calc::kPtgSxName);
  EXPECT_EQ(tokens.value()[0].sx_name, 0U);
  EXPECT_EQ(tokens.value()[1].ptg, pivot_calc::kPtgSxName);
  EXPECT_EQ(tokens.value()[1].sx_name, 1U);
  EXPECT_EQ(tokens.value()[2].ptg, 0x03);

  // Each PtgSxName index names one (field, shared item) pair.
  const pugi::xml_node region = def.child("cacheFields").child("cacheField");
  ASSERT_STREQ(region.attribute("name").value(), "Region");
  std::vector<std::string> shared;
  for (pugi::xml_node s_item : region.child("sharedItems").children()) {
    shared.emplace_back(s_item.attribute("v").value());
  }
  ASSERT_TRUE(item.names.has_value());
  ASSERT_EQ(item.names->size(), 2U);
  const char* const expected_items[] = {"East", "West"};
  for (std::size_t k = 0; k < 2U; ++k) {
    const pivot_calc::PName& name = (*item.names)[k];
    EXPECT_EQ(name.field, pivot_calc::kNoField);
    ASSERT_EQ(name.pairs.size(), 1U);
    EXPECT_EQ(name.pairs[0].field, 0U);
    ASSERT_LT(name.pairs[0].item, shared.size());
    EXPECT_EQ(shared[name.pairs[0].item], expected_items[k]);
  }

  // pivotArea cacheIndex="1" outline="0" fieldPosition="0", reference field 0, x 3.
  const pugi::xml_node area_xml = calc.child("pivotArea");
  ASSERT_TRUE(item.area.has_value());
  const pivot_area::PivotArea& area = *item.area;
  EXPECT_EQ(area.field, pivot_area::kNoField);
  EXPECT_EQ(area.type, pivot_area::kAreaTypeNormal);
  EXPECT_EQ((area.flags & pivot_area::kAreaCacheIndex) != 0, area_xml.attribute("cacheIndex").as_bool(false));
  EXPECT_EQ((area.flags & pivot_area::kAreaDataOnly) != 0, area_xml.attribute("dataOnly").as_bool(true));
  const pugi::xml_node ref = area_xml.child("references").child("reference");
  ASSERT_TRUE(area.filters.has_value());
  ASSERT_EQ(area.filters->size(), 1U);
  const pivot_area::PivotAreaFilter& filter = area.filters->front();
  EXPECT_EQ(filter.field, ref.attribute("field").as_uint());
  EXPECT_EQ((filter.flags & pivot_area::kFilterSelected) != 0, ref.attribute("selected").as_bool(true));
  ASSERT_EQ(filter.items.size(), 1U);
  EXPECT_EQ(filter.items[0], ref.child("x").attribute("v").as_uint());
  EXPECT_EQ(filter.items[0], 3U);
}

TEST(XlsbPivotCacheGroupRecords, CalcItemFormulaWalkRefusesUnknownToken) {
  // PtgNum (0x1F) has no Excel-saved calculated-item sample; PtgInt does.
  auto tokens = pivot_calc::walk_calc_item_formula({0x1f, 0, 0, 0, 0, 0, 0, 0xf0, 0x3f});
  ASSERT_FALSE(tokens);
  EXPECT_EQ(tokens.error().code, FormulonErrorCode::kIoXlsbUnsupportedPtg);
  EXPECT_FALSE(pivot_calc::walk_calc_item_formula({0x18, 0x1d, 0x00}));
  EXPECT_FALSE(pivot_calc::walk_calc_item_formula({0x1e, 0x02}));
  auto integer = pivot_calc::walk_calc_item_formula({0x1e, 0x02, 0x00});
  ASSERT_TRUE(integer);
  ASSERT_EQ(integer.value().size(), 1U);
  EXPECT_EQ(integer.value()[0].integer, 2U);
}

TEST(XlsbPivotCacheGroupRecords, CalcItemRefusesCountMismatch) {
  std::vector<RawRecord> records;
  pivot_calc::encode_calc_items({pivot_calc::CalcItem{}}, &records);
  records[0].payload[0] = 2;
  std::size_t pos = 0;
  auto decoded = pivot_calc::decode_calc_items(records, &pos);
  ASSERT_FALSE(decoded);
  EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
}

// ---- Pivot area under autoSortScope ----------------------------------------

TEST(XlsbPivotCacheGroupRecords, AutoSortScopeAreaMatchesTwin) {
  const std::string path = Fixture("xlsb_pivot_twins/C07_sort_by_data.xlsb");
  std::string entry;
  const Stream s = OnlyPart(path, kTablePrefix, &entry);
  std::size_t pos = IndexOf(s, pivot_records::kBeginAutoSortScope);
  ASSERT_LT(pos, s.raw.size());
  ++pos;
  auto decoded = pivot_area::decode_pivot_area(s.raw, &pos);
  ASSERT_TRUE(decoded) << ErrorText(decoded.error());
  ASSERT_LT(pos, s.raw.size());
  EXPECT_EQ(s.raw[pos].type, pivot_records::kEndAutoSortScope);

  pugi::xml_document xml;
  ASSERT_TRUE(LoadXml(XlsxTwin(path), XmlEntryFor(entry), &xml));
  const pugi::xml_node area_xml =
      xml.child("pivotTableDefinition")
          .child("pivotFields")
          .find_child([](pugi::xml_node n) { return static_cast<bool>(n.child("autoSortScope")); })
          .child("autoSortScope")
          .child("pivotArea");
  ASSERT_TRUE(area_xml);
  const pivot_area::PivotArea& area = decoded.value();
  EXPECT_EQ(area.field, pivot_area::kNoField);
  EXPECT_EQ(area.type, pivot_area::kAreaTypeNormal);
  EXPECT_EQ((area.flags & pivot_area::kAreaDataOnly) != 0, area_xml.attribute("dataOnly").as_bool(true));
  EXPECT_EQ((area.flags & pivot_area::kAreaCacheIndex) != 0, area_xml.attribute("cacheIndex").as_bool(false));
  ASSERT_TRUE(area.filters.has_value());
  ASSERT_EQ(area.filters->size(), 1U);
  const pugi::xml_node ref = area_xml.child("references").child("reference");
  const pivot_area::PivotAreaFilter& filter = area.filters->front();
  EXPECT_EQ(filter.field, pivot_area::kDataField);
  EXPECT_EQ(filter.field, ref.attribute("field").as_uint());
  EXPECT_EQ((filter.flags & pivot_area::kFilterSelected) != 0, ref.attribute("selected").as_bool(true));
  EXPECT_EQ(filter.flags & pivot_area::kFilterSelected, 0U);
  ASSERT_EQ(filter.items.size(), 1U);
  EXPECT_EQ(filter.items[0], ref.child("x").attribute("v").as_uint());
}

TEST(XlsbPivotCacheGroupRecords, AreaRefusesNonEmptyEndRecord) {
  std::vector<RawRecord> records;
  pivot_area::encode_pivot_area(pivot_area::PivotArea{}, &records);
  records.back().payload.push_back(0);
  std::size_t pos = 0;
  EXPECT_FALSE(pivot_area::decode_pivot_area(records, &pos));
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
