// AutoFilter forms Excel 16.115 saves for the blank / non-blank custom criteria, the calendar types and the icon
// filter with no icon
// (tests/fixtures/excel/afilter), read into the model and written back.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "gtest/gtest.h"
#include "io/auto_filter_xml.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/auto_filter_records.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "workbook.h"
#include "writer_test_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using Bytes = std::vector<std::uint8_t>;
using Records = std::vector<std::pair<std::uint16_t, Bytes>>;

Bytes ReadFile(const std::string& name) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/afilter/" + name;
  Bytes out;
  FILE* f = std::fopen(path.c_str(), "rb");
  if (f == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(f, 0, SEEK_END);
  const long size = std::ftell(f);
  std::fseek(f, 0, SEEK_SET);
  out.resize(size > 0 ? static_cast<std::size_t>(size) : 0U);
  if (std::fread(out.data(), 1, out.size(), f) != out.size()) {
    ADD_FAILURE() << "short read on fixture: " << path;
  }
  std::fclose(f);
  return out;
}

/// Zeroes the eight bytes after the operator of a blank / non-blank BrtCustomFilter, which hold no value.
void MaskUnionTail(Records& recs) {
  for (auto& rec : recs) {
    if (rec.first == kBrtCustomFilter && rec.second.size() >= 10U && (rec.second[0] == 0x0C || rec.second[0] == 0x0E)) {
      for (std::size_t i = 2; i < 10U; ++i) {
        rec.second[i] = 0;
      }
    }
  }
}

Records AsRecords(const Bytes& stream) {
  Records out;
  for (const FramedRecord& rec : split_records(stream)) {
    out.emplace_back(rec.type, Bytes(rec.payload.data, rec.payload.data + rec.payload.size));
  }
  MaskUnionTail(out);
  return out;
}

/// The first worksheet's records from BrtBeginAFilter through its BrtEndAFilter.
Records ExcelBlock(const std::string& name) {
  const Bytes package = ReadFile(name + ".xlsb");
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto part = zip.read_entry("xl/worksheets/sheet1.bin");
  EXPECT_TRUE(static_cast<bool>(part));
  Records out;
  if (!part) {
    return out;
  }
  bool inside = false;
  for (auto& rec : AsRecords(part.value())) {
    inside = inside || rec.first == kBeginAFilter;
    if (inside) {
      out.push_back(std::move(rec));
      if (out.back().first == kBrtEndAFilter) {
        break;
      }
    }
  }
  return out;
}

Bytes Stream(const Records& recs) {
  Bytes stream;
  for (const auto& rec : recs) {
    emit_record(stream, rec.first, ByteSpan{rec.second.data(), rec.second.size()});
  }
  return stream;
}

AutoFilter ReadBlock(const Bytes& stream) {
  ByteSpan cursor{stream.data(), stream.size()};
  auto begin = read_record(cursor);
  EXPECT_TRUE(static_cast<bool>(begin));
  auto filter = read_auto_filter_block(cursor, begin.value().payload);
  EXPECT_TRUE(static_cast<bool>(filter)) << filter.error().message;
  return filter ? filter.value() : AutoFilter{};
}

AutoFilter ExcelFilter(const std::string& name) {
  return ReadBlock(Stream(ExcelBlock(name)));
}

AutoFilter SourceFilter(const std::string& name) {
  auto loaded = read_ooxml(SpanOf(ReadFile(name + "_src.xlsx")));
  EXPECT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  if (!loaded || loaded.value().workbook.sheet(0).auto_filter() == nullptr) {
    ADD_FAILURE() << "no sheet AutoFilter in " << name;
    return AutoFilter{};
  }
  return *loaded.value().workbook.sheet(0).auto_filter();
}

Records Written(const AutoFilter& filter, std::uint32_t expect_lost = 0U) {
  Bytes out;
  EXPECT_EQ(emit_auto_filter_block(out, filter), expect_lost);
  return AsRecords(out);
}

const Bytes* FindRecord(const Records& recs, std::uint16_t type, std::size_t nth = 0) {
  for (const auto& rec : recs) {
    if (rec.first == type && nth-- == 0U) {
      return &rec.second;
    }
  }
  return nullptr;
}

std::uint32_t U32At(const Bytes& b, std::size_t at) {
  return static_cast<std::uint32_t>(b[at]) | static_cast<std::uint32_t>(b[at + 1]) << 8U |
         static_cast<std::uint32_t>(b[at + 2]) << 16U | static_cast<std::uint32_t>(b[at + 3]) << 24U;
}

TEST(XlsbAutoFilterForms, CalendarTypeReadsAndWritesAsMeasured) {
  const struct {
    const char* file;
    const char* name;
    std::uint32_t code;
  } cases[] = {{"cal_none", "", 0U},
               {"cal_gregorian", "gregorian", 1U},
               {"cal_japan", "japan", 3U},
               {"cal_hijri", "hijri", 6U},
               {"cal_thai", "thai", 7U}};
  for (const auto& c : cases) {
    SCOPED_TRACE(c.file);
    const Records excel = ExcelBlock(c.file);
    const Bytes* begin = FindRecord(excel, kBrtBeginFilters);
    ASSERT_NE(begin, nullptr);
    EXPECT_EQ(U32At(*begin, 4), c.code);
    const AutoFilter filter = ExcelFilter(c.file);
    ASSERT_EQ(filter.columns.size(), 1U);
    EXPECT_EQ(filter.columns[0].values.calendar_type, c.name);
    EXPECT_EQ(Written(filter), excel);
  }
}

TEST(XlsbAutoFilterForms, EverySchemaCalendarTypeRoundTripsInSchemaOrder) {
  const char* const names[] = {"none",
                               "gregorian",
                               "gregorianUs",
                               "japan",
                               "taiwan",
                               "korea",
                               "hijri",
                               "thai",
                               "hebrew",
                               "gregorianMeFrench",
                               "gregorianArabic",
                               "gregorianXlitEnglish",
                               "gregorianXlitFrench"};
  AutoFilter filter = ExcelFilter("cal_none");
  for (std::uint32_t code = 0; code < sizeof(names) / sizeof(names[0]); ++code) {
    SCOPED_TRACE(names[code]);
    filter.columns[0].values.calendar_type = names[code];
    const Records written = Written(filter);
    const Bytes* begin = FindRecord(written, kBrtBeginFilters);
    ASSERT_NE(begin, nullptr);
    EXPECT_EQ(U32At(*begin, 4), code);
    const AutoFilter back = ReadBlock(Stream(written));
    EXPECT_EQ(back.columns[0].values.calendar_type, code == 0U ? "" : names[code]);
  }
  filter.columns[0].values.calendar_type = "lunar";
  EXPECT_EQ(Written(filter, 1U).size(), ExcelBlock("cal_none").size());
}

TEST(XlsbAutoFilterForms, LoneEqualBlankSavesAsTheBlankValueFilter) {
  const Records excel = ExcelBlock("blank_equal");
  const Bytes* begin = FindRecord(excel, kBrtBeginFilters);
  ASSERT_NE(begin, nullptr);
  EXPECT_EQ(U32At(*begin, 0), 1U);
  EXPECT_EQ(FindRecord(excel, kBrtBeginCustomFilters), nullptr);
  const AutoFilter source = SourceFilter("blank_equal");
  ASSERT_EQ(source.columns[0].kind, FilterKind::kCustom);
  EXPECT_EQ(Written(source), excel);
  const AutoFilter read = ExcelFilter("blank_equal");
  EXPECT_EQ(read.columns[0].kind, FilterKind::kValues);
  EXPECT_TRUE(read.columns[0].values.blank);
  EXPECT_EQ(Written(read), excel);
}

TEST(XlsbAutoFilterForms, NonBlankIsVts0EOperatorNotEqual) {
  const Records excel = ExcelBlock("non_blank");
  const Bytes* custom = FindRecord(excel, kBrtCustomFilter);
  ASSERT_NE(custom, nullptr);
  EXPECT_EQ((*custom)[0], 0x0E);
  EXPECT_EQ((*custom)[1], 5);
  EXPECT_EQ(Written(SourceFilter("non_blank")), excel);
  const AutoFilter read = ExcelFilter("non_blank");
  ASSERT_EQ(read.columns[0].custom.filters.size(), 1U);
  EXPECT_EQ(read.columns[0].custom.filters[0].op, FilterOperator::kNotEqual);
  EXPECT_EQ(read.columns[0].custom.filters[0].val, " ");
  EXPECT_EQ(Written(read), excel);
}

TEST(XlsbAutoFilterForms, BlankAsSecondOfAnAndPairStaysCustom) {
  const Records excel = ExcelBlock("and_number_blank");
  const Bytes* begin = FindRecord(excel, kBrtBeginCustomFilters);
  ASSERT_NE(begin, nullptr);
  EXPECT_EQ(U32At(*begin, 0), 1U);
  const Bytes* number = FindRecord(excel, kBrtCustomFilter, 0);
  const Bytes* blank = FindRecord(excel, kBrtCustomFilter, 1);
  ASSERT_NE(number, nullptr);
  ASSERT_NE(blank, nullptr);
  EXPECT_EQ((*number)[0], 0x04);
  EXPECT_EQ((*number)[1], 4);
  EXPECT_EQ((*blank)[0], 0x0C);
  EXPECT_EQ((*blank)[1], 2);
  EXPECT_EQ(Written(SourceFilter("and_number_blank")), excel);
  const AutoFilter read = ExcelFilter("and_number_blank");
  const CustomFilters& custom = read.columns[0].custom;
  EXPECT_TRUE(custom.and_join);
  ASSERT_EQ(custom.filters.size(), 2U);
  EXPECT_EQ(custom.filters[0].op, FilterOperator::kGreaterThan);
  EXPECT_EQ(custom.filters[0].val, "50");
  EXPECT_EQ(custom.filters[1].op, FilterOperator::kEqual);
  EXPECT_EQ(custom.filters[1].val, "");
  EXPECT_EQ(Written(read), excel);
}

TEST(XlsbAutoFilterForms, BlankAndNonBlankTailIsZeroOnWriteAndIgnoredOnRead) {
  Bytes out;
  emit_auto_filter_block(out, SourceFilter("and_number_blank"));
  Records raw;
  for (const FramedRecord& rec : split_records(out)) {
    raw.emplace_back(rec.type, Bytes(rec.payload.data, rec.payload.data + rec.payload.size));
  }
  const Bytes* blank = FindRecord(raw, kBrtCustomFilter, 1);
  ASSERT_NE(blank, nullptr);
  ASSERT_EQ(blank->size(), 10U);
  for (std::size_t i = 2; i < 10U; ++i) {
    EXPECT_EQ((*blank)[i], 0) << i;
  }
  // Excel's own non-zero tail reads to the same model as the zeroed one.
  Records unmasked;
  const Bytes package = ReadFile("and_number_blank.xlsb");
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto part = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(part));
  bool inside = false;
  for (const FramedRecord& rec : split_records(part.value())) {
    inside = inside || rec.type == kBeginAFilter;
    if (inside) {
      unmasked.emplace_back(rec.type, Bytes(rec.payload.data, rec.payload.data + rec.payload.size));
      if (rec.type == kBrtEndAFilter) {
        break;
      }
    }
  }
  const AutoFilter excel_read = ReadBlock(Stream(unmasked));
  const AutoFilter zero_read = ReadBlock(out);
  EXPECT_EQ(excel_read.columns[0].custom.filters[1].val, zero_read.columns[0].custom.filters[1].val);
  EXPECT_EQ(excel_read.columns[0].custom.filters[1].op, zero_read.columns[0].custom.filters[1].op);
}

TEST(XlsbAutoFilterForms, ThreeCustomCriteriaSaveAsTheLastValue) {
  const Records excel = ExcelBlock("three_custom");
  const AutoFilter read = ExcelFilter("three_custom");
  ASSERT_EQ(read.columns[0].kind, FilterKind::kValues);
  ASSERT_EQ(read.columns[0].values.values.size(), 1U);
  EXPECT_EQ(read.columns[0].values.values[0], "70");
  EXPECT_EQ(Written(read), excel);
}

/// The AUTOFILTER block of `<name>.xlsb`'s first sheet.
Records FlagBlock(const std::string& name) {
  const Bytes package = ReadFile(name + ".xlsb");
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto part = zip.read_entry("xl/worksheets/sheet1.bin");
  Records block;
  if (!part) {
    ADD_FAILURE() << name;
    return block;
  }
  for (auto& rec : AsRecords(part.value())) {
    if (rec.first == kBeginAFilter || !block.empty()) {
      block.push_back(std::move(rec));
      if (block.back().first == kBrtEndAFilter) {
        break;
      }
    }
  }
  return block;
}

// A string criterion's union starts with 1 under an ordering operator on an operand without `*` or `?` and 0 for a
// pattern, an escaped `~*` or `~?` included; a lone `~` does not clear it. The rest is unset in Excel's save and zero
// in ours. Each Excel save reads and writes back to its record outside those unset bytes.
TEST(XlsbAutoFilterForms, StringCriterionFlagFollowsExcel) {
  const struct {
    const char* name;
    FilterOperator op;
    const char* operand;
    std::uint8_t flag;
  } kCases[] = {
      {"A_gt_B", FilterOperator::kGreaterThan, "B", 1},
      {"A_gt_Bq", FilterOperator::kGreaterThan, "B?", 0},
      {"A_gt_Bescq", FilterOperator::kGreaterThan, "B~?", 0},
      {"A_gt_Bescstar", FilterOperator::kGreaterThan, "B~*", 0},
      {"A_gt_BstarC", FilterOperator::kGreaterThan, "B*C", 0},
      {"A_gt_esctildeB", FilterOperator::kGreaterThan, "~~B", 1},
      {"A_gt_star", FilterOperator::kGreaterThan, "*", 0},
      {"A_le_B", FilterOperator::kLessThanOrEqual, "B", 1},
      {"A_le_Bq", FilterOperator::kLessThanOrEqual, "B?", 0},
      {"A_le_Bescq", FilterOperator::kLessThanOrEqual, "B~?", 0},
      {"A_le_Bescstar", FilterOperator::kLessThanOrEqual, "B~*", 0},
      {"A_le_BstarC", FilterOperator::kLessThanOrEqual, "B*C", 0},
      {"A_le_esctildeB", FilterOperator::kLessThanOrEqual, "~~B", 1},
      {"A_le_star", FilterOperator::kLessThanOrEqual, "*", 0},
  };
  for (const auto& c : kCases) {
    SCOPED_TRACE(c.name);
    const Records block = FlagBlock(std::string("flag/") + c.name);
    const Bytes* excel = FindRecord(block, kBrtCustomFilter);
    ASSERT_NE(excel, nullptr);
    ASSERT_GE(excel->size(), 10U);
    EXPECT_EQ((*excel)[2], c.flag);
    const AutoFilter filter = ReadBlock(Stream(block));
    ASSERT_EQ(filter.columns[0].custom.filters.size(), 1U);
    EXPECT_EQ(filter.columns[0].custom.filters[0].op, c.op);
    EXPECT_EQ(filter.columns[0].custom.filters[0].val, c.operand);
    const Records written = Written(filter);
    const Bytes* ours = FindRecord(written, kBrtCustomFilter);
    ASSERT_NE(ours, nullptr);
    ASSERT_EQ(ours->size(), excel->size());
    EXPECT_EQ(Bytes(ours->begin(), ours->begin() + 4), Bytes(excel->begin(), excel->begin() + 4));
    EXPECT_EQ(Bytes(ours->begin() + 10, ours->end()), Bytes(excel->begin() + 10, excel->end()));
    for (std::size_t i = 4; i < 10U; ++i) {
      EXPECT_EQ((*ours)[i], 0U) << i;
    }
  }
}

// A string criterion's union starts with its stored literal-comparison bit under every operator.
TEST(XlsbAutoFilterForms, StringCriterionFlagIsTheStoredBit) {
  Records block = FlagBlock("flag/A_gt_B");
  AutoFilter filter = ReadBlock(Stream(block));
  ASSERT_EQ(filter.columns[0].custom.filters.size(), 1U);
  for (const bool literal : {false, true}) {
    for (const FilterOperator op :
         {FilterOperator::kLessThan, FilterOperator::kEqual, FilterOperator::kLessThanOrEqual,
          FilterOperator::kGreaterThan, FilterOperator::kNotEqual, FilterOperator::kGreaterThanOrEqual}) {
      filter.columns[0].custom.filters[0].op = op;
      filter.columns[0].custom.filters[0].literal = literal;
      const Records written = Written(filter);
      const Bytes* rec = FindRecord(written, kBrtCustomFilter);
      ASSERT_NE(rec, nullptr);
      EXPECT_EQ((*rec)[2], literal ? 1U : 0U) << static_cast<int>(op);
      EXPECT_EQ(ReadBlock(Stream(written)).columns[0].custom.filters[0].literal, literal) << static_cast<int>(op);
    }
  }
}

// Excel's xlsb save of an xlsx it loaded (afilter/carry): the flag is the state the xlsx load gave the criterion, so
// <>~B saves clear where Excel's own entry of it sets the flag, and ordering criteria keep it set.
TEST(XlsbAutoFilterForms, XlsxLoadedCriterionSavesExcelsFlag) {
  for (const char* name : {"A_ne_tildeB", "A_lt_tildeB", "A_gt_B"}) {
    SCOPED_TRACE(name);
    const std::string stem = std::string("carry/") + name;
    const Records block = FlagBlock(stem + "_asloaded");
    const Bytes* excel = FindRecord(block, kBrtCustomFilter);
    ASSERT_NE(excel, nullptr);
    ASSERT_GE(excel->size(), 10U);
    auto loaded = read_ooxml(SpanOf(ReadFile(stem + ".xlsx")));
    ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
    const AutoFilter* filter = loaded.value().workbook.sheet(0).auto_filter();
    ASSERT_NE(filter, nullptr);
    const Records written = Written(*filter);
    const Bytes* ours = FindRecord(written, kBrtCustomFilter);
    ASSERT_NE(ours, nullptr);
    ASSERT_EQ(ours->size(), excel->size());
    EXPECT_EQ(Bytes(ours->begin(), ours->begin() + 4), Bytes(excel->begin(), excel->begin() + 4));
    EXPECT_EQ(Bytes(ours->begin() + 10, ours->end()), Bytes(excel->begin() + 10, excel->end()));
  }
}

// = and <> on a sheet (afilter/equal): an = against B, B~*, B~? or ~B is a value list with ~* and ~? unescaped, and
// = B* a custom filter with a clear flag; <> is a custom filter whose flag is set unless the operand has `*` or `?`.
// Each block reads and writes back to Excel's records outside the unset union bytes.
TEST(XlsbAutoFilterForms, EqualityCriteriaWriteBackExcelsBlock) {
  for (const char* name : {"A_eq_B", "A_eq_Bstar", "A_eq_Bescstar", "A_eq_Bescq", "A_eq_tildeB", "A_eq_esctildeB",
                           "A_ne_B", "A_ne_Bstar", "A_ne_Bescstar", "A_ne_Bescq", "A_ne_tildeB", "A_ne_esctildeB"}) {
    SCOPED_TRACE(name);
    const Records excel = FlagBlock(std::string("equal/") + name);
    ASSERT_FALSE(excel.empty());
    const Records ours = Written(ReadBlock(Stream(excel)));
    ASSERT_EQ(ours.size(), excel.size());
    for (std::size_t i = 0; i < excel.size(); ++i) {
      ASSERT_EQ(ours[i].first, excel[i].first) << i;
      Bytes a = ours[i].second;
      Bytes b = excel[i].second;
      if (b.size() >= 10U && excel[i].first == kBrtCustomFilter && b[0] == 0x06) {
        for (std::size_t k = 4; k < 10U && k < a.size(); ++k) {
          a[k] = 0;
          b[k] = 0;
        }
      }
      EXPECT_EQ(a, b) << "record " << i << " type " << excel[i].first;
    }
  }
}

// An icon filter with no icon keeps no set: Excel saves `<iconFilter iconSet="3Arrows"/>` (for any authored set) and
// xlsb nil set and icon; an icon of a set is kept as authored.
TEST(XlsbAutoFilterForms, IconFilterWithNoIconHasNoSet) {
  const struct {
    const char* name;
    std::uint32_t set;
    std::uint32_t icon;
  } kCases[] = {
      {"icon_no_id", 0xFFFFFFFFU, 0xFFFFFFFFU},
      {"icon_id0", 0U, 0U},
      {"icon_rating4_no_id", 0xFFFFFFFFU, 0xFFFFFFFFU},
  };
  for (const auto& c : kCases) {
    SCOPED_TRACE(c.name);
    const Records excel = ExcelBlock(c.name);
    const Bytes* stored = FindRecord(excel, kBrtIconFilter);
    ASSERT_NE(stored, nullptr);
    ASSERT_EQ(stored->size(), 8U);
    EXPECT_EQ(U32At(*stored, 0), c.set);
    EXPECT_EQ(U32At(*stored, 4), c.icon);
    EXPECT_EQ(Written(ExcelFilter(c.name)), excel);
    const AutoFilter source = SourceFilter(c.name);
    EXPECT_EQ(Written(source), excel);
    ASSERT_EQ(source.columns.size(), 1U);
    EXPECT_EQ(source.columns[0].icon.icon_set, cf::IconSetName::Three_Arrows);
    const std::string xml = auto_filter_xml(&source);
    EXPECT_NE(
        xml.find(c.icon == 0U ? "<iconFilter iconSet=\"3Arrows\" iconId=\"0\"/>" : "<iconFilter iconSet=\"3Arrows\"/>"),
        std::string::npos)
        << xml;
  }
}

// A set with a nil icon (Excel opens one and keeps it on an xlsb resave) reads as the filter with no icon; a nil set
// with an icon is refused.
TEST(XlsbAutoFilterForms, IconFilterSetWithANilIconReadsAsNoIcon) {
  const AutoFilter filter = ExcelFilter("icon_set0_nil");
  ASSERT_EQ(filter.columns.size(), 1U);
  EXPECT_EQ(filter.columns[0].kind, FilterKind::kIcon);
  EXPECT_EQ(filter.columns[0].icon.icon_set, cf::IconSetName::Three_Arrows);
  EXPECT_FALSE(filter.columns[0].icon.icon_id.has_value());

  Records block = ExcelBlock("icon_set0_nil");
  for (auto& rec : block) {
    if (rec.first == kBrtIconFilter) {
      rec.second = {0xFF, 0xFF, 0xFF, 0xFF, 0x01, 0x00, 0x00, 0x00};
    }
  }
  const Bytes stream = Stream(block);
  ByteSpan cursor{stream.data(), stream.size()};
  auto begin = read_record(cursor);
  ASSERT_TRUE(static_cast<bool>(begin));
  EXPECT_FALSE(static_cast<bool>(read_auto_filter_block(cursor, begin.value().payload)));
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
