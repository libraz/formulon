// AutoFilter forms Excel 16.115 saves for the blank / non-blank custom criteria and the calendar types
// (tests/fixtures/excel/afilter), read into the model and written back.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "gtest/gtest.h"
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

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
