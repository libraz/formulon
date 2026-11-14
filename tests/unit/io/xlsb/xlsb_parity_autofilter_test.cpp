// The sheet AutoFilter block the xlsb codec writes from the model, held against Excel 16.115's own xlsb saves of the
// same sources (tests/fixtures/excel/xlsb_parity/autofilter).

#include <algorithm>
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
#include "io/xlsb/writer.h"
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

Bytes ReadFile(const std::string& path) {
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

std::string Fixture(const std::string& name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/autofilter/" + name;
}

Records AsRecords(const Bytes& stream) {
  Records out;
  for (const FramedRecord& rec : split_records(stream)) {
    out.emplace_back(rec.type, Bytes(rec.payload.data, rec.payload.data + rec.payload.size));
  }
  return out;
}

/// The first worksheet's records from BrtBeginAFilter through its BrtEndAFilter.
Records ExcelBlock(const std::string& name) {
  const Bytes package = ReadFile(Fixture(name + ".xlsb"));
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

/// The source each case is saved from. Excel computes a top-ten threshold and an above-average bound from the data
/// when it applies the filter, so those two cases start from its own `.xlsx` save, which carries them.
std::string SourceFile(const std::string& name) {
  return name == "top10" || name == "dynamic" ? name + ".xlsx" : name + "_src.xlsx";
}

AutoFilter SourceFilter(const std::string& name) {
  auto loaded = read_ooxml(SpanOf(ReadFile(Fixture(SourceFile(name)))));
  EXPECT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  if (!loaded || loaded.value().workbook.sheet(0).auto_filter() == nullptr) {
    ADD_FAILURE() << "no sheet AutoFilter in " << name;
    return AutoFilter{};
  }
  return *loaded.value().workbook.sheet(0).auto_filter();
}

const char* const kCases[] = {"values",  "calendar_type",     "filtered_rows_hidden", "custom_two", "top10",
                              "dynamic", "color_dxf",         "color_no_dxf",         "icon",       "phonetic_sort",
                              "x14_ext", "sort_beside_filter"};

TEST(XlsbParityAutoFilterTwins, BlockFromTheSourceMatchesExcel) {
  for (const char* name : kCases) {
    SCOPED_TRACE(name);
    Bytes out;
    EXPECT_EQ(emit_auto_filter_block(out, SourceFilter(name)), 0U);
    EXPECT_EQ(AsRecords(out), ExcelBlock(name));
  }
}

TEST(XlsbParityAutoFilterTwins, ExcelBlockReadsBackToTheSameRecords) {
  for (const char* name : kCases) {
    SCOPED_TRACE(name);
    Bytes stream;
    for (const auto& rec : ExcelBlock(name)) {
      emit_record(stream, rec.first, ByteSpan{rec.second.data(), rec.second.size()});
    }
    ByteSpan cursor{stream.data(), stream.size()};
    auto begin = read_record(cursor);
    ASSERT_TRUE(static_cast<bool>(begin));
    auto filter = read_auto_filter_block(cursor, begin.value().payload);
    ASSERT_TRUE(static_cast<bool>(filter)) << filter.error().message;
    Bytes out;
    EXPECT_EQ(emit_auto_filter_block(out, filter.value()), 0U);
    EXPECT_EQ(AsRecords(out), AsRecords(stream));
  }
}

TEST(XlsbParityAutoFilterTwins, AlternateSortMethodReadsAsStroke) {
  // Excel saves a pinYin sort with the alternate-method bit and reads that bit back as stroke.
  Bytes stream;
  for (const auto& rec : ExcelBlock("phonetic_sort")) {
    emit_record(stream, rec.first, ByteSpan{rec.second.data(), rec.second.size()});
  }
  ByteSpan cursor{stream.data(), stream.size()};
  auto begin = read_record(cursor);
  ASSERT_TRUE(static_cast<bool>(begin));
  auto filter = read_auto_filter_block(cursor, begin.value().payload);
  ASSERT_TRUE(static_cast<bool>(filter)) << filter.error().message;
  ASSERT_TRUE(filter.value().sort.has_value());
  EXPECT_EQ(filter.value().sort->sort_method, SortMethod::kStroke);
}

bool IsFilterRecord(std::uint16_t type) {
  return (type >= kBeginAFilter && type <= kBrtAFilterDateGroupItem) ||
         (type >= kBrtBeginSortState && type <= kBrtEndSortState) ||
         type == static_cast<std::uint16_t>(XlsbRecordType::BrtWsProp);
}

/// The first worksheet's BrtWsProp and AutoFilter / sort-state records, in stream order.
Records FilterRecords(const Bytes& package) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto part = zip.read_entry("xl/worksheets/sheet1.bin");
  EXPECT_TRUE(static_cast<bool>(part));
  Records out;
  if (part) {
    for (auto& rec : AsRecords(part.value())) {
      if (IsFilterRecord(rec.first)) {
        out.push_back(std::move(rec));
      }
    }
  }
  return out;
}

TEST(XlsbParityAutoFilterEdit, SavedSheetMatchesExcel) {
  for (const char* name : kCases) {
    SCOPED_TRACE(name);
    const Bytes excel = ReadFile(Fixture(std::string(name) + ".xlsb"));
    auto loaded = read_ooxml(SpanOf(ReadFile(Fixture(SourceFile(name)))));
    ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
    auto saved = write_xlsb_with_result(loaded.value().workbook);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    EXPECT_EQ(FilterRecords(saved.value().bytes), FilterRecords(excel));
    EXPECT_EQ(unwritten_auto_filter_settings(loaded.value().workbook.sheet(0)), 0U);
    // The Excel save reads into the model and writes back the same records.
    auto reread = read_xlsb(SpanOf(excel));
    ASSERT_TRUE(static_cast<bool>(reread)) << reread.error().message;
    EXPECT_NE(reread.value().workbook.sheet(0).auto_filter(), nullptr);
    auto resaved = write_xlsb_with_result(reread.value().workbook);
    ASSERT_TRUE(static_cast<bool>(resaved)) << resaved.error().message;
    EXPECT_EQ(FilterRecords(resaved.value().bytes), FilterRecords(excel));
  }
}

TEST(XlsbParityAutoFilterTwins, UnmeasuredCriteriaStayUnwritten) {
  // Excel keeps one value of three custom criteria and writes bounds it computes from the data; neither is in the
  // model, so each filter is reported as lost.
  {
    Bytes out;
    EXPECT_GT(emit_auto_filter_block(out, SourceFilter("custom_three")), 0U);
  }
  for (const char* name : {"top10_src.xlsx", "dynamic_src.xlsx"}) {
    SCOPED_TRACE(name);
    auto loaded = read_ooxml(SpanOf(ReadFile(Fixture(name))));
    ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
    EXPECT_GT(unwritten_auto_filter_settings(loaded.value().workbook.sheet(0)), 0U);
  }
}

/// Each worksheet part of a package, keyed by sheet name through the workbook's bundle records and relationships.
std::vector<std::pair<std::string, Bytes>> SheetParts(const Bytes& package) {
  std::vector<std::pair<std::string, Bytes>> out;
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto book = zip.read_entry("xl/workbook.bin");
  auto rels = zip.read_entry("xl/_rels/workbook.bin.rels");
  EXPECT_TRUE(book && rels);
  if (!book || !rels) {
    return out;
  }
  const std::string rel_xml(rels.value().begin(), rels.value().end());
  for (const FramedRecord& rec : split_records(book.value())) {
    if (rec.type != static_cast<std::uint16_t>(XlsbRecordType::BrtBundleSh)) {
      continue;
    }
    ByteSpan payload = rec.payload;
    (void)read_u32(payload);
    (void)read_u32(payload);
    auto rel_id = read_xlnullablewidestring(payload);
    auto name = read_xlwidestring(payload);
    EXPECT_TRUE(rel_id && name);
    if (!rel_id || !name) {
      continue;
    }
    const std::size_t at = rel_xml.find("Id=\"" + rel_id.value() + "\"");
    const std::size_t open = at == std::string::npos ? at : rel_xml.rfind('<', at);
    const std::size_t target = open == std::string::npos ? open : rel_xml.find("Target=\"", open);
    EXPECT_NE(target, std::string::npos) << rel_id.value();
    if (target == std::string::npos) {
      continue;
    }
    const std::size_t begin = target + 8U;
    std::string path = rel_xml.substr(begin, rel_xml.find('"', begin) - begin);
    path = path.rfind("/xl/", 0) == 0 ? path.substr(1) : "xl/" + path;
    auto part = zip.read_entry(path);
    EXPECT_TRUE(static_cast<bool>(part)) << path;
    if (part) {
      out.emplace_back(name.value(), std::move(part.value()));
    }
  }
  return out;
}

/// A part's BrtWsProp and AutoFilter / sort-state records as `type:hex` lines, which keep a failure readable.
std::vector<std::string> PartFilterRecords(const Bytes& part) {
  static const char kHex[] = "0123456789abcdef";
  std::vector<std::string> out;
  for (auto& rec : AsRecords(part)) {
    if (!IsFilterRecord(rec.first)) {
      continue;
    }
    // Excel leaves the numeric slot of a text or non-blank criterion uninitialised.
    if (rec.first == kBrtCustomFilter && rec.second.size() >= 10U && rec.second[0] != 0x04U) {
      std::fill(rec.second.begin() + 2, rec.second.begin() + 10, std::uint8_t{0});
    }
    std::string line = std::to_string(rec.first) + ":";
    for (const std::uint8_t byte : rec.second) {
      line.push_back(kHex[byte >> 4U]);
      line.push_back(kHex[byte & 0x0FU]);
    }
    out.push_back(std::move(line));
  }
  return out;
}

TEST(XlsbParityAutoFilterKinds, EveryWrittenKindMatchesExcel) {
  // Excel's .xlsb save of the AutoFilter kinds book, sheet by sheet against ours from its .xlsx.
  const std::string root = std::string(FORMULON_FIXTURES_DIR) + "/excel/auto_filter_kinds";
  const Bytes excel = ReadFile(root + ".xlsb");
  auto loaded = read_ooxml(SpanOf(ReadFile(root + ".xlsx")));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  const Workbook& book = loaded.value().workbook;
  auto saved = write_xlsb_with_result(book);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  const auto ours = SheetParts(saved.value().bytes);
  const auto theirs = SheetParts(excel);
  ASSERT_EQ(ours.size(), book.sheet_count());
  ASSERT_EQ(theirs.size(), book.sheet_count());
  std::size_t compared = 0;
  for (std::size_t i = 0; i < book.sheet_count(); ++i) {
    SCOPED_TRACE(ours[i].first);
    ASSERT_EQ(ours[i].first, theirs[i].first);
    const AutoFilter* filter = book.sheet(i).auto_filter();
    if (filter != nullptr && unwritten_auto_filter_settings(book.sheet(i)) != 0U) {
      // Only date groups in the rich-data form are left unwritten in this book.
      bool rich_dates = false;
      for (const FilterColumn& col : filter->columns) {
        rich_dates = rich_dates || (!col.values.date_groups.empty() && !col.values.plain_date_groups);
      }
      EXPECT_TRUE(rich_dates);
      continue;
    }
    EXPECT_EQ(PartFilterRecords(ours[i].second), PartFilterRecords(theirs[i].second));
    ++compared;
  }
  EXPECT_GE(compared, 50U);
}

TEST(XlsbParityAutoFilterKinds, ExcelSaveWritesBackUnchanged) {
  const Bytes excel = ReadFile(std::string(FORMULON_FIXTURES_DIR) + "/excel/auto_filter_kinds.xlsb");
  auto loaded = read_xlsb(SpanOf(excel));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  auto saved = write_xlsb_with_result(loaded.value().workbook);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  const auto ours = SheetParts(saved.value().bytes);
  const auto theirs = SheetParts(excel);
  ASSERT_EQ(ours.size(), theirs.size());
  for (std::size_t i = 0; i < ours.size(); ++i) {
    SCOPED_TRACE(theirs[i].first);
    EXPECT_EQ(PartFilterRecords(ours[i].second), PartFilterRecords(theirs[i].second));
  }
}

TEST(XlsbParityAutoFilterDateGroups, EveryPlainGroupingMatchesExcel) {
  // One sheet per grouping, year through second, each a plain `<filters><dateGroupItem>`.
  const Bytes excel = ReadFile(Fixture("date_group_plain.xlsb"));
  auto loaded = read_ooxml(SpanOf(ReadFile(Fixture("date_group_plain_src.xlsx"))));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  const Workbook& book = loaded.value().workbook;
  auto saved = write_xlsb_with_result(book);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  auto reread = read_xlsb(SpanOf(excel));
  ASSERT_TRUE(static_cast<bool>(reread)) << reread.error().message;
  auto resaved = write_xlsb_with_result(reread.value().workbook);
  ASSERT_TRUE(static_cast<bool>(resaved)) << resaved.error().message;
  const auto ours = SheetParts(saved.value().bytes);
  const auto theirs = SheetParts(excel);
  const auto back = SheetParts(resaved.value().bytes);
  ASSERT_EQ(book.sheet_count(), 6U);
  ASSERT_EQ(ours.size(), book.sheet_count());
  ASSERT_EQ(theirs.size(), book.sheet_count());
  ASSERT_EQ(back.size(), book.sheet_count());
  // Excel's own .xlsx save keeps the plain form, which the model writes back as read.
  const Bytes twin_package = ReadFile(Fixture("date_group_plain.xlsx"));
  auto twin = read_ooxml(SpanOf(twin_package));
  ASSERT_TRUE(static_cast<bool>(twin)) << twin.error().message;
  ZipReader twin_zip;
  ASSERT_TRUE(static_cast<bool>(twin_zip.open(SpanOf(twin_package))));
  for (std::size_t i = 0; i < book.sheet_count(); ++i) {
    SCOPED_TRACE(theirs[i].first);
    EXPECT_EQ(unwritten_auto_filter_settings(book.sheet(i)), 0U);
    EXPECT_EQ(PartFilterRecords(ours[i].second), PartFilterRecords(theirs[i].second));
    EXPECT_EQ(PartFilterRecords(back[i].second), PartFilterRecords(theirs[i].second));
    const AutoFilter* from_xlsb = reread.value().workbook.sheet(i).auto_filter();
    const AutoFilter* from_xlsx = twin.value().workbook.sheet(i).auto_filter();
    ASSERT_NE(from_xlsb, nullptr);
    ASSERT_NE(from_xlsx, nullptr);
    EXPECT_TRUE(from_xlsx->columns.at(0).values.plain_date_groups);
    auto sheet_xml = twin_zip.read_entry("xl/worksheets/sheet" + std::to_string(i + 1U) + ".xml");
    ASSERT_TRUE(static_cast<bool>(sheet_xml));
    const std::string text(sheet_xml.value().begin(), sheet_xml.value().end());
    const std::size_t begin = text.find("<autoFilter");
    const std::size_t end = text.find("</autoFilter>");
    ASSERT_NE(begin, std::string::npos);
    ASSERT_NE(end, std::string::npos);
    EXPECT_EQ(serialize_auto_filter(*from_xlsx), text.substr(begin, end + 13U - begin));
    EXPECT_EQ(serialize_auto_filter(*from_xlsb), serialize_auto_filter(*book.sheet(i).auto_filter()));
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
