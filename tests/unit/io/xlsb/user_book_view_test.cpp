// Custom workbook views in XLSB: an `.xlsb` source's BrtUserBookView records
// go back verbatim between BrtCalcProp and BrtFileRecover / BrtEndBook, and
// only together with the sheet custom-view blocks that name them. An `.xlsx`
// source's custom views have no record encoding here, so both sides are left
// out and counted. Fixtures: Excel saves under
// tests/fixtures/excel/retained_children_edit/.

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <set>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/user_book_views.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "workbook.h"
#include "writer_test_helpers.h"
#include "zip_repack_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::uint16_t kBrtEndBook = 132;
constexpr std::uint16_t kBrtFileRecover = 155;
constexpr std::uint16_t kBrtCalcProp = 157;
constexpr std::uint16_t kBrtUserBookView = 397;
constexpr std::uint16_t kBrtBeginUserShViews = 422;
constexpr std::uint16_t kBrtBeginUserShView = 423;
constexpr const char* kWorkbookBin = "xl/workbook.bin";
constexpr const char* kSheetBin = "xl/worksheets/sheet1.bin";

const char* const kChildren[] = {
    "children_0_base",           "children_1_insert_row4", "children_2_delete_row5",
    "children_3_delete_rows3_6", "children_4_delete_row9", "children_5_delete_rows2_10",
    "children_6_insert_colC",    "children_7_delete_colB", "children_8_delete_colsB_D",
};

std::vector<std::uint8_t> Fixture(const std::string& name) {
  return ReadFixtureFile(std::string(FORMULON_FIXTURES_DIR) + "/excel/retained_children_edit/" + name);
}

std::vector<std::uint8_t> Entry(const std::vector<std::uint8_t>& package, const std::string& path) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto entry = zip.read_entry(path);
  EXPECT_TRUE(static_cast<bool>(entry)) << path;
  return entry ? std::move(entry.value()) : std::vector<std::uint8_t>{};
}

/// The framed records of `type` in `body`, concatenated in order.
std::vector<std::uint8_t> RecordsOf(const std::vector<std::uint8_t>& body, std::uint16_t type) {
  std::vector<std::uint8_t> out;
  for (const FramedRecord& rec : split_records(body)) {
    if (rec.type == type) {
      out.insert(out.end(), body.begin() + static_cast<std::ptrdiff_t>(rec.begin),
                 body.begin() + static_cast<std::ptrdiff_t>(rec.end));
    }
  }
  return out;
}

std::vector<std::uint16_t> TypesOf(const std::vector<std::uint8_t>& body) {
  std::vector<std::uint16_t> types;
  for (const FramedRecord& rec : split_records(body)) {
    types.push_back(rec.type);
  }
  return types;
}

/// Index of the first `type` in `types`, or `types.size()`.
std::size_t IndexOf(const std::vector<std::uint16_t>& types, std::uint16_t type) {
  return static_cast<std::size_t>(std::find(types.begin(), types.end(), type) - types.begin());
}

TEST(XlsbUserBookView, XlsbSourceKeepsTheViewsAtExcelsPosition) {
  for (const char* name : kChildren) {
    SCOPED_TRACE(name);
    const std::vector<std::uint8_t> source = Fixture(std::string(name) + ".xlsb");
    const std::vector<std::uint8_t> expected = RecordsOf(Entry(source, kWorkbookBin), kBrtUserBookView);
    ASSERT_FALSE(expected.empty());
    auto read = read_xlsb(SpanOf(source));
    ASSERT_TRUE(static_cast<bool>(read)) << read.error().message;
    auto written = write_xlsb_with_result(read.value().workbook);
    ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
    const std::vector<std::uint8_t> book = Entry(written.value().bytes, kWorkbookBin);
    EXPECT_EQ(RecordsOf(book, kBrtUserBookView), expected);
    const std::vector<std::uint16_t> types = TypesOf(book);
    const std::size_t first = IndexOf(types, kBrtUserBookView);
    const std::size_t last = static_cast<std::size_t>(
        types.size() - 1U - (std::find(types.rbegin(), types.rend(), kBrtUserBookView) - types.rbegin()));
    ASSERT_LT(first, types.size());
    if (IndexOf(types, kBrtCalcProp) < types.size()) {
      EXPECT_LT(IndexOf(types, kBrtCalcProp), first);
    }
    EXPECT_LT(last, IndexOf(types, kBrtEndBook));
    EXPECT_LT(last, IndexOf(types, kBrtFileRecover) < types.size() ? IndexOf(types, kBrtFileRecover) : types.size());
    // The sheet side the views belong to is written with them.
    EXPECT_FALSE(RecordsOf(Entry(written.value().bytes, kSheetBin), kBrtBeginUserShView).empty());
  }
}

TEST(XlsbUserBookView, XlsxSourceLeavesBothSidesOutAndCountsThem) {
  for (const char* name : kChildren) {
    SCOPED_TRACE(name);
    auto read = read_ooxml(SpanOf(Fixture(std::string(name) + ".xlsx")));
    ASSERT_TRUE(static_cast<bool>(read)) << read.error().message;
    const Workbook& wb = read.value().workbook;
    ASSERT_EQ(unwritten_xml_book_views(wb), 1U);
    ASSERT_EQ(unwritten_xml_sheet_views(wb.sheet(0)), 1U);
    auto written = write_xlsb_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
    EXPECT_TRUE(RecordsOf(Entry(written.value().bytes, kWorkbookBin), kBrtUserBookView).empty());
    const std::vector<std::uint8_t> sheet = Entry(written.value().bytes, kSheetBin);
    EXPECT_TRUE(RecordsOf(sheet, kBrtBeginUserShViews).empty());
    EXPECT_TRUE(RecordsOf(sheet, kBrtBeginUserShView).empty());
    EXPECT_GE(written.value().diagnostics.deferred_feature_count, 2U);
  }
}

using Guid = std::array<std::uint8_t, 16>;

/// GUIDs of the records of `type` in `body`, read at payload offset `at`.
std::set<Guid> GuidsOf(const std::vector<std::uint8_t>& body, std::uint16_t type, std::size_t at) {
  std::set<Guid> out;
  for (const FramedRecord& rec : split_records(body)) {
    if (rec.type == type && rec.payload.size >= at + 16U) {
      Guid guid{};
      std::copy_n(rec.payload.data + at, 16U, guid.begin());
      out.insert(guid);
    }
  }
  return out;
}

/// Saves `wb` and checks that every written workbook view has a sheet view
/// naming it and every sheet view a workbook view: Excel refuses the one
/// unpaired form and crashes on the other.
void ExpectPairedViews(const Workbook& wb, std::size_t expected_views) {
  auto written = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  const std::set<Guid> book = GuidsOf(Entry(written.value().bytes, kWorkbookBin), kBrtUserBookView, 24U);
  std::set<Guid> sheets;
  for (std::size_t i = 1; i <= wb.sheet_count(); ++i) {
    const std::set<Guid> one = GuidsOf(Entry(written.value().bytes, "xl/worksheets/sheet" + std::to_string(i) + ".bin"),
                                       kBrtBeginUserShView, 0U);
    sheets.insert(one.begin(), one.end());
  }
  EXPECT_EQ(book, sheets);
  EXPECT_EQ(book.size(), expected_views);
}

Workbook ReadChildren(const char* name) {
  auto read = read_xlsb(SpanOf(Fixture(std::string(name) + ".xlsb")));
  EXPECT_TRUE(static_cast<bool>(read)) << read.error().message;
  return read ? std::move(read.value().workbook) : Workbook::create_empty();
}

TEST(XlsbUserBookView, ReorderedCopiedAndEditedSheetsKeepThePair) {
  {
    Workbook wb = ReadChildren("children_0_base");
    wb.add_sheet("Other");
    ASSERT_TRUE(static_cast<bool>(wb.move_sheet(0U, 1U)));
    ExpectPairedViews(wb, 1U);
  }
  {
    Workbook wb = ReadChildren("children_0_base");
    ASSERT_TRUE(static_cast<bool>(wb.copy_sheet(0U, 1U, std::nullopt)));
    ExpectPairedViews(wb, 1U);
  }
  {
    Workbook wb = ReadChildren("children_0_base");
    ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0U, 3U, 2U)));
    ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0U, 0U, 1U)));
    ExpectPairedViews(wb, 1U);
  }
}

TEST(XlsbUserBookView, OnlyViewsASheetStillCarriesAreWritten) {
  Workbook wb = ReadChildren("children_0_base");
  // A second workbook view no sheet names: it goes, the carried one stays.
  std::vector<std::uint8_t> views = wb.xlsb_user_book_views();
  const std::vector<FramedRecord> recs = split_records(views);
  ASSERT_EQ(recs.size(), 1U);
  std::vector<std::uint8_t> other(views.begin() + static_cast<std::ptrdiff_t>(recs[0].begin),
                                  views.begin() + static_cast<std::ptrdiff_t>(recs[0].end));
  const std::size_t guid_at = recs[0].payload.data - views.data() - recs[0].begin + 24U;
  other[guid_at] ^= 0xFFU;
  views.insert(views.end(), other.begin(), other.end());
  wb.set_xlsb_user_book_views(views);
  ExpectPairedViews(wb, 1U);
  auto written = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(written));
  EXPECT_GE(written.value().diagnostics.deferred_feature_count, 1U);
}

TEST(XlsbUserBookView, ViewWithoutASheetSideIsLeftOut) {
  auto read = read_xlsb(SpanOf(Fixture("children_0_base.xlsb")));
  ASSERT_TRUE(static_cast<bool>(read)) << read.error().message;
  Workbook& wb = read.value().workbook;
  // The only sheet carrying the custom view goes; the workbook view goes with it.
  wb.add_sheet("Other");
  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(0U)));
  auto written = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  EXPECT_TRUE(RecordsOf(Entry(written.value().bytes, kWorkbookBin), kBrtUserBookView).empty());
  EXPECT_TRUE(RecordsOf(Entry(written.value().bytes, kSheetBin), kBrtBeginUserShView).empty());
  std::vector<std::uint8_t> book = Entry(written.value().bytes, kWorkbookBin);
  EXPECT_EQ(splice_book_user_views(book, wb), 1U);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
