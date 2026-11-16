// Custom workbook views in XLSB: an `.xlsb` source's BrtUserBookView records
// go back verbatim between BrtCalcProp and BrtFileRecover / BrtEndBook, and
// only together with the sheet custom-view blocks that name them. An `.xlsx`
// source's custom views are encoded as Excel writes them. Fixtures: Excel
// saves under tests/fixtures/excel/retained_children_edit/ and
// tests/fixtures/custom_views/.

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
#include "io/ooxml_writer.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/user_book_views.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "support/xlsb_package_validator.h"
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

/// The framed records from the first `begin` through the next `end`, and the
/// records outside that span, of `body`.
struct SplitBlock {
  std::vector<std::uint8_t> inside;
  std::vector<std::uint8_t> outside;
};

SplitBlock SplitAtBlock(const std::vector<std::uint8_t>& body, std::uint16_t begin, std::uint16_t end) {
  SplitBlock out;
  bool inside = false;
  bool done = false;
  for (const FramedRecord& rec : split_records(body)) {
    const bool in_block = !done && (inside || rec.type == begin);
    std::vector<std::uint8_t>& dst = in_block ? out.inside : out.outside;
    dst.insert(dst.end(), body.begin() + static_cast<std::ptrdiff_t>(rec.begin),
               body.begin() + static_cast<std::ptrdiff_t>(rec.end));
    inside = in_block && rec.type != end;
    done = done || (in_block && rec.type == end);
  }
  return out;
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
        types.size() - 1U -
        static_cast<std::size_t>(std::find(types.rbegin(), types.rend(), kBrtUserBookView) - types.rbegin()));
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

TEST(XlsbUserBookView, XlsxSourceWritesBothSidesAsExcelDoes) {
  for (const char* name : kChildren) {
    SCOPED_TRACE(name);
    const std::vector<std::uint8_t> twin = Fixture(std::string(name) + ".xlsb");
    auto read = read_ooxml(SpanOf(Fixture(std::string(name) + ".xlsx")));
    ASSERT_TRUE(static_cast<bool>(read)) << read.error().message;
    const Workbook& wb = read.value().workbook;
    ASSERT_EQ(xml_book_view_count(wb), 1U);
    ASSERT_EQ(xml_sheet_view_count(wb.sheet(0)), 1U);
    auto written = write_xlsb_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
    const std::vector<std::uint8_t> book = Entry(written.value().bytes, kWorkbookBin);
    EXPECT_EQ(RecordsOf(book, kBrtUserBookView), RecordsOf(Entry(twin, kWorkbookBin), kBrtUserBookView));
    const std::vector<std::uint16_t> types = TypesOf(book);
    EXPECT_LT(IndexOf(types, kBrtCalcProp), IndexOf(types, kBrtUserBookView));
    EXPECT_LT(IndexOf(types, kBrtUserBookView), IndexOf(types, kBrtEndBook));
    const std::vector<std::uint8_t> sheet = Entry(written.value().bytes, kSheetBin);
    const std::vector<std::uint8_t> twin_sheet = Entry(twin, kSheetBin);
    EXPECT_EQ(SplitAtBlock(sheet, kBrtBeginUserShViews, kBrtEndUserShViews).inside,
              SplitAtBlock(twin_sheet, kBrtBeginUserShViews, kBrtEndUserShViews).inside);
    // The block closes the records ahead of the merged cells, where Excel puts it.
    const std::vector<std::uint16_t> sheet_types = TypesOf(sheet);
    const std::vector<std::uint16_t> twin_types = TypesOf(twin_sheet);
    const std::size_t end = IndexOf(sheet_types, kBrtEndUserShViews);
    const std::size_t twin_end = IndexOf(twin_types, kBrtEndUserShViews);
    ASSERT_LT(end + 1U, sheet_types.size());
    ASSERT_LT(twin_end + 1U, twin_types.size());
    EXPECT_EQ(sheet_types[end + 1U], twin_types[twin_end + 1U]);
    EXPECT_EQ(written.value().diagnostics.deferred_feature_count, 0U);
  }
}

/// The `<tag>...</tag>` element of `xml`, empty when it has none.
std::string ElementOf(const std::string& xml, const std::string& tag) {
  const std::size_t begin = xml.find("<" + tag + ">");
  const std::size_t end = xml.find("</" + tag + ">");
  if (begin == std::string::npos || end == std::string::npos) {
    return std::string();
  }
  return xml.substr(begin, end + tag.size() + 3U - begin);
}

std::string Text(const std::vector<std::uint8_t>& bytes) {
  return std::string(bytes.begin(), bytes.end());
}

TEST(XlsbUserBookView, XlsbSourceWritesBothSidesToXlsxAsExcelDoes) {
  for (const char* name : kChildren) {
    SCOPED_TRACE(name);
    const std::vector<std::uint8_t> twin = Fixture(std::string(name) + ".xlsx");
    auto read = read_xlsb(SpanOf(Fixture(std::string(name) + ".xlsb")));
    ASSERT_TRUE(static_cast<bool>(read)) << read.error().message;
    auto written = write_ooxml_with_result(read.value().workbook);
    ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
    const std::string book = ElementOf(Text(Entry(written.value().bytes, "xl/workbook.xml")), "customWorkbookViews");
    ASSERT_FALSE(book.empty());
    EXPECT_EQ(book, ElementOf(Text(Entry(twin, "xl/workbook.xml")), "customWorkbookViews"));
    EXPECT_EQ(ElementOf(Text(Entry(written.value().bytes, "xl/worksheets/sheet1.xml")), "customSheetViews"),
              ElementOf(Text(Entry(twin, "xl/worksheets/sheet1.xml")), "customSheetViews"));
  }
}

TEST(XlsbUserBookView, XlsbViewWithoutAnXlsxFormLeavesBothSidesOut) {
  auto read = read_xlsb(SpanOf(Fixture("children_0_base.xlsb")));
  ASSERT_TRUE(static_cast<bool>(read)) << read.error().message;
  Workbook& wb = read.value().workbook;
  // An automatic page break inside the view has no `.xlsx` form: the view goes on both sides.
  XlsbSheetTail tail = wb.sheet(0).xlsb_tail();
  std::vector<std::uint8_t> header;
  emit_record(header, kBrtBeginRwBrk, std::vector<std::uint8_t>{0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00});
  emit_record(header, kBrtBrk, std::vector<std::uint8_t>{0x05, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xff, 0x3f,
                                                         0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00});
  emit_record(header, kBrtEndRwBrk, ByteSpan{});
  std::vector<std::uint8_t>& slot = tail.before_merges;
  const std::vector<FramedRecord> recs = split_records(slot);
  const auto end_view =
      std::find_if(recs.begin(), recs.end(), [](const FramedRecord& rec) { return rec.type == kBrtEndUserShView; });
  ASSERT_NE(end_view, recs.end());
  slot.insert(slot.begin() + static_cast<std::ptrdiff_t>(end_view->begin), header.begin(), header.end());
  auto baseline = write_ooxml_with_result(read.value().workbook);
  ASSERT_TRUE(static_cast<bool>(baseline)) << baseline.error().message;
  wb.sheet(0).set_xlsb_tail(tail);
  auto written = write_ooxml_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  // The left-out view counts once, as on an .xlsb save.
  EXPECT_EQ(written.value().diagnostics.deferred_feature_count,
            baseline.value().diagnostics.deferred_feature_count + 1U);
  EXPECT_EQ(Text(Entry(written.value().bytes, "xl/workbook.xml")).find("customWorkbookView"), std::string::npos);
  EXPECT_EQ(Text(Entry(written.value().bytes, "xl/worksheets/sheet1.xml")).find("customSheetView"), std::string::npos);
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

/// Saves `wb` and checks that the package passes the structural validator,
/// which pairs every workbook view with a sheet view naming it, and that
/// `expected_views` workbook views are written.
void ExpectPairedViews(const Workbook& wb, std::size_t expected_views) {
  auto written = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  const std::vector<std::uint8_t>& bytes = written.value().bytes;
  EXPECT_EQ(::formulon::test::ValidateXlsbPackage(SpanOf(bytes)), std::vector<std::string>{});
  EXPECT_EQ(GuidsOf(Entry(bytes, kWorkbookBin), kBrtUserBookView, 24U).size(), expected_views);
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
  const std::size_t guid_at = static_cast<std::size_t>(recs[0].payload.data - views.data()) - recs[0].begin + 24U;
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
  EXPECT_EQ(splice_book_user_views(book, wb, {}), 1U);
}

std::vector<std::uint8_t> CustomViewsFixture(const std::string& name) {
  return ReadFixtureFile(std::string(FORMULON_FIXTURES_DIR) + "/custom_views/" + name);
}

TEST(XlsbUserBookView, ViewMarginsAreNotTheSheetsOnLoad) {
  auto xlsb = read_xlsb(SpanOf(Fixture("children_0_base.xlsb")));
  ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
  auto xlsx = read_ooxml(SpanOf(Fixture("children_0_base.xlsx")));
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  const SheetPrintSettings& print = xlsb.value().workbook.sheet(0).print_settings();
  ASSERT_FALSE(print.page_margins_xml.empty());
  EXPECT_EQ(print.page_margins_xml, xlsx.value().workbook.sheet(0).print_settings().page_margins_xml);
}

TEST(XlsbUserBookView, XlsbSaveKeepsTheViewBlockAndTheSheetsPrintRecords) {
  const std::vector<std::uint8_t> source = Entry(Fixture("children_0_base.xlsb"), kSheetBin);
  const SplitBlock expected = SplitAtBlock(source, kBrtBeginUserShViews, kBrtEndUserShViews);
  ASSERT_FALSE(RecordsOf(expected.inside, kBrtMargins).empty());
  auto read = read_xlsb(SpanOf(Fixture("children_0_base.xlsb")));
  ASSERT_TRUE(static_cast<bool>(read)) << read.error().message;
  auto written = write_xlsb_with_result(read.value().workbook);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  const SplitBlock actual =
      SplitAtBlock(Entry(written.value().bytes, kSheetBin), kBrtBeginUserShViews, kBrtEndUserShViews);
  EXPECT_EQ(actual.inside, expected.inside);
  EXPECT_EQ(RecordsOf(actual.outside, kBrtMargins), RecordsOf(expected.outside, kBrtMargins));
  EXPECT_EQ(RecordsOf(actual.outside, kBrtPrintOptions), RecordsOf(expected.outside, kBrtPrintOptions));
}

/// The text of `xml` from the first `open` through the next `close`, empty when either is missing.
std::string ElementSpan(const std::string& xml, const std::string& open, const std::string& close) {
  const std::size_t begin = xml.find(open);
  const std::size_t end = xml.find(close, begin);
  if (begin == std::string::npos || end == std::string::npos) {
    return std::string();
  }
  return xml.substr(begin, end + close.size() - begin);
}

/// The framed records of `body` except those of the given types.
std::vector<std::uint8_t> WithoutTypes(const std::vector<std::uint8_t>& body, const std::set<std::uint16_t>& types) {
  std::vector<std::uint8_t> out;
  for (const FramedRecord& rec : split_records(body)) {
    if (types.count(rec.type) == 0U) {
      out.insert(out.end(), body.begin() + static_cast<std::ptrdiff_t>(rec.begin),
                 body.begin() + static_cast<std::ptrdiff_t>(rec.end));
    }
  }
  return out;
}

TEST(XlsbUserBookView, ViewAutoFilterIsNotTheSheetsAndStaysInTheView) {
  const std::vector<std::uint8_t> package = CustomViewsFixture("autofilter_view.xlsb");
  const SplitBlock expected = SplitAtBlock(Entry(package, kSheetBin), kBrtBeginUserShViews, kBrtEndUserShViews);
  ASSERT_FALSE(RecordsOf(expected.inside, kBeginAFilter).empty());
  auto read = read_xlsb(SpanOf(package));
  ASSERT_TRUE(static_cast<bool>(read)) << read.error().message;
  auto twin = read_ooxml(SpanOf(CustomViewsFixture("autofilter_view.xlsx")));
  ASSERT_TRUE(static_cast<bool>(twin)) << twin.error().message;
  EXPECT_FALSE(twin.value().workbook.sheet(0).has_auto_filter());
  EXPECT_FALSE(read.value().workbook.sheet(0).has_auto_filter());
  auto written = write_xlsb_with_result(read.value().workbook);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  const SplitBlock actual =
      SplitAtBlock(Entry(written.value().bytes, kSheetBin), kBrtBeginUserShViews, kBrtEndUserShViews);
  EXPECT_EQ(actual.inside, expected.inside);
  EXPECT_TRUE(RecordsOf(actual.outside, kBeginAFilter).empty());
  // The view goes to an .xlsx save with its filter, apart from the sheet's own.
  auto xlsx = write_ooxml_with_result(read.value().workbook);
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  const std::string sheet_xml = Text(Entry(xlsx.value().bytes, "xl/worksheets/sheet1.xml"));
  const std::string views = ElementOf(sheet_xml, "customSheetViews");
  EXPECT_NE(views.find("<customSheetView "), std::string::npos);
  const std::string view = ElementSpan(sheet_xml, "<customSheetView ", "</customSheetView>");
  EXPECT_NE(view.find("<autoFilter ref=\"A1:E13\""), std::string::npos);
  // The fixture's sheet has no AutoFilter of its own, so none sits outside the view.
  std::string outside = sheet_xml;
  outside.erase(outside.find(view), view.size());
  EXPECT_EQ(outside.find("<autoFilter"), std::string::npos);
}

TEST(XlsbUserBookView, XlsxViewAutoFilterIsWrittenToTheXlsbViewOnly) {
  const SplitBlock expected = SplitAtBlock(Entry(CustomViewsFixture("autofilter_view.xlsb"), kSheetBin),
                                           kBrtBeginUserShViews, kBrtEndUserShViews);
  auto twin = read_ooxml(SpanOf(CustomViewsFixture("autofilter_view.xlsx")));
  ASSERT_TRUE(static_cast<bool>(twin)) << twin.error().message;
  auto written = write_xlsb_with_result(twin.value().workbook);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  const SplitBlock actual =
      SplitAtBlock(Entry(written.value().bytes, kSheetBin), kBrtBeginUserShViews, kBrtEndUserShViews);
  ASSERT_FALSE(RecordsOf(actual.inside, kBeginAFilter).empty());
  EXPECT_EQ(RecordsOf(actual.inside, kBeginAFilter), RecordsOf(expected.inside, kBeginAFilter));
  EXPECT_EQ(RecordsOf(actual.inside, kBrtEndAFilter), RecordsOf(expected.inside, kBrtEndAFilter));
  EXPECT_TRUE(RecordsOf(actual.outside, kBeginAFilter).empty());
}

TEST(XlsbUserBookView, FilterCriteriaViewCrossesBothWaysAsExcelSavesIt) {
  const std::vector<std::uint8_t> twin_xlsb = CustomViewsFixture("filtercriteria_view.xlsb");
  const std::vector<std::uint8_t> twin_xlsx = CustomViewsFixture("filtercriteria_view.xlsx");
  constexpr const char* kSheetXml = "xl/worksheets/sheet1.xml";
  const SplitBlock expected = SplitAtBlock(Entry(twin_xlsb, kSheetBin), kBrtBeginUserShViews, kBrtEndUserShViews);
  ASSERT_FALSE(RecordsOf(expected.outside, kBeginAFilter).empty());

  // .xlsx to .xlsb: the view's records, and the sheet's own filter.
  auto from_xlsx = read_ooxml(SpanOf(twin_xlsx));
  ASSERT_TRUE(static_cast<bool>(from_xlsx)) << from_xlsx.error().message;
  auto as_xlsb = write_xlsb_with_result(from_xlsx.value().workbook);
  ASSERT_TRUE(static_cast<bool>(as_xlsb)) << as_xlsb.error().message;
  const SplitBlock actual =
      SplitAtBlock(Entry(as_xlsb.value().bytes, kSheetBin), kBrtBeginUserShViews, kBrtEndUserShViews);
  EXPECT_EQ(WithoutTypes(actual.inside, {kBrtBeginUserShViews, kBrtEndUserShViews}),
            WithoutTypes(expected.inside, {kBrtBeginUserShViews, kBrtEndUserShViews, kAcBegin, kUid, kAcEnd}));
  EXPECT_EQ(RecordsOf(actual.outside, kBeginAFilter), RecordsOf(expected.outside, kBeginAFilter));

  // .xlsb to .xlsx: the view element minus the filter's revision uid, and the sheet's own filter.
  auto from_xlsb = read_xlsb(SpanOf(twin_xlsb));
  ASSERT_TRUE(static_cast<bool>(from_xlsb)) << from_xlsb.error().message;
  auto as_xlsx = write_ooxml_with_result(from_xlsb.value().workbook);
  ASSERT_TRUE(static_cast<bool>(as_xlsx)) << as_xlsx.error().message;
  const std::string saved = Text(Entry(as_xlsx.value().bytes, kSheetXml));
  std::string twin_view = ElementSpan(Text(Entry(twin_xlsx, kSheetXml)), "<customSheetView ", "</customSheetView>");
  ASSERT_FALSE(twin_view.empty());
  const std::size_t uid = twin_view.find(" xr:uid=\"");
  ASSERT_NE(uid, std::string::npos);
  twin_view.erase(uid, twin_view.find('"', uid + 9U) + 1U - uid);
  EXPECT_EQ(ElementSpan(saved, "<customSheetView ", "</customSheetView>"), twin_view);
  const std::size_t views_at = saved.find("<customSheetViews>");
  ASSERT_NE(views_at, std::string::npos);
  EXPECT_NE(saved.substr(0U, views_at).find("<autoFilter ref=\"A1:D13\""), std::string::npos);
}

/// Both directions of a view whose print children Excel saved in an `.xlsx` / `.xlsb` twin pair.
void ExpectPrintViewCrossesBothWays(const std::string& stem, const std::vector<const char*>& children) {
  const std::vector<std::uint8_t> twin_xlsb = CustomViewsFixture(stem + ".xlsb");
  const std::vector<std::uint8_t> twin_xlsx = CustomViewsFixture(stem + ".xlsx");
  constexpr const char* kSheetXml = "xl/worksheets/sheet1.xml";
  const SplitBlock expected = SplitAtBlock(Entry(twin_xlsb, kSheetBin), kBrtBeginUserShViews, kBrtEndUserShViews);
  ASSERT_FALSE(expected.inside.empty());

  // .xlsx to .xlsb: the view's records.
  auto from_xlsx = read_ooxml(SpanOf(twin_xlsx));
  ASSERT_TRUE(static_cast<bool>(from_xlsx)) << from_xlsx.error().message;
  auto as_xlsb = write_xlsb_with_result(from_xlsx.value().workbook);
  ASSERT_TRUE(static_cast<bool>(as_xlsb)) << as_xlsb.error().message;
  const SplitBlock actual =
      SplitAtBlock(Entry(as_xlsb.value().bytes, kSheetBin), kBrtBeginUserShViews, kBrtEndUserShViews);
  EXPECT_EQ(WithoutTypes(actual.inside, {kBrtBeginUserShViews, kBrtEndUserShViews}),
            WithoutTypes(expected.inside, {kBrtBeginUserShViews, kBrtEndUserShViews}));

  // .xlsb keeps the view on an .xlsb save.
  auto from_xlsb = read_xlsb(SpanOf(twin_xlsb));
  ASSERT_TRUE(static_cast<bool>(from_xlsb)) << from_xlsb.error().message;
  auto resaved = write_xlsb_with_result(from_xlsb.value().workbook);
  ASSERT_TRUE(static_cast<bool>(resaved)) << resaved.error().message;
  EXPECT_EQ(SplitAtBlock(Entry(resaved.value().bytes, kSheetBin), kBrtBeginUserShViews, kBrtEndUserShViews).inside,
            expected.inside);

  // .xlsb to .xlsx: the view element as Excel's .xlsx twin states it.
  auto as_xlsx = write_ooxml_with_result(from_xlsb.value().workbook);
  ASSERT_TRUE(static_cast<bool>(as_xlsx)) << as_xlsx.error().message;
  const std::string saved =
      ElementSpan(Text(Entry(as_xlsx.value().bytes, kSheetXml)), "<customSheetView ", "</customSheetView>");
  const std::string twin_view =
      ElementSpan(Text(Entry(twin_xlsx, kSheetXml)), "<customSheetView ", "</customSheetView>");
  ASSERT_FALSE(twin_view.empty());
  EXPECT_EQ(saved, twin_view);
  for (const char* child : children) {
    EXPECT_NE(saved.find(child), std::string::npos) << child;
  }
}

TEST(XlsbUserBookView, HeaderFooterViewCrossesBothWaysAsExcelSavesIt) {
  ExpectPrintViewCrossesBothWays("headerfooter_view",
                                 {"<headerFooter differentOddEven=\"1\" differentFirst=\"1\">",
                                  "<oddHeader>&amp;CODD-H</oddHeader>", "<firstFooter>&amp;CFIRST-F</firstFooter>"});
}

TEST(XlsbUserBookView, PageBreakViewCrossesBothWaysAsExcelSavesIt) {
  ExpectPrintViewCrossesBothWays(
      "pagebreaks_view",
      {"<rowBreaks count=\"2\" manualBreakCount=\"2\"><brk id=\"10\" max=\"16383\" man=\"1\"/><brk id=\"20\" "
       "max=\"16383\" man=\"1\"/></rowBreaks>",
       "<colBreaks count=\"2\" manualBreakCount=\"2\"><brk id=\"3\" max=\"1048575\" man=\"1\"/><brk id=\"6\" "
       "max=\"1048575\" man=\"1\"/></colBreaks>"});
}

TEST(XlsbUserBookView, ViewPageBreaksAreNotTheSheetsOnLoad) {
  auto xlsb = read_xlsb(SpanOf(CustomViewsFixture("pagebreaks_view.xlsb")));
  ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
  auto xlsx = read_ooxml(SpanOf(CustomViewsFixture("pagebreaks_view.xlsx")));
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  const SheetPrintSettings& got = xlsb.value().workbook.sheet(0).print_settings();
  const SheetPrintSettings& want = xlsx.value().workbook.sheet(0).print_settings();
  ASSERT_EQ(want.manual_row_breaks.size(), 1U);
  ASSERT_EQ(got.manual_row_breaks.size(), 1U);
  EXPECT_EQ(got.manual_row_breaks[0].id, 15U);
  ASSERT_EQ(got.manual_col_breaks.size(), 1U);
  EXPECT_EQ(got.manual_col_breaks[0].id, 4U);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
