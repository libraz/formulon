//
// The XLSB writer saves cross-workbook references as formulas: every
// external link becomes an `externalLink<N>.bin` part with its rels, the
// workbook stream lists the links as supporting books, and each XTI names
// the book it qualifies. Excel-written fixtures pin the record layout:
// saving one back must reproduce Excel's own supporting-book table, link
// parts and formula tokens byte for byte.

#include <algorithm>
#include <cstdint>
#include <cstdio>
#include <cstring>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "external_link.h"
#include "gtest/gtest.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/writer.h"
#include "miniz.h"
#include "sheet_passthrough.h"
#include "value.h"
#include "workbook.h"
#include "writer_test_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::uint16_t kBeginExternals = 353;
constexpr std::uint16_t kEndExternals = 354;
constexpr std::uint16_t kSupBookSrc = 355;
constexpr std::uint16_t kSupSelf = 357;
constexpr std::uint16_t kExternSheet = 362;

std::vector<std::uint8_t> ReadFileBytes(const std::string& path) {
  std::vector<std::uint8_t> out;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  if (size > 0) {
    out.resize(static_cast<std::size_t>(size));
    if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
      ADD_FAILURE() << "short read on fixture: " << path;
      out.clear();
    }
  }
  std::fclose(file);
  return out;
}

std::vector<std::uint8_t> FixtureBytes(const std::string& stem) {
  return ReadFileBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/" + stem + ".xlsb");
}

Workbook Load(const std::vector<std::uint8_t>& bytes) {
  auto result_or = read_xlsb(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(result_or)) << (result_or ? "" : result_or.error().message);
  if (!result_or) {
    return Workbook::create_empty();
  }
  return std::move(result_or.value().workbook);
}

std::vector<std::uint8_t> Save(const Workbook& wb) {
  auto result_or = write_xlsb_with_result(wb);
  EXPECT_TRUE(static_cast<bool>(result_or)) << (result_or ? "" : result_or.error().message);
  if (!result_or) {
    return {};
  }
  EXPECT_EQ(result_or.value().diagnostics.downgraded_formula_count, 0U);
  return std::move(result_or.value().bytes);
}

/// The bytes of package entry `name`, or an empty string when absent.
std::string Entry(const std::vector<std::uint8_t>& package, const char* name) {
  mz_zip_archive zip{};
  if (mz_zip_reader_init_mem(&zip, package.data(), package.size(), 0) == MZ_FALSE) {
    ADD_FAILURE() << "not a zip package";
    return {};
  }
  std::size_t size = 0;
  void* data = mz_zip_reader_extract_file_to_heap(&zip, name, &size, 0);
  std::string out;
  if (data != nullptr) {
    out.assign(static_cast<const char*>(data), size);
    mz_free(data);
  }
  mz_zip_reader_end(&zip);
  return out;
}

struct Rec {
  std::uint16_t type;
  std::string payload;
};

std::vector<Rec> Records(const std::string& bytes) {
  std::vector<Rec> out;
  ByteSpan cursor{reinterpret_cast<const std::uint8_t*>(bytes.data()), bytes.size()};
  while (cursor.size > 0U) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      ADD_FAILURE() << rec_or.error().message;
      break;
    }
    const XlsbRecord& rec = rec_or.value();
    out.push_back(Rec{rec.type, std::string(reinterpret_cast<const char*>(rec.payload.data), rec.payload.size)});
  }
  return out;
}

/// The workbook stream's supporting-book block, `BrtBeginExternals` through
/// `BrtEndExternals`.
std::vector<Rec> ExternalsBlock(const std::vector<std::uint8_t>& package) {
  std::vector<Rec> out;
  bool inside = false;
  for (Rec& rec : Records(Entry(package, "xl/workbook.bin"))) {
    inside = inside || rec.type == kBeginExternals;
    if (inside) {
      out.push_back(std::move(rec));
    }
    if (rec.type == kEndExternals) {
      break;
    }
  }
  return out;
}

std::vector<std::uint16_t> Types(const std::vector<Rec>& recs) {
  std::vector<std::uint16_t> out;
  for (const Rec& rec : recs) {
    out.push_back(rec.type);
  }
  return out;
}

struct Xti {
  std::uint32_t sup_book;
  std::int32_t first;
  std::int32_t last;
  bool operator==(const Xti& o) const { return sup_book == o.sup_book && first == o.first && last == o.last; }
};

std::uint32_t U32At(const std::string& p, std::size_t at) {
  std::uint32_t v = 0;
  std::memcpy(&v, p.data() + at, sizeof(v));
  return v;
}

std::vector<Xti> XtiTable(const std::vector<Rec>& block) {
  std::vector<Xti> out;
  for (const Rec& rec : block) {
    if (rec.type != kExternSheet) {
      continue;
    }
    const std::uint32_t count = U32At(rec.payload, 0);
    for (std::uint32_t i = 0; i < count; ++i) {
      const std::size_t at = 4U + i * 12U;
      out.push_back(Xti{U32At(rec.payload, at), static_cast<std::int32_t>(U32At(rec.payload, at + 4U)),
                        static_cast<std::int32_t>(U32At(rec.payload, at + 8U))});
    }
  }
  return out;
}

/// The XLWideString payload of a `BrtSupBookSrc`, narrowed to ASCII.
std::string SupBookRelId(const Rec& rec) {
  const std::uint32_t cch = U32At(rec.payload, 0);
  std::string out;
  for (std::uint32_t i = 0; i < cch; ++i) {
    out.push_back(rec.payload[4U + i * 2U]);
  }
  return out;
}

/// The `rgce` of every formula-cell record of a sheet part, in order.
std::vector<std::string> FormulaTokens(const std::string& sheet) {
  std::vector<std::string> out;
  for (const Rec& rec : Records(sheet)) {
    std::size_t value_bytes = 0;
    switch (rec.type) {
      case 8:  // BrtFmlaString
        value_bytes = 4U + 2U * U32At(rec.payload, 8);
        break;
      case 9:  // BrtFmlaNum
        value_bytes = 8U;
        break;
      case 10:  // BrtFmlaBool
      case 11:  // BrtFmlaError
        value_bytes = 1U;
        break;
      default:
        continue;
    }
    const std::size_t cce_at = 8U + value_bytes + 2U;
    out.push_back(rec.payload.substr(cce_at + 4U, U32At(rec.payload, cce_at)));
  }
  return out;
}

std::string FormulaAt(const Workbook& wb, std::size_t sheet, std::uint32_t row, std::uint32_t col) {
  const Cell* cell = wb.sheet(sheet).cell_at(row, col);
  return cell == nullptr ? std::string() : cell->formula_text;
}

Value Recalculated(Workbook& wb, std::size_t sheet, std::uint32_t row, std::uint32_t col) {
  auto recalc_or = wb.recalc(eval::default_registry());
  EXPECT_TRUE(static_cast<bool>(recalc_or)) << (recalc_or ? "" : recalc_or.error().message);
  const Cell* cell = wb.sheet(sheet).cell_at(row, col);
  return cell == nullptr ? Value::blank() : cell->cached_value;
}

Workbook Host(std::initializer_list<const char*> extra_sheets = {}) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Host");
  for (const char* name : extra_sheets) {
    wb.add_sheet(name);
  }
  return wb;
}

void SetFormula(Workbook& wb, std::uint32_t row, const std::string& formula) {
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, row, 0, formula))) << formula;
}

// ---------------------------------------------------------------------------
// Excel-written fixtures saved back
// ---------------------------------------------------------------------------

class XlsbWriterExternalLinkFixture : public ::testing::TestWithParam<const char*> {};

TEST_P(XlsbWriterExternalLinkFixture, SupBookTableMatchesExcel) {
  const std::vector<std::uint8_t> excel = FixtureBytes(GetParam());
  const std::vector<std::uint8_t> ours = Save(Load(excel));
  const std::vector<Rec> want = ExternalsBlock(excel);
  const std::vector<Rec> got = ExternalsBlock(ours);
  ASSERT_FALSE(want.empty());
  EXPECT_EQ(Types(got), Types(want));
  EXPECT_EQ(XtiTable(got), XtiTable(want));
}

TEST_P(XlsbWriterExternalLinkFixture, LinkPartsMatchExcelBytes) {
  const std::vector<std::uint8_t> excel = FixtureBytes(GetParam());
  const std::vector<std::uint8_t> ours = Save(Load(excel));
  for (const char* part : {"xl/externalLinks/externalLink1.bin", "xl/externalLinks/externalLink2.bin"}) {
    const std::string want = Entry(excel, part);
    EXPECT_EQ(Entry(ours, part), want) << part;
  }
}

TEST_P(XlsbWriterExternalLinkFixture, FormulaTokensMatchExcel) {
  const std::vector<std::uint8_t> excel = FixtureBytes(GetParam());
  const std::vector<std::uint8_t> ours = Save(Load(excel));
  const std::vector<std::string> want = FormulaTokens(Entry(excel, "xl/worksheets/sheet1.bin"));
  const std::vector<std::string> got = FormulaTokens(Entry(ours, "xl/worksheets/sheet1.bin"));
  ASSERT_FALSE(want.empty());
  EXPECT_EQ(got, want);
}

TEST_P(XlsbWriterExternalLinkFixture, SupBookSourcesNameTheirLinkParts) {
  const std::vector<std::uint8_t> ours = Save(Load(FixtureBytes(GetParam())));
  const std::string rels = Entry(ours, "xl/_rels/workbook.bin.rels");
  const std::string types = Entry(ours, "[Content_Types].xml");
  std::uint32_t link = 0;
  for (const Rec& rec : ExternalsBlock(ours)) {
    if (rec.type != kSupBookSrc) {
      continue;
    }
    ++link;
    const std::string part = "externalLinks/externalLink" + std::to_string(link) + ".bin";
    const std::string rel = "Id=\"" + SupBookRelId(rec) +
                            "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
                            "externalLink\" Target=\"" +
                            part + "\"";
    EXPECT_NE(rels.find(rel), std::string::npos) << rel << "\n" << rels;
    const std::string override_entry =
        "<Override PartName=\"/xl/" + part + "\" ContentType=\"application/vnd.ms-excel.externalLink\"/>";
    EXPECT_NE(types.find(override_entry), std::string::npos) << override_entry;
  }
  EXPECT_GE(link, 1U);
  // The source package's own link relationships are not carried twice.
  EXPECT_EQ(rels.find("externalLinks/externalLink" + std::to_string(link + 1U)), std::string::npos);
}

TEST_P(XlsbWriterExternalLinkFixture, LinkRelsKeepTheTargetAndTheAbsoluteUrl) {
  const std::vector<std::uint8_t> ours = Save(Load(FixtureBytes(GetParam())));
  const std::string rels = Entry(ours, "xl/externalLinks/_rels/externalLink1.bin.rels");
  const std::string path_type =
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
      "externalLinkPath\"";
  const std::string book = std::string(GetParam()) == "external_link_mixed" ? "ExtSource.xlsx" : "ExtSource3.xlsx";
  EXPECT_NE(rels.find("Id=\"rId1\" " + path_type + " Target=\"" + book + "\" TargetMode=\"External\""),
            std::string::npos)
      << rels;
  EXPECT_NE(rels.find("Id=\"rId2\" " + path_type + " Target=\"/Users/libraz/Documents/ext_link_probe/" + book +
                      "\" TargetMode=\"External\""),
            std::string::npos)
      << rels;
}

TEST_P(XlsbWriterExternalLinkFixture, ReloadKeepsSpellingAndValues) {
  Workbook source = Load(FixtureBytes(GetParam()));
  Workbook reloaded = Load(Save(source));
  ASSERT_EQ(reloaded.external_links().size(), source.external_links().size());
  for (std::size_t i = 0; i < source.external_links().size(); ++i) {
    EXPECT_EQ(reloaded.external_links()[i].target, source.external_links()[i].target);
    EXPECT_EQ(reloaded.external_links()[i].absolute_target, source.external_links()[i].absolute_target);
  }
  for (std::uint32_t row = 0; row < 6U; ++row) {
    const Cell* cell = source.sheet(0).cell_at(row, 0);
    if (cell == nullptr) {
      continue;
    }
    EXPECT_EQ(FormulaAt(reloaded, 0, row, 0), cell->formula_text) << row;
    const Value excel = cell->cached_value;
    const std::string excel_text = excel.is_text() ? std::string(excel.as_text()) : std::string();
    const Value ours = Recalculated(reloaded, 0, row, 0);
    ASSERT_EQ(ours.kind(), excel.kind()) << row;
    if (excel.is_number()) {
      EXPECT_DOUBLE_EQ(ours.as_number(), excel.as_number()) << row;
    } else if (excel.is_text()) {
      EXPECT_EQ(std::string(ours.as_text()), excel_text) << row;
    } else if (excel.is_error()) {
      EXPECT_EQ(ours.as_error(), excel.as_error()) << row;
    } else if (excel.is_boolean()) {
      EXPECT_EQ(ours.as_boolean(), excel.as_boolean()) << row;
    }
  }
}

std::string FixtureName(const ::testing::TestParamInfo<const char*>& info) {
  return info.param;
}

INSTANTIATE_TEST_SUITE_P(ExcelFixtures, XlsbWriterExternalLinkFixture,
                         ::testing::Values("external_link_mixed", "external_link_cell_kinds"), FixtureName);

// ---------------------------------------------------------------------------
// Links entered in this session
// ---------------------------------------------------------------------------

TEST(XlsbWriterExternalLink, AnEnteredLinkIsTheFirstSupportingBook) {
  Workbook wb = Host();
  SetFormula(wb, 0, "=[Book.xlsx]Sheet1!A1");
  const std::vector<std::uint8_t> bytes = Save(wb);
  const std::vector<Rec> block = ExternalsBlock(bytes);
  // No local qualified reference, so no self entry: index 0 is the link.
  EXPECT_EQ(Types(block), (std::vector<std::uint16_t>{kBeginExternals, kSupBookSrc, kExternSheet, kEndExternals}));
  EXPECT_EQ(XtiTable(block), (std::vector<Xti>{{0U, 0, 0}}));

  const std::string link = Entry(bytes, "xl/externalLinks/externalLink1.bin");
  ASSERT_FALSE(link.empty());
  const std::vector<Rec> recs = Records(link);
  ASSERT_GE(recs.size(), 3U);
  EXPECT_EQ(recs.front().type, 360);  // BrtBeginExternalBook
  EXPECT_EQ(recs.back().type, 588);   // BrtEndExternalBook
  const auto tabs = std::find_if(recs.begin(), recs.end(), [](const Rec& r) { return r.type == 359; });
  ASSERT_NE(tabs, recs.end());
  EXPECT_EQ(tabs->payload, std::string("\x01\x00\x00\x00\x06\x00\x00\x00S\0h\0e\0e\0t\0"
                                       "1\0",
                                       20));
  // A sheet created by entry has no cache.
  EXPECT_TRUE(std::none_of(recs.begin(), recs.end(), [](const Rec& r) { return r.type == 363; }));

  const std::string rels = Entry(bytes, "xl/externalLinks/_rels/externalLink1.bin.rels");
  EXPECT_NE(rels.find("Id=\"rId1\""), std::string::npos) << rels;
  EXPECT_NE(rels.find("Target=\"Book.xlsx\" TargetMode=\"External\""), std::string::npos) << rels;
  EXPECT_EQ(rels.find("rId2"), std::string::npos) << rels;
}

TEST(XlsbWriterExternalLink, LocalAndExternalXtisNameTheirOwnBooks) {
  Workbook wb = Host({"Other"});
  SetFormula(wb, 0, "=Other!A1");
  SetFormula(wb, 1, "=[Book.xlsx]Sheet1!B2");
  SetFormula(wb, 2, "=SUM([Book.xlsx]Sheet2!A1:B3)");
  const std::vector<Rec> block = ExternalsBlock(Save(wb));
  EXPECT_EQ(Types(block),
            (std::vector<std::uint16_t>{kBeginExternals, kSupSelf, kSupBookSrc, kExternSheet, kEndExternals}));
  // The self entry comes first, so the link is supporting book 1.
  EXPECT_EQ(XtiTable(block), (std::vector<Xti>{{0U, 1, 1}, {1U, 0, 0}, {1U, 1, 1}}));
}

TEST(XlsbWriterExternalLink, AnEnteredLinkSurvivesReloadAndReadsRef) {
  Workbook wb = Host();
  SetFormula(wb, 0, "=[Book.xlsx]Sheet1!A1");
  SetFormula(wb, 1, "=SUM([Book.xlsx]Sheet1!A1:B2)");
  SetFormula(wb, 2, "=SUM([Book.xlsx]Sheet1!C:C)");
  SetFormula(wb, 3, "='/Users/x/[Other Book.xlsx]Data'!$B$2");
  // Entered in the formula bar's spelling, which quotes a 3-D span.
  SetFormula(wb, 4, "=SUM('[Book.xlsx]Sheet1:Sheet2'!A1:B2)");
  SetFormula(wb, 5, "=SUM([Book.xlsx]Sheet1!1:2)");
  Workbook reloaded = Load(Save(wb));
  ASSERT_EQ(reloaded.external_links().size(), 2U);
  EXPECT_EQ(reloaded.external_links()[0].target, "Book.xlsx");
  EXPECT_EQ(reloaded.external_links()[1].target, "/Users/x/Other Book.xlsx");
  for (std::uint32_t row = 0; row < 6U; ++row) {
    EXPECT_EQ(FormulaAt(reloaded, 0, row, 0), FormulaAt(wb, 0, row, 0)) << row;
  }
  const Value value = Recalculated(reloaded, 0, 0, 0);
  ASSERT_TRUE(value.is_error());
  EXPECT_EQ(value.as_error(), ErrorCode::Ref);
}

TEST(XlsbWriterExternalLink, EveryRecordIsWrittenInIndexOrderWithoutGaps) {
  Workbook wb = Host();
  std::vector<ExternalLinkRecord> links(2);
  links[0].index = 2;
  links[0].target = "First.xlsx";
  links[0].kind = ExternalLinkRecord::Kind::kExternalBook;
  links[0].book.sheet_names = {"S"};
  links[0].book.sheet_data = {true};
  links[1].index = 5;
  links[1].target = "Second.xlsx";
  links[1].kind = ExternalLinkRecord::Kind::kExternalBook;
  links[1].book.sheet_names = {"T"};
  links[1].book.sheet_data = {true};
  wb.set_external_links(std::move(links));
  // Only the second link is referenced; the first is written all the same.
  SetFormula(wb, 0, "=[Second.xlsx]T!A1");
  const std::vector<std::uint8_t> bytes = Save(wb);
  EXPECT_EQ(XtiTable(ExternalsBlock(bytes)), (std::vector<Xti>{{1U, 0, 0}}));
  Workbook reloaded = Load(bytes);
  ASSERT_EQ(reloaded.external_links().size(), 2U);
  EXPECT_EQ(reloaded.external_links()[0].target, "First.xlsx");
  EXPECT_EQ(reloaded.external_links()[1].target, "Second.xlsx");
  EXPECT_EQ(FormulaAt(reloaded, 0, 0, 0), "=[Second.xlsx]T!A1");
  const Value value = Recalculated(reloaded, 0, 0, 0);
  ASSERT_TRUE(value.is_number());
  EXPECT_DOUBLE_EQ(value.as_number(), 0.0);
}

TEST(XlsbWriterExternalLink, AnOleLinkIsNotASupportingBook) {
  Workbook wb = Host();
  std::vector<ExternalLinkRecord> links(2);
  links[0].index = 1;
  links[0].target = "Embedded.xlsx";
  links[0].kind = ExternalLinkRecord::Kind::kOleLink;
  links[1].index = 2;
  links[1].target = "Book.xlsx";
  links[1].kind = ExternalLinkRecord::Kind::kExternalBook;
  links[1].book.sheet_names = {"S"};
  wb.set_external_links(std::move(links));
  SetFormula(wb, 0, "=[Book.xlsx]S!A1");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 1, 0, "=[Embedded.xlsx]S!A1")));
  auto result_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  // The formula naming the OLE link keeps its cached value.
  EXPECT_EQ(result_or.value().diagnostics.downgraded_formula_count, 1U);
  const std::vector<std::uint8_t>& bytes = result_or.value().bytes;
  const std::vector<Rec> block = ExternalsBlock(bytes);
  EXPECT_EQ(Types(block), (std::vector<std::uint16_t>{kBeginExternals, kSupBookSrc, kExternSheet, kEndExternals}));
  EXPECT_EQ(XtiTable(block), (std::vector<Xti>{{0U, 0, 0}}));
  EXPECT_TRUE(Entry(bytes, "xl/externalLinks/externalLink2.bin").empty());
  Workbook reloaded = Load(bytes);
  ASSERT_EQ(reloaded.external_links().size(), 1U);
  EXPECT_EQ(reloaded.external_links()[0].target, "Book.xlsx");
  EXPECT_EQ(FormulaAt(reloaded, 0, 0, 0), "=[Book.xlsx]S!A1");
}

TEST(XlsbWriterExternalLink, AnExternalSpanNamesBothSheetsOfItsBook) {
  Workbook wb = Host();
  SetFormula(wb, 0, "=SUM([Book.xlsx]S1:S2!A1)");
  EXPECT_EQ(XtiTable(ExternalsBlock(Save(wb))), (std::vector<Xti>{{0U, 0, 1}}));
}

TEST(XlsbWriterExternalLink, ARetainedTailKeepsItsExternalEntries) {
  Workbook wb = Load(FixtureBytes("external_link_mixed"));
  // A retained future-record block may name the source table's entries by
  // index, so they lead the saved table, re-pointed at the saved links.
  XlsbSheetTail tail;
  emit_record(tail.before_merges, 35, std::vector<std::uint8_t>{0x01, 0x00, 0x06, 0x11, 0x00, 0x80});
  emit_record(tail.before_merges, 36, ByteSpan{});
  XlsbExternSheetEntry cell;
  cell.first = "data";
  cell.last = "data";
  cell.unresolved = true;
  cell.external_book = 1;
  XlsbExternSheetEntry name;
  name.unresolved = true;
  name.external_book = 2;
  tail.extern_sheets = {cell, name};
  wb.sheet(0).set_xlsb_tail(tail);
  const std::vector<Xti> xti = XtiTable(ExternalsBlock(Save(wb)));
  ASSERT_GE(xti.size(), 2U);
  EXPECT_EQ(xti[0], (Xti{1U, 0, 0}));
  EXPECT_EQ(xti[1], (Xti{2U, -2, -2}));
}

TEST(XlsbWriterExternalLink, ARetainedEntryOfAnUnknownLinkFailsTheSave) {
  Workbook wb = Load(FixtureBytes("external_link_mixed"));
  XlsbSheetTail tail;
  emit_record(tail.before_merges, 35, std::vector<std::uint8_t>{0x01, 0x00, 0x06, 0x11, 0x00, 0x80});
  emit_record(tail.before_merges, 36, ByteSpan{});
  XlsbExternSheetEntry entry;
  entry.first = "Nope";
  entry.last = "Nope";
  entry.unresolved = true;
  entry.external_book = 1;
  tail.extern_sheets = {entry};
  wb.sheet(0).set_xlsb_tail(tail);
  auto result_or = write_xlsb_with_result(wb);
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoXlsbRetainedPartStale);
}

// ---------------------------------------------------------------------------
// Names the book does not declare, and an extensionless book's names
// ---------------------------------------------------------------------------

ExternalBookName CachedCellName(const char* name, std::uint32_t row) {
  ExternalBookName n;
  n.name = name;
  n.row = row;
  n.row_end = row;
  n.resolvable = true;
  return n;
}

/// A host sheet with one link whose first sheet caches `A1 = value`.
Workbook HostWithLink(const char* target, std::vector<ExternalBookName> names, double value) {
  Workbook wb = Host();
  std::vector<ExternalLinkRecord> links(1);
  links[0].index = 1;
  links[0].target = target;
  links[0].kind = ExternalLinkRecord::Kind::kExternalBook;
  links[0].book.sheet_names = {"Data"};
  links[0].book.sheet_data = {true};
  links[0].book.names = std::move(names);
  links[0].book.cells[ExternalBook::cell_key(0, 0, 0)].value = Value::number(value);
  wb.set_external_links(std::move(links));
  return wb;
}

TEST(XlsbWriterExternalLink, AbsentNameIsWrittenWithoutABodyAndKeepsLaterNamesInPlace) {
  ExternalBookName gone = CachedCellName("Gone", 0);
  gone.exists = false;
  Workbook wb = HostWithLink("Book.xlsx", {gone, CachedCellName("After", 0)}, 10.0);
  SetFormula(wb, 0, "=Book.xlsx!Gone");
  SetFormula(wb, 1, "=Book.xlsx!After");
  SetFormula(wb, 2, "=Book.xlsx!NoSuch");
  const std::vector<std::uint8_t> bytes = Save(wb);

  // Excel's payload for a name with no body: cce = 0, nothing after it.
  const std::vector<Rec> recs = Records(Entry(bytes, "xl/externalLinks/externalLink1.bin"));
  std::vector<std::string> bodies;
  for (const Rec& rec : recs) {
    if (rec.type == 585) {
      bodies.push_back(rec.payload);
    }
  }
  ASSERT_EQ(bodies.size(), 3U);
  EXPECT_EQ(bodies[0], std::string("\0\0\0\0", 4));
  EXPECT_EQ(bodies[1].size(), 4U + 9U);
  EXPECT_EQ(bodies[2], std::string("\0\0\0\0", 4));

  Workbook reloaded = Load(bytes);
  const ExternalBook& book = reloaded.external_links()[0].book;
  ASSERT_EQ(book.names.size(), 3U);
  EXPECT_FALSE(book.names[0].exists);
  EXPECT_TRUE(book.names[1].exists);
  EXPECT_TRUE(book.names[1].resolvable);
  EXPECT_FALSE(book.names[2].exists);
  EXPECT_EQ(book.names[2].name, "NoSuch");
  for (const std::uint32_t row : {0U, 2U}) {
    const Value missing = Recalculated(reloaded, 0, row, 0);
    ASSERT_TRUE(missing.is_error()) << row;
    EXPECT_EQ(missing.as_error(), ErrorCode::Name) << row;
  }
  const Value after = Recalculated(reloaded, 0, 1, 0);
  ASSERT_TRUE(after.is_number());
  EXPECT_DOUBLE_EQ(after.as_number(), 10.0);
}

TEST(XlsbWriterExternalLink, UnresolvableNameKeepsItsRefErrorBody) {
  ExternalBookName constant;
  constant.name = "Rate";
  Workbook wb = HostWithLink("Book.xlsx", {constant}, 1.0);
  SetFormula(wb, 0, "=Book.xlsx!Rate");
  const std::vector<Rec> recs = Records(Entry(Save(wb), "xl/externalLinks/externalLink1.bin"));
  const auto fmla = std::find_if(recs.begin(), recs.end(), [](const Rec& r) { return r.type == 585; });
  ASSERT_NE(fmla, recs.end());
  EXPECT_EQ(fmla->payload, std::string("\x02\0\0\0\x1c\x17", 6));
}

TEST(XlsbWriterExternalLink, ExtensionlessBookScopeNameIsWrittenAsPtgNameX) {
  Workbook wb = HostWithLink("Src2", {CachedCellName("Other", 1), CachedCellName("Total", 0)}, 7.0);
  SetFormula(wb, 0, "=SUM('Src2'!Total)");
  const std::vector<std::uint8_t> bytes = Save(wb);

  const std::vector<std::string> tokens = FormulaTokens(Entry(bytes, "xl/worksheets/sheet1.bin"));
  ASSERT_EQ(tokens.size(), 1U);
  ASSERT_GE(tokens[0].size(), 7U);
  EXPECT_EQ(static_cast<unsigned>(static_cast<std::uint8_t>(tokens[0][0]) & 0x1FU), 0x19U);  // PtgNameX
  EXPECT_EQ(U32At(tokens[0], 3), 2U);  // Total is the second name of the link
  const std::vector<Xti> xti = XtiTable(ExternalsBlock(bytes));
  ASSERT_EQ(xti.size(), 1U);
  EXPECT_EQ(xti[0].sup_book, 0U);

  Workbook reloaded = Load(bytes);
  EXPECT_EQ(FormulaAt(reloaded, 0, 0, 0), "=SUM('Src2'!Total)");
  const Value value = Recalculated(reloaded, 0, 0, 0);
  ASSERT_TRUE(value.is_number());
  EXPECT_DOUBLE_EQ(value.as_number(), 7.0);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
