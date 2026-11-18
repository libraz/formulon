// Excel's save drops the trailing external links nothing references, in both formats; a link that is
// referenced, or that precedes a referenced one, stays. The model keeps every link.

#include <cstdint>
#include <cstdio>
#include <string>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "external_link.h"
#include "gtest/gtest.h"
#include "io/external_link_use.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/writer.h"
#include "miniz.h"
#include "passthrough_part.h"
#include "sheet.h"
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

constexpr std::uint16_t kSupBookSrc = 355;

std::vector<std::uint8_t> FixtureBytes(const std::string& name) {
  std::vector<std::uint8_t> out;
  FILE* file = std::fopen((std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name).c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << name;
    return out;
  }
  for (int c = std::fgetc(file); c != EOF; c = std::fgetc(file)) {
    out.push_back(static_cast<std::uint8_t>(c));
  }
  std::fclose(file);
  return out;
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

std::size_t Count(const std::string& text, const std::string& needle) {
  std::size_t n = 0;
  for (std::size_t at = text.find(needle); at != std::string::npos; at = text.find(needle, at + 1U)) {
    ++n;
  }
  return n;
}

std::size_t SupBookCount(const std::vector<std::uint8_t>& xlsb) {
  const std::string stream = Entry(xlsb, "xl/workbook.bin");
  ByteSpan cursor{reinterpret_cast<const std::uint8_t*>(stream.data()), stream.size()};
  std::size_t n = 0;
  while (cursor.size > 0U) {
    auto rec = read_record(cursor);
    if (!rec) {
      ADD_FAILURE() << rec.error().message;
      break;
    }
    n += rec.value().type == kSupBookSrc ? 1U : 0U;
  }
  return n;
}

Workbook LoadAuthored() {
  auto read_or = read_ooxml(SpanOf(FixtureBytes("extlink_unreferenced.xlsx")));
  EXPECT_TRUE(static_cast<bool>(read_or));
  return read_or ? std::move(read_or.value().workbook) : Workbook::create_empty();
}

std::vector<std::uint8_t> SaveXlsx(const Workbook& wb) {
  auto saved = write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value()) : std::vector<std::uint8_t>{};
}

std::vector<std::uint8_t> SaveXlsb(const Workbook& wb) {
  auto saved = write_xlsb(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value()) : std::vector<std::uint8_t>{};
}

// `[1]Alpha!A1` at C1 after a full recalc.
Value AlphaCell(Workbook& wb) {
  auto recalc_or = wb.recalc_full(eval::default_registry());
  EXPECT_TRUE(static_cast<bool>(recalc_or));
  const Cell* cell = wb.sheet(0).cell_at(0U, 2U);
  return cell == nullptr ? Value::blank() : cell->cached_value;
}

TEST(ExternalLinkUnreferenced, XlsxSaveDropsTheTrailingLinkAsExcelDoes) {
  const Workbook wb = LoadAuthored();
  ASSERT_EQ(wb.external_links().size(), 2U);
  const std::vector<std::uint8_t> saved = SaveXlsx(wb);
  const std::vector<std::uint8_t> excel = FixtureBytes("extlink_unreferenced_excel.xlsx");
  const std::string book = Entry(saved, "xl/workbook.xml");
  EXPECT_EQ(Count(book, "<externalReference "), 1U) << book;
  EXPECT_EQ(Count(book, "<externalReference "), Count(Entry(excel, "xl/workbook.xml"), "<externalReference "));
  EXPECT_FALSE(Entry(saved, "xl/externalLinks/externalLink1.xml").empty());
  EXPECT_TRUE(Entry(saved, "xl/externalLinks/externalLink2.xml").empty());
  EXPECT_TRUE(Entry(saved, "xl/externalLinks/_rels/externalLink2.xml.rels").empty());
  EXPECT_EQ(Count(Entry(saved, "[Content_Types].xml"), "externalLink2"), 0U);
  EXPECT_EQ(Count(Entry(saved, "xl/_rels/workbook.xml.rels"), "externalLink2"), 0U);
  const char* const kRels = "xl/externalLinks/_rels/externalLink1.xml.rels";
  EXPECT_NE(Entry(excel, kRels).find("Target=\"ShortBook1.xlsx\""), std::string::npos);
  EXPECT_NE(Entry(saved, kRels).find("Target=\"ShortBook1.xlsx\""), std::string::npos);
  EXPECT_EQ(wb.external_links().size(), 2U);
}

TEST(ExternalLinkUnreferenced, XlsbSaveDropsTheTrailingLinkAsExcelDoes) {
  const Workbook wb = LoadAuthored();
  const std::vector<std::uint8_t> saved = SaveXlsb(wb);
  EXPECT_EQ(SupBookCount(saved), 1U);
  EXPECT_EQ(SupBookCount(saved), SupBookCount(FixtureBytes("extlink_unreferenced_excel.xlsb")));
  EXPECT_FALSE(Entry(saved, "xl/externalLinks/externalLink1.bin").empty());
  EXPECT_TRUE(Entry(saved, "xl/externalLinks/externalLink2.bin").empty());
  EXPECT_EQ(wb.external_links().size(), 2U);
}

TEST(ExternalLinkUnreferenced, TheReferencedLinkStillEvaluatesAfterAReload) {
  Workbook wb = LoadAuthored();
  const Value expected = AlphaCell(wb);
  auto via_xlsx = read_ooxml(SpanOf(SaveXlsx(wb)));
  ASSERT_TRUE(static_cast<bool>(via_xlsx));
  EXPECT_EQ(AlphaCell(via_xlsx.value().workbook), expected);
  auto via_xlsb = read_xlsb(SpanOf(SaveXlsb(wb)));
  ASSERT_TRUE(static_cast<bool>(via_xlsb));
  EXPECT_EQ(AlphaCell(via_xlsb.value().workbook), expected);
}

ExternalLinkRecord NewLink(std::uint32_t index) {
  ExternalLinkRecord rec;
  rec.index = index;
  rec.target = "Book" + std::to_string(index) + ".xlsx";
  rec.kind = ExternalLinkRecord::Kind::kExternalBook;
  rec.book.sheet_names = {"Sheet1"};
  return rec;
}

std::string Ref(std::uint32_t index) {
  return "[Book" + std::to_string(index) + ".xlsx]Sheet1!A1";
}

Workbook HostBook(std::uint32_t link_count) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Host");
  std::vector<ExternalLinkRecord> links;
  for (std::uint32_t i = 1; i <= link_count; ++i) {
    links.push_back(NewLink(i));
  }
  wb.set_external_links(std::move(links));
  return wb;
}

void AddCellRef(Workbook& wb, std::uint32_t index) {
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, index, 0, "=" + Ref(index))));
}

void ExpectBothFormatsWrite(const Workbook& wb, std::size_t links) {
  const std::string book = Entry(SaveXlsx(wb), "xl/workbook.xml");
  EXPECT_EQ(Count(book, "<externalReference "), links) << book;
  EXPECT_EQ(SupBookCount(SaveXlsb(wb)), links);
}

TEST(ExternalLinkUnreferenced, AReferenceFromAnyHolderKeepsTheLinkBeforeIt) {
  Workbook cell = HostBook(2);
  AddCellRef(cell, 2);
  ExpectBothFormatsWrite(cell, 2U);

  Workbook name = HostBook(2);
  ASSERT_TRUE(static_cast<bool>(name.set_defined_name("Ext", "=" + Ref(2))));
  ExpectBothFormatsWrite(name, 2U);

  Workbook rule_book = HostBook(2);
  cf::ConditionalFormat format;
  format.sqref.push_back(cf::CFCellRange{{0, 0}, {0, 0}});
  cf::CFRule rule;
  rule.type = cf::RuleType::Expression;
  rule.formula1 = Ref(2) + ">0";
  format.rules.push_back(rule);
  rule_book.sheet(0).mutable_conditional_formats().push_back(format);
  ExpectBothFormatsWrite(rule_book, 2U);

  Workbook validation_book = HostBook(2);
  DataValidation validation;
  validation.ranges.push_back(MergeRange{0, 0, 0, 0});
  validation.type = 1;
  validation.formula1 = Ref(2);
  validation_book.sheet(0).mutable_validations().push_back(validation);
  ExpectBothFormatsWrite(validation_book, 2U);

  Workbook chart = HostBook(2);
  const std::string xml = "<c:chartSpace><c:f>[2]Sheet1!$A$1</c:f></c:chartSpace>";
  chart.set_passthrough_parts(
      {PassthroughPart("xl/charts/chart1.xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
                       std::vector<std::uint8_t>(xml.begin(), xml.end()))});
  ExpectBothFormatsWrite(chart, 2U);
}

TEST(ExternalLinkUnreferenced, AnUnreferencedLinkBeforeAReferencedOneStays) {
  Workbook wb = HostBook(3);
  AddCellRef(wb, 1);
  AddCellRef(wb, 3);
  ExpectBothFormatsWrite(wb, 3U);
}

TEST(ExternalLinkUnreferenced, ATrailingOleLinkIsKept) {
  Workbook wb = HostBook(2);
  AddCellRef(wb, 1);
  EXPECT_EQ(last_written_external_link(wb), 1U);
  std::vector<ExternalLinkRecord> links = wb.external_links();
  links[1].kind = ExternalLinkRecord::Kind::kOleLink;
  wb.set_external_links(std::move(links));
  EXPECT_EQ(last_written_external_link(wb), 2U);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
