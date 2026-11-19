// Excel's save drops every external link nothing references and renumbers the rest densely, in both formats;
// a link a verbatim holder names, and every link before it, stays. The model keeps every link.

#include <algorithm>
#include <cstdint>
#include <cstdio>
#include <string>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "defined_name.h"
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

constexpr std::uint16_t kBeginExternals = 353;
constexpr std::uint16_t kEndExternals = 354;
constexpr std::uint16_t kExternSheet = 362;
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

TEST(ExternalLinkUnreferenced, EachHolderKeepsTheLinkItNames) {
  Workbook cell = HostBook(2);
  AddCellRef(cell, 2);
  ExpectBothFormatsWrite(cell, 1U);

  Workbook name = HostBook(2);
  ASSERT_TRUE(static_cast<bool>(name.set_defined_name("Ext", "=" + Ref(2))));
  ExpectBothFormatsWrite(name, 1U);

  Workbook rule_book = HostBook(2);
  cf::ConditionalFormat format;
  format.sqref.push_back(cf::CFCellRange{{0, 0}, {0, 0}});
  cf::CFRule rule;
  rule.type = cf::RuleType::Expression;
  rule.formula1 = Ref(2) + ">0";
  format.rules.push_back(rule);
  rule_book.sheet(0).mutable_conditional_formats().push_back(format);
  ExpectBothFormatsWrite(rule_book, 1U);

  Workbook validation_book = HostBook(2);
  DataValidation validation;
  validation.ranges.push_back(MergeRange{0, 0, 0, 0});
  validation.type = 1;
  validation.formula1 = Ref(2);
  validation_book.sheet(0).mutable_validations().push_back(validation);
  ExpectBothFormatsWrite(validation_book, 1U);

  // A chart writes `[2]` as stored, so the link before it cannot be dropped.
  Workbook chart = HostBook(2);
  const std::string xml = "<c:chartSpace><c:f>[2]Sheet1!$A$1</c:f></c:chartSpace>";
  chart.set_passthrough_parts(
      {PassthroughPart("xl/charts/chart1.xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
                       std::vector<std::uint8_t>(xml.begin(), xml.end()))});
  ExpectBothFormatsWrite(chart, 2U);
}

TEST(ExternalLinkUnreferenced, AnUnreferencedLinkBetweenReferencedOnesIsDroppedAndTheRestRenumbered) {
  Workbook wb = HostBook(3);
  AddCellRef(wb, 1);
  AddCellRef(wb, 3);
  ExpectBothFormatsWrite(wb, 2U);
  EXPECT_EQ(written_external_link_indices(wb), (std::vector<std::uint32_t>{1U, 3U}));
  const std::string sheet = Entry(SaveXlsx(wb), "xl/worksheets/sheet1.xml");
  EXPECT_NE(sheet.find("<f>[2]Sheet1!A1</f>"), std::string::npos) << sheet;
  EXPECT_EQ(sheet.find("[3]"), std::string::npos) << sheet;
}

TEST(ExternalLinkUnreferenced, AnOleLinkKeepsEveryLinkBeforeIt) {
  Workbook wb = HostBook(3);
  AddCellRef(wb, 3);
  EXPECT_EQ(written_external_link_indices(wb), (std::vector<std::uint32_t>{3U}));
  std::vector<ExternalLinkRecord> links = wb.external_links();
  links[1].kind = ExternalLinkRecord::Kind::kOleLink;
  wb.set_external_links(std::move(links));
  EXPECT_EQ(written_external_link_indices(wb), (std::vector<std::uint32_t>{1U, 2U, 3U}));
}

// Excel's saves of one book (sources S1..S3, cells A1..A3 each naming one) after each edit. The nameonly save is
// the input: A1 names S1, A3 names S3, and the name KeepTwo names S2.
Workbook LoadNameOnly() {
  auto read_or = read_ooxml(SpanOf(FixtureBytes("extlink_gap_nameonly_excel.xlsx")));
  EXPECT_TRUE(static_cast<bool>(read_or));
  return read_or ? std::move(read_or.value().workbook) : Workbook::create_empty();
}

// `book`'s defined names with `name` removed.
void RemoveName(Workbook& book, const std::string& name) {
  std::vector<DefinedName> names = book.defined_names();
  names.erase(std::remove_if(names.begin(), names.end(), [&name](const DefinedName& n) { return n.name == name; }),
              names.end());
  book.set_defined_names(std::move(names));
}

// The externals block of `xl/workbook.bin` (BrtBeginExternals .. BrtEndExternals), as record type and payload.
std::vector<std::pair<std::uint16_t, std::string>> ExternalsBlock(const std::vector<std::uint8_t>& xlsb) {
  const std::string stream = Entry(xlsb, "xl/workbook.bin");
  ByteSpan cursor{reinterpret_cast<const std::uint8_t*>(stream.data()), stream.size()};
  std::vector<std::pair<std::uint16_t, std::string>> out;
  bool inside = false;
  while (cursor.size > 0U) {
    auto rec = read_record(cursor);
    if (!rec) {
      ADD_FAILURE() << rec.error().message;
      break;
    }
    inside = inside || rec.value().type == kBeginExternals;
    if (inside) {
      out.emplace_back(rec.value().type,
                       std::string(reinterpret_cast<const char*>(rec.value().payload.data), rec.value().payload.size));
    }
    if (rec.value().type == kEndExternals) {
      break;
    }
  }
  return out;
}

// The `<f>` texts, defined names and link targets of an xlsx package, the parts a renumbering moves together.
std::string XlsxLinkView(const std::vector<std::uint8_t>& xlsx) {
  std::string out;
  const auto collect = [&out](const std::string& text, const std::string& open, const std::string& close) {
    for (std::size_t at = text.find(open); at != std::string::npos; at = text.find(open, at + 1U)) {
      const std::size_t end = text.find(close, at);
      out.append(text, at, end == std::string::npos ? std::string::npos : end + close.size() - at);
      out.push_back('\n');
    }
  };
  collect(Entry(xlsx, "xl/worksheets/sheet1.xml"), "<f>", "</f>");
  collect(Entry(xlsx, "xl/workbook.xml"), "<definedName ", "</definedName>");
  for (const char* k : {"1", "2", "3", "4"}) {
    const std::string rels =
        Entry(xlsx, (std::string("xl/externalLinks/_rels/externalLink") + k + ".xml.rels").c_str());
    if (rels.empty()) {
      continue;
    }
    out.append(std::string("part ") + k + "\n");
    collect(rels, "Target=\"", "\"");
  }
  return out;
}

// `block` with its BrtExternSheet entries sorted: the renumbering decides which SupBook each entry names, not the
// order the entries are listed in.
std::vector<std::pair<std::uint16_t, std::string>> WithSortedXtis(
    std::vector<std::pair<std::uint16_t, std::string>> block) {
  constexpr std::size_t kXtiBytes = 12U;
  for (auto& [type, payload] : block) {
    if (type != kExternSheet || payload.size() < 4U) {
      continue;
    }
    std::vector<std::string> entries;
    for (std::size_t at = 4U; at + kXtiBytes <= payload.size(); at += kXtiBytes) {
      entries.push_back(payload.substr(at, kXtiBytes));
    }
    std::sort(entries.begin(), entries.end());
    std::string sorted = payload.substr(0U, 4U);
    for (const std::string& entry : entries) {
      sorted += entry;
    }
    payload = sorted;
  }
  return block;
}

void ExpectSavesLike(const Workbook& wb, const std::string& excel_stem) {
  const std::vector<std::uint8_t> excel_xlsx = FixtureBytes(excel_stem + ".xlsx");
  EXPECT_EQ(XlsxLinkView(SaveXlsx(wb)), XlsxLinkView(excel_xlsx)) << excel_stem;
  EXPECT_EQ(WithSortedXtis(ExternalsBlock(SaveXlsb(wb))),
            WithSortedXtis(ExternalsBlock(FixtureBytes(excel_stem + ".xlsb"))))
      << excel_stem;
}

// `view` without the link parts' relationship targets: Excel's save of a reopened book rewrites their form and count.
std::string WithoutTargets(const std::string& view) {
  std::string out;
  std::size_t at = 0;
  while (at < view.size()) {
    std::size_t end = view.find('\n', at);
    end = end == std::string::npos ? view.size() : end + 1U;
    if (view.compare(at, 8U, "Target=\"") != 0) {
      out.append(view, at, end - at);
    }
    at = end;
  }
  return out;
}

// The link only a name references stays, and the reopen numbers it first, as Excel's save of the reopened book does.
TEST(ExternalLinkUnreferenced, ALinkOnlyANameReferencesIsKept) {
  const Workbook wb = LoadNameOnly();
  const std::string reopened = "external_link_renumber/name_link_reopened";
  EXPECT_EQ(WithoutTargets(XlsxLinkView(SaveXlsx(wb))), WithoutTargets(XlsxLinkView(FixtureBytes(reopened + ".xlsx"))));
  EXPECT_EQ(WithSortedXtis(ExternalsBlock(SaveXlsb(wb))),
            WithSortedXtis(ExternalsBlock(FixtureBytes(reopened + ".xlsb"))));
}

TEST(ExternalLinkUnreferenced, AMiddleLinkNothingReferencesIsDroppedAsExcelDoes) {
  Workbook wb = LoadNameOnly();
  RemoveName(wb, "KeepTwo");
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("N", "=[S3.xlsx]S!$A$1")));
  ExpectSavesLike(wb, "extlink_gap_mid_excel");
  EXPECT_EQ(wb.external_links().size(), 3U);
}

TEST(ExternalLinkUnreferenced, ALeadingLinkNothingReferencesIsDroppedAsExcelDoes) {
  Workbook wb = LoadNameOnly();
  RemoveName(wb, "KeepTwo");
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("N", "=[S3.xlsx]S!$A$1")));
  wb.sheet(0).erase_cell(0U, 0U);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 1U, 0U, "=[S2.xlsx]S!A1")));
  ExpectSavesLike(wb, "extlink_gap_lead_excel");
}

TEST(ExternalLinkUnreferenced, ASaveLeavesTheModelSoALaterSaveWritesTheLinkAgain) {
  Workbook wb = LoadNameOnly();
  RemoveName(wb, "KeepTwo");
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("N", "=[S3.xlsx]S!$A$1")));
  ExpectBothFormatsWrite(wb, 2U);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 1U, 0U, "=[S2.xlsx]S!A1")));
  ExpectBothFormatsWrite(wb, 3U);
  const std::string sheet = Entry(SaveXlsx(wb), "xl/worksheets/sheet1.xml");
  EXPECT_NE(sheet.find("<f>[2]S!A1</f>"), std::string::npos) << sheet;
  EXPECT_NE(sheet.find("<f>[3]S!A1</f>"), std::string::npos) << sheet;

  // A dropped link's cached values survive in the model and read back after a reload of the later save.
  auto reread = read_ooxml(SpanOf(SaveXlsx(wb)));
  ASSERT_TRUE(static_cast<bool>(reread));
  Workbook& back = reread.value().workbook;
  ASSERT_TRUE(static_cast<bool>(back.recalc_full(eval::default_registry())));
  const Cell* a2 = back.sheet(0).cell_at(1U, 0U);
  ASSERT_NE(a2, nullptr);
  EXPECT_EQ(a2->cached_value, Value::number(22.0));
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
