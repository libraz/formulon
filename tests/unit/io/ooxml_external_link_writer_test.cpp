//
// Saving external links to .xlsx: a link whose loaded body still matches
// the model is written back verbatim, every other link gets a body
// generated from the model, and each written link carries all of its
// package faces (`<externalReference>`, workbook relationship, content
// type Override, its own rels with both the target and the absolute
// path). Formulas store each book as the 1-based position of its link in
// `<externalReferences>`.

#include <cmath>
#include <cstdint>
#include <limits>
#include <regex>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "defined_name.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "external_book.h"
#include "external_link.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "miniz.h"
#include "passthrough_part.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

constexpr std::string_view kCtExternalLink =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";

std::vector<std::uint8_t> Save(const Workbook& wb) {
  auto saved = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value()) : std::vector<std::uint8_t>{};
}

std::string Part(const std::vector<std::uint8_t>& pkg, std::string_view name) {
  std::string body;
  EXPECT_TRUE(test::extract_part(test::span_of(pkg), name, &body));
  return body;
}

bool HasPart(const std::vector<std::uint8_t>& pkg, std::string_view name) {
  std::string body;
  return static_cast<bool>(test::extract_part(test::span_of(pkg), name, &body));
}

Workbook Load(const std::vector<std::uint8_t>& pkg) {
  auto loaded = io::read_ooxml(test::span_of(pkg));
  EXPECT_TRUE(static_cast<bool>(loaded)) << (loaded ? "" : loaded.error().message);
  return loaded ? std::move(loaded.value().workbook) : Workbook::create_empty();
}

std::vector<std::uint8_t> FixtureBytes(const char* name) {
  return test::read_file_bytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name);
}

std::vector<std::uint8_t> ReplaceZipPart(const std::vector<std::uint8_t>& package, std::string_view part_name,
                                         std::string_view from, std::string_view to) {
  mz_zip_archive reader{};
  if (mz_zip_reader_init_mem(&reader, package.data(), package.size(), 0) == MZ_FALSE) {
    ADD_FAILURE() << "could not open fixture package";
    return {};
  }
  mz_zip_archive writer{};
  if (mz_zip_writer_init_heap(&writer, 0, 4096) == MZ_FALSE) {
    ADD_FAILURE() << "could not create edited fixture package";
    mz_zip_reader_end(&reader);
    return {};
  }

  bool replaced = false;
  const mz_uint count = mz_zip_reader_get_num_files(&reader);
  for (mz_uint i = 0; i < count; ++i) {
    char name[256] = {};
    if (mz_zip_reader_get_filename(&reader, i, name, sizeof(name)) == MZ_FALSE) {
      ADD_FAILURE() << "could not read fixture entry name at index " << i;
      continue;
    }
    std::size_t size = 0;
    void* data = mz_zip_reader_extract_to_heap(&reader, i, &size, 0);
    if (data == nullptr) {
      ADD_FAILURE() << "could not extract fixture entry " << name;
      continue;
    }
    std::string bytes(static_cast<const char*>(data), size);
    mz_free(data);
    if (part_name == name) {
      const std::size_t at = bytes.find(from);
      EXPECT_NE(at, std::string::npos) << "missing source text in " << name;
      if (at != std::string::npos) {
        bytes.replace(at, from.size(), to);
        replaced = true;
      }
    }
    EXPECT_NE(
        mz_zip_writer_add_mem(&writer, name, bytes.data(), bytes.size(), static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
        MZ_FALSE)
        << "could not write fixture entry " << name;
  }
  EXPECT_TRUE(replaced) << "fixture entry not found: " << part_name;
  mz_zip_reader_end(&reader);

  void* archive = nullptr;
  std::size_t archive_size = 0;
  if (mz_zip_writer_finalize_heap_archive(&writer, &archive, &archive_size) == MZ_FALSE) {
    ADD_FAILURE() << "could not finish edited fixture package";
    mz_zip_writer_end(&writer);
    return {};
  }
  mz_zip_writer_end(&writer);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive),
                                static_cast<const std::uint8_t*>(archive) + archive_size);
  mz_free(archive);
  return out;
}

pugi::xml_node FindCachedCell(pugi::xml_document& doc, std::string_view ref) {
  const pugi::xml_node set = doc.child("externalLink").child("externalBook").child("sheetDataSet");
  for (pugi::xml_node sheet = set.child("sheetData"); sheet; sheet = sheet.next_sibling("sheetData")) {
    for (pugi::xml_node row = sheet.child("row"); row; row = row.next_sibling("row")) {
      for (pugi::xml_node cell = row.child("cell"); cell; cell = cell.next_sibling("cell")) {
        if (std::string_view(cell.attribute("r").value()) == ref) {
          return cell;
        }
      }
    }
  }
  return {};
}

/// The text of every `<f>` element in `xml`, in document order.
std::vector<std::string> FormulaTexts(const std::string& xml) {
  std::vector<std::string> out;
  pugi::xml_document doc;
  if (!doc.load_buffer(xml.data(), xml.size())) {
    ADD_FAILURE() << "unparseable part";
    return out;
  }
  for (const pugi::xpath_node& f : doc.select_nodes("//*[local-name()='f']")) {
    out.emplace_back(f.node().text().get());
  }
  return out;
}

/// Every `[N]` book index written into `xml`'s formulas.
std::vector<std::uint32_t> StoredBookIndices(const std::string& xml) {
  std::vector<std::uint32_t> out;
  const std::regex index_re("\\[([0-9]+)\\]");
  for (const std::string& f : FormulaTexts(xml)) {
    for (std::sregex_iterator it(f.begin(), f.end(), index_re), end; it != end; ++it) {
      out.push_back(static_cast<std::uint32_t>(std::stoul((*it)[1].str())));
    }
  }
  return out;
}

std::size_t Count(const std::string& haystack, std::string_view needle) {
  std::size_t n = 0;
  for (std::size_t at = haystack.find(needle); at != std::string::npos; at = haystack.find(needle, at + 1U)) {
    ++n;
  }
  return n;
}

ExternalLinkRecord LoadedLink(std::uint32_t index, std::string part_path, std::string target) {
  ExternalLinkRecord rec;
  rec.index = index;
  rec.rel_id = "rId" + std::to_string(index + 10U);
  rec.part_path = std::move(part_path);
  rec.body_rel_id = "rId1";
  rec.target = std::move(target);
  rec.kind = ExternalLinkRecord::Kind::kExternalBook;
  rec.book.sheet_names = {"Data"};
  rec.book.sheet_data = {true};
  return rec;
}

ExternalLinkRecord NewLink(std::uint32_t index, std::string target) {
  ExternalLinkRecord rec;
  rec.index = index;
  rec.target = std::move(target);
  rec.kind = ExternalLinkRecord::Kind::kExternalBook;
  rec.book.sheet_names = {"Data"};
  return rec;
}

PassthroughPart BodyPart(const std::string& path, const std::string& xml) {
  PassthroughPart part;
  part.path = path;
  part.content_type = std::string(kCtExternalLink);
  part.bytes.assign(xml.begin(), xml.end());
  return part;
}

Workbook HostBook() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Host");
  return wb;
}

void SetFormula(Workbook& wb, std::uint32_t row, std::uint32_t col, const std::string& formula) {
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, row, col, formula))) << formula;
}

std::string FormulaAt(const Workbook& wb, std::size_t sheet, std::uint32_t row, std::uint32_t col) {
  const Cell* cell = wb.sheet(sheet).cell_at(row, col);
  return cell == nullptr ? std::string() : cell->formula_text;
}

Value ValueAt(Workbook& wb, std::size_t sheet, std::uint32_t row, std::uint32_t col) {
  auto recalc = wb.recalc(eval::default_registry());
  EXPECT_TRUE(static_cast<bool>(recalc)) << (recalc ? "" : recalc.error().message);
  const Cell* cell = wb.sheet(sheet).cell_at(row, col);
  return cell == nullptr ? Value::blank() : cell->cached_value;
}

// ---------------------------------------------------------------------------
// Generated bodies
// ---------------------------------------------------------------------------

/// A link with two sheets (one cached), a book-scope and a sheet-scope
/// name, and one cached cell of each kind.
Workbook BookWithGeneratedLink() {
  Workbook wb = HostBook();
  ExternalLinkRecord rec = NewLink(1, "Book.xlsx");
  rec.absolute_target = "/Users/x/Book.xlsx";
  rec.book.sheet_names = {"Data", "Empty"};
  rec.book.sheet_data = {true, false};
  ExternalBookName total;
  total.name = "Total";
  total.sheet = 0;
  total.row = 2;
  total.col = 0;
  total.row_end = 2;
  total.col_end = 0;
  total.resolvable = true;
  ExternalBookName local;
  local.name = "Local";
  local.scope_sheet = 0;
  local.sheet = 0;
  local.row = 0;
  local.col = 0;
  local.row_end = 1;
  local.col_end = 1;
  local.is_range = true;
  local.resolvable = true;
  rec.book.names = {total, local};
  rec.book.cells[ExternalBook::cell_key(0, 0, 0)].value = Value::number(10.0);
  ExternalCell text;
  text.value = Value::text({});
  text.text = "hi";
  rec.book.cells[ExternalBook::cell_key(0, 0, 1)] = text;
  rec.book.cells[ExternalBook::cell_key(0, 1, 0)].value = Value::boolean(true);
  rec.book.cells[ExternalBook::cell_key(0, 1, 1)].value = Value::error(ErrorCode::NA);
  wb.set_external_links({rec});
  return wb;
}

TEST(OoxmlExternalLinkWriter, NonFiniteCachedNumbersAreStoredAsNumErrors) {
  Workbook wb = HostBook();
  ExternalLinkRecord rec = NewLink(1, "Book.xlsx");
  rec.book.sheet_data = {true};
  const auto add_number = [&rec](std::uint32_t col, double value) {
    ExternalCell cell;
    cell.value = Value::number(value);
    rec.book.cells[ExternalBook::cell_key(0, 0, col)] = std::move(cell);
  };
  add_number(0, std::numeric_limits<double>::infinity());
  add_number(1, -std::numeric_limits<double>::infinity());
  add_number(2, std::numeric_limits<double>::quiet_NaN());
  add_number(3, 12.5);
  ExternalCell boolean;
  boolean.value = Value::boolean(true);
  rec.book.cells[ExternalBook::cell_key(0, 0, 4)] = std::move(boolean);
  ExternalCell text;
  text.value = Value::text({});
  text.text = "control";
  rec.book.cells[ExternalBook::cell_key(0, 0, 5)] = std::move(text);
  wb.set_external_links({std::move(rec)});

  const std::vector<std::uint8_t> pkg = Save(wb);
  const std::string body = Part(pkg, "xl/externalLinks/externalLink1.xml");
  pugi::xml_document doc;
  ASSERT_TRUE(test::parse_xml(body, &doc));
  for (const char* ref : {"A1", "B1", "C1"}) {
    const pugi::xml_node cell = FindCachedCell(doc, ref);
    ASSERT_TRUE(cell) << ref << ": " << body;
    EXPECT_STREQ(cell.attribute("t").value(), "e") << ref;
    EXPECT_STREQ(cell.child("v").text().get(), "#NUM!") << ref;
  }
  const pugi::xml_node finite = FindCachedCell(doc, "D1");
  ASSERT_TRUE(finite) << body;
  EXPECT_FALSE(finite.attribute("t"));
  EXPECT_STREQ(finite.child("v").text().get(), "12.5");
  const pugi::xml_node boolean_cell = FindCachedCell(doc, "E1");
  ASSERT_TRUE(boolean_cell) << body;
  EXPECT_STREQ(boolean_cell.attribute("t").value(), "b");
  EXPECT_STREQ(boolean_cell.child("v").text().get(), "1");
  const pugi::xml_node text_cell = FindCachedCell(doc, "F1");
  ASSERT_TRUE(text_cell) << body;
  EXPECT_STREQ(text_cell.attribute("t").value(), "str");
  EXPECT_STREQ(text_cell.child("v").text().get(), "control");

  const Workbook reloaded = Load(pkg);
  ASSERT_EQ(reloaded.external_links().size(), 1U);
  const ExternalBook& book = reloaded.external_links()[0].book;
  for (std::uint32_t col = 0; col < 3U; ++col) {
    const Value value = book.cached_cell(0, 0, col);
    ASSERT_TRUE(value.is_error()) << col;
    EXPECT_EQ(value.as_error(), ErrorCode::Num) << col;
  }
  const Value finite_value = book.cached_cell(0, 0, 3);
  ASSERT_TRUE(finite_value.is_number());
  EXPECT_DOUBLE_EQ(finite_value.as_number(), 12.5);
  EXPECT_TRUE(book.cached_cell(0, 0, 4).as_boolean());
  EXPECT_EQ(book.cached_cell(0, 0, 5).as_text(), "control");
}

TEST(OoxmlExternalLinkWriter, NonFiniteNumbersInLoadedBodyStayVerbatimUntilRegenerated) {
  constexpr std::string_view kLinkPart = "xl/externalLinks/externalLink1.xml";
  const std::vector<std::uint8_t> fixture = FixtureBytes("external_link_cell_kinds.xlsx");
  ASSERT_FALSE(fixture.empty());
  const std::vector<std::uint8_t> patched = ReplaceZipPart(fixture, kLinkPart, "<v>42</v>", "<v>1e999</v>");
  ASSERT_FALSE(patched.empty());
  const std::string patched_body = Part(patched, kLinkPart);

  Workbook wb = Load(patched);
  ASSERT_EQ(wb.external_links().size(), 1U);
  const Value cached = wb.external_links()[0].book.cached_cell(1, 1, 1);
  ASSERT_TRUE(cached.is_number());
  ASSERT_TRUE(std::isinf(cached.as_number()));
  EXPECT_GT(cached.as_number(), 0.0);

  const std::vector<std::uint8_t> untouched = Save(wb);
  EXPECT_EQ(Part(untouched, kLinkPart), patched_body);

  std::vector<ExternalLinkRecord> links = wb.external_links();
  links[0].body_stale = true;
  wb.set_external_links(std::move(links));
  const std::vector<std::uint8_t> regenerated = Save(wb);
  const std::string body = Part(regenerated, kLinkPart);
  EXPECT_NE(body, patched_body);
  pugi::xml_document doc;
  ASSERT_TRUE(test::parse_xml(body, &doc));
  const pugi::xml_node cell = FindCachedCell(doc, "B2");
  ASSERT_TRUE(cell) << body;
  EXPECT_STREQ(cell.attribute("t").value(), "e");
  EXPECT_STREQ(cell.child("v").text().get(), "#NUM!");

  const Workbook reloaded = Load(regenerated);
  ASSERT_EQ(reloaded.external_links().size(), 1U);
  const Value error = reloaded.external_links()[0].book.cached_cell(1, 1, 1);
  ASSERT_TRUE(error.is_error());
  EXPECT_EQ(error.as_error(), ErrorCode::Num);
}

TEST(OoxmlExternalLinkWriter, GeneratedBodyCarriesSheetsNamesAndCachedData) {
  const std::vector<std::uint8_t> pkg = Save(BookWithGeneratedLink());
  const std::string body = Part(pkg, "xl/externalLinks/externalLink1.xml");
  pugi::xml_document doc;
  ASSERT_TRUE(test::parse_xml(body, &doc));
  const pugi::xml_node book = doc.child("externalLink").child("externalBook");
  ASSERT_TRUE(book) << body;
  EXPECT_STREQ(book.attribute("r:id").value(), "rId1");
  EXPECT_STREQ(book.child("xxl21:alternateUrls").child("xxl21:absoluteUrl").attribute("r:id").value(), "rId2");
  // The prefixes resolve to the namespaces Excel writes.
  EXPECT_STREQ(doc.child("externalLink").attribute("xmlns:xxl21").value(),
               "http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021");
  EXPECT_STREQ(book.attribute("xmlns:r").value(),
               "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

  std::vector<std::string> sheets;
  for (pugi::xml_node s = book.child("sheetNames").child("sheetName"); s; s = s.next_sibling("sheetName")) {
    sheets.emplace_back(s.attribute("val").value());
  }
  EXPECT_EQ(sheets, (std::vector<std::string>{"Data", "Empty"}));

  const pugi::xml_node total = book.child("definedNames").find_child_by_attribute("name", "Total");
  const pugi::xml_node local = book.child("definedNames").find_child_by_attribute("name", "Local");
  ASSERT_TRUE(total) << body;
  ASSERT_TRUE(local) << body;
  EXPECT_STREQ(total.attribute("refersTo").value(), "='Data'!$A$3");
  EXPECT_FALSE(total.attribute("sheetId"));
  EXPECT_STREQ(local.attribute("refersTo").value(), "='Data'!$A$1:$B$2");
  EXPECT_STREQ(local.attribute("sheetId").value(), "0");

  // Only the sheet that has cached data gets a `<sheetData>`.
  std::size_t sheet_data_count = 0;
  for (pugi::xml_node d = book.child("sheetDataSet").child("sheetData"); d; d = d.next_sibling("sheetData")) {
    ++sheet_data_count;
    EXPECT_STREQ(d.attribute("sheetId").value(), "0");
  }
  EXPECT_EQ(sheet_data_count, 1U);
}

TEST(OoxmlExternalLinkWriter, GeneratedBodyReadsBackToTheSameCache) {
  const Workbook reloaded = Load(Save(BookWithGeneratedLink()));
  ASSERT_EQ(reloaded.external_links().size(), 1U);
  const ExternalLinkRecord& rec = reloaded.external_links()[0];
  EXPECT_EQ(rec.target, "Book.xlsx");
  EXPECT_EQ(rec.absolute_target, "/Users/x/Book.xlsx");
  EXPECT_EQ(rec.book.sheet_names, (std::vector<std::string>{"Data", "Empty"}));
  EXPECT_TRUE(rec.book.sheet_has_data(0));
  EXPECT_FALSE(rec.book.sheet_has_data(1));
  const ExternalBookName* total = rec.book.find_name("Total");
  ASSERT_NE(total, nullptr);
  EXPECT_TRUE(total->resolvable);
  EXPECT_EQ(total->row, 2U);
  const ExternalBookName* local = rec.book.find_name("Local", 0);
  ASSERT_NE(local, nullptr);
  EXPECT_TRUE(local->is_range);
  EXPECT_EQ(local->row_end, 1U);
  EXPECT_EQ(local->col_end, 1U);
  EXPECT_EQ(rec.book.find_name("Local"), nullptr);
  EXPECT_DOUBLE_EQ(rec.book.cached_cell(0, 0, 0).as_number(), 10.0);
  EXPECT_EQ(rec.book.cached_cell(0, 0, 1).as_text(), "hi");
  EXPECT_TRUE(rec.book.cached_cell(0, 1, 0).as_boolean());
  EXPECT_EQ(rec.book.cached_cell(0, 1, 1).as_error(), ErrorCode::NA);
}

TEST(OoxmlExternalLinkWriter, CachedCellsKeepSheetAndRowMajorOrder) {
  Workbook wb = HostBook();
  ExternalLinkRecord rec = NewLink(1, "Book.xlsx");
  rec.book.sheet_names = {"First", "Empty", "Unrefreshed", "Last"};
  rec.book.sheet_data = {true, true, false, true};
  const auto add_number = [&rec](std::uint32_t sheet, std::uint32_t row, std::uint32_t col, double number) {
    ExternalCell cell;
    cell.value = Value::number(number);
    rec.book.cells[ExternalBook::cell_key(sheet, row, col)] = std::move(cell);
  };
  // Insert in reverse sheet and row/column order so the XML ordering comes
  // from the packed-key sort rather than the unordered map's iteration.
  add_number(3U, 1U, 0U, 30.0);
  add_number(2U, 0U, 0U, 20.0);  // Cached values are hidden when sheet_data is false.
  add_number(0U, 8U, 2U, 9.0);
  add_number(0U, 0U, 4U, 5.0);
  add_number(0U, 2U, 0U, 3.0);
  add_number(0U, 0U, 1U, 2.0);
  wb.set_external_links({std::move(rec)});

  const std::string body = Part(Save(wb), "xl/externalLinks/externalLink1.xml");
  pugi::xml_document doc;
  ASSERT_TRUE(test::parse_xml(body, &doc));
  const pugi::xml_node book = doc.child("externalLink").child("externalBook");
  const pugi::xml_node set = book.child("sheetDataSet");
  std::vector<std::string> sheet_ids;
  std::vector<std::vector<std::string>> cell_refs_by_sheet;
  std::vector<std::vector<std::string>> row_refs_by_sheet;
  for (pugi::xml_node sheet = set.child("sheetData"); sheet; sheet = sheet.next_sibling("sheetData")) {
    sheet_ids.emplace_back(sheet.attribute("sheetId").value());
    std::vector<std::string> rows;
    std::vector<std::string> cells;
    for (pugi::xml_node row = sheet.child("row"); row; row = row.next_sibling("row")) {
      rows.emplace_back(row.attribute("r").value());
      for (pugi::xml_node cell = row.child("cell"); cell; cell = cell.next_sibling("cell")) {
        cells.emplace_back(cell.attribute("r").value());
      }
    }
    row_refs_by_sheet.push_back(std::move(rows));
    cell_refs_by_sheet.push_back(std::move(cells));
  }

  EXPECT_EQ(sheet_ids, (std::vector<std::string>{"0", "1", "3"}));
  EXPECT_EQ(row_refs_by_sheet, (std::vector<std::vector<std::string>>{{"1", "3", "9"}, {}, {"2"}}));
  EXPECT_EQ(cell_refs_by_sheet, (std::vector<std::vector<std::string>>{{"B1", "E1", "A3", "C9"}, {}, {"A2"}}));
}

TEST(OoxmlExternalLinkWriter, CachedTextSurvivesTwoOoxmlCycles) {
  const std::string payload = "literal _x0041_ & <猫> \rcontrol \x01 🙂";
  Workbook current = HostBook();
  ExternalLinkRecord rec = NewLink(1, "Book.xlsx");
  rec.book.sheet_names = {"Data"};
  rec.book.sheet_data = {true};
  ExternalCell cached_text;
  cached_text.value = Value::text({});
  cached_text.text = payload;
  rec.book.cells[ExternalBook::cell_key(0, 0, 0)] = std::move(cached_text);
  current.set_external_links({std::move(rec)});

  for (int cycle = 0; cycle < 2; ++cycle) {
    current = Load(Save(current));
    ASSERT_EQ(current.external_links().size(), 1U);
    EXPECT_EQ(current.external_links()[0].book.cached_cell(0, 0, 0).as_text(), payload);
  }
}

TEST(OoxmlExternalLinkWriter, EveryWrittenLinkHasAllItsPackageFaces) {
  Workbook wb = HostBook();
  wb.set_external_links({NewLink(1, "A.xlsx"), NewLink(2, "B.xlsx")});
  const std::vector<std::uint8_t> pkg = Save(wb);
  const std::string workbook_xml = Part(pkg, "xl/workbook.xml");
  const std::string workbook_rels = Part(pkg, "xl/_rels/workbook.xml.rels");
  const std::string content_types = Part(pkg, "[Content_Types].xml");
  EXPECT_EQ(Count(workbook_xml, "<externalReference "), 2U) << workbook_xml;
  for (const char* k : {"1", "2"}) {
    const std::string name = std::string("externalLink") + k + ".xml";
    EXPECT_TRUE(HasPart(pkg, "xl/externalLinks/" + name));
    EXPECT_NE(workbook_rels.find("Target=\"externalLinks/" + name + "\""), std::string::npos) << workbook_rels;
    EXPECT_NE(content_types.find("<Override PartName=\"/xl/externalLinks/" + name + "\" ContentType=\"" +
                                 std::string(kCtExternalLink) + "\"/>"),
              std::string::npos)
        << content_types;
    EXPECT_TRUE(HasPart(pkg, "xl/externalLinks/_rels/" + name + ".rels"));
  }
}

TEST(OoxmlExternalLinkWriter, NewPartsTakeTheSmallestUnusedName) {
  Workbook wb = HostBook();
  const std::string loaded_body = "<externalLink/>";
  wb.set_external_links({LoadedLink(1, "xl/externalLinks/externalLink1.xml", "A.xlsx"),
                         LoadedLink(2, "xl/externalLinks/externalLink3.xml", "B.xlsx"), NewLink(3, "C.xlsx"),
                         NewLink(4, "D.xlsx")});
  wb.set_passthrough_parts({BodyPart("xl/externalLinks/externalLink1.xml", loaded_body),
                            BodyPart("xl/externalLinks/externalLink3.xml", loaded_body)});
  const std::vector<std::uint8_t> pkg = Save(wb);
  EXPECT_EQ(Part(pkg, "xl/externalLinks/externalLink1.xml"), loaded_body);
  EXPECT_EQ(Part(pkg, "xl/externalLinks/externalLink3.xml"), loaded_body);
  EXPECT_NE(Part(pkg, "xl/externalLinks/_rels/externalLink2.xml.rels").find("Target=\"C.xlsx\""), std::string::npos);
  EXPECT_NE(Part(pkg, "xl/externalLinks/_rels/externalLink4.xml.rels").find("Target=\"D.xlsx\""), std::string::npos);
  EXPECT_NE(Part(pkg, "xl/externalLinks/externalLink2.xml").find("<externalBook"), std::string::npos);

  // The written order is the index order whatever the part names are.
  const Workbook reloaded = Load(pkg);
  ASSERT_EQ(reloaded.external_links().size(), 4U);
  EXPECT_EQ(reloaded.external_links()[0].target, "A.xlsx");
  EXPECT_EQ(reloaded.external_links()[1].target, "B.xlsx");
  EXPECT_EQ(reloaded.external_links()[2].target, "C.xlsx");
  EXPECT_EQ(reloaded.external_links()[3].target, "D.xlsx");
}

TEST(OoxmlExternalLinkWriter, StoredIndicesArePositionsInTheWrittenList) {
  // Indices with gaps: the stored `[N]` is the link's position, not its index.
  Workbook wb = HostBook();
  wb.set_external_links({NewLink(2, "A.xlsx"), NewLink(5, "B.xlsx")});
  SetFormula(wb, 0, 0, "=[A.xlsx]Data!A1");
  SetFormula(wb, 1, 0, "=[B.xlsx]Data!A1+1");
  SetFormula(wb, 2, 0, "=SUM([b.xlsx]Data!A1:B2)");
  ASSERT_EQ(wb.external_links().size(), 2U);
  const std::vector<std::uint8_t> pkg = Save(wb);
  const std::string sheet_xml = Part(pkg, "xl/worksheets/sheet1.xml");
  EXPECT_EQ(FormulaTexts(sheet_xml), (std::vector<std::string>{"[1]Data!A1", "[2]Data!A1+1", "SUM([2]Data!A1:B2)"}));
  const std::size_t link_count = Count(Part(pkg, "xl/workbook.xml"), "<externalReference ");
  for (const std::uint32_t n : StoredBookIndices(sheet_xml)) {
    EXPECT_GE(n, 1U);
    EXPECT_LE(n, link_count);
  }
}

TEST(OoxmlExternalLinkWriter, BothRelationshipsAreEmittedForAGeneratedBody) {
  const std::vector<std::uint8_t> pkg = Save(BookWithGeneratedLink());
  const std::string rels = Part(pkg, "xl/externalLinks/_rels/externalLink1.xml.rels");
  pugi::xml_document doc;
  ASSERT_TRUE(test::parse_xml(rels, &doc));
  const pugi::xml_node root = doc.child("Relationships");
  const pugi::xml_node target = root.find_child_by_attribute("Id", "rId1");
  const pugi::xml_node absolute = root.find_child_by_attribute("Id", "rId2");
  ASSERT_TRUE(target) << rels;
  ASSERT_TRUE(absolute) << rels;
  EXPECT_STREQ(target.attribute("Target").value(), "Book.xlsx");
  EXPECT_STREQ(absolute.attribute("Target").value(), "/Users/x/Book.xlsx");
  for (const pugi::xml_node r : {target, absolute}) {
    EXPECT_STREQ(r.attribute("Type").value(),
                 "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath");
    EXPECT_STREQ(r.attribute("TargetMode").value(), "External");
  }
}

TEST(OoxmlExternalLinkWriter, AStaleBodyIsRegeneratedAndAFreshOneKept) {
  const std::string loaded_body =
      "<externalLink xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><externalBook "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\"><sheetNames>"
      "<sheetName val=\"Data\"/></sheetNames></externalBook></externalLink>";
  for (const bool stale : {false, true}) {
    Workbook wb = HostBook();
    ExternalLinkRecord rec = LoadedLink(1, "xl/externalLinks/externalLink1.xml", "A.xlsx");
    if (stale) {
      rec.book.sheet_names.push_back("Added");
      rec.book.sheet_data.push_back(false);
      rec.body_stale = true;
    }
    wb.set_external_links({rec});
    wb.set_passthrough_parts({BodyPart("xl/externalLinks/externalLink1.xml", loaded_body)});
    auto saved = io::write_ooxml_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    const std::string body = Part(saved.value().bytes, "xl/externalLinks/externalLink1.xml");
    if (stale) {
      EXPECT_NE(body, loaded_body);
      EXPECT_NE(body.find("<sheetName val=\"Added\"/>"), std::string::npos) << body;
      // A regenerated body supersedes the loaded copy; nothing was lost.
      EXPECT_EQ(saved.value().diagnostics.dropped_part_count, 0U);
    } else {
      EXPECT_EQ(body, loaded_body);
    }
    const std::string content_types = Part(saved.value().bytes, "[Content_Types].xml");
    EXPECT_EQ(Count(content_types, "/xl/externalLinks/externalLink1.xml\""), 1U) << content_types;
  }
}

// ---------------------------------------------------------------------------
// Formulas in every holder store `[N]`
// ---------------------------------------------------------------------------

TEST(OoxmlExternalLinkWriter, NamesConditionalFormatsAndValidationsStoreTheIndex) {
  Workbook wb = HostBook();
  wb.set_external_links({NewLink(1, "Book.xlsx")});
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Ext", "=[Book.xlsx]Data!A1")));
  cf::ConditionalFormat format;
  format.sqref.push_back(cf::CFCellRange{{0, 0}, {0, 0}});
  cf::CFRule rule;
  rule.type = cf::RuleType::Expression;
  rule.formula1 = "[Book.xlsx]Data!A1>0";
  format.rules.push_back(rule);
  wb.sheet(0).mutable_conditional_formats().push_back(format);
  DataValidation validation;
  validation.ranges.push_back(MergeRange{0, 0, 0, 0});
  validation.type = 1;
  validation.formula1 = "[Book.xlsx]Data!A1";
  wb.sheet(0).mutable_validations().push_back(validation);

  const std::vector<std::uint8_t> pkg = Save(wb);
  EXPECT_NE(Part(pkg, "xl/workbook.xml").find(">[1]Data!A1</definedName>"), std::string::npos);
  const std::string sheet_xml = Part(pkg, "xl/worksheets/sheet1.xml");
  EXPECT_NE(sheet_xml.find("<formula>[1]Data!A1&gt;0</formula>"), std::string::npos) << sheet_xml;
  EXPECT_NE(sheet_xml.find("<formula1>[1]Data!A1</formula1>"), std::string::npos) << sheet_xml;
}

TEST(OoxmlExternalLinkWriter, AnX14DataBarThresholdStoresTheIndex) {
  Workbook wb = HostBook();
  wb.set_external_links({NewLink(1, "Book.xlsx")});
  cf::ConditionalFormat format;
  format.sqref.push_back(cf::CFCellRange{{0, 0}, {9, 0}});
  cf::CFRule rule;
  rule.type = cf::RuleType::DataBar;
  rule.id = "{00000000-0000-0000-0000-000000000001}";
  cf::DataBarSpec bar;
  bar.min.type = cf::CfvoType::Formula;
  bar.min.value = "[Book.xlsx]Data!A1";
  bar.max.type = cf::CfvoType::Max;
  // A negative fill of its own is a setting only the x14 element carries.
  bar.negative_fill = cf::Color{255, 0, 0, 255};
  rule.data_bar = bar;
  format.rules.push_back(rule);
  wb.sheet(0).mutable_conditional_formats().push_back(format);

  const std::string sheet_xml = Part(Save(wb), "xl/worksheets/sheet1.xml");
  EXPECT_NE(sheet_xml.find("<xm:f>[1]Data!A1</xm:f>"), std::string::npos) << sheet_xml;
  EXPECT_NE(sheet_xml.find("val=\"[1]Data!A1\""), std::string::npos) << sheet_xml;
}

// ---------------------------------------------------------------------------
// Entered formulas
// ---------------------------------------------------------------------------

TEST(OoxmlExternalLinkWriter, AnEnteredBookSavesAsItsIndexAndReloads) {
  Workbook wb = HostBook();
  SetFormula(wb, 0, 0, "=[Book.xlsx]Sheet1!A1");
  const Value before = ValueAt(wb, 0, 0, 0);
  const std::vector<std::uint8_t> pkg = Save(wb);
  EXPECT_EQ(FormulaTexts(Part(pkg, "xl/worksheets/sheet1.xml")), (std::vector<std::string>{"[1]Sheet1!A1"}));
  EXPECT_NE(Part(pkg, "xl/externalLinks/_rels/externalLink1.xml.rels").find("Target=\"Book.xlsx\""), std::string::npos);

  Workbook reloaded = Load(pkg);
  EXPECT_EQ(FormulaAt(reloaded, 0, 0, 0), FormulaAt(wb, 0, 0, 0));
  ASSERT_EQ(reloaded.external_links().size(), 1U);
  const Value after = ValueAt(reloaded, 0, 0, 0);
  ASSERT_TRUE(before.is_error());
  EXPECT_EQ(before.as_error(), ErrorCode::Ref);
  ASSERT_TRUE(after.is_error());
  EXPECT_EQ(after.as_error(), before.as_error());
}

TEST(OoxmlExternalLinkWriter, ABareCellShapedSheetIsQuotedOnSave) {
  Workbook wb = HostBook();
  wb.add_sheet("S2");
  wb.sheet(1).set_cell_value(0, 0, Value::number(5.0));
  SetFormula(wb, 0, 0, "=S2!A1");
  EXPECT_DOUBLE_EQ(ValueAt(wb, 0, 0, 0).as_number(), 5.0);
  const std::vector<std::uint8_t> pkg = Save(wb);
  EXPECT_EQ(FormulaTexts(Part(pkg, "xl/worksheets/sheet1.xml")), (std::vector<std::string>{"'S2'!A1"}));

  Workbook reloaded = Load(pkg);
  EXPECT_EQ(FormulaAt(reloaded, 0, 0, 0), "='S2'!A1");
  EXPECT_DOUBLE_EQ(ValueAt(reloaded, 0, 0, 0).as_number(), 5.0);
}

// ---------------------------------------------------------------------------
// Names the cache cannot resolve, and names the book does not declare
// ---------------------------------------------------------------------------

ExternalBookName CellName(const char* name, std::uint32_t row) {
  ExternalBookName n;
  n.name = name;
  n.row = row;
  n.row_end = row;
  n.resolvable = true;
  return n;
}

TEST(OoxmlExternalLinkWriter, UnresolvableNameIsStoredAsARefError) {
  Workbook wb = HostBook();
  ExternalLinkRecord rec = NewLink(1, "Book.xlsx");
  ExternalBookName constant;
  constant.name = "Rate";
  constant.resolvable = false;
  rec.book.names = {constant};
  wb.set_external_links({rec});
  const std::string body = Part(Save(wb), "xl/externalLinks/externalLink1.xml");
  EXPECT_NE(body.find("<definedName name=\"Rate\" refersTo=\"#REF!\"/>"), std::string::npos) << body;

  const Workbook reloaded = Load(Save(wb));
  const ExternalBookName* back = reloaded.external_links()[0].book.find_name("Rate");
  ASSERT_NE(back, nullptr);
  EXPECT_TRUE(back->exists);
  EXPECT_FALSE(back->resolvable);
}

TEST(OoxmlExternalLinkWriter, AbsentNameIsStoredWithoutATargetAndReadsAsNoSuchName) {
  Workbook wb = HostBook();
  ExternalLinkRecord rec = NewLink(1, "Book.xlsx");
  ExternalBookName gone = CellName("NoSuch", 0);
  gone.exists = false;
  rec.book.names = {gone, CellName("After", 0)};
  rec.book.sheet_data = {true};
  rec.book.cells[ExternalBook::cell_key(0, 0, 0)].value = Value::number(10.0);
  wb.set_external_links({rec});
  SetFormula(wb, 0, 0, "=Book.xlsx!NoSuch");
  SetFormula(wb, 1, 0, "=Book.xlsx!After");

  const std::vector<std::uint8_t> pkg = Save(wb);
  const std::string body = Part(pkg, "xl/externalLinks/externalLink1.xml");
  EXPECT_NE(body.find("<definedName name=\"NoSuch\"/>"), std::string::npos) << body;
  EXPECT_EQ(body.find("NoSuch\" refersTo"), std::string::npos) << body;

  Workbook reloaded = Load(pkg);
  const ExternalBook& book = reloaded.external_links()[0].book;
  ASSERT_EQ(book.names.size(), 2U);
  EXPECT_FALSE(book.names[0].exists);
  EXPECT_TRUE(book.names[1].exists);
  EXPECT_TRUE(book.names[1].resolvable);
  const Value missing = ValueAt(reloaded, 0, 0, 0);
  ASSERT_TRUE(missing.is_error());
  EXPECT_EQ(missing.as_error(), ErrorCode::Name);
  const Cell* after = reloaded.sheet(0).cell_at(1, 0);
  ASSERT_NE(after, nullptr);
  ASSERT_TRUE(after->cached_value.is_number());
  EXPECT_DOUBLE_EQ(after->cached_value.as_number(), 10.0);
}

TEST(OoxmlExternalLinkWriter, ExtensionlessBookScopeNameIsStoredWithItsLinkOrdinal) {
  Workbook wb = HostBook();
  ExternalLinkRecord first = NewLink(1, "Book.xlsx");
  ExternalLinkRecord second = NewLink(2, "Src2");
  second.book.names = {CellName("Total", 0)};
  second.book.sheet_data = {true};
  second.book.cells[ExternalBook::cell_key(0, 0, 0)].value = Value::number(7.0);
  wb.set_external_links({first, second});
  SetFormula(wb, 0, 0, "='Src2'!Total");

  const std::vector<std::uint8_t> pkg = Save(wb);
  EXPECT_EQ(FormulaTexts(Part(pkg, "xl/worksheets/sheet1.xml")), (std::vector<std::string>{"[2]!Total"}));
  Workbook reloaded = Load(pkg);
  EXPECT_EQ(FormulaAt(reloaded, 0, 0, 0), "='Src2'!Total");
  const Value value = ValueAt(reloaded, 0, 0, 0);
  ASSERT_TRUE(value.is_number());
  EXPECT_DOUBLE_EQ(value.as_number(), 7.0);
}

// ---------------------------------------------------------------------------
// Excel-produced fixture
// ---------------------------------------------------------------------------

TEST(OoxmlExternalLinkWriter, FixtureFormulasSaveBackToTheirIndexForms) {
  const std::vector<std::uint8_t> original = FixtureBytes("external_link_mixed.xlsx");
  ASSERT_FALSE(original.empty());
  const std::vector<std::uint8_t> saved = Save(Load(original));
  const std::vector<std::string> expected = FormulaTexts(Part(original, "xl/worksheets/sheet1.xml"));
  ASSERT_EQ(expected,
            (std::vector<std::string>{"Local!A1", "SUM(Local!A1:A3)", "[1]Data!A1", "[1]!SrcTotal", "[2]!FarCell"}));
  EXPECT_EQ(FormulaTexts(Part(saved, "xl/worksheets/sheet1.xml")), expected);
}

TEST(OoxmlExternalLinkWriter, FixtureKeepsTheAbsoluteRelationship) {
  const std::vector<std::uint8_t> original = FixtureBytes("external_link_mixed.xlsx");
  ASSERT_FALSE(original.empty());
  const std::vector<std::uint8_t> saved = Save(Load(original));
  for (const char* k : {"1", "2"}) {
    const std::string body = std::string("xl/externalLinks/externalLink") + k + ".xml";
    EXPECT_EQ(Part(saved, body), Part(original, body));
    const std::string rels = Part(saved, std::string("xl/externalLinks/_rels/externalLink") + k + ".xml.rels");
    pugi::xml_document doc;
    ASSERT_TRUE(test::parse_xml(rels, &doc));
    const pugi::xml_node root = doc.child("Relationships");
    EXPECT_TRUE(root.find_child_by_attribute("Id", "rId1")) << rels;
    const pugi::xml_node absolute = root.find_child_by_attribute("Id", "rId2");
    ASSERT_TRUE(absolute) << rels;
    EXPECT_EQ(std::string(absolute.attribute("Target").value()),
              std::string("/Users/libraz/Documents/ext_link_probe/ExtSource") + (k[0] == '1' ? "" : "2") + ".xlsx");
  }
}

}  // namespace
}  // namespace formulon
