// Parts derived from sheet content the reader modelled are the writer's own retained originals, not passthroughs:
// a save that regenerates or leaves them out counts no lost part or relationship, and they are not listed.
//
// A sheet's note drawing is written from its notes; a drawing the reader did not tie to notes stays a passthrough,
// listed and counted. A desktop-written .xlsb's binary index over a sheet's cell records holds nothing the model
// lacks, Excel rebuilds it, and .xlsx has no such part, so leaving it out loses nothing. The calculation chain of
// either format only orders recalculation; the readers consume it and neither save writes or counts it.

#include <cstddef>
#include <cstdint>
#include <map>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/package_diagnostics.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "workbook.h"
#include "xlsb_twin_compare.h"
#include "zip_repack_helpers.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

namespace tc = test;

const char kVml[] = "xl/drawings/vmlDrawing1.vml";
const char kBinaryIndex[] = "xl/worksheets/binaryIndex1.bin";
// Desktop-written without a calculation chain: two sheets, and eight sheets, each with its binary index.
const char* const kIndexedFixtures[] = {"xlsb_parity/pivot/formats_areas.xlsb", "xlsb_pivot_source_sheet.xlsb"};

tc::Bytes Fixture(const char* name) {
  return tc::ReadBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name);
}

Workbook LoadXlsb(const tc::Bytes& bytes) {
  auto read = read_xlsb(ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(read.has_value()) << read.error().message;
  return read.has_value() ? std::move(read.value().workbook) : Workbook::create();
}

void ExpectCounts(const WriteDiagnostics& diagnostics, std::uint32_t parts) {
  EXPECT_EQ(diagnostics.dropped_part_count, parts);
  EXPECT_EQ(diagnostics.dropped_relationship_count, 0U);
}

Workbook NotedWorkbook() {
  Workbook wb = Workbook::create();
  wb.set_cell_value(0, 0, 0, Value::number(1.0));
  wb.sheet(0).mutable_comments().push_back(CellComment{0, 0, "Author", "Note"});
  return wb;
}

std::vector<std::uint8_t> NotedXlsx() {
  auto bytes = write_ooxml(NotedWorkbook());
  EXPECT_TRUE(bytes.has_value()) << bytes.error().message;
  return bytes.has_value() ? bytes.value() : std::vector<std::uint8_t>{};
}

std::vector<std::uint8_t> NotedXlsb() {
  auto bytes = write_xlsb(NotedWorkbook());
  EXPECT_TRUE(bytes.has_value()) << bytes.error().message;
  return bytes.has_value() ? bytes.value() : std::vector<std::uint8_t>{};
}

void ExpectBothSavesCountNothing(const Workbook& wb) {
  auto xlsx = write_ooxml_with_result(wb);
  ASSERT_TRUE(xlsx.has_value()) << xlsx.error().message;
  EXPECT_EQ(xlsx.value().diagnostics.dropped_part_count, 0U);
  auto xlsb = write_xlsb_with_result(wb);
  ASSERT_TRUE(xlsb.has_value()) << xlsb.error().message;
  EXPECT_EQ(xlsb.value().diagnostics.dropped_part_count, 0U);
  auto back = read_xlsb(ByteSpan{xlsb.value().bytes.data(), xlsb.value().bytes.size()});
  ASSERT_TRUE(back.has_value()) << back.error().message;
  EXPECT_EQ(back.value().workbook.sheet(0).comments().size(), 1U);
}

bool Listed(const std::vector<std::uint8_t>& package, const char* path) {
  fm_workbook_t* wb = nullptr;
  EXPECT_EQ(fm_workbook_load(package.data(), package.size(), nullptr, &wb), 0) << fm_get_last_error_message();
  int32_t present = -1;
  EXPECT_EQ(fm_workbook_has_passthrough(wb, path, &present), 0);
  fm_workbook_destroy(wb);
  return present == 1;
}

TEST(NoteDrawingRetained, AnXlsxNoteDrawingIsNotCountedByEitherSave) {
  const std::vector<std::uint8_t> source = NotedXlsx();
  auto read = read_ooxml(ByteSpan{source.data(), source.size()});
  ASSERT_TRUE(read.has_value()) << read.error().message;
  ExpectBothSavesCountNothing(read.value().workbook);
}

TEST(NoteDrawingRetained, AnXlsbNoteDrawingIsNotCountedByEitherSave) {
  const std::vector<std::uint8_t> source = NotedXlsb();
  auto read = read_xlsb(ByteSpan{source.data(), source.size()});
  ASSERT_TRUE(read.has_value()) << read.error().message;
  ExpectBothSavesCountNothing(read.value().workbook);
}

TEST(NoteDrawingRetained, AModelledNoteDrawingIsNotListed) {
  EXPECT_FALSE(Listed(NotedXlsx(), kVml));
  EXPECT_FALSE(Listed(NotedXlsb(), kVml));
}

// Without its comments relationship the sheet has no notes to tie the drawing to, so the drawing stays unmodelled.
TEST(NoteDrawingRetained, ADrawingNotTiedToNotesIsListed) {
  const std::vector<std::uint8_t> source = NotedXlsx();
  const char kRels[] = "xl/worksheets/_rels/sheet1.xml.rels";
  const std::optional<tc::Bytes> rels_bytes = tc::ReadPart(source, kRels);
  ASSERT_TRUE(rels_bytes.has_value());
  std::string rels(rels_bytes->begin(), rels_bytes->end());
  const std::size_t type = rels.find("/relationships/comments\"");
  ASSERT_NE(type, std::string::npos);
  const std::size_t open = rels.rfind("<Relationship ", type);
  const std::size_t close = rels.find("/>", type);
  ASSERT_TRUE(open != std::string::npos && close != std::string::npos);
  rels.erase(open, close + 2U - open);
  const std::vector<std::uint8_t> untied =
      RepackWithEntries(source, {{kRels, std::vector<std::uint8_t>(rels.begin(), rels.end())}});
  EXPECT_TRUE(Listed(untied, kVml));
}

// A part the reader kept without modelling it, under the path the notes' drawing is written to, is lost and counted.
TEST(NoteDrawingRetained, AnUnmodelledDrawingUnderTheNotesPathIsCounted) {
  Workbook wb = NotedWorkbook();
  wb.set_passthrough_parts({PassthroughPart(kVml, "", std::vector<std::uint8_t>{'<', 'x', '/', '>'})});
  auto xlsx = write_ooxml_with_result(wb);
  ASSERT_TRUE(xlsx.has_value()) << xlsx.error().message;
  EXPECT_EQ(xlsx.value().diagnostics.dropped_part_count, 1U);
  auto xlsb = write_xlsb_with_result(wb);
  ASSERT_TRUE(xlsb.has_value()) << xlsb.error().message;
  EXPECT_EQ(xlsb.value().diagnostics.dropped_part_count, 1U);
}

TEST(BinaryIndexRetained, AnUneditedSaveCountsNothingInEitherFormat) {
  for (const char* name : kIndexedFixtures) {
    SCOPED_TRACE(name);
    const Workbook wb = LoadXlsb(Fixture(name));
    auto xlsb = write_xlsb_with_result(wb);
    ASSERT_TRUE(xlsb.has_value()) << xlsb.error().message;
    ExpectCounts(xlsb.value().diagnostics, 0U);
    auto xlsx = write_ooxml_with_result(wb);
    ASSERT_TRUE(xlsx.has_value()) << xlsx.error().message;
    ExpectCounts(xlsx.value().diagnostics, 0U);
  }
}

TEST(BinaryIndexRetained, ABinaryIndexIsNotListed) {
  for (const char* name : kIndexedFixtures) {
    SCOPED_TRACE(name);
    EXPECT_FALSE(Listed(Fixture(name), kBinaryIndex));
  }
}

// An unmodelled part beside the indexes, under a path the save generates, is still lost and counted.
TEST(BinaryIndexRetained, AnUnmodelledPartBesideItIsCounted) {
  const std::vector<std::uint8_t> filler{'<', 'x', '/', '>'};
  {
    Workbook wb = LoadXlsb(Fixture(kIndexedFixtures[0]));
    std::vector<PassthroughPart> parts = wb.passthrough_parts();
    parts.emplace_back("xl/worksheets/sheet1.bin", "", filler);
    wb.set_passthrough_parts(std::move(parts));
    auto xlsb = write_xlsb_with_result(wb);
    ASSERT_TRUE(xlsb.has_value()) << xlsb.error().message;
    ExpectCounts(xlsb.value().diagnostics, 1U);
  }
  {
    Workbook wb = LoadXlsb(Fixture(kIndexedFixtures[0]));
    std::vector<PassthroughPart> parts = wb.passthrough_parts();
    parts.emplace_back("xl/worksheets/sheet1.xml", "", filler);
    wb.set_passthrough_parts(std::move(parts));
    auto xlsx = write_ooxml_with_result(wb);
    ASSERT_TRUE(xlsx.has_value()) << xlsx.error().message;
    ExpectCounts(xlsx.value().diagnostics, 1U);
  }
}

// Desktop-written twins with formulas and a calculation chain.
const char* const kChainedFixtures[] = {"self_book_name.xlsx", "self_book_name.xlsb"};

Workbook LoadEither(const char* name) {
  const tc::Bytes bytes = Fixture(name);
  if (std::string(name).find(".xlsb") != std::string::npos) {
    return LoadXlsb(bytes);
  }
  auto read = read_ooxml(ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(read.has_value()) << read.error().message;
  return read.has_value() ? std::move(read.value().workbook) : Workbook::create();
}

/// Whether `package` has the part at `chain`, or a workbook relationship of the calcChain type in `rels`.
bool HasChain(const tc::Bytes& package, const char* chain, const char* rels) {
  const std::optional<tc::Bytes> rel_part = tc::ReadPart(package, rels);
  EXPECT_TRUE(rel_part.has_value());
  const std::string rel_xml = rel_part.has_value() ? std::string(rel_part->begin(), rel_part->end()) : std::string();
  return tc::ReadPart(package, chain).has_value() || rel_xml.find("/calcChain\"") != std::string::npos;
}

TEST(CalcChainLeftOut, AnUneditedSaveCountsNothingAndWritesNoChain) {
  for (const char* name : kChainedFixtures) {
    SCOPED_TRACE(name);
    const Workbook wb = LoadEither(name);
    auto xlsx = write_ooxml_with_result(wb);
    ASSERT_TRUE(xlsx.has_value()) << xlsx.error().message;
    ExpectCounts(xlsx.value().diagnostics, 0U);
    EXPECT_FALSE(HasChain(xlsx.value().bytes, "xl/calcChain.xml", "xl/_rels/workbook.xml.rels"));
    auto xlsb = write_xlsb_with_result(wb);
    ASSERT_TRUE(xlsb.has_value()) << xlsb.error().message;
    ExpectCounts(xlsb.value().diagnostics, 0U);
    EXPECT_FALSE(HasChain(xlsb.value().bytes, "xl/calcChain.bin", "xl/_rels/workbook.bin.rels"));
  }
}

TEST(CalcChainLeftOut, AChainIsNotListed) {
  EXPECT_FALSE(Listed(Fixture("self_book_name.xlsx"), "xl/calcChain.xml"));
  EXPECT_FALSE(Listed(Fixture("self_book_name.xlsb"), "xl/calcChain.bin"));
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
