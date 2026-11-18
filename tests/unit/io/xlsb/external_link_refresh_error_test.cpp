// An external sheet's cache in each state Excel distinguishes, against Excel's own saves of the same
// books: a sheet the link lists with no cached data, or referenced but not listed, is written as a
// failed refresh (`refreshError="1"`, `BrtBeginExternTable` flag bit 0); an empty cache stays empty.

#include <cstdint>
#include <cstdio>
#include <string>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
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

// Authored books: A1:B7 reference [1]Missing, whose cache is absent from the
// part, listed with no data, listed and empty, or marked refreshError.
const char* const kStates[] = {"absent", "nodata", "empty", "referr"};

std::vector<std::uint8_t> FixtureBytes(const std::string& name) {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  for (int c = std::fgetc(file); c != EOF; c = std::fgetc(file)) {
    out.push_back(static_cast<std::uint8_t>(c));
  }
  std::fclose(file);
  return out;
}

std::string Entry(const std::vector<std::uint8_t>& package, const char* name) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto entry = zip.read_entry(name);
  EXPECT_TRUE(static_cast<bool>(entry)) << name;
  return entry ? std::string(entry.value().begin(), entry.value().end()) : std::string();
}

// The externalBook body after its relationship id, which Excel's save keeps verbatim.
std::string BookBody(const std::string& part) {
  const std::size_t begin = part.find("<sheetNames>");
  const std::size_t end = part.find("</externalBook>");
  EXPECT_NE(begin, std::string::npos) << part;
  return begin == std::string::npos || end == std::string::npos ? std::string() : part.substr(begin, end - begin);
}

Workbook Load(const std::string& name) {
  auto read_or = read_ooxml(SpanOf(FixtureBytes(name)));
  EXPECT_TRUE(static_cast<bool>(read_or)) << name;
  return read_or ? std::move(read_or.value().workbook) : Workbook::create_empty();
}

// `[1]Missing!A1` at B1 after a full recalc.
Value MissingCell(Workbook& wb) {
  auto recalc_or = wb.recalc_full(eval::default_registry());
  EXPECT_TRUE(static_cast<bool>(recalc_or));
  const Cell* cell = wb.sheet(0).cell_at(0U, 1U);
  return cell == nullptr ? Value::blank() : cell->cached_value;
}

// A1:B7 after a full recalc, or as cached when `recalc` is false.
std::vector<Value> Grid(Workbook& wb, bool recalc) {
  if (recalc) {
    EXPECT_TRUE(static_cast<bool>(wb.recalc_full(eval::default_registry())));
  }
  std::vector<Value> out;
  for (std::uint32_t row = 0; row < 7U; ++row) {
    for (std::uint32_t col = 0; col < 2U; ++col) {
      const Cell* cell = wb.sheet(0).cell_at(row, col);
      out.push_back(cell == nullptr ? Value::blank() : cell->cached_value);
    }
  }
  return out;
}

TEST(ExternalLinkRefreshError, XlsxSaveWritesEachCacheStateAsExcelDoes) {
  for (const char* state : kStates) {
    const Workbook wb = Load(std::string("extlink_cache_") + state + ".xlsx");
    auto saved = write_ooxml(wb);
    ASSERT_TRUE(static_cast<bool>(saved)) << state;
    const std::string excel = Entry(FixtureBytes(std::string("extlink_cache_") + state + "_excel.xlsx"),
                                    "xl/externalLinks/externalLink1.xml");
    EXPECT_EQ(BookBody(Entry(saved.value(), "xl/externalLinks/externalLink1.xml")), BookBody(excel)) << state;
  }
}

TEST(ExternalLinkRefreshError, XlsbSaveWritesEachCacheStateAsExcelDoes) {
  for (const char* state : kStates) {
    const Workbook wb = Load(std::string("extlink_cache_") + state + ".xlsx");
    auto saved = write_xlsb(wb);
    ASSERT_TRUE(static_cast<bool>(saved)) << state;
    const std::string excel = Entry(FixtureBytes(std::string("extlink_cache_") + state + "_excel.xlsb"),
                                    "xl/externalLinks/externalLink1.bin");
    EXPECT_EQ(Entry(saved.value(), "xl/externalLinks/externalLink1.bin"), excel) << state;
  }
}

TEST(ExternalLinkRefreshError, ACellOfTheSheetReadsTheSameThroughEitherSave) {
  for (const char* state : kStates) {
    // Excel's saves cache 0 for the empty sheet's blank cell and #REF! for the rest.
    const bool empty = std::string(state) == "empty";
    Workbook wb = Load(std::string("extlink_cache_") + state + ".xlsx");
    const Value expected = empty ? Value::number(0.0) : Value::error(ErrorCode::Ref);
    EXPECT_EQ(MissingCell(wb), expected) << state;

    auto xlsx = write_ooxml(wb);
    ASSERT_TRUE(static_cast<bool>(xlsx)) << state;
    auto via_xlsx = read_ooxml(SpanOf(xlsx.value()));
    ASSERT_TRUE(static_cast<bool>(via_xlsx)) << state;
    EXPECT_EQ(MissingCell(via_xlsx.value().workbook), expected) << state;

    auto xlsb = write_xlsb(wb);
    ASSERT_TRUE(static_cast<bool>(xlsb)) << state;
    auto via_xlsb = read_xlsb(SpanOf(xlsb.value()));
    ASSERT_TRUE(static_cast<bool>(via_xlsb)) << state;
    EXPECT_EQ(MissingCell(via_xlsb.value().workbook), expected) << state;
  }
}

TEST(ExternalLinkRefreshError, EveryFormulaReadsExcelsValuesThroughAnXlsbSave) {
  for (const char* state : kStates) {
    Workbook wb = Load(std::string("extlink_cache_") + state + ".xlsx");
    Workbook excel = Load(std::string("extlink_cache_") + state + "_excel.xlsx");
    auto xlsb = write_xlsb(wb);
    ASSERT_TRUE(static_cast<bool>(xlsb)) << state;
    auto via_xlsb = read_xlsb(SpanOf(xlsb.value()));
    ASSERT_TRUE(static_cast<bool>(via_xlsb)) << state;
    EXPECT_EQ(Grid(via_xlsb.value().workbook, true), Grid(excel, false)) << state;
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
