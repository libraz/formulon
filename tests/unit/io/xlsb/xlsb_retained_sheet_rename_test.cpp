// A sheet rename in a workbook loaded from an Excel-saved .xlsb keeps its retained future records (x14
// conditional formats and validations) savable: the records index the sheet table by position, so the renamed
// entry still names the same sheet, and a reference to it reads the new name after the save.

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
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

Workbook LoadXlsb(const std::vector<std::uint8_t>& bytes) {
  auto r = read_xlsb(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(r)) << (r ? "" : r.error().message);
  return r ? std::move(r.value().workbook) : Workbook::create_empty();
}

Workbook LoadFixture(const char* name) {
  return LoadXlsb(ReadFixtureFile(std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name));
}

// Every worksheet part of `wb` saved as .xlsx, concatenated.
std::string XlsxSheetParts(const Workbook& wb) {
  auto saved = write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  std::string out;
  if (!saved) {
    return out;
  }
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(saved.value()))));
  for (std::size_t i = 1; i <= wb.sheet_count(); ++i) {
    auto entry = zip.read_entry("xl/worksheets/sheet" + std::to_string(i) + ".xml");
    if (entry) {
      out.append(entry.value().begin(), entry.value().end());
    }
  }
  return out;
}

TEST(XlsbRetainedSheetRename, ReferencedSheetRenameSavesAndTheReferenceFollows) {
  Workbook wb = LoadFixture("xlsb_feature_iconbits.xlsb");
  ASSERT_EQ(wb.sheet_count(), 2U);
  ASSERT_EQ(wb.sheet(1).name(), "S2");
  ASSERT_NE(XlsxSheetParts(wb).find("'S2'!"), std::string::npos);

  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(1U, "Renamed")));
  auto written = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;

  const Workbook back = LoadXlsb(written.value());
  ASSERT_EQ(back.sheet(1).name(), "Renamed");
  const std::string parts = XlsxSheetParts(back);
  EXPECT_NE(parts.find("Renamed!"), std::string::npos);
  EXPECT_EQ(parts.find("S2'!"), std::string::npos);
}

TEST(XlsbRetainedSheetRename, RenamingTheOwningSheetSaves) {
  Workbook wb = LoadFixture("xlsb_feature_x14bars.xlsb");
  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(0U, "Renamed")));
  auto written = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  EXPECT_EQ(LoadXlsb(written.value()).sheet(0).name(), "Renamed");
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
