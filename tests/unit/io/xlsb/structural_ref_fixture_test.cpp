//
// Reads Mac Excel 365-saved workbooks whose formulas lost their targets to
// structural edits, and checks each formula against the text Excel's
// formula bar shows for it.
//
//   `structural_ref_p1.xlsb` — Sheet1 and Sheet3 survive; Sheet3 lost cells
//     and the sheet Sheet2 is gone, so its references keep their address.
//   `structural_ref_p6.xlsb` — references to a deleted sheet without an
//     address.

#include <cstdint>
#include <cstdio>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "defined_name.h"
#include "gtest/gtest.h"
#include "io/xlsb/reader.h"
#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace {

std::vector<std::uint8_t> ReadFileBytes(const std::string& path) {
  std::vector<std::uint8_t> out;
  std::FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "cannot open fixture: " << path;
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  if (size > 0) {
    out.resize(static_cast<std::size_t>(size));
    const std::size_t read = std::fread(out.data(), 1, out.size(), file);
    if (read != out.size()) {
      ADD_FAILURE() << "short read on fixture: " << path;
      out.clear();
    }
  }
  std::fclose(file);
  return out;
}

Workbook LoadFixture(const std::string& name) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name;
  const std::vector<std::uint8_t> bytes = ReadFileBytes(path);
  if (bytes.empty()) {
    return Workbook::create_empty();
  }
  auto result_or = io::xlsb::read_xlsb(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(result_or)) << path << ": " << (result_or ? "" : result_or.error().message);
  if (!result_or) {
    return Workbook::create_empty();
  }
  EXPECT_EQ(result_or.value().undecoded_formula_count, 0U) << path;
  return std::move(result_or.value().workbook);
}

std::string FormulaAt(const Workbook& wb, const char* sheet, std::uint32_t row, std::uint32_t col) {
  const std::size_t index = wb.sheet_index_by_name(sheet);
  if (index >= wb.sheet_count()) {
    ADD_FAILURE() << "no sheet " << sheet;
    return {};
  }
  const Cell* cell = wb.sheet(index).cell_at(row, col);
  return cell != nullptr ? cell->formula_text : std::string();
}

std::string NameFormula(const Workbook& wb, const char* name) {
  for (const DefinedName& dn : wb.defined_names()) {
    if (dn.name == name) {
      return dn.formula;
    }
  }
  ADD_FAILURE() << "no defined name " << name;
  return {};
}

TEST(XlsbStructuralRefFixture, QualifiedAndDeletedSheetReferences) {
  const Workbook wb = LoadFixture("structural_ref_p1.xlsb");
  ASSERT_EQ(wb.sheet_count(), 2U);
  EXPECT_EQ(FormulaAt(wb, "Sheet1", 0U, 3U), "=Sheet3!#REF!*2");
  EXPECT_EQ(FormulaAt(wb, "Sheet1", 1U, 3U), "=SUM(Sheet3!#REF!)");
  EXPECT_EQ(FormulaAt(wb, "Sheet1", 2U, 3U), "=Sheet3!#REF!");
  EXPECT_EQ(FormulaAt(wb, "Sheet1", 3U, 3U), "=#REF!A1");
  EXPECT_EQ(FormulaAt(wb, "Sheet1", 4U, 3U), "=SUM(#REF!A1:B2)");
  EXPECT_EQ(FormulaAt(wb, "Sheet1", 5U, 3U), "=SUM(#REF!A:A)");
  EXPECT_EQ(FormulaAt(wb, "Sheet1", 6U, 3U), "=SUM(#REF!1:3)");
  EXPECT_EQ(FormulaAt(wb, "Sheet1", 7U, 3U), "=#REF!$A$1+Sheet3!#REF!");
  EXPECT_EQ(FormulaAt(wb, "Sheet3", 0U, 3U), "=#REF!");
  EXPECT_EQ(NameFormula(wb, "Nm"), "Sheet3!#REF!");
  EXPECT_EQ(NameFormula(wb, "Nm2"), "#REF!$A$1");
  EXPECT_EQ(NameFormula(wb, "Nm3"), "#REF!$A$1:$B$2");
}

TEST(XlsbStructuralRefFixture, DeletedSheetWithoutAddress) {
  const Workbook wb = LoadFixture("structural_ref_p6.xlsb");
  EXPECT_EQ(FormulaAt(wb, "Sheet1", 0U, 3U), "=#REF!#REF!");
  EXPECT_EQ(FormulaAt(wb, "Sheet1", 1U, 3U), "=SUM(#REF!#REF!)+1");
}

}  // namespace
}  // namespace formulon
