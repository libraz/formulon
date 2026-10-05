// In-memory XLSB reader tests grouped by cells, formulas, validation, and passthrough.

#include <array>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "defined_name.h"
#include "gtest/gtest.h"
#include "io/xlsb/reader.h"
#include "miniz.h"
#include "reader_test_helpers.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {
using namespace reader_test_support;
TEST(XlsbReader, DecodesFormulaAndPreservesCachedValue) {
  // rgce = `PtgInt 5` (`=5`): 0x1E 0x05 0x00.
  const std::vector<std::uint8_t> rgce = {0x1E, 0x05, 0x00};
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinFmlaNum(5.0, rgce)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message << " | " << result.error().context;

  const Cell* c = result.value().workbook.sheet(0).cell_at(0, 0);
  ASSERT_NE(c, nullptr);
  // The real formula text is decoded from the Ptg stream.
  EXPECT_EQ(c->formula_text, "=5");
  // The cached value is preserved.
  ASSERT_TRUE(c->cached_value.is_number());
  EXPECT_EQ(c->cached_value.as_number(), 5.0);
}
TEST(XlsbReader, UndecodableFormulaPreservesCachedValueWithoutFakeFormula) {
  // rgce = `PtgName` (0x23) + a 4-byte name index. PtgName is not in the
  // supported common-token set, so the decode fails and the reader must
  // preserve the cached value and store NO formula (never a fake one
  // that would recalc to #NAME?).
  const std::vector<std::uint8_t> rgce = {0x23, 0x01, 0x00, 0x00, 0x00};
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinFmlaNum(42.0, rgce)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message << " | " << result.error().context;

  const Cell* c = result.value().workbook.sheet(0).cell_at(0, 0);
  ASSERT_NE(c, nullptr);
  // No fabricated formula.
  EXPECT_TRUE(c->formula_text.empty());
  // Cached value preserved so the cell still shows the right number.
  ASSERT_TRUE(c->cached_value.is_number());
  EXPECT_EQ(c->cached_value.as_number(), 42.0);
  EXPECT_EQ(result.value().undecoded_formula_count, 1U);
  EXPECT_EQ(result.value().undecoded_defined_name_count, 0U);
}
TEST(XlsbReader, ExternalBookNameInCellFormulaKeepsCachedValueAndCounts) {
  // rgce = `PtgNameX` (0x39) + ixti (u16) + name index (u32): a defined
  // name owned by another workbook. Resolving it needs the supporting-
  // book and external-name tables, which this reader does not decode, so
  // the token stays explicitly unsupported. Reusing the index against
  // this workbook's own name table would silently retarget the formula
  // at a different name, so the cell must instead fall back to the same
  // contract an undecodable stream gets: cached value kept, no formula,
  // one diagnostic counted.
  const std::vector<std::uint8_t> rgce = {0x39, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00};
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinFmlaNum(7.25, rgce)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message << " | " << result.error().context;

  const Cell* c = result.value().workbook.sheet(0).cell_at(0, 0);
  ASSERT_NE(c, nullptr);
  EXPECT_TRUE(c->formula_text.empty());
  ASSERT_TRUE(c->cached_value.is_number());
  EXPECT_EQ(c->cached_value.as_number(), 7.25);
  EXPECT_EQ(result.value().undecoded_formula_count, 1U);
  EXPECT_EQ(result.value().undecoded_defined_name_count, 0U);
}
TEST(XlsbReader, DecodableDefinedNameRegistersAndCountsNothing) {
  // Positive control for the record builder the next test relies on: a
  // `BrtName` whose rgce is `PtgInt 5` must reach the workbook intact,
  // so a zero counter there means "nothing was undecodable" rather than
  // "no name record was ever parsed".
  const std::vector<std::uint8_t> rgce = {0x1E, 0x05, 0x00};
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin({NameRecord("Rate", rgce)})});
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinReal(1.0)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message << " | " << result.error().context;

  const std::vector<DefinedName>& names = result.value().workbook.defined_names();
  ASSERT_EQ(names.size(), 1U);
  EXPECT_EQ(names[0].name, "Rate");
  EXPECT_EQ(names[0].formula, "5");
  EXPECT_EQ(names[0].local_sheet_id, -1);
  EXPECT_EQ(result.value().undecoded_defined_name_count, 0U);
}
TEST(XlsbReader, DefinedNameCommentSurvivesDecode) {
  // The Name Manager comment is the trailing `BrtName` string after
  // `rgce`/`rgcb`; a reader that stops at the formula body (or
  // misreads `cb`) never reaches it.
  const std::vector<std::uint8_t> rgce = {0x1E, 0x05, 0x00};
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin({NameRecord("Rate", rgce, "The annual interest rate")})});
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinReal(1.0)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message << " | " << result.error().context;

  const std::vector<DefinedName>& names = result.value().workbook.defined_names();
  ASSERT_EQ(names.size(), 1U);
  EXPECT_EQ(names[0].comment, "The annual interest rate");
}
TEST(XlsbReader, ExternalBookNameInDefinedNameIsSkippedAndCounted) {
  // Same unsupported token, this time as a defined name's own body. The
  // name is dropped rather than registered with a fabricated formula,
  // and the read still succeeds so one external reference cannot cost
  // the caller the whole workbook.
  const std::vector<std::uint8_t> rgce = {0x39, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00};
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin({NameRecord("ExtRate", rgce)})});
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinReal(1.0)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message << " | " << result.error().context;

  EXPECT_TRUE(result.value().workbook.defined_names().empty());
  EXPECT_EQ(result.value().undecoded_defined_name_count, 1U);
  EXPECT_EQ(result.value().undecoded_formula_count, 0U);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
