// Formulas whose sheet qualifier names no sheet of the workbook, against Mac Excel 365's saves in
// `external_link_path_missing/`. Excel turns each unknown qualifier into a link with a missing path and stores:
//   one_link    `[1]Nope!A1`               PtgRef3d through XTI (link 0, 0, 0)
//   two_links   `[1]Nope!A1`, `[2]Other!B3` XTIs (0, 0, 0) and (1, 0, 0), one link per qualifier
//   span_name   `SUM(Data:'[1]Nope'!A1)`  PtgMemFunc { PtgName -> `Data` stub with an empty formula,
//                                          PtgRef3d through the link XTI, PtgRange }
//   book_name   `[1]!LocalName`            PtgNameX through XTI (link, -2, -2); the link lists the name with no
//                                          formula, and the self book appears for the book's own `LocalName`
// `<stem>_input.xlsx` is the authored book Excel opened (`Nope!A1`, `SUM(Data:Nope!A1)`, ...), with no link. Its
// cached values are 0 under calcId 0, so Excel recalculated on open; the engine's save matches after a full recalc.

#include <cstdint>
#include <set>
#include <string>
#include <utility>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest-spi.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "support/roundtrip_symmetry.h"
#include "workbook.h"
#include "xlsb_twin_compare.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::uint16_t kBrtExternNameBits = 586;

constexpr const char* kStems[] = {"one_link", "two_links", "span_name", "book_name"};

std::uint16_t Type(XlsbRecordType type) {
  return static_cast<std::uint16_t>(type);
}

/// The supporting-book list, the XTI table, formula cells and every link-part record. Defined names are added
/// for span_name only: Excel's .xlsb of book_name carries a hidden name its .xlsx twin does not.
std::set<std::uint16_t> ComparedTypes(const std::string& stem) {
  std::set<std::uint16_t> types = {
      Type(XlsbRecordType::BrtSupBookSrc),
      Type(XlsbRecordType::BrtSupSelf),
      Type(XlsbRecordType::BrtSupSame),
      Type(XlsbRecordType::BrtSupTabs),
      Type(XlsbRecordType::BrtExternSheet),
      Type(XlsbRecordType::BrtFmlaString),
      Type(XlsbRecordType::BrtFmlaNum),
      Type(XlsbRecordType::BrtFmlaBool),
      Type(XlsbRecordType::BrtFmlaError),
      Type(XlsbRecordType::BrtBeginExternalBook),
      Type(XlsbRecordType::BrtEndExternalBook),
      Type(XlsbRecordType::BrtExternNameStart),
      Type(XlsbRecordType::BrtExternNameFmla),
      Type(XlsbRecordType::BrtExternNameEnd),
      Type(XlsbRecordType::BrtBeginExternTable),
      Type(XlsbRecordType::BrtEndExternTable),
      kBrtExternNameBits,
  };
  if (stem == "span_name") {
    types.insert(Type(XlsbRecordType::BrtName));
  }
  return types;
}

test::Bytes Fixture(const std::string& stem, const char* extension) {
  return test::ReadBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/external_link_path_missing/" + stem + extension);
}

Workbook Load(const test::Bytes& package, bool xlsb) {
  if (xlsb) {
    auto result = read_xlsb(formulon::test::span_of(package));
    EXPECT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
    return result ? std::move(result.value().workbook) : Workbook::create_empty();
  }
  auto result = read_ooxml(formulon::test::span_of(package));
  EXPECT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  return result ? std::move(result.value().workbook) : Workbook::create_empty();
}

/// One line per record difference between our .xlsb save of `wb` and Excel's save of `stem`.
std::vector<std::string> SaveMismatches(const Workbook& wb, const std::string& stem) {
  auto saved = write_xlsb(wb);
  if (!saved) {
    return {"write_xlsb failed: " + saved.error().message};
  }
  const test::Bytes excel = Fixture(stem, ".xlsb");
  test::ByteMask mask;
  test::AddUninitialisedFields(excel, test::DefaultUninitialisedHooks(), &mask);
  return test::CompareRecordSets(saved.value(), excel, ComparedTypes(stem), mask);
}

TEST(PathMissingLinkRecords, ExcelSavesReSaveToTheSameRecords) {
  for (const char* stem : kStems) {
    for (const char* source : {".xlsx", ".xlsb"}) {
      SCOPED_TRACE(std::string(stem) + source);
      for (const std::string& line :
           SaveMismatches(Load(Fixture(stem, source), std::string(source) == ".xlsb"), stem)) {
        ADD_FAILURE() << line;
      }
    }
  }
}

/// Every difference between the save of the authored `<stem>_input.xlsx`, loaded and fully recalculated as Excel
/// does on open, and Excel's save.
std::vector<std::string> AuthoredSaveMismatches(const std::string& stem) {
  Workbook wb = Load(Fixture(stem + "_input", ".xlsx"), false);
  auto recalculated = wb.recalc_full(eval::default_registry());
  if (!recalculated) {
    return {"recalc_full failed: " + recalculated.error().message};
  }
  return SaveMismatches(wb, stem);
}

TEST(PathMissingLinkRecords, AuthoredBooksRecalculatedSaveAsExcel) {
  for (const char* stem : {"one_link", "two_links", "span_name"}) {
    SCOPED_TRACE(stem);
    for (const std::string& line : AuthoredSaveMismatches(stem)) {
      ADD_FAILURE() << line;
    }
  }
}

void ExpectAuthoredBookNameSavesAsExcel() {
  std::string report;
  for (const std::string& line : AuthoredSaveMismatches("book_name")) {
    report += line + "\n";
  }
  if (!report.empty()) {
    ADD_FAILURE() << report;
  }
}

// Excel's link part for `[1]!LocalName` lists the sheet `Nope` and the name with no formula; the engine's save of the
// authored book carries neither. Fails once the saves match.
TEST(PathMissingLinkRecords, AuthoredBookNameSaveDiffersFromExcelSave) {
  EXPECT_NONFATAL_FAILURE(ExpectAuthoredBookNameSavesAsExcel(), "");
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
