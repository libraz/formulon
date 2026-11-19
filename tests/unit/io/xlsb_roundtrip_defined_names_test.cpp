// Cross-format XLSB/OOXML symmetry tests.

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <cstring>
#include <iterator>
#include <limits>
#include <string>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "cell.h"
#include "cf/cf_types.h"
#include "defined_name.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "styles.h"
#include "support/roundtrip_symmetry.h"
#include "value.h"
#include "xlsb_roundtrip_symmetry_test_helpers.h"

namespace formulon {
namespace {
using namespace xlsb_roundtrip_test_support;
TEST(XlsbWriteReadSymmetry, CellStyleAndNumberFormatSurviveRoundTrip) {
  const std::vector<std::uint8_t> bytes = test::read_file_bytes(FixturePath("xlsb_fidelity_base.xlsb"));
  ASSERT_FALSE(bytes.empty());
  auto loaded = io::xlsb::read_xlsb(test::span_of(bytes));
  ASSERT_TRUE(static_cast<bool>(loaded)) << "read_xlsb failed: " << loaded.error().message;
  const Workbook& before = loaded.value().workbook;

  // Baseline: the fixture resolves the expected number formats.
  ASSERT_EQ(ResolvedNumFmtId(before, 0, 3), 179U);  // D1 yyyy/mm/dd
  ASSERT_EQ(ResolvedNumFmtId(before, 1, 3), 4U);    // D2 #,##0.00
  ASSERT_EQ(ResolvedNumFmtId(before, 4, 3), 180U);  // D5 0.0%
  const Cell* d3_before = before.sheet(0).cell_at(2, 3);
  ASSERT_NE(d3_before, nullptr);
  ASSERT_NE(d3_before->xf_index, 0U) << "fixture D3 should carry a non-default style";

  auto saved = io::xlsb::write_xlsb(before);
  ASSERT_TRUE(static_cast<bool>(saved)) << "write_xlsb failed: " << saved.error().message;
  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << "reload failed: " << reloaded.error().message;
  const Workbook& after = reloaded.value().workbook;

  // Number formats still resolve after the round-trip (index + style table).
  EXPECT_EQ(ResolvedNumFmtId(after, 0, 3), 179U);
  EXPECT_EQ(ResolvedNumFmtId(after, 1, 3), 4U);
  EXPECT_EQ(ResolvedNumFmtId(after, 4, 3), 180U);
  // D3's font style index is preserved.
  const Cell* d3_after = after.sheet(0).cell_at(2, 3);
  ASSERT_NE(d3_after, nullptr);
  EXPECT_EQ(d3_after->xf_index, d3_before->xf_index);
}
TEST(XlsbWriteReadSymmetry, LetParameterReferencedInDifferentCaseKeepsItsBinding) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=LET(x,10,X+1)")));  // A1
  auto before_recalc = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(before_recalc)) << before_recalc.error().message;
  const Value a1_before = wb.sheet(0).resolve_cell_value(0U, 0U);
  ASSERT_TRUE(a1_before.is_number());
  ASSERT_EQ(a1_before.as_number(), 11.0);

  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << "write_xlsb failed: " << saved.error().message;
  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << "read_xlsb failed: " << reloaded.error().message;
  Workbook after = std::move(reloaded.value().workbook);

  // No free defined name was invented for `X`: the parameter reference was
  // recognised as in-scope, not treated as an unbound NameRef.
  EXPECT_TRUE(after.defined_names().empty());

  auto after_recalc = after.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(after_recalc)) << after_recalc.error().message;
  const Value a1_after = after.sheet(0).resolve_cell_value(0U, 0U);
  ASSERT_TRUE(a1_after.is_number());
  EXPECT_EQ(a1_after.as_number(), 11.0);
}
TEST(XlsbWriteReadSymmetry, SpillRefSurvivesRoundTrip) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=SEQUENCE(3)")));  // A1, spills to A1:A3
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=SUM(A1#)")));     // B1
  auto before_recalc = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(before_recalc)) << before_recalc.error().message;
  const Value b1_before = wb.sheet(0).resolve_cell_value(0U, 1U);
  ASSERT_TRUE(b1_before.is_number());
  ASSERT_EQ(b1_before.as_number(), 6.0);

  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << "write_xlsb failed: " << saved.error().message;
  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << "read_xlsb failed: " << reloaded.error().message;
  Workbook after = std::move(reloaded.value().workbook);

  auto after_recalc = after.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(after_recalc)) << after_recalc.error().message;
  const Value b1_after = after.sheet(0).resolve_cell_value(0U, 1U);
  ASSERT_TRUE(b1_after.is_number());
  EXPECT_EQ(b1_after.as_number(), 6.0);
}
TEST(XlsbWriteReadSymmetry, NamesSharingTextAcrossScopesKeepTheirOwnOrdinals) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Foo", "Sheet1!$A$1", -1)));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Foo", "Sheet1!$B$1", 0)));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Bar", "Sheet1!$C$1", -1)));
  // Route the edits through the workbook so the dep graph sees them.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(10.0))));  // A1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 1U, Value::number(20.0))));  // B1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 2U, Value::number(30.0))));  // C1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 3U, "=Bar+1")));           // D1
  auto before_recalc = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(before_recalc)) << before_recalc.error().message;
  const Value d1_before = wb.sheet(0).resolve_cell_value(0U, 3U);
  ASSERT_TRUE(d1_before.is_number()) << d1_before.debug_to_string();
  ASSERT_EQ(d1_before.as_number(), 31.0);

  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << "write_xlsb failed: " << saved.error().message;
  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << "read_xlsb failed: " << reloaded.error().message;
  Workbook after = std::move(reloaded.value().workbook);

  // All three names survive, in declaration order, with their scopes.
  const std::vector<DefinedName>& names = after.defined_names();
  ASSERT_EQ(names.size(), 3U);
  EXPECT_EQ(names[0].name, "Foo");
  EXPECT_EQ(names[0].formula, "Sheet1!$A$1");
  EXPECT_EQ(names[0].local_sheet_id, -1);
  EXPECT_EQ(names[1].name, "Foo");
  EXPECT_EQ(names[1].formula, "Sheet1!$B$1");
  EXPECT_EQ(names[1].local_sheet_id, 0);
  EXPECT_EQ(names[2].name, "Bar");
  EXPECT_EQ(names[2].formula, "Sheet1!$C$1");
  EXPECT_EQ(names[2].local_sheet_id, -1);

  // The referencing formula still names `Bar`, and recalculating it on
  // the reloaded workbook lands on C1 + 1 rather than on whatever name
  // ordinal 2 would otherwise have become.
  const Cell* d1 = after.sheet(0).cell_at(0U, 3U);
  ASSERT_NE(d1, nullptr);
  EXPECT_EQ(d1->formula_text, "=Bar+1");
  auto after_recalc = after.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(after_recalc)) << after_recalc.error().message;
  const Value d1_after = after.sheet(0).resolve_cell_value(0U, 3U);
  ASSERT_TRUE(d1_after.is_number());
  EXPECT_EQ(d1_after.as_number(), d1_before.as_number());
}
TEST(XlsbWriteReadSymmetry, UnqualifiedNameEncodesTheOrdinalOfTheScopeExcelResolves) {
  RunScopeResolutionCase(/*local_first=*/false);
}
TEST(XlsbWriteReadSymmetry, UnqualifiedNameScopeIgnoresDeclarationOrder) {
  RunScopeResolutionCase(/*local_first=*/true);
}
TEST(XlsbWriteReadSymmetry, SheetQualifiedNameKeepsItsScope) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_sheet("Sheet2");
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Local", "Sheet2!$A$1", 1)));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1U, 0U, 0U, Value::number(4.0))));   // Sheet2!A1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=Sheet2!Local*2")));  // Sheet1!A1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=ISERROR(Local)")));  // Sheet1!B1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1U, 0U, 1U, "=Local+1")));         // Sheet2!B1

  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << "write_xlsb failed: " << saved.error().message;
  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << "read_xlsb failed: " << reloaded.error().message;
  EXPECT_EQ(reloaded.value().undecoded_formula_count, 0U);
  Workbook after = std::move(reloaded.value().workbook);
  EXPECT_EQ(after.sheet(0).cell_at(0U, 0U)->formula_text, "=Sheet2!Local*2");
  EXPECT_EQ(after.sheet(0).cell_at(0U, 1U)->formula_text, "=ISERROR(Local)");
  EXPECT_EQ(after.sheet(1).cell_at(0U, 1U)->formula_text, "=Local+1");

  auto recalc_or = after.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(recalc_or)) << recalc_or.error().message;
  EXPECT_EQ(after.sheet(0).resolve_cell_value(0U, 0U).as_number(), 8.0);
  EXPECT_TRUE(after.sheet(0).resolve_cell_value(0U, 1U).as_boolean());
  EXPECT_EQ(after.sheet(1).resolve_cell_value(0U, 1U).as_number(), 5.0);
}
TEST(XlsbWriteReadSymmetry, SheetQualifiedLocalBodyScopeRoundTrips) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_sheet("Sheet2");
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Inner", "1")));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Inner", "100", 1)));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Outer", "Inner+1", 1)));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=Sheet2!Outer")));  // Sheet1!A1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=Inner")));         // Sheet1!B1

  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << "write_xlsb failed: " << saved.error().message;
  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << "read_xlsb failed: " << reloaded.error().message;
  EXPECT_EQ(reloaded.value().undecoded_formula_count, 0U);
  Workbook after = std::move(reloaded.value().workbook);
  EXPECT_EQ(after.sheet(0).cell_at(0U, 0U)->formula_text, "=Sheet2!Outer");
  EXPECT_EQ(after.sheet(0).cell_at(0U, 1U)->formula_text, "=Inner");

  auto recalc_or = after.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(recalc_or)) << recalc_or.error().message;
  EXPECT_EQ(after.sheet(0).resolve_cell_value(0U, 0U).as_number(), 101.0);
  EXPECT_EQ(after.sheet(0).resolve_cell_value(0U, 1U).as_number(), 1.0);
}
TEST(XlsbWriteReadSymmetry, ExcelSheetQualifiedLocalNameFixture) {
  const std::vector<std::uint8_t> bytes = test::read_file_bytes(FixturePath("sheet_qualified_local_name.xlsb"));
  ASSERT_FALSE(bytes.empty());
  auto loaded = io::xlsb::read_xlsb(test::span_of(bytes));
  ASSERT_TRUE(static_cast<bool>(loaded)) << "read_xlsb failed: " << loaded.error().message;
  EXPECT_EQ(loaded.value().undecoded_formula_count, 0U);
  Workbook wb = std::move(loaded.value().workbook);
  ASSERT_EQ(wb.sheet_count(), 2U);
  EXPECT_EQ(wb.sheet(0).cell_at(0U, 0U)->formula_text, "=Sheet2!Local");
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(wb.sheet(0).resolve_cell_value(0U, 0U).as_number(), 20.0);

  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << "write_xlsb failed: " << saved.error().message;
  std::string sheet1;
  ASSERT_TRUE(test::extract_part(test::span_of(saved.value()), "xl/worksheets/sheet1.bin", &sheet1));
  // Value-class PtgNameX, ixti 1 (after Local's own Sheet2 entry), ilbl 1:
  // the fixture's own bytes.
  const std::string name_x("\x59\x01\x00\x01\x00\x00\x00", 7);
  EXPECT_NE(sheet1.find(name_x), std::string::npos);

  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << "read_xlsb failed: " << reloaded.error().message;
  EXPECT_EQ(reloaded.value().undecoded_formula_count, 0U);
  Workbook after = std::move(reloaded.value().workbook);
  EXPECT_EQ(after.sheet(0).cell_at(0U, 0U)->formula_text, "=Sheet2!Local");
  ASSERT_TRUE(static_cast<bool>(after.recalc(eval::default_registry())));
  EXPECT_EQ(after.sheet(0).resolve_cell_value(0U, 0U).as_number(), 20.0);
}
TEST(XlsbWriteReadSymmetry, ExcelLambdaNameFixture) {
  const std::vector<std::uint8_t> bytes = test::read_file_bytes(FixturePath("lambda_name.xlsb"));
  ASSERT_FALSE(bytes.empty());
  auto loaded = io::xlsb::read_xlsb(test::span_of(bytes));
  ASSERT_TRUE(static_cast<bool>(loaded)) << "read_xlsb failed: " << loaded.error().message;
  EXPECT_EQ(loaded.value().undecoded_formula_count, 0U);
  Workbook wb = std::move(loaded.value().workbook);

  const auto expect_workbook = [](Workbook& book) {
    const Sheet& sheet = book.sheet(0);
    EXPECT_EQ(sheet.cell_at(0U, 0U)->formula_text, "=Fn(3)");
    EXPECT_EQ(sheet.cell_at(0U, 1U)->formula_text, "=Plus(1,2)");
    EXPECT_EQ(sheet.cell_at(0U, 3U)->formula_text, "=LET(f,LAMBDA(y,y+1),f(2))");
    EXPECT_EQ(sheet.cell_at(0U, 4U)->formula_text, "=LAMBDA(z,z*z)(4)");
    ASSERT_EQ(book.defined_names().size(), 3U);
    EXPECT_EQ(book.defined_names()[0].name, "Fn");
    EXPECT_EQ(book.defined_names()[0].formula, "LAMBDA(x,x*2)");
    EXPECT_EQ(book.defined_names()[2].formula, "LAMBDA(a,b,a+b)");
    ASSERT_TRUE(static_cast<bool>(book.recalc(eval::default_registry())));
    const double expected[] = {6.0, 3.0, 5.0, 3.0, 16.0};
    for (std::uint32_t col = 0; col < 5U; ++col) {
      const Value v = book.sheet(0).resolve_cell_value(0U, col);
      ASSERT_TRUE(v.is_number()) << "col=" << col;
      EXPECT_EQ(v.as_number(), expected[col]) << "col=" << col;
    }
  };
  expect_workbook(wb);

  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << "write_xlsb failed: " << saved.error().message;
  std::string workbook_bin;
  ASSERT_TRUE(test::extract_part(test::span_of(saved.value()), "xl/workbook.bin", &workbook_bin));
  // BrtName for Fn: flags 0x10 (fCalcExp), chKey 0, workbook scope, "Fn".
  const std::string fn_header("\x10\x00\x00\x00\x00\xff\xff\xff\xff\x02\x00\x00\x00\x46\x00\x6e\x00", 17);
  EXPECT_NE(workbook_bin.find(fn_header), std::string::npos);
  std::string sheet1;
  ASSERT_TRUE(test::extract_part(test::span_of(saved.value()), "xl/worksheets/sheet1.bin", &sheet1));
  // A1: PtgName(Fn, ilbl 1), PtgInt 3, PtgFuncVar(cparams 2, iftab 255).
  const std::string fn_call("\x23\x01\x00\x00\x00\x1e\x03\x00\x42\x02\xff\x00", 12);
  EXPECT_NE(sheet1.find(fn_call), std::string::npos);

  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << "read_xlsb failed: " << reloaded.error().message;
  EXPECT_EQ(reloaded.value().undecoded_formula_count, 0U);
  Workbook after = std::move(reloaded.value().workbook);
  expect_workbook(after);
}
TEST(XlsbWriteReadSymmetry, LambdaNamesAndCallsRoundTrip) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Twice", "LAMBDA(x,x*2)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Local", "LAMBDA(x,Twice(x)+Sheet2!$A$1)", 1)));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1U, 0U, 0U, Value::number(100.0))));  // Sheet2!A1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=Twice(3)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=Sheet2!Local(1)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=LAMBDA(a,LAMBDA(b,a-b))(10)(4)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 3U, "=LET(sq,LAMBDA(n,n*n),sq(5)+Twice(1))")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1U, 1U, 0U, "=Local(2)")));  // Sheet2!A2

  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << "write_xlsb failed: " << saved.error().message;
  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << "read_xlsb failed: " << reloaded.error().message;
  EXPECT_EQ(reloaded.value().undecoded_formula_count, 0U);
  Workbook after = std::move(reloaded.value().workbook);
  EXPECT_EQ(after.sheet(0).cell_at(0U, 0U)->formula_text, "=Twice(3)");
  EXPECT_EQ(after.sheet(0).cell_at(0U, 1U)->formula_text, "=Sheet2!Local(1)");
  EXPECT_EQ(after.sheet(0).cell_at(0U, 2U)->formula_text, "=LAMBDA(a,LAMBDA(b,a-b))(10)(4)");
  EXPECT_EQ(after.sheet(0).cell_at(0U, 3U)->formula_text, "=LET(sq,LAMBDA(n,n*n),sq(5)+Twice(1))");
  EXPECT_EQ(after.sheet(1).cell_at(1U, 0U)->formula_text, "=Local(2)");
  ASSERT_EQ(after.defined_names().size(), 2U);
  EXPECT_EQ(after.defined_names()[1].formula, "LAMBDA(x,Twice(x)+Sheet2!$A$1)");
  EXPECT_EQ(after.defined_names()[1].local_sheet_id, 1);

  ASSERT_TRUE(static_cast<bool>(after.recalc(eval::default_registry())));
  EXPECT_EQ(after.sheet(0).resolve_cell_value(0U, 0U).as_number(), 6.0);
  EXPECT_EQ(after.sheet(0).resolve_cell_value(0U, 1U).as_number(), 102.0);
  EXPECT_EQ(after.sheet(0).resolve_cell_value(0U, 2U).as_number(), 6.0);
  EXPECT_EQ(after.sheet(0).resolve_cell_value(0U, 3U).as_number(), 27.0);
  EXPECT_EQ(after.sheet(1).resolve_cell_value(1U, 0U).as_number(), 104.0);
}
TEST(XlsbWriteReadSymmetry, CallToAnotherSheetsLocalLambdaStaysUnresolved) {
  // Fn exists only as Sheet2's local name, so `Fn(3)` on Sheet1 is #NAME?;
  // the encode must not borrow Sheet2's record, which would read back as a
  // resolvable `Sheet2!Fn(3)`.
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Fn", "LAMBDA(x,x*2)", 1)));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=Fn(3)")));

  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << "write_xlsb failed: " << saved.error().message;
  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << "read_xlsb failed: " << reloaded.error().message;
  Workbook after = std::move(reloaded.value().workbook);
  EXPECT_EQ(after.sheet(0).cell_at(0U, 0U)->formula_text, "=Fn(3)");
  ASSERT_TRUE(static_cast<bool>(after.recalc(eval::default_registry())));
  const Value v = after.sheet(0).resolve_cell_value(0U, 0U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}
TEST(XlsbWriteReadSymmetry, ExcelSelfBookNameFixtures) {
  auto from_xlsx = io::read_ooxml(test::span_of(test::read_file_bytes(FixturePath("self_book_name.xlsx"))));
  ASSERT_TRUE(static_cast<bool>(from_xlsx)) << from_xlsx.error().message;
  auto from_xlsb = io::xlsb::read_xlsb(test::span_of(test::read_file_bytes(FixturePath("self_book_name.xlsb"))));
  ASSERT_TRUE(static_cast<bool>(from_xlsb)) << from_xlsb.error().message;
  EXPECT_EQ(from_xlsb.value().undecoded_formula_count, 0U);
  Workbook xlsx = std::move(from_xlsx.value().workbook);
  Workbook xlsb = std::move(from_xlsb.value().workbook);
  ExpectSelfBookFixture(xlsx, "xlsx");
  ExpectSelfBookFixture(xlsb, "xlsb");

  for (Workbook* source : {&xlsx, &xlsb}) {
    auto saved_xlsb = io::xlsb::write_xlsb(*source);
    ASSERT_TRUE(static_cast<bool>(saved_xlsb)) << saved_xlsb.error().message;
    auto reread_xlsb = io::xlsb::read_xlsb(test::span_of(saved_xlsb.value()));
    ASSERT_TRUE(static_cast<bool>(reread_xlsb)) << reread_xlsb.error().message;
    EXPECT_EQ(reread_xlsb.value().undecoded_formula_count, 0U);
    ExpectSelfBookFixture(reread_xlsb.value().workbook, "xlsb round trip");

    auto saved_xlsx = io::write_ooxml(*source);
    ASSERT_TRUE(static_cast<bool>(saved_xlsx)) << saved_xlsx.error().message;
    auto reread_xlsx = io::read_ooxml(test::span_of(saved_xlsx.value()));
    ASSERT_TRUE(static_cast<bool>(reread_xlsx)) << reread_xlsx.error().message;
    ExpectSelfBookFixture(reread_xlsx.value().workbook, "xlsx round trip");
  }
}
TEST(XlsbWriteReadSymmetry, SheetQualifiedWorkbookNameReadsBackAsSelfBook) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("G", "5")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=Sheet1!G")));
  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << reloaded.error().message;
  Workbook after = std::move(reloaded.value().workbook);
  EXPECT_EQ(after.sheet(0).cell_at(0U, 0U)->formula_text, "=[0]!G");
  ASSERT_TRUE(static_cast<bool>(after.recalc(eval::default_registry())));
  EXPECT_EQ(after.sheet(0).resolve_cell_value(0U, 0U).as_number(), 5.0);
}

// Each `#REF!` site form, its xlsb text, and its xlsx (storage) text.
struct RefErrorSiteCase {
  const char* xlsb;
  const char* xlsx;
};

constexpr RefErrorSiteCase kRefErrorSiteCells[] = {
    {"=Sheet3!#REF!*2", "=Sheet3!#REF!*2"},
    {"=SUM(Sheet3!#REF!)", "=SUM(Sheet3!#REF!)"},
    {"=Sheet3!#REF!", "=Sheet3!#REF!"},
    {"=#REF!A1", "=#REF!"},
    {"=SUM(#REF!A1:B2)", "=SUM(#REF!)"},
    {"=SUM(#REF!A:A)", "=SUM(#REF!)"},
    {"=SUM(#REF!1:3)", "=SUM(#REF!)"},
    {"=#REF!$A$1+Sheet3!#REF!", "=#REF!+Sheet3!#REF!"},
    {"=#REF!#REF!", "=#REF!"},
    {"=SUM(#REF!#REF!)+1", "=SUM(#REF!)+1"},
    {"=#REF!", "=#REF!"},
};

struct RefErrorSiteName {
  const char* name;
  const char* xlsb;
  const char* xlsx;
};

constexpr RefErrorSiteName kRefErrorSiteNames[] = {
    {"Nm", "Sheet3!#REF!", "Sheet3!#REF!"},
    {"Nm_2", "#REF!$A$1", "#REF!"},
    {"Nm_3", "#REF!$A$1:$B$2", "#REF!"},
    {"Nm_4", "#REF!#REF!", "#REF!"},
};

void ExpectRefErrorSites(const Workbook& wb, bool storage, const char* label) {
  SCOPED_TRACE(label);
  std::uint32_t row = 0;
  for (const RefErrorSiteCase& c : kRefErrorSiteCells) {
    const Cell* cell = wb.sheet(0).cell_at(row, 3U);
    ASSERT_NE(cell, nullptr) << row;
    EXPECT_EQ(cell->formula_text, storage ? c.xlsx : c.xlsb) << row;
    ++row;
  }
  for (const RefErrorSiteName& n : kRefErrorSiteNames) {
    const auto it = std::find_if(wb.defined_names().begin(), wb.defined_names().end(),
                                 [&n](const DefinedName& dn) { return dn.name == n.name; });
    ASSERT_NE(it, wb.defined_names().end()) << n.name;
    EXPECT_EQ(it->formula, storage ? n.xlsx : n.xlsb) << n.name;
  }
}

TEST(XlsbWriteReadSymmetry, RefErrorSitesRoundTrip) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_sheet("Sheet3");
  std::uint32_t row = 0;
  for (const RefErrorSiteCase& c : kRefErrorSiteCells) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 3U, c.xlsb))) << c.xlsb;
    ++row;
  }
  for (const RefErrorSiteName& n : kRefErrorSiteNames) {
    ASSERT_TRUE(static_cast<bool>(wb.set_defined_name(n.name, n.xlsb))) << n.name;
  }

  auto saved = io::xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  EXPECT_EQ(saved.value().diagnostics.downgraded_formula_count, 0U);
  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value().bytes));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << reloaded.error().message;
  EXPECT_EQ(reloaded.value().undecoded_formula_count, 0U);
  const Workbook after = std::move(reloaded.value().workbook);
  ExpectRefErrorSites(after, /*storage=*/false, "xlsb round trip");

  // A deleted sheet's address does not survive xlsx: its text is bare `#REF!`.
  auto saved_xlsx = io::write_ooxml(after);
  ASSERT_TRUE(static_cast<bool>(saved_xlsx)) << saved_xlsx.error().message;
  auto reread_xlsx = io::read_ooxml(test::span_of(saved_xlsx.value()));
  ASSERT_TRUE(static_cast<bool>(reread_xlsx)) << reread_xlsx.error().message;
  ExpectRefErrorSites(reread_xlsx.value().workbook, /*storage=*/true, "xlsx re-save");
}

// `[Book]Sheet` with no `!` is a name with no definition: #NAME?, no link,
// and the same text after either format's round trip.
void ExpectBracketNames(Workbook& wb, const char* label) {
  SCOPED_TRACE(label);
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_TRUE(wb.external_links().empty());
  EXPECT_TRUE(wb.defined_names().empty());
  for (std::uint32_t row : {0U, 1U}) {
    const Value v = wb.sheet(0).resolve_cell_value(row, 0U);
    ASSERT_TRUE(v.is_error()) << row;
    EXPECT_EQ(v.as_error(), ErrorCode::Name) << row;
  }
  const Value isref = wb.sheet(0).resolve_cell_value(2U, 0U);
  ASSERT_TRUE(isref.is_boolean());
  EXPECT_FALSE(isref.as_boolean());
  EXPECT_EQ(wb.sheet(0).cell_at(0U, 0U)->formula_text, "=[Book.xlsx]Sheet");
  EXPECT_EQ(wb.sheet(0).cell_at(1U, 0U)->formula_text, "=[Book.xlsx]Sheet+1");
  EXPECT_EQ(wb.sheet(0).cell_at(2U, 0U)->formula_text, "=ISREF([Src.xlsx]Data)");
}

TEST(XlsbRoundTripDefinedNames, BracketNameWithoutBangRoundTripsAsAnUndefinedName) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=[Book.xlsx]Sheet")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=[Book.xlsx]Sheet+1")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 2U, 0U, "=ISREF([Src.xlsx]Data)")));
  ExpectBracketNames(wb, "entered");

  auto saved_xlsb = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved_xlsb)) << saved_xlsb.error().message;
  auto from_xlsb = io::xlsb::read_xlsb(test::span_of(saved_xlsb.value()));
  ASSERT_TRUE(static_cast<bool>(from_xlsb)) << from_xlsb.error().message;
  EXPECT_EQ(from_xlsb.value().undecoded_formula_count, 0U);
  ExpectBracketNames(from_xlsb.value().workbook, "xlsb round trip");

  auto saved_xlsx = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(saved_xlsx)) << saved_xlsx.error().message;
  auto from_xlsx = io::read_ooxml(test::span_of(saved_xlsx.value()));
  ASSERT_TRUE(static_cast<bool>(from_xlsx)) << from_xlsx.error().message;
  ExpectBracketNames(from_xlsx.value().workbook, "xlsx round trip");
}

// A book whose name, CF and DV formulas hold `formula` as stored text.
Workbook BookWithFeatureFormula(const std::string& formula) {
  Workbook wb = Workbook::create();
  EXPECT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Rate", formula, -1)));
  cf::ConditionalFormat format;
  format.sqref.push_back(cf::CFCellRange{{0U, 0U}, {9U, 0U}});
  cf::CFRule rule;
  rule.type = cf::RuleType::Expression;
  rule.formula1 = formula;
  format.rules.push_back(std::move(rule));
  wb.sheet(0).mutable_conditional_formats().push_back(std::move(format));
  DataValidation validation;
  validation.ranges.push_back(MergeRange{0U, 1U, 9U, 1U});
  validation.type = 7;  // custom
  validation.formula1 = formula;
  wb.sheet(0).mutable_validations().push_back(std::move(validation));
  return wb;
}

void ExpectFeatureFormula(const Workbook& wb, const std::string& expected, const char* context) {
  SCOPED_TRACE(context);
  ASSERT_EQ(wb.defined_names().size(), 1U);
  EXPECT_EQ(wb.defined_names()[0].formula, expected);
  ASSERT_EQ(wb.sheet(0).conditional_formats().size(), 1U);
  ASSERT_EQ(wb.sheet(0).conditional_formats()[0].rules.size(), 1U);
  EXPECT_EQ(wb.sheet(0).conditional_formats()[0].rules[0].formula1.value_or(""), expected);
  ASSERT_EQ(wb.sheet(0).validations().size(), 1U);
  EXPECT_EQ(wb.sheet(0).validations()[0].formula1, expected);
}

TEST(FeatureFormulaSpelling, XlsxLoadRespellsNamesConditionalFormatsAndValidations) {
  const Workbook wb = BookWithFeatureFormula("SUM( +1.50 , 1E3 )-0+ +5");
  auto saved = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  auto loaded = io::read_ooxml(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  ExpectFeatureFormula(loaded.value().workbook, "SUM( 1.5, 1000 )-0+ 5", "xlsx");
}

TEST(FeatureFormulaSpelling, XlsbLoadKeepsTheRenderedSigns) {
  // A sign Excel kept against a literal is not absorbed again on load.
  const Workbook wb = BookWithFeatureFormula("-0+1");
  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  auto loaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  ExpectFeatureFormula(loaded.value().workbook, "-0+1", "xlsb");
}

// `BrtName` payloads of `xl/workbook.bin` in `pkg`, keyed by name.
std::vector<std::pair<std::string, std::string>> NamePayloads(const std::vector<std::uint8_t>& pkg) {
  std::vector<std::pair<std::string, std::string>> out;
  std::string bin;
  EXPECT_TRUE(test::extract_part(test::span_of(pkg), "xl/workbook.bin", &bin));
  const std::vector<std::uint8_t> bytes(bin.begin(), bin.end());
  io::ByteSpan cursor = test::span_of(bytes);
  while (cursor.size > 0U) {
    auto rec_or = io::xlsb::read_record(cursor);
    if (!rec_or) {
      break;
    }
    if (rec_or.value().type != static_cast<std::uint16_t>(io::xlsb::XlsbRecordType::BrtName)) {
      continue;
    }
    const io::ByteSpan payload = rec_or.value().payload;
    io::ByteSpan p = payload;
    p.data += 9;
    p.size -= 9;
    auto name = io::xlsb::read_xlwidestring(p);
    out.emplace_back(name ? name.value() : std::string(),
                     std::string(reinterpret_cast<const char*>(payload.data), payload.size));
  }
  return out;
}

std::string PayloadOf(const std::vector<std::pair<std::string, std::string>>& payloads, const std::string& name) {
  for (const auto& entry : payloads) {
    if (entry.first == name) {
      return entry.second;
    }
  }
  ADD_FAILURE() << "no BrtName " << name;
  return {};
}

// The defined name `name` of `wb`; a failure, and the first name, when it has none.
const DefinedName& NamedDefinedName(const Workbook& wb, const char* name) {
  for (const DefinedName& entry : wb.defined_names()) {
    if (entry.name == name) {
      return entry;
    }
  }
  ADD_FAILURE() << "no defined name " << name;
  return wb.defined_names().front();
}

// Excel's own .xlsb saves of a hidden name Foo (=Sheet1!$A$1, A1 = 7) flagged as a macro name, beside a plain name
// Bar; Z1 holds =Foo. A macro name is live, keeps its kind and its record through an .xlsb save, and is left out of
// an .xlsx save, as Excel's own saves do.
TEST(XlsbWriteReadSymmetry, ExcelMacroNamesKeepTheirKindInXlsbAndLeaveXlsx) {
  struct Case {
    const char* fixture;
    NameMacroKind kind;
    bool future;
  };
  for (const Case& c : {Case{"macro_name_command.xlsb", NameMacroKind::kCommand, false},
                        Case{"macro_name_function.xlsb", NameMacroKind::kFunction, false},
                        Case{"macro_name_command_future.xlsb", NameMacroKind::kCommand, true}}) {
    SCOPED_TRACE(c.fixture);
    const std::vector<std::uint8_t> bytes = test::read_file_bytes(FixturePath(c.fixture));
    ASSERT_FALSE(bytes.empty());
    auto loaded = io::xlsb::read_xlsb(test::span_of(bytes));
    ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
    EXPECT_EQ(loaded.value().skipped_feature_count, 0U);
    Workbook wb = std::move(loaded.value().workbook);
    ASSERT_EQ(wb.defined_names().size(), 2U);
    const DefinedName& foo = NamedDefinedName(wb, "Foo");
    EXPECT_TRUE(foo.hidden);
    EXPECT_EQ(foo.macro_kind, c.kind);
    EXPECT_EQ(foo.macro_future_function, c.future);
    EXPECT_EQ(NamedDefinedName(wb, "Bar").macro_kind, NameMacroKind::kNone);
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(9.0))));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(wb.sheet(0).resolve_cell_value(0U, 25U).as_number(), 9.0);

    auto saved = io::xlsb::write_xlsb(wb);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    const auto excel = NamePayloads(bytes);
    const auto ours = NamePayloads(saved.value());
    EXPECT_EQ(PayloadOf(ours, "Foo"), PayloadOf(excel, "Foo"));
    EXPECT_EQ(PayloadOf(ours, "Bar"), PayloadOf(excel, "Bar"));
    auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
    ASSERT_TRUE(static_cast<bool>(reloaded)) << reloaded.error().message;
    EXPECT_EQ(NamedDefinedName(reloaded.value().workbook, "Foo").macro_kind, c.kind);

    auto xlsx = io::write_ooxml_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
    EXPECT_EQ(xlsx.value().diagnostics.deferred_feature_count, 1U);
    std::string workbook_xml;
    ASSERT_TRUE(test::extract_part(test::span_of(xlsx.value().bytes), "xl/workbook.xml", &workbook_xml));
    EXPECT_EQ(workbook_xml.find("\"Foo\""), std::string::npos) << workbook_xml;
    EXPECT_NE(workbook_xml.find("<definedName name=\"Bar\">"), std::string::npos) << workbook_xml;
  }
}

// Editing a macro name's formula or hidden flag keeps its kind, as Excel's RefersTo and Visible edits do; removing
// it and setting it again leaves a plain name, as Excel's Names.Add does.
TEST(XlsbWriteReadSymmetry, MacroNameEditsKeepItsKindAndRemoveThenSetLeavesAPlainName) {
  const std::vector<std::uint8_t> bytes = test::read_file_bytes(FixturePath("macro_name_command.xlsb"));
  ASSERT_FALSE(bytes.empty());
  auto loaded = io::xlsb::read_xlsb(test::span_of(bytes));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  Workbook wb = std::move(loaded.value().workbook);
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_hidden("Foo", -1, false)));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_comment("Foo", -1, "note")));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Foo", "Sheet1!$B$1")));
  EXPECT_EQ(NamedDefinedName(wb, "Foo").macro_kind, NameMacroKind::kCommand);
  EXPECT_FALSE(NamedDefinedName(wb, "Foo").hidden);
  EXPECT_EQ(NamedDefinedName(wb, "Foo").comment, "note");
  // Excel's save after that edit writes fHidden off, fProc and fGrp 14: 0x388.
  auto edited = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(edited)) << edited.error().message;
  const std::string edited_record = PayloadOf(NamePayloads(edited.value()), "Foo");
  ASSERT_GE(edited_record.size(), 4U);
  EXPECT_EQ(edited_record.substr(0, 4), std::string("\x88\x03\x00\x00", 4));
  auto kept = io::write_ooxml_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(kept)) << kept.error().message;
  EXPECT_EQ(kept.value().diagnostics.deferred_feature_count, 1U);

  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Foo", "")));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Foo", "Sheet1!$B$1")));
  EXPECT_EQ(NamedDefinedName(wb, "Foo").macro_kind, NameMacroKind::kNone);
  auto xlsb = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
  const std::string record = PayloadOf(NamePayloads(xlsb.value()), "Foo");
  ASSERT_GE(record.size(), 4U);
  EXPECT_EQ(record.substr(0, 4), std::string("\x00\x00\x00\x00", 4));
  EXPECT_EQ(record.size(), PayloadOf(NamePayloads(bytes), "Bar").size());
  auto xlsx = io::write_ooxml_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  EXPECT_EQ(xlsx.value().diagnostics.deferred_feature_count, 0U);
  std::string workbook_xml;
  ASSERT_TRUE(test::extract_part(test::span_of(xlsx.value().bytes), "xl/workbook.xml", &workbook_xml));
  EXPECT_NE(workbook_xml.find("<definedName name=\"Foo\">Sheet1!$B$1</definedName>"), std::string::npos)
      << workbook_xml;
}

// A lone fFunc flag names no macro: Excel opens such a book and its save drops the flag, so the load drops it and
// counts it.
TEST(XlsbWriteReadSymmetry, FunctionFlagWithoutAMacroIsDroppedAndCounted) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Foo", "Sheet1!$A$1")));
  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  std::vector<std::uint8_t> patched = saved.value();
  std::string bin;
  ASSERT_TRUE(test::extract_part(test::span_of(patched), "xl/workbook.bin", &bin));
  // The BrtName record of Foo: grbit 0, chKey 0, workbook scope, cch 3, "Foo".
  const std::string plain(
      "\x00\x00\x00\x00\x00\xff\xff\xff\xff\x03\x00\x00\x00"
      "F\x00o\x00o\x00",
      19);
  const std::size_t at = bin.find(plain);
  ASSERT_NE(at, std::string::npos);
  bin[at] = '\x02';  // fFunc
  patched = test::replace_part(test::span_of(patched), "xl/workbook.bin", bin);
  ASSERT_FALSE(patched.empty());
  auto loaded = io::xlsb::read_xlsb(test::span_of(patched));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  EXPECT_EQ(loaded.value().skipped_feature_count, 1U);
  ASSERT_EQ(loaded.value().workbook.defined_names().size(), 1U);
  EXPECT_EQ(loaded.value().workbook.defined_names()[0].macro_kind, NameMacroKind::kNone);
}

// The first four bytes of a BrtName payload, the flags word.
std::uint32_t NameFlags(const std::string& payload) {
  EXPECT_GE(payload.size(), 4U);
  if (payload.size() < 4U) {
    return 0U;
  }
  return static_cast<std::uint32_t>(static_cast<unsigned char>(payload[0])) |
         static_cast<std::uint32_t>(static_cast<unsigned char>(payload[1])) << 8U |
         static_cast<std::uint32_t>(static_cast<unsigned char>(payload[2])) << 16U |
         static_cast<std::uint32_t>(static_cast<unsigned char>(payload[3])) << 24U;
}

Workbook LoadXlsbFixture(const char* name) {
  const std::vector<std::uint8_t> bytes = test::read_file_bytes(FixturePath(name));
  EXPECT_FALSE(bytes.empty()) << name;
  auto loaded = io::xlsb::read_xlsb(test::span_of(bytes));
  EXPECT_TRUE(static_cast<bool>(loaded)) << name << ": " << (loaded ? "" : loaded.error().message);
  return loaded ? std::move(loaded.value().workbook) : Workbook::create_empty();
}

Workbook LoadXlsxFixture(const char* name) {
  auto loaded = io::read_ooxml(test::span_of(test::read_file_bytes(FixturePath(name))));
  EXPECT_TRUE(static_cast<bool>(loaded)) << name << ": " << (loaded ? "" : loaded.error().message);
  return loaded ? std::move(loaded.value().workbook) : Workbook::create_empty();
}

std::vector<std::string> DefinedNameTexts(const Workbook& wb) {
  std::vector<std::string> out;
  for (const DefinedName& entry : wb.defined_names()) {
    out.push_back(entry.name);
  }
  return out;
}

// Excel's saves of a book with Ctl_8 and the print area, E1 = ROWS(Print_Area): the xlsb holds the print area as
// the English text with fBuiltin. Both formats load the built-in under one name, and the xlsb save writes the
// record as Excel does.
TEST(XlsbBuiltinName, BuiltinFlagMakesTheEnglishTextTheBuiltin) {
  const Workbook xlsb = LoadXlsbFixture("xlsb_builtin_name/builtin.xlsb");
  const Workbook xlsx = LoadXlsxFixture("xlsb_builtin_name/builtin.xlsx");
  const std::vector<std::string> names{"Ctl_8", "_xlnm.Print_Area"};
  EXPECT_EQ(DefinedNameTexts(xlsx), names);
  EXPECT_EQ(DefinedNameTexts(xlsb), names);
  ASSERT_NE(xlsb.sheet(0).cell_at(0U, 4U), nullptr);
  ASSERT_NE(xlsx.sheet(0).cell_at(0U, 4U), nullptr);
  EXPECT_EQ(xlsb.sheet(0).cell_at(0U, 4U)->formula_text, xlsx.sheet(0).cell_at(0U, 4U)->formula_text);

  auto saved = io::xlsb::write_xlsb(xlsb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  const auto excel = NamePayloads(test::read_file_bytes(FixturePath("xlsb_builtin_name/builtin.xlsb")));
  const auto ours = NamePayloads(saved.value());
  EXPECT_EQ(NameFlags(PayloadOf(ours, "Print_Area")), 0x20U);
  EXPECT_EQ(PayloadOf(ours, "Print_Area"), PayloadOf(excel, "Print_Area"));
  EXPECT_EQ(PayloadOf(ours, "Ctl_8"), PayloadOf(excel, "Ctl_8"));
  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << reloaded.error().message;
  EXPECT_EQ(DefinedNameTexts(reloaded.value().workbook), names);
}

// The same English text with fBuiltin cleared is an ordinary name of that spelling, in Excel's saves and ours.
TEST(XlsbBuiltinName, EnglishTextWithoutTheFlagIsAnOrdinaryName) {
  const Workbook xlsb = LoadXlsbFixture("xlsb_builtin_name/ordinary.xlsb");
  const Workbook xlsx = LoadXlsxFixture("xlsb_builtin_name/ordinary.xlsx");
  const std::vector<std::string> names{"Ctl_8", "Print_Area"};
  EXPECT_EQ(DefinedNameTexts(xlsx), names);
  EXPECT_EQ(DefinedNameTexts(xlsb), names);
  ASSERT_NE(xlsb.sheet(0).cell_at(0U, 4U), nullptr);
  ASSERT_NE(xlsx.sheet(0).cell_at(0U, 4U), nullptr);
  EXPECT_EQ(xlsb.sheet(0).cell_at(0U, 4U)->formula_text, xlsx.sheet(0).cell_at(0U, 4U)->formula_text);

  auto saved = io::xlsb::write_xlsb(xlsb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  const auto excel = NamePayloads(test::read_file_bytes(FixturePath("xlsb_builtin_name/ordinary.xlsb")));
  EXPECT_EQ(PayloadOf(NamePayloads(saved.value()), "Print_Area"), PayloadOf(excel, "Print_Area"));
}

// fBuiltin on a name that is no built-in is dropped on load, and the save writes the record without it, as Excel's
// save of the same book does.
TEST(XlsbBuiltinName, FlagOnAnotherNameIsDropped) {
  const Workbook wb = LoadXlsbFixture("xlsb_builtin_name/flag_on_other_name.xlsb");
  EXPECT_EQ(DefinedNameTexts(wb), (std::vector<std::string>{"Ctl_8", "_xlnm.Print_Area"}));
  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  const auto excel =
      NamePayloads(test::read_file_bytes(FixturePath("xlsb_builtin_name/flag_on_other_name_saved.xlsb")));
  EXPECT_EQ(PayloadOf(NamePayloads(saved.value()), "Ctl_8"), PayloadOf(excel, "Ctl_8"));
}

// Excel stores the AutoFilter's `_FilterDatabase` with fHidden and fBuiltin: it loads as the model's name, so
// setting the filter again updates that name rather than adding a second one, and the save keeps Excel's record.
TEST(XlsbBuiltinName, FilterDatabaseLoadsAsTheAutoFilterName) {
  const char* fixture = "afilter/equal/A_eq_B.xlsb";
  Workbook wb = LoadXlsbFixture(fixture);
  ASSERT_EQ(DefinedNameTexts(wb), (std::vector<std::string>{"_xlnm._FilterDatabase"}));
  EXPECT_TRUE(wb.defined_names()[0].hidden);
  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  const auto excel = NamePayloads(test::read_file_bytes(FixturePath(fixture)));
  EXPECT_EQ(NameFlags(PayloadOf(NamePayloads(saved.value()), "_FilterDatabase")), 0x21U);
  EXPECT_EQ(PayloadOf(NamePayloads(saved.value()), "_FilterDatabase"), PayloadOf(excel, "_FilterDatabase"));

  ASSERT_NE(wb.sheet(0).auto_filter(), nullptr);
  AutoFilter filter = *wb.sheet(0).auto_filter();
  ASSERT_TRUE(static_cast<bool>(wb.set_sheet_auto_filter(0U, std::move(filter))));
  EXPECT_EQ(DefinedNameTexts(wb), (std::vector<std::string>{"_xlnm._FilterDatabase"}));
}

// A name record spelled with the `_xlnm.` prefix and no fBuiltin, as engine saves wrote built-ins before the flag
// was written: Excel lists an ordinary name of that literal text, keeps the record as read in its xlsb save, and
// writes its xlsx save in the built-in's spelling. The engine loads it as the built-in, which the xlsx save cannot
// tell apart (divergence xlsb_prefixed_name_without_builtin_flag_loads_as_builtin).
TEST(XlsbBuiltinName, PrefixedTextWithoutTheFlagLoadsAsTheBuiltin) {
  struct Case {
    const char* input;
    const char* excel_xlsx;
    const char* excel_xlsb;
    const char* builtin;
    std::uint32_t excel_flags;
    std::uint32_t our_flags;
  };
  for (const Case& c :
       {Case{"xlsb_builtin_name/prefixed_print_area.xlsb", "xlsb_builtin_name/prefixed_print_area_saved.xlsx",
             "xlsb_builtin_name/prefixed_print_area_saved.xlsb", "Print_Area", 0x00U, 0x20U},
        Case{"xlsb_builtin_name/prefixed_filter.xlsb", "xlsb_builtin_name/prefixed_filter_saved.xlsx",
             "xlsb_builtin_name/prefixed_filter_saved.xlsb", "_FilterDatabase", 0x01U, 0x21U}}) {
    SCOPED_TRACE(c.input);
    const Workbook wb = LoadXlsbFixture(c.input);
    std::vector<std::string> ours = DefinedNameTexts(wb);
    std::vector<std::string> excel = DefinedNameTexts(LoadXlsxFixture(c.excel_xlsx));
    std::sort(ours.begin(), ours.end());
    std::sort(excel.begin(), excel.end());
    EXPECT_EQ(ours, excel);
    const std::string stored = std::string("_xlnm.") + c.builtin;
    EXPECT_NE(std::find(ours.begin(), ours.end(), stored), ours.end());

    auto saved = io::xlsb::write_xlsb(wb);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    EXPECT_EQ(NameFlags(PayloadOf(NamePayloads(saved.value()), c.builtin)), c.our_flags);
    EXPECT_EQ(NameFlags(PayloadOf(NamePayloads(test::read_file_bytes(FixturePath(c.excel_xlsb))), stored)),
              c.excel_flags);
  }
}

// A print area set through the engine is written as Excel writes a built-in, and reads back as the built-in.
TEST(XlsbBuiltinName, EnginePrintAreaIsWrittenAsABuiltin) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Print_Area", "Sheet1!$A$1:$B$2")));
  ASSERT_EQ(DefinedNameTexts(wb), (std::vector<std::string>{"_xlnm.Print_Area"}));
  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  const std::string record = PayloadOf(NamePayloads(saved.value()), "Print_Area");
  EXPECT_EQ(NameFlags(record), 0x20U);
  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << reloaded.error().message;
  const Workbook& after = reloaded.value().workbook;
  ASSERT_EQ(DefinedNameTexts(after), (std::vector<std::string>{"_xlnm.Print_Area"}));
  EXPECT_EQ(after.defined_names()[0].local_sheet_id, 0);
}

// A bare union keeps Excel's stored text through both save formats, and a
// union written in parentheses keeps them.
TEST(XlsbWriteReadSymmetry, UnionNameBodiesKeepTheirSpelling) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("u", "Sheet1!$A$1,Sheet1!$B$1,Sheet1!$A$2")));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("pu", "(Sheet1!$A$1,Sheet1!$B$1)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("x", "Sheet1!$A$1:$B$2 Sheet1!$B$1:$B$3,Sheet1!$A$1")));
  const auto check = [](const Workbook& after, const char* format) {
    std::size_t seen = 0;
    for (const DefinedName& name : after.defined_names()) {
      if (name.name == "u") {
        EXPECT_EQ(name.formula, "Sheet1!$A$1,Sheet1!$B$1,Sheet1!$A$2") << format;
      } else if (name.name == "pu") {
        EXPECT_EQ(name.formula, "(Sheet1!$A$1,Sheet1!$B$1)") << format;
      } else if (name.name == "x") {
        EXPECT_EQ(name.formula, "Sheet1!$A$1:$B$2 Sheet1!$B$1:$B$3,Sheet1!$A$1") << format;
      } else {
        continue;
      }
      ++seen;
    }
    EXPECT_EQ(seen, 3U) << format;
  };
  auto saved_xlsb = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved_xlsb)) << saved_xlsb.error().message;
  auto from_xlsb = io::xlsb::read_xlsb(test::span_of(saved_xlsb.value()));
  ASSERT_TRUE(static_cast<bool>(from_xlsb)) << from_xlsb.error().message;
  check(from_xlsb.value().workbook, "xlsb");
  auto saved_xlsx = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(saved_xlsx)) << saved_xlsx.error().message;
  auto from_xlsx = io::read_ooxml(test::span_of(saved_xlsx.value()));
  ASSERT_TRUE(static_cast<bool>(from_xlsx)) << from_xlsx.error().message;
  check(from_xlsx.value().workbook, "xlsx");
}

}  // namespace
}  // namespace formulon
