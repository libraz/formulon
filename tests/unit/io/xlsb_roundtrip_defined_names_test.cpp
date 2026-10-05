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

#include "cell.h"
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

}  // namespace
}  // namespace formulon
