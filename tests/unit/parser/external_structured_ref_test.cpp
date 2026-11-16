//
// Table references into another workbook (`Book.xlsx!Table[Col]`): parsing,
// display and storage spelling, and how the xlsx and xlsb writers carry them.
// A link caches cells and defined names, never table definitions, so such a
// reference keeps its text and evaluates to #REF! (see external_ref_eval_test).

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "external_link.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "parser/ast.h"
#include "parser/ast_dump.h"
#include "parser/ast_format.h"
#include "parser/ast_format_r1c1.h"
#include "parser/parser.h"
#include "support/roundtrip_symmetry.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace parser {
namespace {

bool HasErrorCode(const std::vector<ParseError>& errors, ParseErrorCode code) {
  for (const ParseError& e : errors) {
    if (e.code == code) {
      return true;
    }
  }
  return false;
}

std::string SpellAsIs(std::string_view name) {
  return std::string(name);
}

std::uint32_t IndexBook(const void* /*ctx*/, std::string_view /*path*/, std::string_view book) {
  return book == "Book.xlsx" ? 1U : 0U;
}

struct Row {
  const char* entered;
  const char* stored;
};

// The spellings Excel stores and shows for a table of another book.
constexpr Row kRows[] = {
    {"=Book.xlsx!Table3[[#This Row],[Resource Name]]", "[1]!Table3[[#This Row],[Resource Name]]"},
    {"=Book.xlsx!Table3[#Data]", "[1]!Table3[#Data]"},
    {"=Book.xlsx!Table1[Interest]", "[1]!Table1[Interest]"},
    {"=Book.xlsx!Table1[[A]:[B]]", "[1]!Table1[[A]:[B]]"},
    {"=SUM(Book.xlsx!Table1[Interest],1)", "SUM([1]!Table1[Interest],1)"},
};

const AstNode* ParseClean(std::string_view formula, Parser& p) {
  const AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << formula;
  EXPECT_TRUE(p.errors().empty()) << formula;
  return root;
}

TEST(ExternalStructuredRef, DisplayAndStorageSpellingsRoundTrip) {
  const ExternalBookIndexer indexer{&IndexBook, nullptr};
  for (const Row& row : kRows) {
    Arena arena;
    Parser p(row.entered, arena);
    const AstNode* root = ParseClean(row.entered, p);
    ASSERT_NE(root, nullptr);
    EXPECT_EQ(format_formula(*root), std::string(row.entered).substr(1)) << row.entered;
    EXPECT_EQ(format_formula_r1c1(*root, 1, 1), std::string(row.entered).substr(1)) << row.entered;
    EXPECT_EQ(format_formula_storage(*root, &SpellAsIs, nullptr, &indexer), row.stored) << row.entered;
  }
}

TEST(ExternalStructuredRef, NodeKeepsTheTableAndThePayloadVerbatim) {
  Arena arena;
  Parser p("=Book.xlsx!Table3[[#This Row],[Resource Name]]", arena);
  const AstNode* root = ParseClean("this row", p);
  ASSERT_NE(root, nullptr);
  ASSERT_EQ(root->kind(), NodeKind::ExternalRef);
  EXPECT_EQ(root->as_external_ref_name(), "Table3");
  EXPECT_EQ(root->as_external_ref_structured(), "[#This Row],[Resource Name]");
  EXPECT_EQ(root->as_external_ref_book(), "Book.xlsx");
  EXPECT_EQ(dump_sexpr(*root), "(external-ref [Book.xlsx] name Table3 structured [#This Row],[Resource Name])");
}

TEST(ExternalStructuredRef, DumpShape) {
  Arena arena;
  Parser p("=Book.xlsx!Tbl[#Data]", arena);
  const AstNode* root = ParseClean("data", p);
  ASSERT_NE(root, nullptr);
  EXPECT_EQ(dump_sexpr(*root), "(external-ref [Book.xlsx] name Tbl structured #Data)");
}

TEST(ExternalStructuredRef, ASheetWrittenBeforeTheTableIsDropped) {
  const ExternalBookIndexer indexer{&IndexBook, nullptr};
  Arena arena;
  Parser p("=[Book.xlsx]Sheet1!Tbl[v]", arena);
  const AstNode* root = ParseClean("sheet", p);
  ASSERT_NE(root, nullptr);
  ASSERT_EQ(root->kind(), NodeKind::ExternalRef);
  EXPECT_TRUE(root->as_external_ref_sheet().empty());
  EXPECT_EQ(format_formula(*root), "Book.xlsx!Tbl[v]");
  EXPECT_EQ(format_formula_storage(*root, &SpellAsIs, nullptr, &indexer), "[1]!Tbl[v]");
}

TEST(ExternalStructuredRef, TheOwnBookTableIsAStructuredRef) {
  Arena arena;
  Parser p("=[0]!Table1[Col]", arena);
  const AstNode* root = ParseClean("own", p);
  ASSERT_NE(root, nullptr);
  ASSERT_EQ(root->kind(), NodeKind::StructuredRef);
  EXPECT_EQ(root->as_structured_ref_table(), "Table1");
  EXPECT_EQ(root->as_structured_ref_column(), "Col");
}

TEST(ExternalStructuredRef, UnbalancedBracketIsReported) {
  Arena arena;
  Parser p("=Book.xlsx!Tbl[#Data", arena);
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::UnbalancedBrackets));
}

TEST(ExternalStructuredRef, SpecifierKeywordsRaiseNoDiagnostic) {
  Arena arena;
  Parser p("=Book.xlsx!Tbl[[#Headers],[#Data],[v]]", arena);
  const AstNode* root = ParseClean("keywords", p);
  ASSERT_NE(root, nullptr);
  EXPECT_EQ(root->as_external_ref_structured(), "[#Headers],[#Data],[v]");
}

}  // namespace
}  // namespace parser

namespace {

Workbook HostWithTableFormula() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Host");
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 0, "=Src.xlsx!Tbl[#Data]")));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 1, 0, "=SUM(Src.xlsx!Tbl[[A]:[B]])")));
  return wb;
}

std::string CellFormula(const Workbook& wb, std::uint32_t row) {
  const Cell* cell = wb.sheet(0).cell_at(row, 0);
  return cell == nullptr ? std::string() : cell->formula_text;
}

// Entry creates the link as it does for a name, and a save and load keep the text.
TEST(ExternalStructuredRefWorkbook, XlsxSaveKeepsTheFormulaText) {
  Workbook wb = HostWithTableFormula();
  ASSERT_EQ(wb.external_links().size(), 1U);
  EXPECT_EQ(wb.external_links()[0].target, "Src.xlsx");
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  auto saved = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  std::string sheet;
  ASSERT_TRUE(test::extract_part(test::span_of(saved.value()), "xl/worksheets/sheet1.xml", &sheet));
  EXPECT_NE(sheet.find(">[1]!Tbl[#Data]</f>"), std::string::npos) << sheet;
  EXPECT_NE(sheet.find(">SUM([1]!Tbl[[A]:[B]])</f>"), std::string::npos) << sheet;

  auto loaded = io::read_ooxml(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  Workbook& reloaded = loaded.value().workbook;
  EXPECT_EQ(CellFormula(reloaded, 0), CellFormula(wb, 0));
  EXPECT_EQ(CellFormula(reloaded, 1), CellFormula(wb, 1));
  ASSERT_TRUE(static_cast<bool>(reloaded.recalc(eval::default_registry())));
  const Cell* cell = reloaded.sheet(0).cell_at(0, 0);
  ASSERT_NE(cell, nullptr);
  ASSERT_TRUE(cell->cached_value.is_error());
  EXPECT_EQ(cell->cached_value.as_error(), ErrorCode::Ref);
}

// A table has no stored form in the link part, which lists names and sheets
// only. It fails like the sibling "external name not in its link's name list":
// the cell is written as its cached literal and the save counts one downgraded
// formula, as for every formula the encoder refuses.
TEST(ExternalStructuredRefWorkbook, XlsbSaveDowngradesTheFormulaToItsValue) {
  Workbook wb = HostWithTableFormula();
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  auto saved = io::xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  EXPECT_EQ(saved.value().diagnostics.downgraded_formula_count, 2U);

  auto loaded = io::xlsb::read_xlsb(test::span_of(saved.value().bytes));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  const Cell* cell = loaded.value().workbook.sheet(0).cell_at(0, 0);
  ASSERT_NE(cell, nullptr);
  EXPECT_TRUE(cell->formula_text.empty());
  ASSERT_TRUE(cell->cached_value.is_error());
  EXPECT_EQ(cell->cached_value.as_error(), ErrorCode::Ref);
}

}  // namespace
}  // namespace formulon
