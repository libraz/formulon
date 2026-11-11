//
// CELL("filename") and CELL("address") against the workbook's location. The
// mac rows are Mac Excel 365 measurements (a book in `.../a b Ü/` with sheets
// Sheet1, Sheet2 and `S 1`, `S 1` active); the win rows are an estimate that
// renders the folder with Windows separators.

#include <cstdint>
#include <initializer_list>
#include <optional>
#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_profile_scope.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// "a b Ü" / "Book Ü.xlsx", Ü as NFC U+00DC.
constexpr const char* kDirectory = "/U/x/a b \xC3\x9C";
constexpr const char* kFile = "Book \xC3\x9C.xlsx";

// Evaluates `src` on `current` with the formula anchored at A1 and returns
// the text result ("<not text>" otherwise).
std::string TextOf(std::string_view src, const Workbook& wb, const Sheet& current) {
  Arena parse_arena;
  Arena eval_arena;
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << src;
  if (root == nullptr) {
    return "<parse failed>";
  }
  EvalState state;
  const EvalContext ctx = EvalContext(wb, current, state).with_formula_cell(0U, 0U);
  const Value v = evaluate(*root, eval_arena, default_registry(), ctx);
  return v.is_text() ? std::string(v.as_text()) : std::string("<not text>");
}

// Sheet1, Sheet2 and `S 1`, with `S 1` active.
Workbook ThreeSheetBook(const ExcelProfile& profile) {
  Workbook wb = Workbook::create_for_profile(profile);
  wb.add_sheet("Sheet2");
  wb.add_sheet("S 1");
  wb.set_active_sheet(2U);
  return wb;
}

TEST(CellLocation, FilenameIsEmptyWithoutAFolder) {
  const ExcelProfile profile = mac_365_ja_jp_profile();
  const ScopedEvalProfile scope(profile);
  Workbook wb = ThreeSheetBook(profile);
  EXPECT_EQ(TextOf("=CELL(\"filename\")", wb, wb.sheet(0)), "");
  EXPECT_EQ(TextOf("=CELL(\"filename\",A1)", wb, wb.sheet(0)), "");

  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{std::nullopt, "b.xlsx"})));
  EXPECT_EQ(TextOf("=CELL(\"filename\")", wb, wb.sheet(0)), "");
  EXPECT_EQ(TextOf("=CELL(\"filename\",Sheet2!A1)", wb, wb.sheet(0)), "");
}

TEST(CellLocation, FilenameNamesFolderFileAndSheetOnMac) {
  const ExcelProfile profile = mac_365_ja_jp_profile();
  const ScopedEvalProfile scope(profile);
  Workbook wb = ThreeSheetBook(profile);
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{std::string(kDirectory), kFile})));
  const std::string prefix = std::string(kDirectory) + "/[" + kFile + "]";
  EXPECT_EQ(TextOf("=CELL(\"filename\",A1)", wb, wb.sheet(0)), prefix + "Sheet1");
  // No reference: the active sheet, not the formula's.
  EXPECT_EQ(TextOf("=CELL(\"filename\")", wb, wb.sheet(0)), prefix + "S 1");
  EXPECT_EQ(TextOf("=CELL(\"filename\",Sheet2!A1)", wb, wb.sheet(0)), prefix + "Sheet2");
  // The sheet name is not quoted.
  EXPECT_EQ(TextOf("=CELL(\"filename\",'S 1'!A1)", wb, wb.sheet(0)), prefix + "S 1");
  EXPECT_EQ(TextOf("=CELL(\"filename\",A1)", wb, wb.sheet(2)), prefix + "S 1");

  // The root is not followed by a second separator.
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{std::string("/"), "b.xlsx"})));
  EXPECT_EQ(TextOf("=CELL(\"filename\",A1)", wb, wb.sheet(0)), "/[b.xlsx]Sheet1");
}

// Estimated Windows form (not measured): win-* profiles render `\` separators.
TEST(CellLocation, FilenameUsesWindowsSeparatorsOnWin) {
  const ExcelProfile profile = win_365_ja_jp_profile();
  const ScopedEvalProfile scope(profile);
  Workbook wb = ThreeSheetBook(profile);
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{std::string("C:\\d"), "b.xlsx"})));
  EXPECT_EQ(TextOf("=CELL(\"filename\",A1)", wb, wb.sheet(0)), "C:\\d\\[b.xlsx]Sheet1");
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{std::string("C:\\"), "b.xlsx"})));
  EXPECT_EQ(TextOf("=CELL(\"filename\",A1)", wb, wb.sheet(0)), "C:\\[b.xlsx]Sheet1");
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{std::string("\\\\srv\\sh\\d"), "b.xlsx"})));
  EXPECT_EQ(TextOf("=CELL(\"filename\",A1)", wb, wb.sheet(0)), "\\\\srv\\sh\\d\\[b.xlsx]Sheet1");
}

TEST(CellLocation, FilenameIsEmptyWithoutAWorkbook) {
  const ScopedEvalProfile scope(mac_365_ja_jp_profile());
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{std::string("/d"), "b.xlsx"})));
  Arena parse_arena;
  Arena eval_arena;
  parser::Parser p("=CELL(\"filename\",A1)", parse_arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  EvalState state;
  const EvalContext ctx = EvalContext(wb.sheet(0), state).with_formula_cell(0U, 0U);
  const Value v = evaluate(*root, eval_arena, default_registry(), ctx);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "");
}

struct AddressRow {
  const char* formula;
  const char* expected;
};

// Formulas on Sheet1 of a book with Sheet1, Sheet2 and `S 2`.
void ExpectAddresses(const std::optional<WorkbookLocation>& location, std::initializer_list<AddressRow> rows) {
  const ExcelProfile profile = mac_365_ja_jp_profile();
  const ScopedEvalProfile scope(profile);
  Workbook wb = Workbook::create_for_profile(profile);
  wb.add_sheet("Sheet2");
  wb.add_sheet("S 2");
  ASSERT_TRUE(static_cast<bool>(wb.set_location(location)));
  for (const AddressRow& row : rows) {
    EXPECT_EQ(TextOf(row.formula, wb, wb.sheet(0)), row.expected)
        << row.formula << " in " << (location ? location->file_name : std::string("<unset>"));
  }
}

TEST(CellLocation, AddressPrefixesOtherSheetsWithTheFileName) {
  // Unset: no window-name prefix.
  ExpectAddresses(std::nullopt, {
                                    {"=CELL(\"address\",A1)", "$A$1"},
                                    {"=CELL(\"address\",Sheet1!A1)", "$A$1"},
                                    {"=CELL(\"address\",Sheet2!A1)", "Sheet2!$A$1"},
                                    {"=CELL(\"address\",'S 2'!A1)", "'S 2'!$A$1"},
                                });
  ExpectAddresses(WorkbookLocation{std::string("/p q"), "Book.xlsx"},
                  {
                      {"=CELL(\"address\",A1)", "$A$1"},
                      {"=CELL(\"address\",Sheet1!A1)", "$A$1"},
                      {"=CELL(\"address\",Sheet2!A1)", "[Book.xlsx]Sheet2!$A$1"},
                      {"=CELL(\"address\",'S 2'!A1)", "'[Book.xlsx]S 2'!$A$1"},
                  });
  // A book name that needs quoting quotes the whole `[book]sheet`.
  ExpectAddresses(WorkbookLocation{std::string("/p q"), "Book 1.xlsx"},
                  {
                      {"=CELL(\"address\",A1)", "$A$1"},
                      {"=CELL(\"address\",Sheet1!A1)", "$A$1"},
                      {"=CELL(\"address\",Sheet2!A1)", "'[Book 1.xlsx]Sheet2'!$A$1"},
                      {"=CELL(\"address\",'S 2'!A1)", "'[Book 1.xlsx]S 2'!$A$1"},
                  });
  // A name without a folder still spells the file.
  ExpectAddresses(WorkbookLocation{std::nullopt, "Book.xlsx"},
                  {
                      {"=CELL(\"address\",Sheet1!A1)", "$A$1"},
                      {"=CELL(\"address\",Sheet2!A1)", "[Book.xlsx]Sheet2!$A$1"},
                  });
}

// "Same sheet" is the formula's sheet, whichever sheet is active.
TEST(CellLocation, AddressSameSheetIsTheFormulasSheetNotTheActiveOne) {
  const ExcelProfile profile = mac_365_ja_jp_profile();
  const ScopedEvalProfile scope(profile);
  Workbook wb = Workbook::create_for_profile(profile);
  wb.add_sheet("Sheet2");
  wb.set_active_sheet(0U);
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{std::string("/p q"), "Book.xlsx"})));
  EXPECT_EQ(TextOf("=CELL(\"address\",Sheet2!A1)", wb, wb.sheet(1)), "$A$1");
  EXPECT_EQ(TextOf("=CELL(\"address\",Sheet1!A1)", wb, wb.sheet(1)), "[Book.xlsx]Sheet1!$A$1");
}

}  // namespace
}  // namespace eval
}  // namespace formulon
