//
// End-to-end tests for the CELL(info_type, [reference]) builtin. CELL
// is a lazy impl (see `eval/cell_lazy.{h,cpp}`) because the optional
// reference and multi-cell-range top-left selection both need raw AST
// access. The tests cover:
//
//   * Reference-dependent keys (address, col, row, contents, type) for
//     various AST shapes (single Ref, RangeOp, omitted reference using
//     the formula cell anchor).
//   * MVP fixed-stub keys (filename, format, color, parentheses,
//     prefix, width) - these will become live once the style subsystem
//     lands.
//   * `protect` - live: reads the referenced cell's xf `locked` flag.
//   * Case-insensitivity of the info_type key.
//   * Error-propagation rules: error info_type, error reference,
//     unknown info_type, arity violations.

#include <cstdint>
#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "io/styles_reader.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Parses `src` and evaluates it through the default function registry
// against a bound workbook + current sheet, with the formula cell
// anchored at `(row, col)`. Many CELL tests need the anchor so the
// 1-arg form (`=CELL("address")`) has a meaningful answer.
Value EvalSourceAt(std::string_view src, const Workbook& wb, const Sheet& current, std::uint32_t row,
                   std::uint32_t col) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  EvalState state;
  const EvalContext ctx = EvalContext(wb, current, state).with_formula_cell(row, col);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

// Convenience overload anchored at A1 (0, 0) for tests that don't care
// about the anchor location.
Value EvalSourceIn(std::string_view src, const Workbook& wb, const Sheet& current) {
  return EvalSourceAt(src, wb, current, 0U, 0U);
}

// Convenience overload with no bound workbook. Used to confirm that the
// 1-arg form surfaces #REF! when no formula cell is anchored.
Value EvalSourceNoCtx(std::string_view src) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, eval_arena);
}

// ---------------------------------------------------------------------------
// row / col / address - reference-dependent keys
// ---------------------------------------------------------------------------

TEST(BuiltinsCellRow, RefAtA5IsFive) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"row\", A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsCellCol, RefAtC5IsThree) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"col\", C5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsCellAddress, SingleRefIsAbsoluteA1) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"address\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "$A$1");
}

TEST(BuiltinsCellAddress, RangeUsesTopLeft) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"address\", B5:D7)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "$B$5");
}

TEST(BuiltinsCellAddress, OmittedReferenceUsesFormulaCell) {
  // Formula cell at B7 (row=6, col=1 in 0-based) -> $B$7.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceAt("=CELL(\"address\")", wb, wb.sheet(0), 6U, 1U);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "$B$7");
}

// ---------------------------------------------------------------------------
// contents - reference-dependent
// ---------------------------------------------------------------------------

TEST(BuiltinsCellContents, NumberCell) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  const Value v = EvalSourceIn("=CELL(\"contents\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 42.0);
}

TEST(BuiltinsCellContents, BlankCellReturnsZero) {
  // Excel's blank-as-zero rule: CELL("contents", blank_cell) -> 0.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"contents\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsCellContents, TextCell) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("hello"));
  const Value v = EvalSourceIn("=CELL(\"contents\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "hello");
}

// ---------------------------------------------------------------------------
// type - reference-dependent: "b" / "l" / "v"
// ---------------------------------------------------------------------------

TEST(BuiltinsCellType, NumberCellIsV) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  const Value v = EvalSourceIn("=CELL(\"type\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "v");
}

TEST(BuiltinsCellType, TextCellIsL) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  const Value v = EvalSourceIn("=CELL(\"type\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "l");
}

TEST(BuiltinsCellType, BlankCellIsB) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"type\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "b");
}

TEST(BuiltinsCellType, EmptyStringFormulaIsB) {
  // Mac Excel folds an empty string `""` to "b" (blank), not "l" -- a
  // non-empty text value is required to surface "l".
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text(""));
  const Value v = EvalSourceIn("=CELL(\"type\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "b");
}

// ---------------------------------------------------------------------------
// MVP fixed-stub keys: filename / format / color / parentheses / prefix /
// width. (`protect` is live below, driven by the xf locked flag.)
// ---------------------------------------------------------------------------

TEST(BuiltinsCellFilename, AlwaysEmpty) {
  // Mac surfaces blank for an unsaved workbook, but Formulon returns
  // empty text so the top-level blank-as-zero rule does not collapse
  // the result to 0. The oracle case uses empty_string_readback because
  // xlwings reads Excel's "" result as blank.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"filename\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "");
}

TEST(BuiltinsCellFormat, GeneralReturnsG) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"format\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "G");
}

TEST(BuiltinsCellFormat, BuiltinNumberFormatCodes) {
  Workbook wb = Workbook::create();
  io::StylesTable styles;
  styles.cell_xfs.push_back(io::CellXf{});
  io::CellXf percent{};
  percent.num_fmt_id = 10U;
  styles.cell_xfs.push_back(percent);
  io::CellXf date{};
  date.num_fmt_id = 15U;
  styles.cell_xfs.push_back(date);
  wb.set_styles(std::move(styles));
  wb.sheet(0).set_cell_xf_index(0U, 0U, 1U);
  wb.sheet(0).set_cell_xf_index(1U, 0U, 2U);
  EXPECT_EQ(EvalSourceIn("=CELL(\"format\", A1)", wb, wb.sheet(0)).as_text(), "P2");
  EXPECT_EQ(EvalSourceIn("=CELL(\"format\", A2)", wb, wb.sheet(0)).as_text(), "D1");
}

TEST(BuiltinsCellColor, DefaultFormatReturnsZero) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"color\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsCellColor, NegativeColorDirectiveReturnsOne) {
  Workbook wb = Workbook::create();
  io::StylesTable styles;
  styles.cell_xfs.push_back(io::CellXf{});
  styles.num_fmt_strings.push_back("0;[Red]0");
  styles.num_fmts.push_back(io::NumFmtRecord{164U, 0U});
  io::CellXf colored{};
  colored.num_fmt_id = 164U;
  styles.cell_xfs.push_back(colored);
  wb.set_styles(std::move(styles));
  wb.sheet(0).set_cell_xf_index(0U, 0U, 1U);
  EXPECT_DOUBLE_EQ(EvalSourceIn("=CELL(\"color\", A1)", wb, wb.sheet(0)).as_number(), 1.0);
}

TEST(BuiltinsCellColor, PositiveOnlyColorDirectiveReturnsZero) {
  Workbook wb = Workbook::create();
  io::StylesTable styles;
  styles.cell_xfs.push_back(io::CellXf{});
  styles.num_fmt_strings.push_back("[Red]0;0");
  styles.num_fmts.push_back(io::NumFmtRecord{164U, 0U});
  io::CellXf colored{};
  colored.num_fmt_id = 164U;
  styles.cell_xfs.push_back(colored);
  wb.set_styles(std::move(styles));
  wb.sheet(0).set_cell_xf_index(0U, 0U, 1U);
  EXPECT_DOUBLE_EQ(EvalSourceIn("=CELL(\"color\", A1)", wb, wb.sheet(0)).as_number(), 0.0);
}

TEST(BuiltinsCellParentheses, AlwaysZero) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"parentheses\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsCellPrefix, AlwaysEmpty) {
  // Mac surfaces blank for a cell with no alignment prefix, but Formulon
  // returns empty text so the top-level blank-as-zero rule does not
  // collapse the result to 0. The oracle case uses empty_string_readback
  // because xlwings reads Excel's "" result as blank.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"prefix\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "");
}

TEST(BuiltinsCellPrefix, QuotePrefixReturnsApostrophe) {
  Workbook wb = Workbook::create();
  io::StylesTable styles;
  styles.cell_xfs.push_back(io::CellXf{});
  io::CellXf quote_prefixed{};
  quote_prefixed.quote_prefix = true;
  styles.cell_xfs.push_back(quote_prefixed);
  wb.set_styles(std::move(styles));
  wb.sheet(0).set_cell_xf_index(0U, 0U, 1U);
  const Value v = EvalSourceIn("=CELL(\"prefix\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "'");
}

TEST(BuiltinsCellPrefix, AlignmentPrefixesAreReported) {
  Workbook wb = Workbook::create();
  io::StylesTable styles;
  styles.cell_xfs.push_back(io::CellXf{});
  for (uint8_t alignment : {uint8_t{1}, uint8_t{2}, uint8_t{3}}) {
    io::CellXf xf{};
    xf.horizontal_align = alignment;
    styles.cell_xfs.push_back(xf);
  }
  wb.set_styles(std::move(styles));
  wb.sheet(0).set_cell_xf_index(0U, 0U, 1U);
  wb.sheet(0).set_cell_xf_index(1U, 0U, 2U);
  wb.sheet(0).set_cell_xf_index(2U, 0U, 3U);
  EXPECT_EQ(EvalSourceIn("=CELL(\"prefix\", A1)", wb, wb.sheet(0)).as_text(), "\\");
  EXPECT_EQ(EvalSourceIn("=CELL(\"prefix\", A2)", wb, wb.sheet(0)).as_text(), "^");
  EXPECT_EQ(EvalSourceIn("=CELL(\"prefix\", A3)", wb, wb.sheet(0)).as_text(), "\"");
}

TEST(BuiltinsCellProtect, DefaultCellIsLocked) {
  // A default cell (xf 0, no <protection>) is locked -> 1.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"protect\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsCellProtect, UnlockedCellReturnsZero) {
  // A cell whose xf carries `<protection locked="0"/>` -> 0.
  Workbook wb = Workbook::create();
  io::StylesTable styles;
  styles.cell_xfs.push_back(io::CellXf{});  // xf 0: default (locked).
  io::CellXf unlocked{};
  unlocked.has_protection = true;
  unlocked.locked = false;
  styles.cell_xfs.push_back(unlocked);  // xf 1: unlocked.
  wb.set_styles(std::move(styles));
  wb.sheet(0).set_cell_value(0, 0, Value::number(5.0));
  wb.sheet(0).set_cell_xf_index(0, 0, 1U);  // A1 -> unlocked xf.
  const Value v = EvalSourceIn("=CELL(\"protect\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsCellProtect, ProtectionElementAbsentIsLocked) {
  // An xf whose `<protection>` element is absent defaults to locked -> 1,
  // even when the cell references that xf explicitly.
  Workbook wb = Workbook::create();
  io::StylesTable styles;
  styles.cell_xfs.push_back(io::CellXf{});  // xf 0.
  io::CellXf no_protection{};               // has_protection=false, locked=true.
  styles.cell_xfs.push_back(no_protection);
  wb.set_styles(std::move(styles));
  wb.sheet(0).set_cell_value(0, 0, Value::number(5.0));
  wb.sheet(0).set_cell_xf_index(0, 0, 1U);
  const Value v = EvalSourceIn("=CELL(\"protect\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsCellProtect, AbsentCellIsLocked) {
  // A never-populated cell uses xf 0 -> locked -> 1.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"protect\", Z9)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsCellWidth, OneByTwoArrayDefaultEightTrue) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"width\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  const ArrayValue* arr = v.as_array();
  ASSERT_NE(arr, nullptr);
  EXPECT_EQ(arr->rows, 1U);
  EXPECT_EQ(arr->cols, 2U);
  ASSERT_TRUE(arr->cells[0].is_number());
  EXPECT_DOUBLE_EQ(arr->cells[0].as_number(), 8.0);
  ASSERT_TRUE(arr->cells[1].is_boolean());
  EXPECT_TRUE(arr->cells[1].as_boolean());
}

TEST(BuiltinsCellWidth, ExplicitColumnLayoutIsReported) {
  Workbook wb = Workbook::create();
  wb.sheet(0).mutable_layout().columns.push_back(ColumnLayout{0U, 0U, 12.5, false, 0U});
  const Value v = EvalSourceIn("=CELL(\"width\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_DOUBLE_EQ(v.as_array()->cells[0].as_number(), 12.5);
  EXPECT_FALSE(v.as_array()->cells[1].as_boolean());
}

// ---------------------------------------------------------------------------
// Case insensitivity of info_type
// ---------------------------------------------------------------------------

TEST(BuiltinsCellCase, UppercaseRowKey) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"ROW\", A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsCellCase, MixedCaseAddressKey) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"Address\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "$A$1");
}

// ---------------------------------------------------------------------------
// Error / arity rules
// ---------------------------------------------------------------------------

TEST(BuiltinsCellError, UnknownKeyIsValue) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"foo\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsCellError, ErrorInfoTypePropagates) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(#N/A, A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsCellError, ErrorReferencePropagates) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"row\", #REF!)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsCellError, ZeroArityIsValue) {
  const Value v = EvalSourceNoCtx("=CELL()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsCellError, ThreeArityIsValue) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"row\", A1, B1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsCellError, OmittedRefWithoutAnchorIsRef) {
  // Without a formula cell anchor (CLI ad-hoc eval) the 1-arg form has
  // no meaningful row/col to report; surface #REF!.
  const Value v = EvalSourceNoCtx("=CELL(\"row\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
