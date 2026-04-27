// Copyright 2026 libraz. Licensed under the MIT License.
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
//     prefix, protect, width) - these will become live once the style
//     subsystem lands.
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

TEST(BuiltinsCellType, EmptyStringFormulaIsL) {
  // Empty text "" is NOT blank in Excel: CELL("type") -> "l".
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text(""));
  const Value v = EvalSourceIn("=CELL(\"type\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "l");
}

// ---------------------------------------------------------------------------
// MVP fixed-stub keys: filename / format / color / parentheses / prefix /
// protect / width
// ---------------------------------------------------------------------------

TEST(BuiltinsCellFilename, AlwaysEmpty) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"filename\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "");
}

TEST(BuiltinsCellFormat, AlwaysG) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"format\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "G");
}

TEST(BuiltinsCellColor, AlwaysZero) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"color\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsCellParentheses, AlwaysZero) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"parentheses\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsCellPrefix, AlwaysEmpty) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"prefix\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "");
}

TEST(BuiltinsCellProtect, AlwaysOne) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=CELL(\"protect\", A1)", wb, wb.sheet(0));
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
