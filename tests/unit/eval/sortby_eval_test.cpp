//
// Tests for the lazy `SORTBY(array, by_array1, [order1], [by_array2, ...])`
// builtin. Shares its TU and helpers (`sort_kind_rank`, `sort_lane_less`)
// with SORT. SORTBY differs from SORT in that the keys are out-of-band
// 1D vectors (not columns inside `array`), and that multiple keys with
// per-key orders are supported. Axis is inferred from the shape of the
// first key vector.

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <string_view>

#include "eval/collation.h"
#include "eval/dynamic_array/common.h"
#include "eval/eval_context.h"
#include "eval/eval_profile_scope.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "test_eval_helpers.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

Value EvalUnder(std::string_view src, Arena* parse_arena, Arena* eval_arena, const EvalContext& ctx) {
  parser::Parser p(src, *parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty());
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, *eval_arena, default_registry(), ctx);
}

// ---------------------------------------------------------------------------
// Single key
// ---------------------------------------------------------------------------

TEST(BuiltinsSortby, SingleKeyAscending) {
  // A1:A4 = {"a","b","c","d"}; key B1:B4 = {3,1,4,2}. SORTBY ascending
  // by the key permutes rows -> {"b","d","a","c"}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::text("a"));
  sheet.set_cell_value(1, 0, Value::text("b"));
  sheet.set_cell_value(2, 0, Value::text("c"));
  sheet.set_cell_value(3, 0, Value::text("d"));
  sheet.set_cell_value(0, 1, Value::number(3));
  sheet.set_cell_value(1, 1, Value::number(1));
  sheet.set_cell_value(2, 1, Value::number(4));
  sheet.set_cell_value(3, 1, Value::number(2));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORTBY(A1:A4, B1:B4)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 4U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "b");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "d");
  EXPECT_EQ(v.as_array_cells()[2].as_text(), "a");
  EXPECT_EQ(v.as_array_cells()[3].as_text(), "c");
}

TEST(BuiltinsSortby, SingleKeyDescending) {
  // Same data as above; order = -1 reverses.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::text("a"));
  sheet.set_cell_value(1, 0, Value::text("b"));
  sheet.set_cell_value(2, 0, Value::text("c"));
  sheet.set_cell_value(3, 0, Value::text("d"));
  sheet.set_cell_value(0, 1, Value::number(3));
  sheet.set_cell_value(1, 1, Value::number(1));
  sheet.set_cell_value(2, 1, Value::number(4));
  sheet.set_cell_value(3, 1, Value::number(2));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORTBY(A1:A4, B1:B4, -1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 4U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "c");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "a");
  EXPECT_EQ(v.as_array_cells()[2].as_text(), "d");
  EXPECT_EQ(v.as_array_cells()[3].as_text(), "b");
}

// ---------------------------------------------------------------------------
// Multi-key tiebreaker
// ---------------------------------------------------------------------------

TEST(BuiltinsSortby, TwoKeysTiebreaker) {
  // Rows A1:A4 = {"x","y","z","w"}; primary key B1:B4 = {1,1,2,2};
  // secondary key C1:C4 = {20,10,30,40}. With order 1, 1 -> primary ties
  // resolve via C: rows with B=1 sort by C ascending (10, 20) -> ("y","x"),
  // then rows with B=2 by C ascending (30, 40) -> ("z","w").
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::text("x"));
  sheet.set_cell_value(1, 0, Value::text("y"));
  sheet.set_cell_value(2, 0, Value::text("z"));
  sheet.set_cell_value(3, 0, Value::text("w"));
  sheet.set_cell_value(0, 1, Value::number(1));
  sheet.set_cell_value(1, 1, Value::number(1));
  sheet.set_cell_value(2, 1, Value::number(2));
  sheet.set_cell_value(3, 1, Value::number(2));
  sheet.set_cell_value(0, 2, Value::number(20));
  sheet.set_cell_value(1, 2, Value::number(10));
  sheet.set_cell_value(2, 2, Value::number(30));
  sheet.set_cell_value(3, 2, Value::number(40));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORTBY(A1:A4, B1:B4, 1, C1:C4, 1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 4U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "y");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "x");
  EXPECT_EQ(v.as_array_cells()[2].as_text(), "z");
  EXPECT_EQ(v.as_array_cells()[3].as_text(), "w");
}

TEST(BuiltinsSortby, TwoKeysWithMixedOrders) {
  // Same as above but secondary key descending. Rows with B=1 then sort
  // by C descending -> ("x","y"); rows with B=2 -> ("w","z").
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::text("x"));
  sheet.set_cell_value(1, 0, Value::text("y"));
  sheet.set_cell_value(2, 0, Value::text("z"));
  sheet.set_cell_value(3, 0, Value::text("w"));
  sheet.set_cell_value(0, 1, Value::number(1));
  sheet.set_cell_value(1, 1, Value::number(1));
  sheet.set_cell_value(2, 1, Value::number(2));
  sheet.set_cell_value(3, 1, Value::number(2));
  sheet.set_cell_value(0, 2, Value::number(20));
  sheet.set_cell_value(1, 2, Value::number(10));
  sheet.set_cell_value(2, 2, Value::number(30));
  sheet.set_cell_value(3, 2, Value::number(40));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORTBY(A1:A4, B1:B4, 1, C1:C4, -1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 4U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "x");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "y");
  EXPECT_EQ(v.as_array_cells()[2].as_text(), "w");
  EXPECT_EQ(v.as_array_cells()[3].as_text(), "z");
}

TEST(BuiltinsSortby, AllKeysEqualPreservesInputOrder) {
  // Stable sort guarantee: when every key matches across rows, input order
  // is preserved.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::text("first"));
  sheet.set_cell_value(1, 0, Value::text("second"));
  sheet.set_cell_value(2, 0, Value::text("third"));
  sheet.set_cell_value(0, 1, Value::number(7));
  sheet.set_cell_value(1, 1, Value::number(7));
  sheet.set_cell_value(2, 1, Value::number(7));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORTBY(A1:A3, B1:B3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "first");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "second");
  EXPECT_EQ(v.as_array_cells()[2].as_text(), "third");
}

// ---------------------------------------------------------------------------
// Axis inference and shape validation
// ---------------------------------------------------------------------------

TEST(BuiltinsSortby, RowVectorKeySortsColumns) {
  // 2-row, 4-column array; row-vector key (1 x 4) infers col-sort axis.
  // A1:D2 = {{"a","b","c","d"},{1,2,3,4}}. Key A4:D4 = {3,1,4,2}. SORTBY
  // ascending -> columns reordered to (B,D,A,C).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::text("a"));
  sheet.set_cell_value(0, 1, Value::text("b"));
  sheet.set_cell_value(0, 2, Value::text("c"));
  sheet.set_cell_value(0, 3, Value::text("d"));
  sheet.set_cell_value(1, 0, Value::number(1));
  sheet.set_cell_value(1, 1, Value::number(2));
  sheet.set_cell_value(1, 2, Value::number(3));
  sheet.set_cell_value(1, 3, Value::number(4));
  sheet.set_cell_value(3, 0, Value::number(3));
  sheet.set_cell_value(3, 1, Value::number(1));
  sheet.set_cell_value(3, 2, Value::number(4));
  sheet.set_cell_value(3, 3, Value::number(2));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORTBY(A1:D2, A4:D4)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 4U);
  const Value* cells = v.as_array_cells();
  EXPECT_EQ(cells[0].as_text(), "b");
  EXPECT_EQ(cells[1].as_text(), "d");
  EXPECT_EQ(cells[2].as_text(), "a");
  EXPECT_EQ(cells[3].as_text(), "c");
  EXPECT_DOUBLE_EQ(cells[4].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[5].as_number(), 4.0);
  EXPECT_DOUBLE_EQ(cells[6].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[7].as_number(), 3.0);
}

TEST(BuiltinsSortby, KeyLengthMismatchReturnsValue) {
  // 4-row data, 3-element key -> #VALUE!.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  for (std::uint32_t r = 0; r < 4U; ++r) {
    sheet.set_cell_value(r, 0, Value::number(static_cast<double>(r)));
  }
  for (std::uint32_t r = 0; r < 3U; ++r) {
    sheet.set_cell_value(r, 1, Value::number(0));
  }

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORTBY(A1:A4, B1:B3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSortby, SecondKeyAxisMismatchReturnsValue) {
  // First key is column-vector (4 x 1); second key is row-vector (1 x 4).
  // The second key's axis disagrees with the first -> #VALUE!.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  for (std::uint32_t r = 0; r < 4U; ++r) {
    sheet.set_cell_value(r, 0, Value::number(static_cast<double>(r)));
  }
  for (std::uint32_t r = 0; r < 4U; ++r) {
    sheet.set_cell_value(r, 1, Value::number(0));
  }
  // Row-vector key in row 5 (0-based row 5).
  for (std::uint32_t c = 0; c < 4U; ++c) {
    sheet.set_cell_value(5, c, Value::number(static_cast<double>(c)));
  }

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORTBY(A1:A4, B1:B4, 1, A6:D6)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Argument validation
// ---------------------------------------------------------------------------

TEST(BuiltinsSortby, OneArgRejected) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORTBY(A1:A1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSortby, InvalidOrderReturnsValue) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(0, 1, Value::number(1));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORTBY(A1:A1, B1:B1, 2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Locale collation (dynamic-array strength)
// ---------------------------------------------------------------------------

struct CollationPair {
  const char* a;
  const char* b;
  bool sort_first_is_b;        // SORT({a;b}) puts b first
  std::uint32_t sortby_first;  // SORTBY({1;2}, {a;b}) first element
  std::uint32_t unique_rows;   // ROWS(UNIQUE({a;b}))
};

const CollationPair kMacJa[] = {
    {"ß", "ss", false, 1U, 1U},      {"ｶﾞ", "ガ", false, 1U, 1U}, {"ア", "あ", false, 1U, 1U},
    {"Ａ", "A", false, 1U, 1U},      {"①", "1", false, 1U, 1U},   {"ä", "a", true, 2U, 2U},
    {"かー", "かあ", false, 1U, 1U}, {"เก", "กา", true, 2U, 2U},
};

const CollationPair kMacEn[] = {
    {"ß", "ss", false, 1U, 1U},      {"ｶﾞ", "ガ", false, 1U, 1U}, {"ア", "あ", false, 1U, 1U},
    {"Ａ", "A", false, 1U, 1U},      {"①", "1", false, 1U, 1U},   {"ä", "a", true, 2U, 2U},
    {"かー", "かあ", false, 1U, 2U}, {"เก", "กา", true, 2U, 2U},
};

const CollationPair kMacZh[] = {
    {"ß", "ss", false, 1U, 1U},     {"ｶﾞ", "ガ", false, 1U, 1U}, {"ア", "あ", false, 1U, 2U},
    {"Ａ", "A", false, 1U, 1U},     {"①", "1", false, 1U, 1U},   {"ä", "a", true, 2U, 2U},
    {"かー", "かあ", true, 2U, 2U}, {"เก", "กา", false, 1U, 2U},
};

const CollationPair kWinJa[] = {
    {"ß", "ss", true, 2U, 2U}, {"ｶﾞ", "ガ", false, 1U, 1U}, {"ア", "あ", true, 2U, 2U},     {"Ａ", "A", true, 2U, 2U},
    {"①", "1", true, 2U, 2U},  {"ä", "a", true, 2U, 2U},    {"かー", "かあ", true, 2U, 2U}, {"เก", "กา", true, 2U, 2U},
};

struct CollationCase {
  ExcelProfile profile;
  const CollationPair* pairs;
  std::size_t count;
};

const CollationCase kCollationCases[] = {
    {mac_365_ja_jp_profile(), kMacJa, sizeof(kMacJa) / sizeof(kMacJa[0])},
    {mac_365_en_us_profile(), kMacEn, sizeof(kMacEn) / sizeof(kMacEn[0])},
    {ExcelProfile{ExcelHost::kMac365, ExcelLocale::kZhCN}, kMacZh, sizeof(kMacZh) / sizeof(kMacZh[0])},
    {win_365_ja_jp_profile(), kWinJa, sizeof(kWinJa) / sizeof(kWinJa[0])},
};

// Owns what an EvalPair result points into; it lives until the next call.
struct PairEval {
  Workbook wb = Workbook::create();
  EvalState state;
  Arena parse_arena;
  Arena eval_arena;
};

Value EvalPair(const std::string& formula, ExcelProfile profile) {
  static std::unique_ptr<PairEval> last;
  last = std::make_unique<PairEval>();
  const EvalContext ctx = test::mac_context(last->wb, last->wb.sheet(0), last->state).with_excel_profile(profile);
  return EvalUnder(formula, &last->parse_arena, &last->eval_arena, ctx);
}

std::string PairArray(const CollationPair& p) {
  return std::string("{\"") + p.a + "\";\"" + p.b + "\"}";
}

TEST(BuiltinsSortbyCollation, FirstOfPairFollowsMeasuredOrder) {
  for (const CollationCase& c : kCollationCases) {
    ScopedEvalProfile scope(c.profile);
    for (std::size_t i = 0; i < c.count; ++i) {
      const CollationPair& p = c.pairs[i];
      SCOPED_TRACE(std::string(p.a) + " / " + p.b);
      const Value v = EvalPair("=SORTBY({1;2}," + PairArray(p) + ")", c.profile);
      ASSERT_TRUE(v.is_array());
      EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), static_cast<double>(p.sortby_first));
    }
  }
}

TEST(BuiltinsSortbyCollation, TiedFirstKeyFallsThroughToSecond) {
  ScopedEvalProfile scope(mac_365_ja_jp_profile());
  const Value v = EvalPair("=SORTBY({1;2},{\"ア\";\"あ\"},1,{2;1},1)", mac_365_ja_jp_profile());
  ASSERT_TRUE(v.is_array());
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 1.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
