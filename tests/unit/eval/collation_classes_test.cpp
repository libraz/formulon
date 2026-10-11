//
// Tests for `dynamic_array::collation_classes`: class ids agree with
// `collate` for every profile and strength, keep type boundaries, and order
// like `=SORT`.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/collation.h"
#include "eval/dynamic_array/common.h"
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
#include "test_eval_helpers.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

constexpr CollationStrength kStrengths[] = {CollationStrength::kCompare, CollationStrength::kDynamic,
                                            CollationStrength::kClassicExact};

const char* StrengthName(CollationStrength s) {
  switch (s) {
    case CollationStrength::kCompare:
      return "kCompare";
    case CollationStrength::kDynamic:
      return "kDynamic";
    default:
      return "kClassicExact";
  }
}

int Sign(long long v) {
  return v < 0 ? -1 : (v > 0 ? 1 : 0);
}

std::vector<std::uint32_t> Classify(const std::vector<Value>& values, ExcelProfile profile, CollationStrength strength,
                                    std::uint32_t* class_count = nullptr) {
  std::vector<std::uint32_t> ids(values.size());
  const std::uint32_t n = dynamic_array::collation_classes(values.data(), values.size(), profile, strength, ids.data());
  if (class_count != nullptr) {
    *class_count = n;
  }
  return ids;
}

TEST(CollationClasses, AgreesWithCollateForEveryProfile) {
  const std::vector<std::string> words = {
      "",     "a",    "A",    "b",  "é",    "e",    "ß",    "ss",     "あ",  "ア",  "ｱ",    "か",
      "カ",   "ｶ",    "が",   "ガ", "ｶﾞ",   "かー", "かあ", "ＡＢＣ", "ABC", "abc", "１２", "12",
      "漢字", "東京", "한글", "가", "中文", "เก",   "กา",   "x-y",    "xy",  "Ａ",
  };
  std::vector<Value> values;
  for (const std::string& w : words) {
    values.push_back(Value::text(w));
  }
  for (const ::formulon::detail::ExcelProfileIdEntry& entry : ::formulon::detail::kExcelProfileIds) {
    for (const CollationStrength strength : kStrengths) {
      const std::vector<std::uint32_t> ids = Classify(values, entry.profile, strength);
      for (std::size_t i = 0; i < words.size(); ++i) {
        for (std::size_t j = 0; j < words.size(); ++j) {
          const int order = collate(words[i], words[j], entry.profile, strength);
          const int id_order = Sign(static_cast<long long>(ids[i]) - static_cast<long long>(ids[j]));
          EXPECT_EQ(ids[i] == ids[j], order == 0)
              << entry.id << " " << StrengthName(strength) << " a=" << words[i] << " b=" << words[j];
          EXPECT_EQ(id_order, Sign(order))
              << entry.id << " " << StrengthName(strength) << " a=" << words[i] << " b=" << words[j];
        }
      }
    }
  }
}

TEST(CollationClasses, TypeBoundaries) {
  const std::vector<Value> values = {
      Value::number(1),
      Value::text("1"),
      Value::boolean(true),
      Value::text("TRUE"),
      Value::blank(),
      Value::text(""),
      Value::number(0),
      Value::number(-0.0),
      Value::boolean(false),
      Value::error(ErrorCode::NA),
      Value::error(ErrorCode::Value),
  };
  for (const CollationStrength strength : kStrengths) {
    SCOPED_TRACE(StrengthName(strength));
    std::uint32_t count = 0;
    const std::vector<std::uint32_t> ids = Classify(values, mac_365_ja_jp_profile(), strength, &count);
    EXPECT_EQ(count, 10U);
    EXPECT_EQ(ids[6], ids[7]);
    std::vector<std::uint32_t> distinct = ids;
    std::sort(distinct.begin(), distinct.end());
    distinct.erase(std::unique(distinct.begin(), distinct.end()), distinct.end());
    EXPECT_EQ(distinct.size(), 10U);
    EXPECT_EQ(distinct.back(), 9U);
  }
}

// Blanks are left out: a bare-range SORT spill returns them as 0.
std::vector<Value> SortInputs() {
  return {
      Value::text("b"),    Value::number(3),
      Value::text("A"),    Value::error(ErrorCode::NA),
      Value::text("ア"),   Value::boolean(true),
      Value::text("あ"),   Value::number(-1),
      Value::text("ＡＢ"), Value::error(ErrorCode::Div0),
      Value::text("a"),    Value::boolean(false),
      Value::text("ｶﾞ"),   Value::number(3),
      Value::text("ガ"),   Value::error(ErrorCode::Value),
      Value::text("AB"),   Value::text("é"),
      Value::text("e"),    Value::number(0),
  };
}

TEST(CollationClasses, OrderMatchesSort) {
  const ExcelProfile profiles[] = {mac_365_ja_jp_profile(), ExcelProfile{ExcelHost::kMac365, ExcelLocale::kZhCN},
                                   win_365_en_us_profile()};
  const std::vector<Value> inputs = SortInputs();
  for (const ExcelProfile& profile : profiles) {
    SCOPED_TRACE(excel_profile_id(profile));
    const std::vector<std::uint32_t> ids = Classify(inputs, profile, CollationStrength::kDynamic);
    std::vector<std::size_t> expected(inputs.size());
    for (std::size_t i = 0; i < expected.size(); ++i) {
      expected[i] = i;
    }
    std::stable_sort(expected.begin(), expected.end(), [&](std::size_t a, std::size_t b) { return ids[a] < ids[b]; });

    Workbook wb = Workbook::create();
    Sheet& sheet = wb.sheet(0);
    for (std::size_t i = 0; i < inputs.size(); ++i) {
      sheet.set_cell_value(static_cast<std::uint32_t>(i), 0, inputs[i]);
    }
    ScopedEvalProfile scope(profile);
    EvalState state;
    const EvalContext ctx = test::mac_context(wb, sheet, state).with_excel_profile(profile);
    Arena parse_arena;
    Arena eval_arena;
    parser::Parser p("=SORT(A1:A20)", parse_arena);
    parser::AstNode* root = p.parse();
    ASSERT_NE(root, nullptr);
    const Value v = evaluate(*root, eval_arena, default_registry(), ctx);
    ASSERT_TRUE(v.is_array());
    ASSERT_EQ(v.as_array_rows(), inputs.size());
    for (std::size_t i = 0; i < expected.size(); ++i) {
      const Value& want = inputs[expected[i]];
      const Value& got = v.as_array_cells()[i];
      ASSERT_EQ(got.kind(), want.kind()) << "row " << i;
      switch (want.kind()) {
        case ValueKind::Number:
          EXPECT_DOUBLE_EQ(got.as_number(), want.as_number()) << "row " << i;
          break;
        case ValueKind::Text:
          EXPECT_EQ(got.as_text(), want.as_text()) << "row " << i;
          break;
        case ValueKind::Bool:
          EXPECT_EQ(got.as_boolean(), want.as_boolean()) << "row " << i;
          break;
        default:
          EXPECT_EQ(got.as_error(), want.as_error()) << "row " << i;
          break;
      }
    }
  }
}

TEST(CollationClasses, EmptyInput) {
  EXPECT_EQ(dynamic_array::collation_classes(nullptr, 0, mac_365_ja_jp_profile(), CollationStrength::kDynamic, nullptr),
            0U);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
