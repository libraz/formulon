//
// Tests for the per-function columns of the lazy dispatch table.

#include <cstddef>
#include <cstdint>
#include <set>
#include <string>
#include <string_view>

#include "eval/function_registry.h"
#include "eval/range_resolvers.h"
#include "eval/tree_walker.h"
#include "eval/tree_walker_lazy_table.h"
#include "gtest/gtest.h"
#include "io/xlsb/func_id_table.h"

namespace formulon {
namespace eval {
namespace {

constexpr std::uint32_t kMaxProbedIndex = 8U;

TEST(LazyTableColumns, StructureSizesAreFixed) {
  EXPECT_EQ(sizeof(LazyEntry), sizeof(void*) == 8 ? 32U : 20U);
  EXPECT_EQ(sizeof(FunctionDef), sizeof(void*) == 8 ? 72U : 48U);
}

TEST(LazyTableColumns, ReferenceResultSetIsTheReferenceReturningCalls) {
  const std::set<std::string> expected = {"INDEX",        "INDIRECT",      "OFFSET",    "IF",     "CHOOSE", "XLOOKUP",
                                          "IFS",          "SWITCH",        "TRIMRANGE", "SINGLE", "TAKE",   "DROP",
                                          "_TRO_LEADING", "_TRO_TRAILING", "_TRO_ALL"};
  std::set<std::string> actual;
  for (const char* const* p = lazy_form_names(); *p != nullptr; ++p) {
    if (find_lazy_entry(*p)->ref_result != RefResult::kNone) {
      actual.insert(*p);
    }
  }
  EXPECT_EQ(actual, expected);
}

TEST(LazyTableColumns, ReferenceParamsAreTheReferenceOnlyParameters) {
  struct Row {
    const char* name;
    ParamSet params;
  };
  const Row expected[] = {
      {"AREAS", ParamSet::kFirst},
      {"AVERAGEIF", ParamSet::kFirstAndThird},
      {"AVERAGEIFS", ParamSet::kFirstAndOdd},
      {"CELL", ParamSet::kSecond},
      {"COLUMN", ParamSet::kFirst},
      {"COUNTBLANK", ParamSet::kFirst},
      {"COUNTIF", ParamSet::kFirst},
      {"COUNTIFS", ParamSet::kEven},
      {"DAVERAGE", ParamSet::kFirst},
      {"DCOUNT", ParamSet::kFirst},
      {"DCOUNTA", ParamSet::kFirst},
      {"DGET", ParamSet::kFirst},
      {"DMAX", ParamSet::kFirst},
      {"DMIN", ParamSet::kFirst},
      {"DPRODUCT", ParamSet::kFirst},
      {"DSTDEV", ParamSet::kFirst},
      {"DSTDEVP", ParamSet::kFirst},
      {"DSUM", ParamSet::kFirst},
      {"DVAR", ParamSet::kFirst},
      {"DVARP", ParamSet::kFirst},
      {"FORMULATEXT", ParamSet::kFirst},
      {"ISFORMULA", ParamSet::kFirst},
      {"MAXIFS", ParamSet::kFirstAndOdd},
      {"MINIFS", ParamSet::kFirstAndOdd},
      {"OFFSET", ParamSet::kFirst},
      {"PHONETIC", ParamSet::kFirst},
      {"RANK", ParamSet::kSecond},
      {"RANK.AVG", ParamSet::kSecond},
      {"RANK.EQ", ParamSet::kSecond},
      {"ROW", ParamSet::kFirst},
      {"SUBTOTAL", ParamSet::kAllButFirst},
      {"SUMIF", ParamSet::kFirstAndThird},
      {"SUMIFS", ParamSet::kFirstAndOdd},
  };
  std::size_t with_params = 0;
  for (const char* const* p = lazy_form_names(); *p != nullptr; ++p) {
    if (find_lazy_entry(*p)->ref_params != ParamSet::kNone) {
      ++with_params;
    }
  }
  EXPECT_EQ(with_params, sizeof(expected) / sizeof(expected[0]));
  for (const Row& row : expected) {
    const LazyEntry* entry = find_lazy_entry(row.name);
    ASSERT_NE(entry, nullptr) << row.name << " has no lazy entry";
    EXPECT_EQ(entry->ref_params, row.params) << row.name;
  }
}

TEST(LazyTableColumns, ReferenceOnlyParametersAreReferenceClassInTheStorageTable) {
  for (const char* const* p = lazy_form_names(); *p != nullptr; ++p) {
    const LazyEntry* entry = find_lazy_entry(*p);
    for (std::uint32_t i = 0; i < kMaxProbedIndex; ++i) {
      if (param_in_set(entry->ref_params, i)) {
        EXPECT_EQ(io::xlsb::xlsb_parameter_class(*p, i), 'R') << *p << " parameter " << i;
      }
    }
  }
}

TEST(LazyTableColumns, ShortCircuitFormsDeclareNoClass) {
  for (const char* name : {"IF", "IFS", "SWITCH", "CHOOSE", "IFERROR", "IFNA"}) {
    const LazyEntry* entry = find_lazy_entry(name);
    ASSERT_NE(entry, nullptr) << name;
    EXPECT_EQ(entry->ref_params, ParamSet::kNone) << name;
    EXPECT_EQ(entry->grid_params, ParamSet::kNone) << name;
    EXPECT_EQ(entry->grid_sel, GridArg::kElementError) << name;
    EXPECT_EQ(entry->grid_rest, GridArg::kElementError) << name;
    EXPECT_EQ(entry->scalar_params, ParamSet::kNone) << name;
    EXPECT_EQ(entry->scalar_error, ScalarArg::kAccept) << name;
    EXPECT_EQ(entry->lift_params, ParamSet::kNone) << name;
  }
}

// The functions that read a reference or an array and refuse a lone value there, with the positions and the answer.
TEST(LazyTableColumns, LoneValuePolicyRows) {
  struct Row {
    const char* name;
    ParamSet params;
    ScalarArg answer;
  };
  const Row rows[] = {
      {"MATCH", ParamSet::kSecond, ScalarArg::kRejectNAOrTextValue},
      {"IRR", ParamSet::kFirst, ScalarArg::kRejectValue},
      {"MIRR", ParamSet::kFirst, ScalarArg::kRejectValue},
      {"XIRR", ParamSet::kFirstAndSecond, ScalarArg::kRejectValue},
      {"XNPV", ParamSet::kAllButFirst, ScalarArg::kRejectValue},
      {"RANK", ParamSet::kSecond, ScalarArg::kRejectValue},
      {"RANK.EQ", ParamSet::kSecond, ScalarArg::kRejectValue},
      {"RANK.AVG", ParamSet::kSecond, ScalarArg::kRejectValue},
      {"T.TEST", ParamSet::kFirstAndSecond, ScalarArg::kRejectNA},
      {"TTEST", ParamSet::kFirstAndSecond, ScalarArg::kRejectNA},
      {"F.TEST", ParamSet::kFirstAndSecond, ScalarArg::kRejectNA},
      {"FTEST", ParamSet::kFirstAndSecond, ScalarArg::kRejectNA},
      {"Z.TEST", ParamSet::kFirst, ScalarArg::kRejectNA},
      {"ZTEST", ParamSet::kFirst, ScalarArg::kRejectNA},
  };
  std::set<std::string> expected;
  for (const Row& row : rows) {
    const LazyEntry* entry = find_lazy_entry(row.name);
    ASSERT_NE(entry, nullptr) << row.name;
    EXPECT_EQ(entry->scalar_params, row.params) << row.name;
    EXPECT_EQ(entry->scalar_error, row.answer) << row.name;
    expected.insert(row.name);
  }
  std::set<std::string> actual;
  for (const char* const* p = lazy_form_names(); *p != nullptr; ++p) {
    if (find_lazy_entry(*p)->scalar_params != ParamSet::kNone) {
      actual.insert(*p);
    }
  }
  EXPECT_EQ(actual, expected);
}

TEST(LazyTableColumns, ReferenceCallsAreTheResultsCarryingARectangle) {
  // SINGLE projects onto the formula cell, so it has no static rectangle.
  for (const char* const* p = lazy_form_names(); *p != nullptr; ++p) {
    const RefResult result = find_lazy_entry(*p)->ref_result;
    EXPECT_EQ(is_reference_call_name(*p), result != RefResult::kNone && result != RefResult::kProjected) << *p;
  }
}

TEST(LazyTableColumns, VariadicArityIsReportedAsVariadic) {
  std::uint32_t min_arity = 0U;
  std::uint32_t max_arity = 0U;
  ASSERT_TRUE(find_lazy_arity("SUBTOTAL", &min_arity, &max_arity));
  EXPECT_EQ(min_arity, 2U);
  EXPECT_EQ(max_arity, kVariadic);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
