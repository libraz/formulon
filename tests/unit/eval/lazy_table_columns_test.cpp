//
// Tests for the per-function columns of the lazy dispatch table.

#include <cstddef>
#include <cstdint>
#include <set>
#include <string>
#include <string_view>
#include <vector>

#include "eval/function_registry.h"
#include "eval/range_resolvers.h"
#include "eval/spill_potential.h"
#include "eval/tree_walker.h"
#include "eval/tree_walker_lazy_table.h"
#include "gtest/gtest.h"
#include "io/xlsb/func_id_table.h"
#include "parser/parser.h"
#include "util/test_eval_helpers.h"
#include "utils/arena.h"
#include "value.h"

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
      {"IRR", ParamSet::kFirst, ScalarArg::kRejectNum},
      {"MIRR", ParamSet::kFirst, ScalarArg::kRejectDiv0},
      {"XIRR", ParamSet::kFirstAndSecond, ScalarArg::kRejectTextValue},
      {"XNPV", ParamSet::kAllButFirst, ScalarArg::kRejectTextValue},
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

// Every reducing or one-value row that lifts over an array argument, with valid scalar arguments and, per slot, a
// two-element array valid there. An array in a slot lifts the call, and makes it a possible spill, exactly when the
// row's lift_params or value_lifts holds the slot.
TEST(LazyTableColumns, LiftedSlotsAreTheDeclaredOnes) {
  struct Row {
    const char* name;
    std::vector<const char*> scalars;
    std::vector<const char*> arrays;
  };
  const char* const kDatabase = "{\"a\";1;2}";
  const char* const kCriteria = "{\"a\";1}";
  const std::vector<const char*> kDatabaseScalars = {kDatabase, "\"a\"", kCriteria};
  const std::vector<const char*> kDatabaseArrays = {kDatabase, "{1,1}", kCriteria};
  const std::vector<const char*> kRankScalars = {"{1,2,3}", "2", "3"};
  const std::vector<const char*> kRankArrays = {"{1,2,3}", "{1,2}", "{1,2}"};
  const std::vector<const char*> kIfsScalars = {"{1,2}", "{1,2}", "1"};
  const std::vector<const char*> kIfsArrays = {"{1,2}", "{1,2}", "{1,2}"};
  const std::vector<const char*> kIfScalars = {"{1,2}", "1", "{1,2}"};
  const std::vector<const char*> kIfArrays = {"{1,2}", "{1,2}", "{1,2}"};
  const Row rows[] = {
      {"AVERAGEIF", kIfScalars, kIfArrays},
      {"AVERAGEIFS", kIfsScalars, kIfsArrays},
      {"COUNTIF", {"{1,2}", "1"}, {"{1,2}", "{1,2}"}},
      {"COUNTIFS", {"{1,2}", "1", "{1,2}", "1"}, {"{1,2}", "{1,2}", "{1,2}", "{1,2}"}},
      {"DAVERAGE", kDatabaseScalars, kDatabaseArrays},
      {"DCOUNT", kDatabaseScalars, kDatabaseArrays},
      {"DCOUNTA", kDatabaseScalars, kDatabaseArrays},
      {"DGET", kDatabaseScalars, kDatabaseArrays},
      {"DMAX", kDatabaseScalars, kDatabaseArrays},
      {"DMIN", kDatabaseScalars, kDatabaseArrays},
      {"DPRODUCT", kDatabaseScalars, kDatabaseArrays},
      {"DSTDEV", kDatabaseScalars, kDatabaseArrays},
      {"DSTDEVP", kDatabaseScalars, kDatabaseArrays},
      {"DSUM", kDatabaseScalars, kDatabaseArrays},
      {"DVAR", kDatabaseScalars, kDatabaseArrays},
      {"DVARP", kDatabaseScalars, kDatabaseArrays},
      {"IRR", {"{-10,6,6}", "0.1"}, {"{-10,6,6}", "{0.1,0.2}"}},
      {"MAXIFS", kIfsScalars, kIfsArrays},
      {"MINIFS", kIfsScalars, kIfsArrays},
      {"MIRR", {"{-10,6,6}", "0.1", "0.1"}, {"{-10,6,6}", "{0.1,0.2}", "{0.1,0.2}"}},
      {"PERCENTRANK", kRankScalars, kRankArrays},
      {"PERCENTRANK.EXC", kRankScalars, kRankArrays},
      {"PERCENTRANK.INC", kRankScalars, kRankArrays},
      {"SERIESSUM", {"1", "1", "1", "{1,2}"}, {"{1,2}", "{1,2}", "{1,2}", "{1,2}"}},
      {"SUMIF", kIfScalars, kIfArrays},
      {"SUMIFS", kIfsScalars, kIfsArrays},
      {"XIRR", {"{-10,6}", "{1,400}", "0.1"}, {"{-10,6}", "{1,400}", "{0.1,0.2}"}},
      {"XNPV", {"0.1", "{-10,6}", "{1,400}"}, {"{0.1,0.2}", "{-10,6}", "{1,400}"}},
  };
  // Rows whose lifted slots depend on the call's function number are checked by their own tests.
  const std::set<std::string> by_form = {"AGGREGATE"};
  std::set<std::string> expected;
  for (const char* const* p = lazy_form_names(); *p != nullptr; ++p) {
    const LazyEntry* entry = find_lazy_entry(*p);
    const bool one_value = entry->shape == LazyResultShape::kScalar || entry->shape == LazyResultShape::kReduce;
    if (one_value && (entry->lift_params != ParamSet::kNone || entry->value_lifts != ParamSet::kNone) &&
        by_form.count(*p) == 0U) {
      expected.insert(*p);
    }
  }
  std::set<std::string> actual;
  for (const Row& row : rows) {
    actual.insert(row.name);
    const LazyEntry* entry = find_lazy_entry(row.name);
    ASSERT_NE(entry, nullptr) << row.name;
    ASSERT_EQ(row.scalars.size(), row.arrays.size()) << row.name;
    for (std::size_t slot = 0; slot < row.arrays.size(); ++slot) {
      std::string formula = std::string("=") + row.name + "(";
      for (std::size_t i = 0; i < row.arrays.size(); ++i) {
        formula += (i == 0U ? "" : ",");
        formula += i == slot ? row.arrays[i] : row.scalars[i];
      }
      formula += ")";
      const auto i = static_cast<std::uint32_t>(slot);
      const bool lifts = param_in_set(entry->lift_params, i) || param_in_set(entry->value_lifts, i);
      const Value v = formulon::test::EvalSource(formula);
      EXPECT_EQ(v.is_array() && v.as_array_rows() * v.as_array_cols() == 2U, lifts) << formula;
      Arena arena;
      const parser::AstNode* root = parser::parse_strict(std::string_view(formula).substr(1), arena);
      ASSERT_NE(root, nullptr) << formula;
      EXPECT_EQ(spill_potential(*root, default_registry()) == SpillPotential::kMaySpill, lifts) << formula;
    }
  }
  EXPECT_EQ(actual, expected);
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
