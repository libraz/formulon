//
// The supplementary-plane text cases of the win-365-ja_JP formula goldens,
// captured on Windows Excel in a book it created (stamped Compatibility
// Version 2), evaluated in a new win-365-ja_JP workbook. That target is not an
// active ctest oracle, so this test holds the Windows half of the text
// position rule: LEN, MID, FIND, SEARCH and REPLACE count code points in a
// version-2 book on Windows as on Mac, and LEFT and RIGHT on every version.

#include <cstdint>
#include <string>
#include <string_view>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "tests/oracle/json_reader.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_ORACLE_GOLDEN_DIR
#error "FORMULON_ORACLE_GOLDEN_DIR must be defined by the build"
#endif

namespace formulon {
namespace tests {
namespace oracle {
namespace {

// The win-365-ja_JP golden beside the primary target's golden directory.
std::string WinTextGoldenPath() {
  std::string dir = FORMULON_ORACLE_GOLDEN_DIR;
  const std::string primary = "mac-365-ja_JP";
  const std::size_t at = dir.rfind(primary);
  if (at != std::string::npos) {
    dir.replace(at, primary.size(), "win-365-ja_JP");
  }
  return dir + "/text.golden.json";
}

TEST(WinTextUnitGolden, SupplementaryPlaneCasesMatchWindowsExcel) {
  auto parsed = parse_json_file(WinTextGoldenPath());
  ASSERT_TRUE(static_cast<bool>(parsed)) << WinTextGoldenPath();
  const JsonValue* cases = parsed.value().find("cases");
  ASSERT_NE(cases, nullptr);
  ASSERT_TRUE(cases->is_array());
  int checked = 0;
  for (const JsonValue& c : cases->as_array()) {
    const JsonValue* id = c.find("id");
    if (id == nullptr || !id->is_string() || id->as_string().rfind("supp_", 0) != 0) {
      continue;
    }
    const std::string& name = id->as_string();
    const JsonValue* formula = c.find("formula");
    const JsonValue* expect = c.find("expect");
    ASSERT_NE(formula, nullptr) << name;
    ASSERT_NE(expect, nullptr) << name;
    const JsonValue* setup = c.find("setup");
    ASSERT_TRUE(setup == nullptr || (setup->is_object() && setup->as_object().empty())) << name << " has a setup";

    Workbook wb = Workbook::create_for_profile(win_365_ja_jp_profile());
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, formula->as_string()))) << name;
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry()))) << name;
    const Cell* cell = wb.sheet(0).cell_at(0U, 0U);
    ASSERT_NE(cell, nullptr) << name;
    const Value& got = cell->cached_value;
    const std::string kind = expect->find("kind")->as_string();
    const JsonValue* want = expect->find("value");
    ASSERT_NE(want, nullptr) << name;
    if (kind == "number") {
      ASSERT_TRUE(got.is_number()) << name;
      EXPECT_EQ(got.as_number(), want->as_number()) << name << " " << formula->as_string();
    } else if (kind == "text") {
      ASSERT_TRUE(got.is_text()) << name;
      EXPECT_EQ(std::string(got.as_text()), want->as_string()) << name << " " << formula->as_string();
    } else {
      ADD_FAILURE() << name << ": unexpected expect kind " << kind;
    }
    ++checked;
  }
  EXPECT_EQ(checked, 30);
}

}  // namespace
}  // namespace oracle
}  // namespace tests
}  // namespace formulon
