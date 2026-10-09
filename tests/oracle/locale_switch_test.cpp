// Switches one live workbook through every gated formula-track locale and
// checks each recalc against that locale's own golden.
//
// The per-target oracle binaries each evaluate a fresh workbook under a single
// profile, so they cannot see state that survives a profile change: a cached
// value, a criteria cache or a parsed format left over from the previous
// locale. Here every `locale_tokens` case gets one workbook, built once
// through the public `Workbook` API, and is recalculated after each step of a
// profile sequence that crosses every ordered pair of locales exactly once.

#include <cstddef>
#include <cstdint>
#include <map>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "tests/oracle/oracle_test_support.h"

#ifndef FORMULON_ORACLE_SWITCH_TARGETS
#define FORMULON_ORACLE_SWITCH_TARGETS ""
#endif
#ifndef FORMULON_ORACLE_TARGETS_DIR
#define FORMULON_ORACLE_TARGETS_DIR ""
#endif

namespace formulon {
namespace tests {
namespace oracle {
namespace {

using namespace support;

constexpr char kSuite[] = "locale_tokens";
constexpr std::uint32_t kCaseRow = 0U;
constexpr std::uint32_t kCaseCol = 25U;

struct SwitchTarget {
  std::string id;
  ExcelProfile profile;
  std::map<std::string, const JsonValue*> expect_by_case;  // absent = skipped there
  std::vector<OracleCase> cases;                           // owns the JSON the map points into
};

std::map<std::string, std::string> load_skip_file(const std::string& path) {
  std::map<std::string, std::string> out;
  if (path.empty()) {
    return out;
  }
  auto parsed = parse_json_file(path);
  if (!parsed.has_value()) {
    return out;
  }
  if (const JsonValue* skips = parsed.value().find("skips"); skips != nullptr && skips->is_object()) {
    for (const auto& [case_id, reason] : skips->as_object()) {
      out.emplace(case_id, reason.is_string() ? reason.as_string() : std::string());
    }
  }
  return out;
}

// FORMULON_ORACLE_SWITCH_TARGETS is `<target>|<skip file>` entries joined by `;`.
std::vector<SwitchTarget> load_targets() {
  std::vector<SwitchTarget> out;
  std::stringstream entries(FORMULON_ORACLE_SWITCH_TARGETS);
  std::string entry;
  while (std::getline(entries, entry, ';')) {
    if (entry.empty()) {
      continue;
    }
    const std::size_t bar = entry.find('|');
    SwitchTarget t;
    t.id = entry.substr(0, bar);
    if (!parse_excel_profile_id(t.id, &t.profile)) {
      ADD_FAILURE() << "switch target is not a profile id: " << t.id;
      continue;
    }
    const std::map<std::string, std::string> skips =
        load_skip_file(bar == std::string::npos ? std::string() : entry.substr(bar + 1));
    t.cases = load_oracle_cases(std::string(FORMULON_ORACLE_TARGETS_DIR) + "/" + t.id + "/golden");
    for (const OracleCase& c : t.cases) {
      if (c.suite != kSuite || !c.skipped_reason.empty() || skips.count(c.case_id) != 0U) {
        continue;
      }
      if (const JsonValue* expect = c.raw_case.find("expect"); expect != nullptr && expect->is_object()) {
        t.expect_by_case.emplace(c.case_id, expect);
      }
    }
    out.push_back(std::move(t));
  }
  return out;
}

// Steps of 1..n-1 around a prime-sized ring each return to the start and
// together traverse every ordered pair once (n = 7 gated locales).
std::vector<std::size_t> pair_covering_sequence(std::size_t n) {
  std::vector<std::size_t> seq{0U};
  for (std::size_t step = 1U; step < n; ++step) {
    std::size_t at = 0U;
    do {
      at = (at + step) % n;
      seq.push_back(at);
    } while (at != 0U);
  }
  return seq;
}

TEST(OracleLocaleSwitch, EveryLocaleTransitionMatchesTheTargetGolden) {
  const std::vector<SwitchTarget> targets = load_targets();
  ASSERT_GE(targets.size(), 2U) << "FORMULON_ORACLE_SWITCH_TARGETS names fewer than two targets";
  const std::vector<std::size_t> sequence = pair_covering_sequence(targets.size());

  // The case list and setup come from the first target; every target captured
  // the same suite inputs.
  std::size_t checked = 0U;
  for (const OracleCase& c : targets.front().cases) {
    if (c.suite != kSuite || c.case_id == "<load-error>") {
      continue;
    }
    const JsonValue* formula_v = c.raw_case.find("formula");
    ASSERT_TRUE(formula_v != nullptr && formula_v->is_string()) << c.case_id;

    Workbook wb = Workbook::create();
    if (const JsonValue* setup = c.raw_case.find("setup"); setup != nullptr && setup->is_object()) {
      for (const auto& [addr, spec] : setup->as_object()) {
        std::uint32_t row = 0U;
        std::uint32_t col = 0U;
        ASSERT_TRUE(a1_to_row_col(addr, &row, &col)) << c.case_id << ": setup address " << addr;
        const JsonValue* kind = spec.find("kind");
        const JsonValue* value = spec.find(kind != nullptr && kind->as_string() == "formula" ? "formula" : "value");
        ASSERT_TRUE(kind != nullptr && value != nullptr) << c.case_id << ": setup " << addr;
        if (kind->as_string() == "formula") {
          ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, col, value->as_string())));
        } else if (kind->as_string() == "text") {
          ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(0U, row, col, value->as_string())));
        } else if (kind->as_string() == "number") {
          ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, row, col, Value::number(value->as_number()))));
        } else if (kind->as_string() == "bool") {
          ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, row, col, Value::boolean(value->as_bool()))));
        } else {
          FAIL() << c.case_id << ": unsupported setup kind " << kind->as_string();
        }
      }
    }
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, kCaseRow, kCaseCol, formula_v->as_string())));

    for (std::size_t i = 0U; i < sequence.size(); ++i) {
      const SwitchTarget& t = targets[sequence[i]];
      wb.set_excel_profile(t.profile);
      ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
      const auto expect = t.expect_by_case.find(c.case_id);
      if (expect == t.expect_by_case.end()) {
        continue;
      }
      const Cell* cell = wb.sheet(0U).cell_at(kCaseRow, kCaseCol);
      ASSERT_NE(cell, nullptr) << c.case_id;
      const std::string diff =
          compare_value(*expect->second, cell->cached_value, c.tolerance_abs, c.tolerance_rel, c.compare_mode);
      EXPECT_TRUE(diff.empty()) << c.case_id << " under " << t.id << " (step " << i << ", from "
                                << (i == 0U ? std::string("a new workbook") : targets[sequence[i - 1U]].id)
                                << "): " << diff << "\n  formula: " << formula_v->as_string();
      ++checked;
    }
  }
  EXPECT_GT(checked, 0U) << "no locale_tokens cases were checked";
}

}  // namespace
}  // namespace oracle
}  // namespace tests
}  // namespace formulon
