#include <filesystem>
#include <fstream>
#include <map>
#include <random>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "tests/oracle/json_reader.h"
#include "tests/oracle/oracle_runner.h"

namespace formulon {
namespace tests {
namespace oracle {
namespace {

// ---------------------------------------------------------------------------
// Divergence registry
// ---------------------------------------------------------------------------

// Creates an empty scratch directory for one test's throwaway golden.
//
// The name carries a random token because these tests are linked into more
// than one oracle binary (the primary and the IronCalc one) and CTest may
// run those concurrently: a fixed path would let one run delete the other's
// fixture mid-read.
std::filesystem::path make_scratch_dir(const std::string& label) {
  std::error_code ec;
  std::random_device entropy;
  const std::filesystem::path dir =
      std::filesystem::temp_directory_path(ec) / ("formulon_oracle_" + label + "_" + std::to_string(entropy()));
  std::filesystem::remove_all(dir, ec);
  std::filesystem::create_directories(dir, ec);
  return dir;
}

// A case the registry names is skipped even though its golden says nothing
// about the skip -- which is the state of every golden captured before the
// divergence was adjudicated. Without this the entry would only take effect
// once someone with an Excel host re-ran the capture, so the case would stay
// red on every other machine in the meantime.
TEST(OracleSkipRegistry, SkipsAGoldenCapturedBeforeTheEntryWasAdjudicated) {
  const std::map<std::string, std::string>& registry = oracle_skip_registry();
  if (registry.empty()) {
    GTEST_SKIP() << "no divergence projection is wired into this build (FORMULON_ORACLE_SKIP_FILE)";
  }
  const std::string& case_id = registry.begin()->first;
  const std::string& reason = registry.begin()->second;

  const std::filesystem::path dir = make_scratch_dir("skip_registry");
  std::error_code ec;
  ASSERT_TRUE(std::filesystem::is_directory(dir, ec)) << dir.string();

  {
    // Deliberately no `"skipped"` field, and an `expect` the engine cannot
    // satisfy: if the registry were not consulted the case would be
    // verified, and fail.
    std::ofstream out(dir / "skip_registry_probe.golden.json");
    ASSERT_TRUE(out.good());
    out << "{\"suite\": \"skip_registry_probe\", \"cases\": [{\"id\": \"" << case_id
        << "\", \"formula\": \"=1\", \"expect\": {\"kind\": \"number\", \"value\": 2}}]}\n";
  }

  const std::vector<OracleCase> cases = load_oracle_cases(dir.string());
  ASSERT_EQ(cases.size(), 1U);
  EXPECT_EQ(cases[0].case_id, case_id);
  EXPECT_EQ(cases[0].skipped_reason, reason);

  std::filesystem::remove_all(dir, ec);
}

// The baked field still wins for goldens that carry one, and it is what a
// case with no registry entry relies on.
TEST(OracleSkipRegistry, BakedSkippedFieldIsHonouredOnItsOwn) {
  const std::filesystem::path dir = make_scratch_dir("baked_skip");
  std::error_code ec;
  ASSERT_TRUE(std::filesystem::is_directory(dir, ec)) << dir.string();

  {
    std::ofstream out(dir / "baked_probe.golden.json");
    ASSERT_TRUE(out.good());
    out << "{\"suite\": \"baked_probe\", \"cases\": [{\"id\": \"baked_probe_case\", "
           "\"skipped\": \"captured with the skip already baked in\"}]}\n";
  }

  const std::vector<OracleCase> cases = load_oracle_cases(dir.string());
  ASSERT_EQ(cases.size(), 1U);
  EXPECT_EQ(cases[0].skipped_reason, "captured with the skip already baked in");

  std::filesystem::remove_all(dir, ec);
}

TEST(OracleJsonReader, AcceptsSubnormalNumbers) {
  auto parsed = parse_json("[5e-324, -5e-324, 2.2250738585072009e-308, 1e-400]");
  ASSERT_TRUE(parsed) << parsed.error().message;
  const auto& items = parsed.value().as_array();
  ASSERT_EQ(items.size(), 4U);
  EXPECT_EQ(items[0].as_number(), 4.9406564584124654e-324);
  EXPECT_EQ(items[1].as_number(), -4.9406564584124654e-324);
  EXPECT_GT(items[2].as_number(), 0.0);
  EXPECT_EQ(items[3].as_number(), 0.0);
}

TEST(OracleJsonReader, RejectsOverflowingNumbers) {
  EXPECT_FALSE(parse_json("[1e999]"));
}

}  // namespace
}  // namespace oracle
}  // namespace tests
}  // namespace formulon
