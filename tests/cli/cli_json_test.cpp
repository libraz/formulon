// `formulon_cli` `--json` output of `dump` and `paginate`.
#include "cli_test_support.h"

using namespace formulon::cli_test;

namespace {

constexpr const char* kFixtureDumpJson =
    "{\"format\":1,\"sheets\":[{\"name\":\"Sheet1\",\"cells\":["
    "{\"ref\":\"A1\",\"kind\":\"number\",\"value\":7},"
    "{\"ref\":\"B1\",\"kind\":\"number\",\"value\":8,\"formula\":\"=A1+1\"},"
    "{\"ref\":\"A2\",\"kind\":\"text\",\"value\":\"q\\\"t\\nz\"}"
    "]}]}\n";

}  // namespace

TEST(FormulonCli, DumpJsonListsCellsWithKindValueAndFormula) {
  const std::string path = temp_path("dump_json.xlsx");
  PathGuard guard(path);
  ASSERT_TRUE(write_fixture_workbook(path, "q\"t\nz"));

  const CliRun r = run_cli({"dump", "--json", path});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  EXPECT_EQ(r.stdout_text, kFixtureDumpJson);
}

TEST(FormulonCli, DumpJsonValuesFlagKeepsTheSameShape) {
  const std::string path = temp_path("dump_json_values.xlsx");
  PathGuard guard(path);
  ASSERT_TRUE(write_fixture_workbook(path, "q\"t\nz"));

  const CliRun r = run_cli({"dump", "--values", "--json", path});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  EXPECT_EQ(r.stdout_text, kFixtureDumpJson);

  const CliRun explicit_formulas = run_cli({"dump", "--json", "--formulas", path});
  ASSERT_EQ(explicit_formulas.exit_code, 0) << explicit_formulas.stderr_text;
  EXPECT_EQ(explicit_formulas.stdout_text, kFixtureDumpJson);
}

TEST(FormulonCli, DumpJsonRejectsModesWithoutCells) {
  const std::string path = temp_path("dump_json_modes.xlsx");
  PathGuard guard(path);
  ASSERT_TRUE(write_fixture_workbook(path));

  for (const char* mode : {"--sheets", "--metadata"}) {
    const CliRun r = run_cli({"dump", mode, "--json", path});
    EXPECT_EQ(r.exit_code, 64) << mode;
    EXPECT_EQ(r.stdout_text, "") << mode;
  }
}

TEST(FormulonCli, DumpJsonFailureLeavesStdoutEmpty) {
  const CliRun r = run_cli({"dump", "--json", temp_path("dump_json_missing.xlsx")});
  EXPECT_EQ(r.exit_code, 1);
  EXPECT_EQ(r.stdout_text, "");
  EXPECT_FALSE(r.stderr_text.empty());
}

TEST(FormulonCli, PaginateJsonListsOneA1RangePerPage) {
  const std::string path = temp_path("paginate_json.xlsx");
  PathGuard guard(path);
  ASSERT_TRUE(write_fixture_workbook(path));

  const CliRun r = run_cli({"paginate", "--json", path});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  // The fixture fills A1:B1 only, so the used range fits on a single page.
  EXPECT_EQ(r.stdout_text, "{\"format\":1,\"sheets\":[{\"name\":\"Sheet1\",\"pages\":[{\"range\":\"A1:B1\"}]}]}\n");

  const CliRun plain = run_cli({"paginate", path});
  ASSERT_EQ(plain.exit_code, 0) << plain.stderr_text;
  EXPECT_NE(plain.stdout_text.find("pages=1\n"), std::string::npos) << plain.stdout_text;
}

TEST(FormulonCli, PaginateJsonFailureLeavesStdoutEmpty) {
  const std::string path = temp_path("paginate_json_bad_sheet.xlsx");
  PathGuard guard(path);
  ASSERT_TRUE(write_fixture_workbook(path));

  const CliRun r = run_cli({"paginate", "--json", "--sheet", "99", path});
  EXPECT_EQ(r.exit_code, 1);
  EXPECT_EQ(r.stdout_text, "");
}
