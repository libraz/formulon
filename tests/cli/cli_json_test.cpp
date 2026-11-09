// `formulon_cli` `--json` output of `dump` and `paginate`.
#include <fstream>
#include <string>

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

  const CliRun r = run_cli({"dump", "--view", "values", "--json", path});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  EXPECT_EQ(r.stdout_text, kFixtureDumpJson);

  const CliRun explicit_formulas = run_cli({"dump", "--json", "--view", "formulas", path});
  ASSERT_EQ(explicit_formulas.exit_code, 0) << explicit_formulas.stderr_text;
  EXPECT_EQ(explicit_formulas.stdout_text, kFixtureDumpJson);
}

TEST(FormulonCli, DumpJsonFailureLeavesStdoutEmpty) {
  const CliRun r = run_cli({"dump", "--json", temp_path("dump_json_missing.xlsx")});
  EXPECT_EQ(r.exit_code, 66);
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
  EXPECT_EQ(r.exit_code, 64);
  EXPECT_EQ(r.stdout_text, "");
}

TEST(FormulonCli, JsonEnvelopeEvalScalarIsATwoDimensionalRowsArray) {
  const CliRun r = run_cli({"eval", "--json", "=SUM(1,2,3)"});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "{\"format\":1,\"rows\":[[{\"kind\":\"number\",\"value\":6}]]}\n");
}

TEST(FormulonCli, JsonEnvelopeEvalErrorValueIsACell) {
  const CliRun r = run_cli({"eval", "--json", "=1/0"});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  EXPECT_EQ(r.stdout_text.rfind("{\"format\":1,\"rows\":[[{\"kind\":\"error\"", 0), 0U) << r.stdout_text;
  EXPECT_NE(r.stdout_text.find("}]]}\n"), std::string::npos) << r.stdout_text;
}

TEST(FormulonCli, JsonEnvelopeDumpSheetsListsSheetNames) {
  const std::string path = temp_path("json_envelope_sheets.xlsx");
  PathGuard guard(path);
  ASSERT_TRUE(write_fixture_workbook(path));
  const CliRun r = run_cli({"dump", "--view", "sheets", "--json", path});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "{\"format\":1,\"sheets\":[{\"name\":\"Sheet1\"}]}\n");

  const CliRun json_first = run_cli({"dump", "--json", "--view", "sheets", path});
  EXPECT_EQ(json_first.stdout_text, r.stdout_text);
}

TEST(FormulonCli, JsonEnvelopeDumpMetadataListsNamesTablesAndParts) {
  const std::string path = temp_path("json_envelope_metadata.xlsx");
  PathGuard guard(path);
  ASSERT_TRUE(write_out_of_range_defined_names_fixture(path));
  const CliRun r = run_cli({"dump", "--view", "metadata", "--json", path});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  EXPECT_EQ(r.stdout_text,
            "{\"format\":1,\"defined_names\":["
            "{\"name\":\"Before\",\"formula\":\"=1\",\"scope\":null},"
            "{\"name\":\"Bad\",\"formula\":\"=Sheet1!$A$1\",\"scope\":\"#99\"},"
            "{\"name\":\"After\",\"formula\":\"=2\",\"scope\":null}],"
            "\"tables\":[],\"passthrough_parts\":[]}\n");
}

TEST(FormulonCli, JsonEnvelopeEveryStdoutFormStartsWithFormatOne) {
  const std::string path = temp_path("json_envelope_all.xlsx");
  PathGuard guard(path);
  ASSERT_TRUE(write_fixture_workbook(path));
  for (const std::vector<std::string>& args : {
           std::vector<std::string>{"eval", "--json", "=1"},
           std::vector<std::string>{"eval", "--json", "=SEQUENCE(2,2)"},
           std::vector<std::string>{"dump", "--json", path},
           std::vector<std::string>{"dump", "--json", "--view", "values", path},
           std::vector<std::string>{"dump", "--json", "--view", "sheets", path},
           std::vector<std::string>{"dump", "--json", "--view", "metadata", path},
           std::vector<std::string>{"paginate", "--json", path},
           std::vector<std::string>{"profiles", "--json"},
       }) {
    const CliRun r = run_cli(args);
    ASSERT_EQ(r.exit_code, 0) << args.front() << " stderr=" << r.stderr_text;
    EXPECT_EQ(r.stdout_text.rfind("{\"format\":1,", 0), 0U) << args.front() << " stdout=" << r.stdout_text;
    EXPECT_EQ(r.stdout_text.back(), '\n');
  }
}

TEST(FormulonCli, RecalcJsonReportsOutputAndByteCount) {
  const std::string path = temp_path("recalc_json.xlsx");
  const std::string out_path = temp_path("recalc_json_out.xlsx");
  PathGuard guard(path);
  PathGuard out_guard(out_path);
  ASSERT_TRUE(write_fixture_workbook(path));
  const CliRun r = run_cli({"recalc", "--json", "--quiet", path, "-o", out_path});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  std::ifstream written(out_path, std::ios::binary | std::ios::ate);
  ASSERT_TRUE(written.good());
  const std::string expected = "{\"format\":1,\"status\":\"ok\",\"output\":\"" + out_path +
                               "\",\"bytes\":" + std::to_string(static_cast<std::streamoff>(written.tellg())) +
                               ",\"warnings\":[]}\n";
  EXPECT_EQ(r.stdout_text, expected);
}

TEST(FormulonCli, RecalcJsonListsLossCountersInsteadOfWarningOnStderr) {
  const std::string input = temp_path("recalc_json_lossy.xlsx");
  const std::string output = temp_path("recalc_json_lossy.xlsb");
  PathGuard input_guard(input);
  PathGuard output_guard(output);
  ASSERT_TRUE(write_lossy_xlsx_fixture(input));
  const CliRun r = run_cli({"recalc", "--json", "--quiet", input, "-o", output});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  EXPECT_NE(r.stdout_text.find("{\"kind\":\"downgraded_formula_count\",\"count\":1}"), std::string::npos)
      << r.stdout_text;
  EXPECT_NE(r.stdout_text.find("{\"kind\":\"deferred_feature_count\",\"count\":1}"), std::string::npos)
      << r.stdout_text;
  EXPECT_EQ(r.stderr_text.find("warning"), std::string::npos) << r.stderr_text;
}

TEST(FormulonCli, RecalcJsonIsRefusedWhenTheWorkbookGoesToStdout) {
  const std::string path = temp_path("recalc_json_stdout.xlsx");
  PathGuard guard(path);
  ASSERT_TRUE(write_fixture_workbook(path));
  const CliRun r = run_cli({"recalc", "--json", path, "-o", "-"});
  EXPECT_EQ(r.exit_code, 64);
  EXPECT_EQ(r.stdout_text, "");
}
