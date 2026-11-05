// Plain CLI output escapes only newline, carriage return, tab and backslash,
// so a line pasted back reproduces the original text.
#include "cli_test_support.h"

using namespace formulon::cli_test;

namespace {

bool write_plain_escape_fixture(const std::string& path) {
  fm_workbook_t* wb = nullptr;
  if (fm_workbook_create(&wb) != 0) {
    return false;
  }
  bool ok = fm_workbook_set_formula(wb, 0, 0, 0, "=\"a\"&\"x\"") == 0 &&
            fm_workbook_set_text(wb, 0, 1, 0, "q\"t\tb\\s\r\nz") == 0 && fm_workbook_recalc(wb) == 0;
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ok = ok && fm_workbook_save(wb, &bytes, &len) == 0;
  if (ok) {
    std::ofstream out(path, std::ios::binary | std::ios::trunc);
    out.write(reinterpret_cast<const char*>(bytes), static_cast<std::streamsize>(len));
    ok = static_cast<bool>(out);
  }
  fm_buffer_free(bytes);
  fm_workbook_destroy(wb);
  return ok;
}

}  // namespace

TEST(FormulonCli, PlainEscapeDumpFormulasKeepsQuotes) {
  const std::string path = temp_path("plain_escape_formulas.xlsx");
  PathGuard guard(path);
  ASSERT_TRUE(write_plain_escape_fixture(path));
  CliRun r = run_cli({"dump", "--formulas", path});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "Sheet1!A1 =\"a\"&\"x\"\n");
}

TEST(FormulonCli, PlainEscapeDumpValuesEscapesOnlyFourCharacters) {
  const std::string path = temp_path("plain_escape_values.xlsx");
  PathGuard guard(path);
  ASSERT_TRUE(write_plain_escape_fixture(path));
  CliRun r = run_cli({"dump", "--values", path});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "Sheet1!A1 ax\nSheet1!A2 q\"t\\tb\\\\s\\r\\nz\n");
}

TEST(FormulonCli, PlainEscapeEvalPlainKeepsQuotesButJsonStaysJson) {
  CliRun plain = run_cli({"eval", "=CHAR(9)&\"\\\"&CHAR(34)"});
  ASSERT_EQ(plain.exit_code, 0) << plain.stderr_text;
  EXPECT_EQ(plain.stdout_text, "\\t\\\\\"\n");

  CliRun json = run_cli({"eval", "--json", "=CHAR(34)"});
  ASSERT_EQ(json.exit_code, 0) << json.stderr_text;
  EXPECT_EQ(json.stdout_text, "{\"kind\":\"text\",\"value\":\"\\\"\"}\n");
}

TEST(FormulonCli, PlainEscapeLeavesOtherControlCharactersRaw) {
  CliRun r = run_cli({"eval", "=CHAR(1)"});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "\x01\n");
}

TEST(FormulonCli, PaginatePrintAreaUsesA1Notation) {
  const std::string path = temp_path("paginate_a1.xlsx");
  PathGuard guard(path);
  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_create(&wb), 0);
  ASSERT_EQ(fm_workbook_set_number(wb, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_print_area(wb, 0, "B2:D5,F1:G3"), 0);
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ASSERT_EQ(fm_workbook_save(wb, &bytes, &len), 0);
  {
    std::ofstream out(path, std::ios::binary | std::ios::trunc);
    out.write(reinterpret_cast<const char*>(bytes), static_cast<std::streamsize>(len));
  }
  fm_buffer_free(bytes);
  fm_workbook_destroy(wb);
  CliRun r = run_cli({"paginate", path});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  EXPECT_NE(r.stdout_text.find("print_area=B2:D5,F1:G3\n"), std::string::npos) << r.stdout_text;
}

TEST(FormulonCli, EvalMissingFormulaMessage) {
  CliRun r = run_cli({"eval"});
  EXPECT_EQ(r.exit_code, 64);
  EXPECT_NE(r.stderr_text.find("formulon: eval: missing formula\n"), std::string::npos) << r.stderr_text;
}

TEST(FormulonCli, HelpDocumentsFlagsExitCodesAndInputFormats) {
  CliRun r = run_cli({"--help"});
  ASSERT_EQ(r.exit_code, 0);
  for (const char* needle :
       {"--repeat", "--json", "--iterative", "--threads", "--quiet", "--sheet", "--output", "--formulas", "--values",
        "--sheets", "--metadata", "Exit codes", "64", "recalculates", "xlsx or xlsb"}) {
    EXPECT_NE(r.stdout_text.find(needle), std::string::npos) << needle;
  }
}
