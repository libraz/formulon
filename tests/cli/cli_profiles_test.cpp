// `formulon profiles` and the CLI's exit-code vocabulary.
#include <string>

#include "c_api/formulon_c.h"
#include "cli_test_support.h"
#include "excel_profile.h"

using namespace formulon::cli_test;

TEST(FormulonCli, ProfilesCommandPrintsOneIdPerLine) {
  std::string expected;
  for (const formulon::detail::ExcelProfileIdEntry& entry : formulon::detail::kExcelProfileIds) {
    expected += entry.id;
    expected += '\n';
  }
  const CliRun r = run_cli({"profiles"});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  EXPECT_EQ(r.stdout_text, expected);
}

TEST(FormulonCli, ProfilesCommandListsTheSameIdsAsTheCAbi) {
  std::size_t count = 0;
  ASSERT_EQ(fm_get_profile_id_count(&count), 0);
  std::string expected;
  for (std::size_t i = 0; i < count; ++i) {
    const char* id = nullptr;
    ASSERT_EQ(fm_get_profile_id_at(i, &id), 0);
    expected += id;
    expected += '\n';
  }
  const CliRun r = run_cli({"profiles"});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  EXPECT_EQ(r.stdout_text, expected);
}

TEST(FormulonCli, ProfilesCommandRejectsUnexpectedArguments) {
  const CliRun r = run_cli({"profiles", "extra"});
  EXPECT_EQ(r.exit_code, 64);
  EXPECT_EQ(r.stdout_text, "");
}

TEST(FormulonCli, ProfilesCommandIsReferencedByEveryUsage) {
  for (const char* command : {"eval", "recalc", "dump", "paginate"}) {
    const CliRun help = run_cli({command, "--help"});
    ASSERT_EQ(help.exit_code, 0) << command;
    EXPECT_NE(help.stdout_text.find("Profiles: see 'formulon profiles'"), std::string::npos) << command;
    EXPECT_EQ(help.stdout_text.find("mac-365-de_DE"), std::string::npos) << command;
  }
  const CliRun top = run_cli({"--help"});
  ASSERT_EQ(top.exit_code, 0);
  EXPECT_NE(top.stdout_text.find("Profiles: see 'formulon profiles'"), std::string::npos);
  EXPECT_EQ(top.stdout_text.find("mac-365-de_DE"), std::string::npos);
}

TEST(FormulonCli, ProfilesJsonListsTheSameIdsInTheSameOrder) {
  std::string expected = "{\"format\":1,\"profiles\":[";
  const char* separator = "";
  for (const formulon::detail::ExcelProfileIdEntry& entry : formulon::detail::kExcelProfileIds) {
    expected += separator;
    expected += std::string("\"") + entry.id + "\"";
    separator = ",";
  }
  expected += "]}\n";
  const CliRun r = run_cli({"profiles", "--json"});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  EXPECT_EQ(r.stdout_text, expected);
}

TEST(FormulonCli, ProfilesJsonWriteFailureExitsIoError) {
  const CliRun r = run_cli({"profiles", "--json"}, /*merge_streams=*/false, /*close_stdout=*/true);
  EXPECT_EQ(r.exit_code, 74);
}

TEST(FormulonCli, EvalEngineFailureExitsOne) {
  // A formula that is not valid UTF-8 is refused by the engine rather than
  // evaluated to a cell-level error.
  const CliRun plain = run_cli({"eval", "=\xff\xfe"});
  EXPECT_EQ(plain.exit_code, 1);
  EXPECT_EQ(plain.stdout_text, "");
}

TEST(FormulonCli, EvalEngineFailureExitsOneWithJson) {
  const CliRun json = run_cli({"eval", "--json", "=\xff\xfe"});
  EXPECT_EQ(json.exit_code, 1);
  EXPECT_EQ(json.stdout_text, "");
}

TEST(FormulonCli, EvalHelpStatesThatUnparsableFormulasExit65) {
  const CliRun r = run_cli({"eval", "--help"});
  ASSERT_EQ(r.exit_code, 0);
  EXPECT_NE(r.stdout_text.find("exits 65"), std::string::npos) << r.stdout_text;
  EXPECT_EQ(r.stdout_text.find("Malformed input is not an error"), std::string::npos) << r.stdout_text;
}

TEST(FormulonCli, ProfileHelpStatesThatFormulaTextStaysEnUsUnlessLocal) {
  const CliRun r = run_cli({"eval", "--help"});
  ASSERT_EQ(r.exit_code, 0);
  EXPECT_NE(r.stdout_text.find("Excel profile for evaluation semantics"), std::string::npos) << r.stdout_text;
  EXPECT_NE(r.stdout_text.find("(separators, function names); otherwise en-US"), std::string::npos) << r.stdout_text;
}

TEST(FormulonCli, DumpHelpDoesNotOfferTheEvalOnlyLocalOption) {
  const CliRun r = run_cli({"dump", "--help"});
  ASSERT_EQ(r.exit_code, 0);
  EXPECT_EQ(r.stdout_text.find("unless --local"), std::string::npos) << r.stdout_text;
}
