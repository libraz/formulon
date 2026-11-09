// `formulon_cli` stdin / stdout `-` forms and `--opt=value` parsing.
#include <sys/stat.h>

#include "cli_test_support.h"

using namespace formulon::cli_test;

namespace {

std::string slurp_file(const std::string& path) {
  std::ifstream in(path, std::ios::binary);
  return std::string((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
}

bool loads_as_workbook(const std::string& bytes) {
  fm_workbook_t* wb = nullptr;
  const fm_status_t rc = fm_workbook_load(reinterpret_cast<const std::uint8_t*>(bytes.data()), bytes.size(), &wb);
  fm_workbook_destroy(wb);
  return rc == 0;
}

}  // namespace

TEST(FormulonCli, StdinDashFeedsEveryInputSubcommand) {
  const std::string path = temp_path("stdin_dash.xlsx");
  PathGuard guard(path);
  ASSERT_TRUE(write_fixture_workbook(path));

  const CliRun dump = run_cli({"dump", "--view", "sheets", "-"}, false, false, path);
  EXPECT_EQ(dump.exit_code, 0) << dump.stderr_text;
  EXPECT_EQ(dump.stdout_text, "Sheet1\n");

  const CliRun dump_after_terminator = run_cli({"dump", "--view", "sheets", "--", "-"}, false, false, path);
  EXPECT_EQ(dump_after_terminator.exit_code, 0) << dump_after_terminator.stderr_text;
  EXPECT_EQ(dump_after_terminator.stdout_text, "Sheet1\n");

  const CliRun paginate = run_cli({"paginate", "-"}, false, false, path);
  EXPECT_EQ(paginate.exit_code, 0) << paginate.stderr_text;
  EXPECT_EQ(paginate.stdout_text, "name=Sheet1\npageCount=1\nprintArea=\nrowBreaks=\ncolumnBreaks=\npages=A1:B1\n");

  const std::string output = temp_path("stdin_dash_out.xlsx");
  PathGuard output_guard(output);
  const CliRun recalc = run_cli({"recalc", "-", "-o", output}, false, false, path);
  EXPECT_EQ(recalc.exit_code, 0) << recalc.stderr_text;
  EXPECT_TRUE(loads_as_workbook(slurp_file(output)));
}

TEST(FormulonCli, StdinDashEmptyInputIsAnEngineFailure) {
  const std::string empty = temp_path("stdin_dash_empty.xlsx");
  PathGuard guard(empty);
  ASSERT_TRUE(write_bytes(empty, ""));

  const CliRun from_stdin = run_cli({"dump", "--view", "sheets", "-"}, false, false, empty);
  EXPECT_EQ(from_stdin.exit_code, 1);
  EXPECT_EQ(from_stdin.stdout_text, "");

  const CliRun from_file = run_cli({"dump", "--view", "sheets", empty});
  EXPECT_EQ(from_file.exit_code, 1);
}

TEST(FormulonCli, StdinDashCorruptInputIsAnEngineFailure) {
  const std::string junk = temp_path("stdin_dash_junk.xlsx");
  PathGuard guard(junk);
  ASSERT_TRUE(write_bytes(junk, "this is not a workbook"));
  const CliRun r = run_cli({"dump", "--view", "sheets", "-"}, false, false, junk);
  EXPECT_EQ(r.exit_code, 1);
}

TEST(FormulonCli, StdinDashOnATerminalIsAUsageError) {
  SpawnOptions options;
  options.stdin_tty = true;
  for (const std::vector<std::string>& args : {
           std::vector<std::string>{"dump", "--view", "sheets", "-"},
           std::vector<std::string>{"paginate", "-"},
           std::vector<std::string>{"recalc", "-", "-o", temp_path("stdin_tty_unused.xlsx")},
       }) {
    const CliRun r = run_cli_spawn(args, options);
    EXPECT_EQ(r.exit_code, 64) << args.front() << " stderr=" << r.stderr_text;
    EXPECT_EQ(r.stdout_text, "") << args.front();
    EXPECT_NE(r.stderr_text.find("terminal"), std::string::npos) << args.front() << " stderr=" << r.stderr_text;
  }
}

TEST(FormulonCli, StdinDashDotSlashNamesAFileCalledDash) {
  const std::string name = "-";
  PathGuard guard(name);
  ASSERT_TRUE(write_fixture_workbook(name));
  const CliRun r = run_cli({"dump", "--view", "sheets", "./-"}, false, false, "/dev/null");
  EXPECT_EQ(r.exit_code, 0) << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "Sheet1\n");
}

TEST(FormulonCli, StdinDashIsTheFormulaSoAnEmptyStdinIsMalformed) {
  const CliRun r = run_cli({"eval", "-"}, false, false, "/dev/null");
  EXPECT_EQ(r.exit_code, 65) << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "");
}

TEST(FormulonCli, StdoutDashWritesTheWorkbookToStdout) {
  const std::string in = temp_path("stdout_dash_in.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));

  const CliRun r = run_cli({"recalc", "--quiet", in, "-o", "-"});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  ASSERT_GT(r.stdout_text.size(), 4U);
  EXPECT_EQ(r.stdout_text.compare(0, 4, "PK\x03\x04"), 0);
  EXPECT_TRUE(loads_as_workbook(r.stdout_text));
  EXPECT_EQ(r.stderr_text, "");
}

TEST(FormulonCli, StdoutDashStatusLineGoesToStderr) {
  const std::string in = temp_path("stdout_dash_status.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  const CliRun r = run_cli({"recalc", in, "--output=-"});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  EXPECT_EQ(r.stdout_text.compare(0, 2, "PK"), 0);
  EXPECT_NE(r.stderr_text.find("recalc: ok"), std::string::npos) << r.stderr_text;
}

TEST(FormulonCli, StdoutDashPipesStdinToStdout) {
  const std::string in = temp_path("stdout_dash_both.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  const CliRun r = run_cli({"recalc", "--quiet", "-", "-o", "-"}, false, false, in);
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  EXPECT_TRUE(loads_as_workbook(r.stdout_text));
}

TEST(FormulonCli, StdoutDashDotSlashNamesAFileCalledDash) {
  const std::string in = temp_path("stdout_dash_file.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  const std::string name = "-";
  PathGuard out_guard(name);
  std::remove(name.c_str());

  const CliRun r = run_cli({"recalc", "--quiet", "--format", "xlsx", in, "-o", "./-"});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "");
  EXPECT_TRUE(loads_as_workbook(slurp_file(name)));
}

TEST(FormulonCli, StdoutDashIsRefusedBySubcommandsWithoutOutputFlag) {
  const std::string in = temp_path("stdout_dash_no_o.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  const CliRun r = run_cli({"dump", "-o", "-", in});
  EXPECT_EQ(r.exit_code, 64);
}

TEST(FormulonCli, StdoutTtyRefusedForBinaryOutput) {
  const std::string in = temp_path("stdout_tty_in.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));

  SpawnOptions options;
  options.stdout_tty = true;
  const CliRun r = run_cli_spawn({"recalc", in, "-o", "-"}, options);
  EXPECT_EQ(r.exit_code, 64) << r.stderr_text;
  EXPECT_NE(r.stderr_text.find("terminal"), std::string::npos) << r.stderr_text;
}

TEST(FormulonCli, StdoutTtyRefusedBeforeReadingInput) {
  // The refusal is an argument-level check, so it wins over an input that
  // would otherwise fail to load.
  SpawnOptions options;
  options.stdout_tty = true;
  const CliRun r = run_cli_spawn({"recalc", temp_path("stdout_tty_absent.xlsx"), "-o", "-"}, options);
  EXPECT_EQ(r.exit_code, 64) << r.stderr_text;
}

TEST(FormulonCli, StdoutTtyDoesNotAffectTextOutput) {
  SpawnOptions options;
  options.stdout_tty = true;
  const CliRun r = run_cli_spawn({"profiles"}, options);
  EXPECT_EQ(r.exit_code, 0) << r.stderr_text;
}

TEST(FormulonCli, OptEqualsMatchesTheSpacedForm) {
  const std::string in = temp_path("opt_equals.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));

  const CliRun spaced = run_cli({"eval", "--repeat", "2", "--now", "2026-10-09T09:30:00", "=TODAY()"});
  const CliRun equals = run_cli({"eval", "--repeat=2", "--now=2026-10-09T09:30:00", "=TODAY()"});
  ASSERT_EQ(spaced.exit_code, 0) << spaced.stderr_text;
  ASSERT_EQ(equals.exit_code, 0) << equals.stderr_text;
  EXPECT_EQ(equals.stdout_text, spaced.stdout_text);
  EXPECT_NE(equals.stderr_text.find("2 iterations"), std::string::npos) << equals.stderr_text;

  const CliRun dump = run_cli({"dump", "--profile=win-365-en_US", "--view", "sheets", in});
  EXPECT_EQ(dump.exit_code, 0) << dump.stderr_text;
  EXPECT_EQ(dump.stdout_text, "Sheet1\n");

  const CliRun paginate = run_cli({"paginate", "--sheet=0", in});
  EXPECT_EQ(paginate.exit_code, 0) << paginate.stderr_text;
  EXPECT_EQ(paginate.stdout_text, run_cli({"paginate", "--sheet", "0", in}).stdout_text);
}

TEST(FormulonCli, OptEqualsAcceptsRecalcValueOptions) {
  const std::string in = temp_path("opt_equals_recalc_in.xlsx");
  const std::string out = temp_path("opt_equals_recalc_out.xlsx");
  PathGuard guard_in(in);
  PathGuard guard_out(out);
  ASSERT_TRUE(write_fixture_workbook(in));

  const CliRun r = run_cli({"recalc", "--quiet", "--threads=1", "--output=" + out, in});
  EXPECT_EQ(r.exit_code, 0) << r.stderr_text;
  EXPECT_TRUE(loads_as_workbook(slurp_file(out)));
}

TEST(FormulonCli, OptEqualsOnAValuelessFlagIsAUsageError) {
  const std::string in = temp_path("opt_equals_valueless.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));

  for (const std::vector<std::string>& args : {
           std::vector<std::string>{"eval", "--json=1", "=1"},
           std::vector<std::string>{"paginate", "--json=1", in},
           std::vector<std::string>{"dump", "--json=", in},
           std::vector<std::string>{"recalc", "--iterative=yes", in, "-o", temp_path("unused.xlsx")},
           std::vector<std::string>{"recalc", "--quiet=1", in, "-o", temp_path("unused.xlsx")},
           std::vector<std::string>{"paginate", "--help=1"},
           std::vector<std::string>{"profiles", "--json=true"},
       }) {
    const CliRun r = run_cli(args);
    EXPECT_EQ(r.exit_code, 64) << args.front() << ' ' << args[1] << " stderr=" << r.stderr_text;
    EXPECT_EQ(r.stdout_text, "") << args.front() << ' ' << args[1];
    EXPECT_NE(r.stderr_text.find("does not take a value"), std::string::npos) << r.stderr_text;
  }
}

TEST(FormulonCli, OptEqualsEmptyValueIsAnEmptyValue) {
  const std::string in = temp_path("opt_equals_empty.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));

  const CliRun output = run_cli({"recalc", in, "--output="});
  EXPECT_EQ(output.exit_code, 64);
  EXPECT_NE(output.stderr_text.find("non-empty path"), std::string::npos) << output.stderr_text;

  EXPECT_EQ(run_cli({"paginate", "--sheet=", in}).exit_code, 64);
  EXPECT_EQ(run_cli({"eval", "--repeat=", "=1"}).exit_code, 64);
  EXPECT_EQ(run_cli({"eval", "--profile=", "=1"}).exit_code, 64);
}

TEST(FormulonCli, OptEqualsInvalidValuesAreStillRejected) {
  EXPECT_EQ(run_cli({"eval", "--profile=bogus", "=1"}).exit_code, 64);
  EXPECT_EQ(run_cli({"eval", "--now=yesterday", "=1"}).exit_code, 64);
  EXPECT_EQ(run_cli({"eval", "--repeat=0", "=1"}).exit_code, 64);
  EXPECT_EQ(run_cli({"eval", "--repeat=+5", "=1"}).exit_code, 64);
  EXPECT_EQ(run_cli({"eval", "--repeat= 5", "=1"}).exit_code, 64);
  EXPECT_EQ(run_cli({"eval", "--repeat=2x", "=1"}).exit_code, 64);
}

TEST(FormulonCli, OptEqualsStopsAtTheOptionTerminatorAndInFormulas) {
  const CliRun after = run_cli({"eval", "--", "--repeat=2"});
  EXPECT_EQ(after.exit_code, 0) << after.stderr_text;
  EXPECT_EQ(after.stderr_text.find("iterations"), std::string::npos) << after.stderr_text;

  const CliRun formula = run_cli({"eval", "=1+1=2"});
  EXPECT_EQ(formula.exit_code, 0) << formula.stderr_text;
  EXPECT_EQ(formula.stdout_text, "TRUE\n");

  const CliRun unknown = run_cli({"eval", "--nonsense=1", "=1"});
  EXPECT_EQ(unknown.exit_code, 64);
  EXPECT_NE(unknown.stderr_text.find("--nonsense=1"), std::string::npos) << unknown.stderr_text;
}
