// `formulon_cli` exit codes: 66 input unavailable, 73 output cannot be
// created, 74 I/O error while reading or writing.
#include <sys/stat.h>

#include "cli_test_support.h"

using namespace formulon::cli_test;

namespace {

// A directory the test can chmod without leaving debris; root ignores mode
// bits, so tests that depend on a permission refusal skip there.
bool running_as_root() {
  return ::geteuid() == 0;
}

const std::vector<std::string>& input_commands() {
  static const std::vector<std::string> commands{"dump", "paginate", "recalc"};
  return commands;
}

std::vector<std::string> command_for(const std::string& subcommand, const std::string& input) {
  if (subcommand == "recalc") {
    return {"recalc", "--quiet", input, "-o", temp_path("exit_code_unused.xlsx")};
  }
  return {subcommand, input};
}

}  // namespace

TEST(FormulonCli, ExitCodeInputMissingFileIs66) {
  for (const std::string& subcommand : input_commands()) {
    const CliRun r = run_cli(command_for(subcommand, temp_path("exit_code_absent.xlsx")));
    EXPECT_EQ(r.exit_code, 66) << subcommand;
    EXPECT_EQ(r.stdout_text, "") << subcommand;
    EXPECT_NE(r.stderr_text.find("No such file"), std::string::npos) << subcommand << " stderr=" << r.stderr_text;
  }
}

TEST(FormulonCli, ExitCodeInputDirectoryIs66WithItsOwnMessage) {
  const std::string dir = temp_path("exit_code_input_dir");
  ASSERT_EQ(::mkdir(dir.c_str(), 0700), 0);
  DirGuard guard(dir);
  for (const std::string& subcommand : input_commands()) {
    const CliRun r = run_cli(command_for(subcommand, dir));
    EXPECT_EQ(r.exit_code, 66) << subcommand;
    EXPECT_NE(r.stderr_text.find("Is a directory"), std::string::npos) << subcommand << " stderr=" << r.stderr_text;
    EXPECT_EQ(r.stderr_text.find("No such file"), std::string::npos) << subcommand;
  }
}

TEST(FormulonCli, ExitCodeInputUnreadableFileIs66WithItsOwnMessage) {
  if (running_as_root()) {
    GTEST_SKIP() << "mode bits do not restrict root";
  }
  const std::string path = temp_path("exit_code_unreadable.xlsx");
  PathGuard guard(path);
  ASSERT_TRUE(write_fixture_workbook(path));
  ASSERT_EQ(::chmod(path.c_str(), 0), 0);
  for (const std::string& subcommand : input_commands()) {
    const CliRun r = run_cli(command_for(subcommand, path));
    EXPECT_EQ(r.exit_code, 66) << subcommand;
    EXPECT_NE(r.stderr_text.find("Permission denied"), std::string::npos) << subcommand << " stderr=" << r.stderr_text;
  }
  ::chmod(path.c_str(), 0600);
}

TEST(FormulonCli, ExitCodeInputEmptyAndCorruptFilesAreEngineFailures) {
  const std::string empty = temp_path("exit_code_empty.xlsx");
  const std::string junk = temp_path("exit_code_junk.xlsx");
  PathGuard empty_guard(empty);
  PathGuard junk_guard(junk);
  ASSERT_TRUE(write_bytes(empty, ""));
  ASSERT_TRUE(write_bytes(junk, "not a workbook"));
  for (const std::string& subcommand : input_commands()) {
    EXPECT_EQ(run_cli(command_for(subcommand, empty)).exit_code, 1) << subcommand;
    EXPECT_EQ(run_cli(command_for(subcommand, junk)).exit_code, 1) << subcommand;
  }
}

TEST(FormulonCli, ExitCodeInputUsageErrorsStayAt64) {
  EXPECT_EQ(run_cli({"dump"}).exit_code, 64);
  EXPECT_EQ(run_cli({"bogus"}).exit_code, 64);
  EXPECT_EQ(run_cli({"recalc", "--json", "x.xlsx", "-o", "-"}).exit_code, 64);
}

TEST(FormulonCli, ExitCodeOutputMissingDirectoryIs73) {
  const std::string in = temp_path("exit_code_out_in.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));

  const CliRun r = run_cli({"recalc", "--quiet", in, "-o", temp_path("no_such_dir/out.xlsx")});
  EXPECT_EQ(r.exit_code, 73);
  EXPECT_NE(r.stderr_text.find("cannot write"), std::string::npos) << r.stderr_text;
}

TEST(FormulonCli, ExitCodeOutputUnwritableDirectoryIs73) {
  if (running_as_root()) {
    GTEST_SKIP() << "mode bits do not restrict root";
  }
  const std::string in = temp_path("exit_code_ro_in.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  const std::string dir = temp_path("exit_code_ro_dir");
  ASSERT_EQ(::mkdir(dir.c_str(), 0500), 0);
  DirGuard dir_guard(dir);

  const CliRun r = run_cli({"recalc", "--quiet", in, "-o", dir + "/out.xlsx"});
  EXPECT_EQ(r.exit_code, 73);
  EXPECT_NE(r.stderr_text.find("Permission denied"), std::string::npos) << r.stderr_text;
}

TEST(FormulonCli, ExitCodeOutputDirectoryTargetIs73) {
  const std::string in = temp_path("exit_code_dir_target_in.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  const std::string dir = temp_path("exit_code_dir_target");
  ASSERT_EQ(::mkdir(dir.c_str(), 0700), 0);
  DirGuard dir_guard(dir);

  const CliRun r = run_cli({"recalc", "--quiet", "--format", "xlsx", in, "-o", dir});
  EXPECT_EQ(r.exit_code, 73);
}

TEST(FormulonCli, ExitCodeIoClosedStdoutIs74) {
  const std::string in = temp_path("exit_code_io_closed.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  for (const std::vector<std::string>& args : {
           std::vector<std::string>{"dump", "--view", "sheets", in},
           std::vector<std::string>{"paginate", in},
           std::vector<std::string>{"profiles"},
           std::vector<std::string>{"recalc", "--quiet", in, "-o", "-"},
       }) {
    const CliRun r = run_cli(args, /*merge_streams=*/false, /*close_stdout=*/true);
    EXPECT_EQ(r.exit_code, 74) << args.front() << " stderr=" << r.stderr_text;
  }
}

TEST(FormulonCli, ExitCodeIoBrokenPipeOnStdoutIs74) {
  const std::string in = temp_path("exit_code_io_pipe.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));

  SpawnOptions options;
  options.stdout_broken_pipe = true;
  for (const std::vector<std::string>& args : {
           std::vector<std::string>{"dump", "--view", "sheets", in},
           std::vector<std::string>{"dump", "--json", in},
           std::vector<std::string>{"eval", "=SEQUENCE(10,10)"},
           std::vector<std::string>{"profiles"},
           std::vector<std::string>{"recalc", "--quiet", in, "-o", "-"},
       }) {
    const CliRun r = run_cli_spawn(args, options);
    EXPECT_EQ(r.exit_code, 74) << args.front() << " stderr=" << r.stderr_text;
    EXPECT_NE(r.stderr_text.find("formulon:"), std::string::npos) << args.front();
  }
}
