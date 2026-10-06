// `formulon_cli` end-to-end tests grouped by command surface.
#include <cstring>

#include "cli/cli.h"
#include "cli_test_support.h"

using namespace formulon::cli_test;

TEST(FormulonCli, RecalcInPlaceOverwritesSamePathAndLeavesNoTemp) {
  // Input and output are the same path: the atomic write must serialize the
  // recalculated workbook fully before replacing the original, and must not
  // leave its sibling temp file behind on success.
  std::string path = temp_path("inplace.xlsx");
  std::string tmp_sidecar = path + ".formulon-tmp";
  PathGuard g_path(path);
  PathGuard g_tmp(tmp_sidecar);
  ASSERT_TRUE(write_fixture_workbook(path));

  CliRun r = run_cli({"recalc", path, "-o", path, "--quiet"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;

  // The temp sidecar must be gone (renamed into place).
  {
    std::ifstream leftover(tmp_sidecar, std::ios::binary);
    EXPECT_FALSE(leftover.good()) << "temp sidecar left behind after in-place recalc";
  }

  // The overwritten file still loads and recalculates correctly.
  std::ifstream f(path, std::ios::binary);
  ASSERT_TRUE(f);
  f.seekg(0, std::ios::end);
  const std::streamsize size = f.tellg();
  ASSERT_GT(size, 0);
  f.seekg(0, std::ios::beg);
  std::vector<std::uint8_t> bytes(static_cast<std::size_t>(size));
  f.read(reinterpret_cast<char*>(bytes.data()), size);
  ASSERT_TRUE(f);

  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &wb), 0);
  ASSERT_EQ(fm_workbook_recalc(wb), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb, 0, 0, 1, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 8.0);
  fm_workbook_destroy(wb);
}

TEST(FormulonCli, RecalcDoesNotReusePredictableLegacyTempPath) {
  const std::string input = temp_path("safe_temp_in.xlsx");
  const std::string output = temp_path("safe_temp_out.xlsx");
  const std::string legacy_temp = output + ".formulon-tmp";
  PathGuard g_input(input);
  PathGuard g_output(output);
  PathGuard g_legacy_temp(legacy_temp);
  ASSERT_TRUE(write_fixture_workbook(input));
  {
    std::ofstream sentinel(legacy_temp, std::ios::binary);
    ASSERT_TRUE(sentinel);
    sentinel << "unrelated file";
  }

  CliRun r = run_cli({"recalc", input, "-o", output, "--quiet"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  std::ifstream sentinel(legacy_temp, std::ios::binary);
  ASSERT_TRUE(sentinel);
  std::string contents;
  std::getline(sentinel, contents);
  EXPECT_EQ(contents, "unrelated file");
}

TEST(FormulonCli, RecalcThroughSymlinkUpdatesTargetAndKeepsLink) {
  // Saving through a symlink must update the file the link names. A
  // rename onto the link path would replace the link with a regular
  // file and leave the real workbook holding stale values -- the same
  // silent-loss shape the atomic write exists to prevent.
  const std::string input = temp_path("symlink_in.xlsx");
  const std::string target = temp_path("symlink_target.xlsx");
  const std::string link = temp_path("symlink_link.xlsx");
  PathGuard g_input(input);
  PathGuard g_target(target);
  PathGuard g_link(link);
  ASSERT_TRUE(write_fixture_workbook(input));
  ASSERT_TRUE(write_fixture_workbook(target));
  std::remove(link.c_str());
  ASSERT_EQ(::symlink(target.c_str(), link.c_str()), 0);

  CliRun r = run_cli({"recalc", input, "-o", link, "--quiet"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;

  // The link survives as a link...
  struct stat link_stat {};
  ASSERT_EQ(::lstat(link.c_str(), &link_stat), 0);
  EXPECT_TRUE(S_ISLNK(link_stat.st_mode)) << "symlink was replaced by a regular file";

  // ...and the file it names is the one that got rewritten.
  std::ifstream written(target, std::ios::binary);
  ASSERT_TRUE(written);
  written.seekg(0, std::ios::end);
  EXPECT_GT(written.tellg(), 0);

  // No temp sidecar is left next to either path.
  for (const std::string& sidecar : {target + ".formulon-tmp", link + ".formulon-tmp"}) {
    std::ifstream leftover(sidecar, std::ios::binary);
    EXPECT_FALSE(leftover.good()) << "temp sidecar left behind: " << sidecar;
  }
}

TEST(FormulonCli, RecalcDanglingSymlinkIsReplacedInPlace) {
  // A link with no target has nothing to preserve, so the write lands
  // on the link path itself rather than failing.
  const std::string input = temp_path("dangling_in.xlsx");
  const std::string link = temp_path("dangling_link.xlsx");
  PathGuard g_input(input);
  PathGuard g_link(link);
  ASSERT_TRUE(write_fixture_workbook(input));
  std::remove(link.c_str());
  ASSERT_EQ(::symlink(temp_path("dangling_absent.xlsx").c_str(), link.c_str()), 0);

  CliRun r = run_cli({"recalc", input, "-o", link, "--quiet"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  std::ifstream written(link, std::ios::binary);
  ASSERT_TRUE(written);
  written.seekg(0, std::ios::end);
  EXPECT_GT(written.tellg(), 0);
}

TEST(FormulonCli, RecalcLeavesOutputIntactWhenTheWriteCannotStart) {
  // Stand-in for a disk-full failure: an unwritable directory makes the
  // temp file impossible to create. The pre-existing output must be left
  // exactly as it was, with no partial content and no sidecar.
  const std::string dir = temp_path("readonly_dir");
  const std::string input = temp_path("readonly_in.xlsx");
  const std::string output = dir + "/out.xlsx";
  PathGuard g_input(input);
  ASSERT_TRUE(write_fixture_workbook(input));
  ASSERT_EQ(::mkdir(dir.c_str(), 0700), 0) << dir << ": " << std::strerror(errno);
  DirGuard g_dir(dir);
  ASSERT_TRUE(write_fixture_workbook(output));

  std::vector<std::uint8_t> before;
  {
    std::ifstream original(output, std::ios::binary);
    ASSERT_TRUE(original);
    before.assign(std::istreambuf_iterator<char>(original), std::istreambuf_iterator<char>());
  }
  ASSERT_FALSE(before.empty());
  ASSERT_EQ(::chmod(dir.c_str(), 0500), 0);

  CliRun r = run_cli({"recalc", input, "-o", output, "--quiet"});
  EXPECT_NE(r.exit_code, 0) << "write into an unwritable directory should fail";

  // Restore write permission so the fixture can be inspected; `g_dir`
  // removes it either way.
  ASSERT_EQ(::chmod(dir.c_str(), 0700), 0);
  std::vector<std::uint8_t> after;
  {
    std::ifstream survivor(output, std::ios::binary);
    ASSERT_TRUE(survivor) << "existing output was destroyed by a failed write";
    after.assign(std::istreambuf_iterator<char>(survivor), std::istreambuf_iterator<char>());
  }
  EXPECT_EQ(before, after) << "existing output was modified by a failed write";

  // "Unchanged" covers the whole directory, not just the output's bytes:
  // a sidecar the failed write forgot to clean up would still be sitting
  // next to it.
  EXPECT_EQ(list_directory(dir), std::vector<std::string>{"out.xlsx"});
}

TEST(FormulonCli, RecalcPreservesExistingOutputPermissions) {
  std::string in = temp_path("mode_in.xlsx");
  std::string out_path = temp_path("mode_out.xlsx");
  PathGuard g_in(in);
  PathGuard g_out(out_path);
  ASSERT_TRUE(write_fixture_workbook(in));
  ASSERT_TRUE(write_fixture_workbook(out_path));
  ASSERT_EQ(::chmod(out_path.c_str(), 0600), 0);

  CliRun r = run_cli({"recalc", in, "-o", out_path, "--quiet"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;

  struct stat saved_stat {};
  ASSERT_EQ(::stat(out_path.c_str(), &saved_stat), 0);
  EXPECT_EQ(saved_stat.st_mode & 0777U, 0600U);
}

TEST(FormulonCli, RecalcCreatesNewOutputWithUmaskDefaultPermissions) {
  // A fresh output path must come out with the mode an ordinary file
  // creation would produce, not the private mode the atomic-write
  // temporary is created with. The reference file below is created the
  // ordinary way, so the expectation follows whatever umask is in effect.
  std::string in = temp_path("newmode_in.xlsx");
  std::string out_path = temp_path("newmode_out.xlsx");
  std::string reference = temp_path("newmode_reference");
  PathGuard g_in(in);
  PathGuard g_out(out_path);
  PathGuard g_reference(reference);
  ASSERT_TRUE(write_fixture_workbook(in));
  std::remove(out_path.c_str());
  std::remove(reference.c_str());

  const int reference_fd = ::open(reference.c_str(), O_CREAT | O_WRONLY | O_EXCL, 0666);
  ASSERT_GE(reference_fd, 0);
  ASSERT_EQ(::close(reference_fd), 0);
  struct stat reference_stat {};
  ASSERT_EQ(::stat(reference.c_str(), &reference_stat), 0);

  CliRun r = run_cli({"recalc", in, "-o", out_path, "--quiet"});
  ASSERT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;

  struct stat created_stat {};
  ASSERT_EQ(::stat(out_path.c_str(), &created_stat), 0);
  EXPECT_EQ(created_stat.st_mode & 0777U, reference_stat.st_mode & 0777U);

  // Re-running against the now-existing path keeps that same mode, so the
  // permissions do not depend on whether the output already existed.
  CliRun again = run_cli({"recalc", in, "-o", out_path, "--quiet"});
  ASSERT_EQ(again.exit_code, 0) << "stderr=" << again.stderr_text;
  struct stat rerun_stat {};
  ASSERT_EQ(::stat(out_path.c_str(), &rerun_stat), 0);
  EXPECT_EQ(rerun_stat.st_mode & 0777U, created_stat.st_mode & 0777U);
}

TEST(FormulonCli, RecalcMissingOutputExits64) {
  std::string in = temp_path("in_missing_o.xlsx");
  PathGuard g_in(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  CliRun r = run_cli({"recalc", in});
  EXPECT_EQ(r.exit_code, 64);
}

TEST(FormulonCli, RecalcEmptyOutputPathIsDiagnosedApartFromAMissingOne) {
  // `-o "$OUT"` with an unset variable supplies the flag and an empty
  // value. Reporting that as a missing flag sends the user looking at the
  // wrong line of their script, so the two states get their own wording.
  std::string in = temp_path("empty_output_flag.xlsx");
  PathGuard g_in(in);
  ASSERT_TRUE(write_fixture_workbook(in));

  CliRun empty = run_cli({"recalc", in, "-o", ""});
  EXPECT_EQ(empty.exit_code, 64);
  EXPECT_NE(empty.stderr_text.find("non-empty path"), std::string::npos) << empty.stderr_text;

  CliRun missing = run_cli({"recalc", in});
  EXPECT_EQ(missing.exit_code, 64);
  EXPECT_NE(missing.stderr_text.find("missing -o/--output"), std::string::npos) << missing.stderr_text;

  EXPECT_NE(empty.stderr_text, missing.stderr_text);
}

TEST(FormulonCli, RecalcMissingInputExits64) {
  CliRun r = run_cli({"recalc", "-o", temp_path("unused.xlsx")});
  EXPECT_EQ(r.exit_code, 64);
}

TEST(FormulonCli, RecalcNonexistentFileFailsCleanly) {
  CliRun r = run_cli({"recalc", temp_path("absent_input.xlsx"), "-o", temp_path("unused.xlsx")});
  EXPECT_EQ(r.exit_code, 1);
  EXPECT_EQ(r.stderr_text.find("warning:"), std::string::npos);
}

TEST(FormulonCli, RecalcOptionTerminatorAcceptsDashLeadingRelativePaths) {
  const std::string input = "-fm_cli_recalc_dash_input.xlsx";
  const std::string output = "-fm_cli_recalc_dash_output.xlsx";
  PathGuard input_guard(input);
  PathGuard output_guard(output);
  std::remove(output.c_str());
  ASSERT_TRUE(write_fixture_workbook(input));

  const CliRun result = run_cli({"recalc", "--threads", "1", "-o", output, "--", input});
  ASSERT_EQ(result.exit_code, 0) << result.stderr_text;

  std::ifstream saved(output, std::ios::binary);
  ASSERT_TRUE(saved);
  const std::vector<std::uint8_t> bytes((std::istreambuf_iterator<char>(saved)), std::istreambuf_iterator<char>());
  ASSERT_FALSE(bytes.empty());
  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &wb), 0);
  ASSERT_EQ(fm_workbook_recalc(wb), 0);
  fm_value_t value{};
  ASSERT_EQ(fm_workbook_get_value(wb, 0U, 0U, 1U, &value), 0);
  EXPECT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 8.0);
  fm_workbook_destroy(wb);
}

TEST(FormulonCli, RecalcDashLeadingInputWithoutTerminatorIsUsageError) {
  const std::string input = "-fm_cli_recalc_dash_without_terminator_input.xlsx";
  const std::string output = "-fm_cli_recalc_dash_without_terminator_output.xlsx";
  PathGuard input_guard(input);
  PathGuard output_guard(output);
  std::remove(output.c_str());
  ASSERT_TRUE(write_fixture_workbook(input));

  const CliRun result = run_cli({"recalc", input, "-o", output});
  EXPECT_EQ(result.exit_code, 64);
  EXPECT_TRUE(result.stdout_text.empty());
  EXPECT_FALSE(std::ifstream(output, std::ios::binary).good());
}

TEST(FormulonCli, RecalcAcceptsDoubleDashAsOutputPathValue) {
  const std::string input = "fm_cli_recalc_double_dash_input.xlsx";
  const std::string output = "--";
  PathGuard input_guard(input);
  PathGuard output_guard(output);
  std::remove(output.c_str());
  ASSERT_TRUE(write_fixture_workbook(input));

  const CliRun result = run_cli({"recalc", input, "-o", output});
  ASSERT_EQ(result.exit_code, 0) << result.stderr_text;

  std::ifstream saved(output, std::ios::binary);
  ASSERT_TRUE(saved);
  const std::vector<std::uint8_t> bytes((std::istreambuf_iterator<char>(saved)), std::istreambuf_iterator<char>());
  ASSERT_FALSE(bytes.empty());
  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &wb), 0);
  ASSERT_EQ(fm_workbook_recalc(wb), 0);
  fm_value_t value{};
  ASSERT_EQ(fm_workbook_get_value(wb, 0U, 0U, 1U, &value), 0);
  EXPECT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 8.0);
  fm_workbook_destroy(wb);
}

TEST(FormulonCli, OptionTerminatorRequiresExactlyOnePositional) {
  const std::string output = temp_path("terminator_recalc_output.xlsx");
  PathGuard output_guard(output);

  for (const std::vector<std::string>& args : {
           std::vector<std::string>{"eval", "--"},
           std::vector<std::string>{"dump", "--"},
           std::vector<std::string>{"paginate", "--"},
       }) {
    const CliRun result = run_cli(args);
    EXPECT_EQ(result.exit_code, 64) << "bare terminator command=" << args.front();
  }

  std::remove(output.c_str());
  CliRun recalc_bare = run_cli({"recalc", "-o", output, "--"});
  EXPECT_EQ(recalc_bare.exit_code, 64);
  EXPECT_FALSE(std::ifstream(output, std::ios::binary).good());

  for (const std::vector<std::string>& args : {
           std::vector<std::string>{"eval", "--", "-1+2", "--json"},
           std::vector<std::string>{"dump", "--", "-fm_cli_missing.xlsx", "--sheets"},
           std::vector<std::string>{"paginate", "--", "-fm_cli_missing.xlsx", "--sheet"},
       }) {
    const CliRun result = run_cli(args);
    EXPECT_EQ(result.exit_code, 64) << "post-terminator extra command=" << args.front();
  }

  std::remove(output.c_str());
  CliRun recalc_extra = run_cli({"recalc", "-o", output, "--", "-fm_cli_missing.xlsx", "--quiet"});
  EXPECT_EQ(recalc_extra.exit_code, 64);
  EXPECT_FALSE(std::ifstream(output, std::ios::binary).good());
}

TEST(FormulonCliFileIo, AtomicWriteLeavesNoDescriptorOrTemporaryBehind) {
  const std::string dir = temp_path("atomic_write_dir");
  ASSERT_EQ(::mkdir(dir.c_str(), 0700), 0) << dir << ": " << std::strerror(errno);
  DirGuard guard(dir);
  const std::string output = dir + "/out.bin";
  const std::vector<std::uint8_t> payload{'f', 'o', 'r', 'm'};

  const int baseline = open_descriptor_count();
  for (int attempt = 0; attempt < 8; ++attempt) {
    ASSERT_EQ(formulon::cli::write_file_atomically(output, payload.data(), payload.size()), 0) << attempt;
  }
  EXPECT_EQ(open_descriptor_count(), baseline) << "successful writes leaked a descriptor";
  EXPECT_EQ(list_directory(dir), std::vector<std::string>{"out.bin"});

  // A destination that is a directory runs the whole sequence -- create,
  // write, sync, close -- and only then fails, at the rename. Repeating
  // it makes a per-call leak visible as growth rather than as a single
  // descriptor that could be noise, and the temporary it created has to
  // be gone from the parent directory afterwards.
  const std::string occupied = dir + "/sub";
  ASSERT_EQ(::mkdir(occupied.c_str(), 0700), 0) << occupied << ": " << std::strerror(errno);
  for (int attempt = 0; attempt < 8; ++attempt) {
    EXPECT_NE(formulon::cli::write_file_atomically(occupied, payload.data(), payload.size()), 0) << attempt;
  }
  EXPECT_EQ(open_descriptor_count(), baseline) << "failed writes leaked a descriptor";
  EXPECT_EQ(list_directory(dir), (std::vector<std::string>{"out.bin", "sub"}));

  // An unwritable directory covers the other end: the failure happens
  // before anything is created at all.
  ASSERT_EQ(::chmod(dir.c_str(), 0500), 0);
  for (int attempt = 0; attempt < 8; ++attempt) {
    EXPECT_NE(formulon::cli::write_file_atomically(output, payload.data(), payload.size()), 0) << attempt;
  }
  EXPECT_EQ(open_descriptor_count(), baseline) << "failed writes leaked a descriptor";
  ASSERT_EQ(::chmod(dir.c_str(), 0700), 0);
  EXPECT_EQ(list_directory(dir), (std::vector<std::string>{"out.bin", "sub"}));

  std::ifstream written(output, std::ios::binary);
  ASSERT_TRUE(written);
  const std::vector<std::uint8_t> contents((std::istreambuf_iterator<char>(written)), std::istreambuf_iterator<char>());
  EXPECT_EQ(contents, payload);
}
