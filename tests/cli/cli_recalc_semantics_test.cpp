// `formulon_cli` end-to-end tests grouped by command surface.
#include "cli/cli.h"
#include "cli_test_support.h"

using namespace formulon::cli_test;

TEST(FormulonCli, RecalcRoundTripsFormulae) {
  // Build a fixture in /tmp.
  std::string in = temp_path("in.xlsx");
  std::string out_path = temp_path("out.xlsx");
  PathGuard g_in(in);
  PathGuard g_out(out_path);
  ASSERT_TRUE(write_fixture_workbook(in));

  CliRun r = run_cli({"recalc", in, "-o", out_path, "--quiet"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;

  // Load the output back through the C ABI and assert B1 still
  // recalculates correctly.
  std::ifstream f(out_path, std::ios::binary);
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

TEST(FormulonCli, RecalcWithoutThreadsKeepsSerialStatusShape) {
  const std::string input = temp_path("default_recalc_input.xlsx");
  const std::string output = temp_path("default_recalc_output.xlsx");
  PathGuard input_guard(input);
  PathGuard output_guard(output);
  ASSERT_TRUE(write_fixture_workbook(input));

  const CliRun result = run_cli({"recalc", input, "-o", output});
  ASSERT_EQ(result.exit_code, 0) << result.stderr_text;
  EXPECT_EQ(result.stderr_text.find("threads="), std::string::npos);
  EXPECT_NE(result.stderr_text.find("formulon: recalc: ok, wrote "), std::string::npos);
  EXPECT_NE(result.stderr_text.find(" bytes to '" + output + "'"), std::string::npos);
}

TEST(FormulonCli, RecalcThreadsUsesParallelSchedulerAndSavesWideDag) {
  const std::string input = temp_path("parallel_recalc_input.xlsx");
  const std::string output = temp_path("parallel_recalc_output.xlsx");
  PathGuard input_guard(input);
  PathGuard output_guard(output);
  ASSERT_TRUE(write_wide_recalc_fixture(input));

  const CliRun result = run_cli({"recalc", "--threads", "4", input, "-o", output});
  ASSERT_EQ(result.exit_code, 0) << result.stderr_text;
  EXPECT_TRUE(result.stdout_text.empty());
  EXPECT_NE(result.stderr_text.find("threads=4"), std::string::npos);
  EXPECT_NE(result.stderr_text.find("cells_evaluated=48"), std::string::npos);
  EXPECT_NE(result.stderr_text.find("sccs_processed="), std::string::npos);
  EXPECT_NE(result.stderr_text.find("parallel_steps="), std::string::npos);
  EXPECT_NE(result.stderr_text.find("worker_threads_started="), std::string::npos);
  EXPECT_NE(result.stderr_text.find("worker_threads_used="), std::string::npos);

  std::ifstream in(output, std::ios::binary);
  ASSERT_TRUE(in);
  const std::vector<std::uint8_t> bytes((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
  ASSERT_FALSE(bytes.empty());
  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &wb), 0);
  fm_value_t value{};
  ASSERT_EQ(fm_workbook_get_value(wb, 0U, 47U, 1U, &value), 0);
  EXPECT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 97.0);
  fm_workbook_destroy(wb);
}

TEST(FormulonCli, RecalcThreadsOneStartsNoWorkers) {
  const std::string input = temp_path("serial_threads_input.xlsx");
  const std::string output = temp_path("serial_threads_output.xlsx");
  PathGuard input_guard(input);
  PathGuard output_guard(output);
  ASSERT_TRUE(write_fixture_workbook(input));

  const CliRun result = run_cli({"recalc", "--threads", "1", input, "-o", output});
  ASSERT_EQ(result.exit_code, 0) << result.stderr_text;
  EXPECT_NE(result.stderr_text.find("threads=1"), std::string::npos);
  EXPECT_NE(result.stderr_text.find("worker_threads_started=0"), std::string::npos);
  EXPECT_NE(result.stderr_text.find("worker_threads_used=0"), std::string::npos);
}

TEST(FormulonCli, RecalcThreadsRejectsInvalidValuesBeforeWriting) {
  const std::string input = temp_path("invalid_threads_input.xlsx");
  const std::string output = temp_path("invalid_threads_output.xlsx");
  PathGuard input_guard(input);
  PathGuard output_guard(output);
  ASSERT_TRUE(write_fixture_workbook(input));
  std::remove(output.c_str());

  for (const char* value : {"9", "-1", "abc", "4294967296", "--"}) {
    const CliRun result = run_cli({"recalc", "--threads", value, input, "-o", output});
    EXPECT_EQ(result.exit_code, 64) << value << " stderr=" << result.stderr_text;
    EXPECT_NE(result.stderr_text.find("--threads"), std::string::npos);
    std::ifstream absent(output, std::ios::binary);
    EXPECT_FALSE(absent.good()) << "invalid value wrote output: " << value;
  }
}

TEST(FormulonCli, RecalcLossWarningsAreNonfatalAndNotSuppressedByQuiet) {
  const std::string input = temp_path("lossy_input.xlsx");
  const std::string output = temp_path("lossy_output.xlsb");
  PathGuard input_guard(input);
  PathGuard output_guard(output);
  ASSERT_TRUE(write_lossy_xlsx_fixture(input));

  CliRun r = run_cli({"recalc", "--quiet", input, "-o", output});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_TRUE(r.stdout_text.empty());
  EXPECT_NE(r.stderr_text.find("warning: XLSB write diagnostics"), std::string::npos);
  EXPECT_NE(r.stderr_text.find("downgraded_formula_count=1"), std::string::npos);
  // Only the autoFilter is deferred; the data validation is written to .xlsb.
  EXPECT_NE(r.stderr_text.find("deferred_feature_count=1"), std::string::npos);
}

TEST(FormulonCli, RecalcReportsOoxmlReadDiagnosticsSeparatelyFromXlsbOnes) {
  const std::string input = temp_path("lossy_ooxml_input.xlsx");
  const std::string output = temp_path("lossy_ooxml_output.xlsx");
  PathGuard input_guard(input);
  PathGuard output_guard(output);
  ASSERT_TRUE(write_lossy_ooxml_fixture(input));

  CliRun r = run_cli({"recalc", "--quiet", input, "-o", output});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_NE(r.stderr_text.find("warning: OOXML read diagnostics"), std::string::npos) << r.stderr_text;
  // One unparseable merge ref plus one conditional-formatting block with no
  // `sqref`; the unrecognised workbook content type is reported separately.
  EXPECT_NE(r.stderr_text.find("skipped_feature_count=2"), std::string::npos) << r.stderr_text;
  EXPECT_NE(r.stderr_text.find("unknown_content_type_count=1"), std::string::npos) << r.stderr_text;
  // The XLSB line must not appear: an `.xlsx` load produces none of its
  // counters, and a zero counter is never printed.
  EXPECT_EQ(r.stderr_text.find("XLSB read diagnostics"), std::string::npos) << r.stderr_text;
}

TEST(FormulonCli, RecalcDroppedPartWarningIsNonfatalAndQuietStillReportsIt) {
  const std::string input = temp_path("dropped_input.xlsb");
  const std::string output = temp_path("dropped_output.xlsx");
  PathGuard input_guard(input);
  PathGuard output_guard(output);
  ASSERT_TRUE(write_dropped_xlsb_fixture(input));

  CliRun r = run_cli({"recalc", "--quiet", input, "-o", output});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_NE(r.stderr_text.find("undecoded_part_count=1"), std::string::npos);
  EXPECT_NE(r.stderr_text.find("warning: XLSB read diagnostics"), std::string::npos);
}

TEST(FormulonCli, RecalcIterativePreservesExistingIterationSettings) {
  std::string in = temp_path("iterative_in.xlsx");
  std::string out_path = temp_path("iterative_out.xlsx");
  PathGuard g_in(in);
  PathGuard g_out(out_path);
  ASSERT_TRUE(write_fixture_workbook(in, {}, /*iterative=*/true, /*max_iterations=*/500, /*max_change=*/0.01));

  CliRun r = run_cli({"recalc", "--iterative", in, "-o", out_path, "--quiet"});
  ASSERT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;

  std::ifstream f(out_path, std::ios::binary);
  ASSERT_TRUE(f);
  f.seekg(0, std::ios::end);
  const std::streamsize size = f.tellg();
  ASSERT_GT(size, 0);
  f.seekg(0, std::ios::beg);
  std::vector<std::uint8_t> bytes(static_cast<std::size_t>(size));
  ASSERT_TRUE(f.read(reinterpret_cast<char*>(bytes.data()), size));

  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &wb), 0);
  int32_t enabled = 0;
  uint32_t max_iterations = 0;
  double max_change = 0.0;
  ASSERT_EQ(fm_workbook_get_iterative(wb, &enabled, &max_iterations, &max_change), 0);
  EXPECT_EQ(enabled, 1);
  EXPECT_EQ(max_iterations, 500U);
  EXPECT_DOUBLE_EQ(max_change, 0.01);
  fm_workbook_destroy(wb);
}

TEST(FormulonCli, RecalcXlsbOutputExtensionWritesXlsbContainer) {
  // `-o out.xlsb` must select the MS-XLSB writer, not silently emit an
  // OOXML package under an `.xlsb` name.
  std::string in = temp_path("in_for_xlsb.xlsx");
  std::string out_path = temp_path("out.xlsb");
  PathGuard g_in(in);
  PathGuard g_out(out_path);
  ASSERT_TRUE(write_fixture_workbook(in));

  CliRun r = run_cli({"recalc", in, "-o", out_path, "--quiet"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;

  std::ifstream f(out_path, std::ios::binary);
  ASSERT_TRUE(f);
  f.seekg(0, std::ios::end);
  const std::streamsize size = f.tellg();
  ASSERT_GT(size, 0);
  f.seekg(0, std::ios::beg);
  std::vector<std::uint8_t> bytes(static_cast<std::size_t>(size));
  f.read(reinterpret_cast<char*>(bytes.data()), size);
  ASSERT_TRUE(f);

  // The written package must declare `xl/workbook.bin` (xlsb), not
  // `xl/workbook.xml` (ooxml).
  formulon::io::ByteSpan span{bytes.data(), bytes.size()};
  EXPECT_EQ(formulon::io::detect_workbook_format(span), formulon::WorkbookFormat::Xlsb);

  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &wb), 0);
  ASSERT_EQ(fm_workbook_recalc(wb), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb, 0, 0, 1, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 8.0);
  fm_workbook_destroy(wb);
}
