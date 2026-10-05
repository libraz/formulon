// `formulon_cli` end-to-end tests grouped by command surface.
#include "cli/cli.h"
#include "cli_test_support.h"

using namespace formulon::cli_test;

TEST(FormulonCli, VersionPrintsNonEmpty) {
  CliRun r = run_cli({"--version"});
  EXPECT_EQ(r.exit_code, 0);
  EXPECT_GT(r.stdout_text.size(), 0U);
}

TEST(FormulonCli, HelpExitsZero) {
  CliRun r = run_cli({"--help"});
  EXPECT_EQ(r.exit_code, 0);
  EXPECT_NE(r.stdout_text.find("Usage"), std::string::npos);
}

TEST(FormulonCli, VersionAndHelpFailWhenTheirOutputCannotBeWritten) {
  // A release script doing `ver=$(formulon --version)` must not be handed
  // an empty string with a success status, so the version and usage
  // banners go through the same flush check `dump` and `eval` use for
  // their results.
  for (const char* flag : {"--version", "--help"}) {
    CliRun r = run_cli({flag}, /*merge_streams=*/false, /*close_stdout=*/true);
    EXPECT_EQ(r.exit_code, 1) << flag;
    EXPECT_NE(r.stderr_text.find("failed to write output"), std::string::npos) << flag << " stderr=" << r.stderr_text;
  }
}

TEST(FormulonCli, SubcommandVersionAndHelpFailWhenTheirOutputCannotBeWritten) {
  for (const char* subcommand : {"eval", "recalc", "dump", "paginate"}) {
    for (const char* flag : {"--version", "--help"}) {
      CliRun r = run_cli({subcommand, flag}, /*merge_streams=*/false, /*close_stdout=*/true);
      EXPECT_EQ(r.exit_code, 1) << subcommand << ' ' << flag;
      EXPECT_NE(r.stderr_text.find("failed to write output"), std::string::npos)
          << subcommand << ' ' << flag << " stderr=" << r.stderr_text;
    }
  }
}

TEST(FormulonCli, VersionIsAcceptedBySubcommandsLikeHelp) {
  // The usage banner lists `--version` under the same "Common options"
  // heading as `-h`, so every subcommand must accept it and print the
  // same line the top-level flag prints.
  CliRun top = run_cli({"--version"});
  ASSERT_EQ(top.exit_code, 0);
  ASSERT_FALSE(top.stdout_text.empty());

  for (const char* subcommand : {"eval", "recalc", "dump", "paginate"}) {
    CliRun r = run_cli({subcommand, "--version"});
    EXPECT_EQ(r.exit_code, 0) << subcommand << " stderr=" << r.stderr_text;
    EXPECT_EQ(r.stdout_text, top.stdout_text) << subcommand;
  }
}

TEST(FormulonCli, RecalcHelpDocumentsActualSuccessStatus) {
  CliRun r = run_cli({"recalc", "--help"});
  EXPECT_EQ(r.exit_code, 0);
  EXPECT_NE(r.stdout_text.find("formulon: recalc: ok, wrote M bytes to 'OUT'"), std::string::npos);
  EXPECT_NE(r.stdout_text.find("--threads N"), std::string::npos);
  EXPECT_NE(r.stdout_text.find("recalc remains serial"), std::string::npos);
}

TEST(FormulonCli, NoArgsExits64) {
  CliRun r = run_cli({});
  EXPECT_EQ(r.exit_code, 64);
}

TEST(FormulonCli, UnknownCommandExits64) {
  CliRun r = run_cli({"frobnicate"});
  EXPECT_EQ(r.exit_code, 64);
}

TEST(FormulonCli, ExitCodesNeverEncodeLargeInternalStatusValues) {
  EXPECT_EQ(formulon::cli::exit_code_for_status(0), 0);
  EXPECT_EQ(formulon::cli::exit_code_for_status(formulon::cli::kExitUsage), 64);
  EXPECT_EQ(formulon::cli::exit_code_for_status(128), 1);
  EXPECT_EQ(formulon::cli::exit_code_for_status(7000), 1);
  EXPECT_EQ(formulon::cli::exit_code_for_status(8000), 1);
}

TEST(FormulonCli, EvalSumLiterals) {
  CliRun r = run_cli({"eval", "=SUM(1,2,3)"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "6\n");
}

TEST(FormulonCli, EvalIfTrue) {
  CliRun r = run_cli({"eval", "=IF(TRUE,1,2)"});
  EXPECT_EQ(r.exit_code, 0);
  EXPECT_EQ(r.stdout_text, "1\n");
}

TEST(FormulonCli, EvalConcatTextNoQuoting) {
  CliRun r = run_cli({"eval", "=CONCAT(\"hello\",\" \",\"world\")"});
  EXPECT_EQ(r.exit_code, 0);
  EXPECT_EQ(r.stdout_text, "hello world\n");
}

TEST(FormulonCli, EvalCellLevelErrorIsStillExitZero) {
  // The ad-hoc evaluator reads A1 without first writing the expression to
  // A1, so it sees the empty cell rather than making a self-reference cycle.
  CliRun r = run_cli({"eval", "=A1+1"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "1\n");
}

TEST(FormulonCli, EvalMalformedSyntaxYieldsTheSameNameErrorEveryBindingReturns) {
  // The ad-hoc evaluator resolves a parser recovery placeholder to a
  // cell-level `#NAME?`. That is what the WASM, Node and Python bindings
  // return over the same C ABI entry point, so the CLI returns it too:
  // a formula prototyped here has to behave identically when it is moved
  // into a script driving one of the other surfaces.
  for (const char* formula : {"=1+", "=1+*2"}) {
    CliRun r = run_cli({"eval", formula});
    EXPECT_EQ(r.exit_code, 0) << formula << " stderr=" << r.stderr_text;
    EXPECT_EQ(r.stdout_text, "#NAME?\n") << formula;
  }
}

TEST(FormulonCli, EvalMalformedSyntaxMatchesTheAdHocCApiEntryPoint) {
  // Same formula, same C ABI call the CLI makes, in-process: a future
  // pre-parse gate on either side would break this comparison.
  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_create_empty(&wb), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb, "Sheet1"), 0);
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  ASSERT_EQ(fm_workbook_evaluate_formula_array(wb, 0, 0, 0, "=1+", &rows, &cols), 0);
  EXPECT_EQ(rows, 1U);
  EXPECT_EQ(cols, 1U);
  fm_value_t value{};
  ASSERT_EQ(fm_workbook_evaluate_formula_array_cell(wb, 0, &value), 0);
  EXPECT_EQ(value.kind, FM_VAL_ERROR);
  fm_workbook_destroy(wb);

  CliRun r = run_cli({"eval", "=1+"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "#NAME?\n");
}

TEST(FormulonCli, EvalUnknownFunctionRemainsCellLevelNameError) {
  CliRun r = run_cli({"eval", "=NOTAREALFUNC(1,2)"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "#NAME?\n");
}

TEST(FormulonCli, EvalDynamicArrayPrintsWholeGrid) {
  CliRun r = run_cli({"eval", "=SEQUENCE(2,3)"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "1\t2\t3\n4\t5\t6\n");
}

TEST(FormulonCli, EvalPlainGridKeepsItsShapeWhenCellsCarryTabsAndNewlines) {
  // Plain output is a TAB-separated grid, so a cell holding a TAB or a
  // newline would add fields and rows that a `cut`/`awk`/TSV consumer
  // reads as real ones. Exactly `rows` lines of `cols` fields must
  // survive any payload.
  CliRun r = run_cli({"eval", "=HSTACK(CHAR(9),CHAR(10),\"tail\")"});
  ASSERT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "\\t\t\\n\ttail\n");

  const auto field_count = [](const std::string& line) {
    return 1 + static_cast<int>(std::count(line.begin(), line.end(), '\t'));
  };
  std::vector<std::string> lines;
  for (std::size_t start = 0; start < r.stdout_text.size();) {
    const std::size_t end = r.stdout_text.find('\n', start);
    ASSERT_NE(end, std::string::npos) << "grid output must be newline-terminated";
    lines.push_back(r.stdout_text.substr(start, end - start));
    start = end + 1;
  }
  ASSERT_EQ(lines.size(), 1U);
  EXPECT_EQ(field_count(lines[0]), 3);
}

TEST(FormulonCli, EvalPlainGridEscapesMultiRowPayloadsToo) {
  // The two-row case pins the row count as well: an unescaped newline in
  // the first row would make the result look like three rows.
  CliRun r = run_cli({"eval", "=VSTACK(HSTACK(CHAR(10),1),HSTACK(2,3))"});
  ASSERT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "\\n\t1\n2\t3\n");
  EXPECT_EQ(std::count(r.stdout_text.begin(), r.stdout_text.end(), '\n'), 2);
}

TEST(FormulonCli, EvalDynamicArrayJsonPrintsNestedArrays) {
  CliRun r = run_cli({"eval", "--json", "=SEQUENCE(2,2)"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text,
            "[[{\"kind\":\"number\",\"value\":1},{\"kind\":\"number\",\"value\":2}],"
            "[{\"kind\":\"number\",\"value\":3},{\"kind\":\"number\",\"value\":4}]]\n");
}

TEST(FormulonCli, EvalDivByZeroSurfacesAsExcelErrorString) {
  CliRun r = run_cli({"eval", "=1/0"});
  EXPECT_EQ(r.exit_code, 0);
  EXPECT_EQ(r.stdout_text, "#DIV/0!\n");
}

TEST(FormulonCli, EvalJsonOutputShape) {
  CliRun r = run_cli({"eval", "--json", "=SUM(1,2,3)"});
  EXPECT_EQ(r.exit_code, 0);
  // Compact form: `{"kind":"number","value":6}`.
  EXPECT_NE(r.stdout_text.find("\"kind\":\"number\""), std::string::npos) << "stdout=" << r.stdout_text;
  EXPECT_NE(r.stdout_text.find("\"value\":6"), std::string::npos) << "stdout=" << r.stdout_text;
}

TEST(FormulonCli, EvalMissingFormulaExits64) {
  CliRun r = run_cli({"eval"});
  EXPECT_EQ(r.exit_code, 64);
}

TEST(FormulonCli, EvalUnknownFlagExits64) {
  CliRun r = run_cli({"eval", "--bogus"});
  EXPECT_EQ(r.exit_code, 64);
}

TEST(FormulonCli, EvalNegativeFormulaWithoutTerminatorExits64) {
  CliRun r = run_cli({"eval", "-1+2"});
  EXPECT_EQ(r.exit_code, 64);
  EXPECT_TRUE(r.stdout_text.empty());
}

TEST(FormulonCli, EvalOptionTerminatorAcceptsNegativeFormula) {
  CliRun r = run_cli({"eval", "--", "-1+2"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "1\n");
}

TEST(FormulonCli, EvalRepeatRunsMultipleTimes) {
  CliRun r = run_cli({"eval", "--repeat", "3", "=SUM(1,2,3)"});
  EXPECT_EQ(r.exit_code, 0);
  EXPECT_EQ(r.stdout_text, "6\n");
  // The timing line goes to stderr.
  EXPECT_NE(r.stderr_text.find("3 iterations"), std::string::npos);
}

TEST(FormulonCli, EvalRepeatReportsTimingForEveryAcceptedCount) {
  // `--repeat` is an explicit request to measure. Every count the flag
  // accepts gets the report, including the smallest one -- a user timing
  // a single evaluation must not be answered with silence and status 0.
  for (const char* count : {"1", "2", "5"}) {
    CliRun r = run_cli({"eval", "--repeat", count, "=1"});
    EXPECT_EQ(r.exit_code, 0) << count << " stderr=" << r.stderr_text;
    EXPECT_EQ(r.stdout_text, "1\n") << count;
    EXPECT_NE(r.stderr_text.find(std::string(count) + " iterations in"), std::string::npos)
        << count << " stderr=" << r.stderr_text;
  }
}

TEST(FormulonCli, EvalWithoutRepeatFlagReportsNoTiming) {
  // Only the default single evaluation is silent, and it stays silent
  // even though it runs the same one pass `--repeat 1` runs.
  CliRun r = run_cli({"eval", "=1"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "1\n");
  EXPECT_EQ(r.stderr_text.find("iterations"), std::string::npos) << "stderr=" << r.stderr_text;
}

TEST(FormulonCli, EvalWriteFailureReturnsOutputError) {
  CliRun r = run_cli({"eval", "=SUM(1,2,3)"}, /*merge_streams=*/false, /*close_stdout=*/true);
  EXPECT_EQ(r.exit_code, 1);
  EXPECT_NE(r.stderr_text.find("failed to write output"), std::string::npos);
}

TEST(FormulonCli, EvalRepeatWriteFailureSkipsTimingOutput) {
  CliRun r = run_cli({"eval", "--repeat", "3", "=SUM(1,2,3)"}, /*merge_streams=*/false,
                     /*close_stdout=*/true);
  EXPECT_EQ(r.exit_code, 1);
  EXPECT_NE(r.stderr_text.find("failed to write output"), std::string::npos);
  EXPECT_EQ(r.stderr_text.find("iterations"), std::string::npos);
}

TEST(FormulonCli, EvalRepeatOverflowExits64) {
  CliRun r = run_cli({"eval", "--repeat", "999999999999999999999999999999", "=1"});
  EXPECT_EQ(r.exit_code, 64);
  EXPECT_NE(r.stderr_text.find("positive integer"), std::string::npos);
}

TEST(FormulonCli, EvalRepeatHighCountStillCorrect) {
  CliRun r = run_cli({"eval", "--repeat", "100", "=SUM(1,2,3)"});
  EXPECT_EQ(r.exit_code, 0);
  EXPECT_EQ(r.stdout_text, "6\n");
  EXPECT_NE(r.stderr_text.find("100 iterations"), std::string::npos);
}
