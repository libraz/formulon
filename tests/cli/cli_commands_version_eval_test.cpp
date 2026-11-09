// `formulon_cli` end-to-end tests grouped by command surface.
#include "cli/cli.h"
#include "cli_test_support.h"
#include "support/json_cases.h"

using namespace formulon::cli_test;
using formulon::test_support::Json;

TEST(FormulonCli, VersionPrintsNonEmpty) {
  CliRun r = run_cli({"--version"});
  EXPECT_EQ(r.exit_code, 0);
  EXPECT_EQ(r.stdout_text, std::string("formulon ") + fm_get_version() + "\n");
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
    EXPECT_EQ(r.exit_code, 74) << flag;
    EXPECT_NE(r.stderr_text.find("failed to write output"), std::string::npos) << flag << " stderr=" << r.stderr_text;
  }
}

TEST(FormulonCli, SubcommandVersionAndHelpFailWhenTheirOutputCannotBeWritten) {
  for (const char* subcommand : {"eval", "recalc", "dump", "paginate"}) {
    for (const char* flag : {"--version", "--help"}) {
      CliRun r = run_cli({subcommand, flag}, /*merge_streams=*/false, /*close_stdout=*/true);
      EXPECT_EQ(r.exit_code, 74) << subcommand << ' ' << flag;
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
  EXPECT_EQ(formulon::cli::exit_code_for_status(formulon::cli::kExitDataErr), 65);
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

TEST(FormulonCli, EvalMalformedSyntaxExits65WithThePosition) {
  for (const char* formula : {"=1+", "=1+*2"}) {
    CliRun r = run_cli({"eval", formula});
    EXPECT_EQ(r.exit_code, 65) << formula << " stderr=" << r.stderr_text;
    EXPECT_EQ(r.stdout_text, "") << formula;
    EXPECT_EQ(r.stderr_text.rfind("formulon: eval: ", 0), 0U) << r.stderr_text;
    EXPECT_NE(r.stderr_text.find(" at 1:"), std::string::npos) << r.stderr_text;
  }
}

TEST(FormulonCli, EvalMalformedSyntaxMatchesTheAdHocCApiEntryPoint) {
  // The CLI reports what the C ABI records for the same call.
  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_create_empty(&wb), 0);
  std::size_t sheet = 0;
  ASSERT_EQ(fm_workbook_add_sheet(wb, "Sheet1", &sheet), 0);
  fm_value_array_t* result = nullptr;
  EXPECT_NE(fm_sheet_evaluate_formula_array(wb, sheet, 0, 0, "=1+", &result), 0);
  fm_parse_error_info_t position = FM_STRUCT_INIT(fm_parse_error_info_t);
  int32_t present = 0;
  ASSERT_EQ(fm_get_last_parse_error(&present, &position), 0);
  ASSERT_EQ(present, 1);
  const std::string expected = std::string("formulon: eval: ") + fm_get_last_error_message() + " at " +
                               std::to_string(position.line) + ":" + std::to_string(position.column) + "\n";
  fm_value_array_destroy(result);
  fm_workbook_destroy(wb);

  CliRun r = run_cli({"eval", "=1+"});
  EXPECT_EQ(r.exit_code, 65) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stderr_text, expected);
}

TEST(FormulonCli, EvalMalformedSyntaxMatchesTheSharedCaseTable) {
  Json cases;
  ASSERT_TRUE(formulon::test_support::LoadRepoJson(__FILE__, "tests/cli/", "tools/api/parse_error_cases.json", &cases));
  const Json& rows = cases.At("cases");
  std::size_t checked = 0;
  for (const Json& row : rows.items) {
    if (row.At("entry").text != "cli_eval") {
      continue;
    }
    std::string formula = row.At("formula").text;
    CliRun r = run_cli({"eval", "--", formula});
    EXPECT_EQ(r.exit_code, static_cast<int>(row.At("exit").number)) << row.At("id").text << " stderr=" << r.stderr_text;
    EXPECT_EQ(r.stdout_text, "") << row.At("id").text;
    EXPECT_EQ(r.stderr_text.rfind("formulon: eval: ", 0), 0U) << row.At("id").text << " stderr=" << r.stderr_text;
    const Json& suffix = row.At("stderr_suffix");
    if (suffix.kind == Json::Kind::kString) {
      const std::string line = r.stderr_text.substr(0, r.stderr_text.size() - (r.stderr_text.empty() ? 0 : 1));
      ASSERT_GE(line.size(), suffix.text.size()) << row.At("id").text;
      EXPECT_EQ(line.substr(line.size() - suffix.text.size()), suffix.text) << row.At("id").text;
    }
    ++checked;
  }
  EXPECT_GT(checked, 0U);
}

TEST(FormulonCli, EvalReadsTheFormulaFromStdinWithoutItsTrailingNewline) {
  const std::string path = temp_path("eval_stdin_formula.txt");
  PathGuard guard(path);
  ASSERT_TRUE(write_bytes(path, "=SUM(1,2,3)\n"));
  const CliRun r = run_cli({"eval", "-"}, false, false, path);
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "6\n");
}

TEST(FormulonCli, EvalStdinSyntaxErrorExits65) {
  const std::string path = temp_path("eval_stdin_bad_formula.txt");
  PathGuard guard(path);
  ASSERT_TRUE(write_bytes(path, "=1+\n"));
  const CliRun r = run_cli({"eval", "-"}, false, false, path);
  EXPECT_EQ(r.exit_code, 65) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "");
}

TEST(FormulonCli, EvalStdinIsRefusedOnATerminal) {
  SpawnOptions options;
  options.stdin_tty = true;
  const CliRun r = run_cli_spawn({"eval", "-"}, options);
  EXPECT_EQ(r.exit_code, 64) << "stderr=" << r.stderr_text;
}

TEST(FormulonCli, EvalFormulaAndWorkbookCannotBothReadStdin) {
  EXPECT_EQ(run_cli({"eval", "--in", "-", "-"}).exit_code, 64);
}

TEST(FormulonCli, EvalInEvaluatesAtA1OfTheFirstSheet) {
  const std::string in = temp_path("eval_in_workbook.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));  // A1 = 7, B1 = =A1+1
  const CliRun r = run_cli({"eval", "--in", in, "=A1*2"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "14\n");
}

TEST(FormulonCli, EvalInSelectsTheSheetByNameOrIndex) {
  const std::string in = temp_path("eval_in_sheets.xlsx");
  PathGuard guard(in);
  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_create_default(&wb), 0);
  std::size_t second = 0;
  ASSERT_EQ(fm_workbook_add_sheet(wb, "Second", &second), 0);
  ASSERT_EQ(fm_sheet_set_number(wb, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_number(wb, second, 0, 0, 5.0), 0);
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ASSERT_EQ(fm_workbook_save(wb, &bytes, &len), 0);
  ASSERT_TRUE(write_bytes(in, std::string_view(reinterpret_cast<const char*>(bytes), len)));
  fm_buffer_destroy(bytes);
  fm_workbook_destroy(wb);

  EXPECT_EQ(run_cli({"eval", "--in", in, "--sheet", "Second", "=A1"}).stdout_text, "5\n");
  EXPECT_EQ(run_cli({"eval", "--in", in, "--sheet", "1", "=A1"}).stdout_text, "5\n");
  EXPECT_EQ(run_cli({"eval", "--in", in, "--sheet", "0", "=A1"}).stdout_text, "1\n");
  EXPECT_EQ(run_cli({"eval", "--in", in, "--sheet", "Nope", "=A1"}).exit_code, 64);
  EXPECT_EQ(run_cli({"eval", "--sheet", "0", "=A1"}).exit_code, 64);
}

TEST(FormulonCli, EvalLocalCanonicalizesTheFormulaBeforeEvaluating) {
  // de-DE spells the argument separator `;`.
  const CliRun local = run_cli({"eval", "--profile", "win-365-de_DE", "--local", "=SUMME(1;2;3)"});
  EXPECT_EQ(local.exit_code, 0) << "stderr=" << local.stderr_text;
  EXPECT_EQ(local.stdout_text, "6\n");
  const CliRun invariant = run_cli({"eval", "--profile", "win-365-de_DE", "=SUM(1,2,3)"});
  EXPECT_EQ(invariant.exit_code, 0) << "stderr=" << invariant.stderr_text;
  EXPECT_EQ(invariant.stdout_text, "6\n");
}

TEST(FormulonCli, EvalLocalReportsSyntaxErrorsInTheTypedText) {
  const CliRun r = run_cli({"eval", "--profile", "mac-365-de_DE", "--local", "=SUMME(1;"});
  EXPECT_EQ(r.exit_code, 65) << "stderr=" << r.stderr_text;
  EXPECT_NE(r.stderr_text.find(" at 1:10"), std::string::npos) << r.stderr_text;
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
            "{\"format\":1,\"rows\":[[{\"kind\":\"number\",\"value\":1},{\"kind\":\"number\",\"value\":2}],"
            "[{\"kind\":\"number\",\"value\":3},{\"kind\":\"number\",\"value\":4}]]}\n");
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
  EXPECT_EQ(r.exit_code, 74);
  EXPECT_NE(r.stderr_text.find("failed to write output"), std::string::npos);
}

TEST(FormulonCli, EvalRepeatWriteFailureSkipsTimingOutput) {
  CliRun r = run_cli({"eval", "--repeat", "3", "=SUM(1,2,3)"}, /*merge_streams=*/false,
                     /*close_stdout=*/true);
  EXPECT_EQ(r.exit_code, 74);
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
