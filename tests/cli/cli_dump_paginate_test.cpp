// `formulon_cli` end-to-end tests grouped by command surface.
#include "cli/cli.h"
#include "cli_test_support.h"

using namespace formulon::cli_test;

TEST(FormulonCli, PaginatePrintsResolvedGeometry) {
  const std::string path = temp_path("paginate.xlsx");
  PathGuard guard(path);
  ASSERT_TRUE(write_fixture_workbook(path));

  CliRun r = run_cli({"paginate", path});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "sheet=0\npages=1\nprint_area=\nhorizontal_breaks=\nvertical_breaks=\n");
}

TEST(FormulonCli, PaginateWriteFailureReturnsOutputError) {
  const std::string path = temp_path("paginate_write_failure.xlsx");
  PathGuard guard(path);
  ASSERT_TRUE(write_fixture_workbook(path));

  CliRun r = run_cli({"paginate", path}, /*merge_streams=*/false, /*close_stdout=*/true);
  EXPECT_EQ(r.exit_code, 1);
  EXPECT_NE(r.stderr_text.find("failed to write output"), std::string::npos);
}

TEST(FormulonCli, DumpAndPaginateReportLoadLossesTheWayRecalcDoes) {
  // Whoever reads a `dump` snapshot or a `paginate` geometry treats it as
  // a faithful account of the input file, so both have to say what the
  // load could not carry across -- in the same words `recalc` uses, since
  // the counters come from one emitter.
  const std::string ooxml = temp_path("shared_read_diag_ooxml.xlsx");
  const std::string xlsb = temp_path("shared_read_diag.xlsb");
  const std::string recalc_output = temp_path("shared_read_diag_out.xlsx");
  PathGuard g_ooxml(ooxml);
  PathGuard g_xlsb(xlsb);
  PathGuard g_output(recalc_output);
  ASSERT_TRUE(write_lossy_ooxml_fixture(ooxml));
  ASSERT_TRUE(write_dropped_xlsb_fixture(xlsb));

  for (const std::string& fixture : {ooxml, xlsb}) {
    const CliRun recalc = run_cli({"recalc", "--quiet", fixture, "-o", recalc_output});
    ASSERT_EQ(recalc.exit_code, 0) << fixture << " stderr=" << recalc.stderr_text;
    const std::vector<std::string> expected = read_diagnostic_lines(recalc.stderr_text);
    ASSERT_FALSE(expected.empty()) << "fixture loads losslessly: " << fixture;

    const CliRun dump = run_cli({"dump", "--sheets", fixture});
    EXPECT_EQ(dump.exit_code, 0) << fixture << " stderr=" << dump.stderr_text;
    EXPECT_EQ(read_diagnostic_lines(dump.stderr_text), expected) << fixture;
    EXPECT_NE(dump.stderr_text.find("formulon: dump: warning:"), std::string::npos) << dump.stderr_text;

    const CliRun paginate = run_cli({"paginate", fixture});
    EXPECT_EQ(paginate.exit_code, 0) << fixture << " stderr=" << paginate.stderr_text;
    EXPECT_EQ(read_diagnostic_lines(paginate.stderr_text), expected) << fixture;
    EXPECT_NE(paginate.stderr_text.find("formulon: paginate: warning:"), std::string::npos) << paginate.stderr_text;
  }
}

TEST(FormulonCli, DumpReportsUndecodedCountsBeforeItsSnapshot) {
  // The warning has to precede the snapshot on the wire, not trail it:
  // a user piping stdout into `diff` still needs to see the caveat while
  // the comparison is being set up.
  const std::string fixture = temp_path("dump_read_diag.xlsb");
  PathGuard guard(fixture);
  ASSERT_TRUE(write_dropped_xlsb_fixture(fixture));

  const CliRun merged = run_cli({"dump", "--sheets", fixture}, /*merge_streams=*/true);
  ASSERT_EQ(merged.exit_code, 0) << merged.stdout_text;
  const std::size_t warning = merged.stdout_text.find("undecoded_part_count=1");
  const std::size_t snapshot = merged.stdout_text.find("Sheet1");
  ASSERT_NE(warning, std::string::npos) << merged.stdout_text;
  ASSERT_NE(snapshot, std::string::npos) << merged.stdout_text;
  EXPECT_LT(warning, snapshot);
}

TEST(FormulonCli, DumpSheetsListsInOrder) {
  std::string in = temp_path("dump_sheets.xlsx");
  PathGuard g_in(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  CliRun r = run_cli({"dump", "--sheets", in});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "Sheet1\n");
}

TEST(FormulonCli, DumpFormulasListsFormulaCells) {
  std::string in = temp_path("dump_formulas.xlsx");
  PathGuard g_in(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  CliRun r = run_cli({"dump", "--formulas", in});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  // B1 = `=A1+1` is the only formula; A1 is a literal.
  EXPECT_NE(r.stdout_text.find("Sheet1!B1"), std::string::npos) << "stdout=" << r.stdout_text;
  EXPECT_EQ(r.stdout_text.find("Sheet1!A1"), std::string::npos) << "literal A1 must not appear in --formulas dump";
}

TEST(FormulonCli, DumpMetadataPrintsSectionHeaders) {
  std::string in = temp_path("dump_metadata.xlsx");
  PathGuard g_in(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  CliRun r = run_cli({"dump", "--metadata", in});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_NE(r.stdout_text.find("[defined_names]"), std::string::npos);
  EXPECT_NE(r.stdout_text.find("[tables]"), std::string::npos);
  EXPECT_NE(r.stdout_text.find("[passthrough_parts]"), std::string::npos);
}

TEST(FormulonCli, DumpWriteFailureReturnsOutputError) {
  std::string in = temp_path("dump_write_failure.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  CliRun r = run_cli({"dump", "--sheets", in}, /*merge_streams=*/false, /*close_stdout=*/true);
  EXPECT_EQ(r.exit_code, 1);
  EXPECT_NE(r.stderr_text.find("failed to write output"), std::string::npos);
}

TEST(FormulonCli, DumpValuesEscapesEmbeddedNewlines) {
  std::string in = temp_path("dump_escaped.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in, "first\nsecond"));
  CliRun r = run_cli({"dump", "--values", in});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_NE(r.stdout_text.find("Sheet1!A2 first\\nsecond\n"), std::string::npos) << r.stdout_text;
  EXPECT_EQ(r.stdout_text.find("first\nsecond"), std::string::npos) << r.stdout_text;
}

TEST(FormulonCli, DumpMissingInputExits64) {
  CliRun r = run_cli({"dump", "--sheets"});
  EXPECT_EQ(r.exit_code, 64);
}

TEST(FormulonCli, DumpOptionTerminatorAcceptsDashLeadingRelativePath) {
  const std::string input = "-fm_cli_dump_dash_input.xlsx";
  PathGuard guard(input);
  ASSERT_TRUE(write_fixture_workbook(input));

  const CliRun without_terminator = run_cli({"dump", "--sheets", input});
  EXPECT_EQ(without_terminator.exit_code, 64);

  const CliRun with_terminator = run_cli({"dump", "--sheets", "--", input});
  EXPECT_EQ(with_terminator.exit_code, 0) << with_terminator.stderr_text;
  EXPECT_EQ(with_terminator.stdout_text, "Sheet1\n");
}

TEST(FormulonCli, DumpPostTerminatorHelpIsAnInputPath) {
  const std::string input = "-h";
  PathGuard guard(input);
  std::remove(input.c_str());

  const CliRun result = run_cli({"dump", "--", input});
  EXPECT_EQ(result.exit_code, 1);
  EXPECT_TRUE(result.stdout_text.empty());
  EXPECT_NE(result.stderr_text.find("cannot read"), std::string::npos);
}

TEST(FormulonCli, SecondLiteralTerminatorIsAnExtraPositional) {
  const std::string input = "-fm_cli_second_literal_terminator.xlsx";
  PathGuard guard(input);
  ASSERT_TRUE(write_fixture_workbook(input));

  const CliRun dump = run_cli({"dump", "--", input, "--"});
  EXPECT_EQ(dump.exit_code, 64);
  const CliRun paginate = run_cli({"paginate", "--", input, "--"});
  EXPECT_EQ(paginate.exit_code, 64);
}

TEST(FormulonCli, DumpMetadataDistinguishesDefinedNameScope) {
  std::string in = temp_path("dump_metadata_scoped.xlsx");
  PathGuard g_in(in);

  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_create(&wb), 0);
  ASSERT_EQ(fm_workbook_set_number(wb, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_workbook_set_defined_name(wb, "WorkbookConst", "=1"), 0);
  ASSERT_EQ(fm_workbook_set_defined_name_scoped(wb, "SheetLocal", "=Sheet1!$A$1", 0), 0);
  ASSERT_EQ(fm_workbook_recalc(wb), 0);
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ASSERT_EQ(fm_workbook_save(wb, &bytes, &len), 0);
  std::ofstream out(in, std::ios::binary | std::ios::trunc);
  ASSERT_TRUE(out);
  out.write(reinterpret_cast<const char*>(bytes), static_cast<std::streamsize>(len));
  out.close();
  fm_buffer_free(bytes);
  fm_workbook_destroy(wb);

  CliRun r = run_cli({"dump", "--metadata", in});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  // Workbook-scoped names print bare; sheet-scoped names are prefixed
  // with `SheetName!` so the two scopes don't collide in the dump.
  EXPECT_NE(r.stdout_text.find("WorkbookConst =1"), std::string::npos) << "stdout=" << r.stdout_text;
  EXPECT_NE(r.stdout_text.find("Sheet1!SheetLocal ="), std::string::npos) << "stdout=" << r.stdout_text;
}

TEST(FormulonCli, DumpMetadataRecoversFromOutOfRangeDefinedNameScope) {
  std::string in = temp_path("dump_metadata_out_of_range_scope.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_out_of_range_defined_names_fixture(in));

  CliRun r = run_cli({"dump", "--metadata", in});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_TRUE(r.stderr_text.empty());
  EXPECT_EQ(r.stdout_text,
            "[defined_names]\n"
            "Before =1\n"
            "#99!Bad =Sheet1!$A$1\n"
            "After =2\n"
            "[tables]\n"
            "[passthrough_parts]\n");
}

TEST(FormulonCli, PaginateOptionTerminatorAcceptsDashLeadingRelativePath) {
  const std::string input = "-fm_cli_paginate_dash_input.xlsx";
  PathGuard guard(input);
  ASSERT_TRUE(write_fixture_workbook(input));

  const CliRun without_terminator = run_cli({"paginate", "--sheet", "0", input});
  EXPECT_EQ(without_terminator.exit_code, 64);

  const CliRun with_terminator = run_cli({"paginate", "--sheet", "0", "--", input});
  EXPECT_EQ(with_terminator.exit_code, 0) << with_terminator.stderr_text;
  EXPECT_EQ(with_terminator.stdout_text, "sheet=0\npages=1\nprint_area=\nhorizontal_breaks=\nvertical_breaks=\n");
}

TEST(FormulonCli, PaginatePostTerminatorHelpIsAnInputPath) {
  const std::string input = "-h";
  PathGuard guard(input);
  std::remove(input.c_str());

  const CliRun result = run_cli({"paginate", "--", input});
  EXPECT_EQ(result.exit_code, 1);
  EXPECT_TRUE(result.stdout_text.empty());
  EXPECT_NE(result.stderr_text.find("cannot read"), std::string::npos);
}

TEST(FormulonCli, PaginateConsumesTerminatorAsSheetValue) {
  const std::string input = "-fm_cli_paginate_sheet_value_terminator.xlsx";
  PathGuard guard(input);
  ASSERT_TRUE(write_fixture_workbook(input));

  const CliRun result = run_cli({"paginate", "--sheet", "--", input});
  EXPECT_EQ(result.exit_code, 64);
  EXPECT_TRUE(result.stdout_text.empty());
}
