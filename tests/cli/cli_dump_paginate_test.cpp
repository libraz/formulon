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
  EXPECT_EQ(r.stdout_text, "name=Sheet1\npageCount=1\nprintArea=\nrowBreaks=\ncolumnBreaks=\npages=A1:B1\n");
}

namespace {

// Sheet1 with print area B2:D5 and a manual row break before 0-based row 3, so it prints as two pages.
void write_print_area_fixture(const std::string& path) {
  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_create_default(&wb), 0);
  const fm_cell_rect_t area = {1, 1, 4, 3};
  ASSERT_EQ(fm_sheet_set_print_area(wb, 0, &area, 1), 0) << fm_get_last_error_message();
  fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
  brk.manual = FM_TRISTATE_TRUE;
  ASSERT_EQ(fm_sheet_set_row_break(wb, 0, 3, &brk), 0) << fm_get_last_error_message();
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ASSERT_EQ(fm_workbook_save(wb, FM_WORKBOOK_FORMAT_XLSX, &bytes, &len), 0);
  ASSERT_TRUE(write_bytes(path, std::string_view(reinterpret_cast<const char*>(bytes), len)));
  fm_buffer_destroy(bytes);
  fm_workbook_destroy(wb);
}

}  // namespace

TEST(FormulonCli, PaginateJsonCarriesCountAreaBreaksAndPages) {
  const std::string path = temp_path("paginate_json_geometry.xlsx");
  PathGuard guard(path);
  write_print_area_fixture(path);

  const CliRun r = run_cli({"paginate", "--json", path});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text,
            "{\"format\":1,\"sheets\":[{\"name\":\"Sheet1\",\"pageCount\":2,\"printArea\":[\"B2:D5\"],"
            "\"rowBreaks\":[3],\"columnBreaks\":[],\"pages\":[{\"range\":\"B2:D3\"},{\"range\":\"B4:D5\"}]}]}\n");
}

TEST(FormulonCli, PaginatePlainPrintsTheJsonKeys) {
  const std::string path = temp_path("paginate_plain_geometry.xlsx");
  PathGuard guard(path);
  write_print_area_fixture(path);

  const CliRun r = run_cli({"paginate", path});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text,
            "name=Sheet1\npageCount=2\nprintArea=B2:D5\nrowBreaks=3\ncolumnBreaks=\npages=B2:D3,B4:D5\n");
}

TEST(FormulonCli, PaginateHelpDescribesTheJsonSchema) {
  const CliRun r = run_cli({"paginate", "--help"});
  EXPECT_EQ(r.exit_code, 0);
  EXPECT_NE(r.stdout_text.find("{\"format\":1,\"sheets\":[...]}"), std::string::npos) << r.stdout_text;
  EXPECT_NE(r.stdout_text.find("\"format\" is the schema version"), std::string::npos) << r.stdout_text;
}

TEST(FormulonCli, DumpHelpDescribesTheJsonCells) {
  const CliRun r = run_cli({"dump", "--help"});
  EXPECT_EQ(r.exit_code, 0);
  EXPECT_NE(r.stdout_text.find("--view formulas reports cached values"), std::string::npos) << r.stdout_text;
  EXPECT_NE(r.stdout_text.find("\"format\" is the schema version"), std::string::npos) << r.stdout_text;
}

TEST(FormulonCli, PaginateWriteFailureReturnsOutputError) {
  const std::string path = temp_path("paginate_write_failure.xlsx");
  PathGuard guard(path);
  ASSERT_TRUE(write_fixture_workbook(path));

  CliRun r = run_cli({"paginate", path}, /*merge_streams=*/false, /*close_stdout=*/true);
  EXPECT_EQ(r.exit_code, 74);
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

    const CliRun dump = run_cli({"dump", "--view", "sheets", fixture});
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

  const CliRun merged = run_cli({"dump", "--view", "sheets", fixture}, /*merge_streams=*/true);
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
  CliRun r = run_cli({"dump", "--view", "sheets", in});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "Sheet1\n");
}

TEST(FormulonCli, DumpFormulasListsFormulaCells) {
  std::string in = temp_path("dump_formulas.xlsx");
  PathGuard g_in(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  CliRun r = run_cli({"dump", "--view", "formulas", in});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  // B1 = `=A1+1` is the only formula; A1 is a literal.
  EXPECT_NE(r.stdout_text.find("Sheet1!B1"), std::string::npos) << "stdout=" << r.stdout_text;
  EXPECT_EQ(r.stdout_text.find("Sheet1!A1"), std::string::npos) << "literal A1 must not appear in the formulas view";
}

TEST(FormulonCli, DumpMetadataKeysAreCamelCase) {
  const std::string in = temp_path("dump_metadata_keys.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));

  const CliRun json = run_cli({"dump", "--view", "metadata", "--json", in});
  EXPECT_EQ(json.exit_code, 0) << "stderr=" << json.stderr_text;
  EXPECT_NE(json.stdout_text.find("\"definedNames\":"), std::string::npos) << json.stdout_text;
  EXPECT_NE(json.stdout_text.find("\"passthroughParts\":"), std::string::npos) << json.stdout_text;
  EXPECT_EQ(json.stdout_text.find("defined_names"), std::string::npos) << json.stdout_text;
  EXPECT_EQ(json.stdout_text.find("passthrough_parts"), std::string::npos) << json.stdout_text;
}

TEST(FormulonCli, DumpMetadataPrintsSectionHeaders) {
  std::string in = temp_path("dump_metadata.xlsx");
  PathGuard g_in(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  CliRun r = run_cli({"dump", "--view", "metadata", in});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_NE(r.stdout_text.find("[definedNames]"), std::string::npos);
  EXPECT_NE(r.stdout_text.find("[tables]"), std::string::npos);
  EXPECT_NE(r.stdout_text.find("[passthroughParts]"), std::string::npos);
}

TEST(FormulonCli, DumpWriteFailureReturnsOutputError) {
  std::string in = temp_path("dump_write_failure.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  CliRun r = run_cli({"dump", "--view", "sheets", in}, /*merge_streams=*/false, /*close_stdout=*/true);
  EXPECT_EQ(r.exit_code, 74);
  EXPECT_NE(r.stderr_text.find("failed to write output"), std::string::npos);
}

TEST(FormulonCli, DumpValuesEscapesEmbeddedNewlines) {
  std::string in = temp_path("dump_escaped.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in, "first\nsecond"));
  CliRun r = run_cli({"dump", "--view", "values", in});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_NE(r.stdout_text.find("Sheet1!A2 first\\nsecond\n"), std::string::npos) << r.stdout_text;
  EXPECT_EQ(r.stdout_text.find("first\nsecond"), std::string::npos) << r.stdout_text;
}

TEST(FormulonCli, DumpMissingInputExits64) {
  CliRun r = run_cli({"dump", "--view", "sheets"});
  EXPECT_EQ(r.exit_code, 64);
}

TEST(FormulonCli, DumpOptionTerminatorAcceptsDashLeadingRelativePath) {
  const std::string input = "-fm_cli_dump_dash_input.xlsx";
  PathGuard guard(input);
  ASSERT_TRUE(write_fixture_workbook(input));

  const CliRun without_terminator = run_cli({"dump", "--view", "sheets", input});
  EXPECT_EQ(without_terminator.exit_code, 64);

  const CliRun with_terminator = run_cli({"dump", "--view", "sheets", "--", input});
  EXPECT_EQ(with_terminator.exit_code, 0) << with_terminator.stderr_text;
  EXPECT_EQ(with_terminator.stdout_text, "Sheet1\n");
}

TEST(FormulonCli, DumpPostTerminatorHelpIsAnInputPath) {
  const std::string input = "-h";
  PathGuard guard(input);
  std::remove(input.c_str());

  const CliRun result = run_cli({"dump", "--", input});
  EXPECT_EQ(result.exit_code, 66);
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
  ASSERT_EQ(fm_workbook_create_default(&wb), 0);
  ASSERT_EQ(fm_sheet_set_number(wb, 0, 0, 0, 1.0), 0);
  fm_defined_name_t workbook_const = FM_STRUCT_INIT(fm_defined_name_t);
  workbook_const.name = "WorkbookConst";
  workbook_const.formula = "=1";
  workbook_const.local_sheet_id = FM_DEFINED_NAME_SCOPE_WORKBOOK;
  fm_defined_name_t sheet_local = FM_STRUCT_INIT(fm_defined_name_t);
  sheet_local.name = "SheetLocal";
  sheet_local.formula = "=Sheet1!$A$1";
  sheet_local.local_sheet_id = 0;
  ASSERT_EQ(fm_workbook_set_defined_name(wb, &workbook_const), 0);
  ASSERT_EQ(fm_workbook_set_defined_name(wb, &sheet_local), 0);
  ASSERT_EQ(fm_workbook_recalc(wb), 0);
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ASSERT_EQ(fm_workbook_save(wb, FM_WORKBOOK_FORMAT_XLSX, &bytes, &len), 0);
  std::ofstream out(in, std::ios::binary | std::ios::trunc);
  ASSERT_TRUE(out);
  out.write(reinterpret_cast<const char*>(bytes), static_cast<std::streamsize>(len));
  out.close();
  fm_buffer_destroy(bytes);
  fm_workbook_destroy(wb);

  CliRun r = run_cli({"dump", "--view", "metadata", in});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  // Workbook-scoped names print bare; sheet-scoped names are prefixed
  // with `SheetName!` so the two scopes don't collide in the dump.
  // The C getter returns defined-name formulas with a leading '='.
  EXPECT_NE(r.stdout_text.find("WorkbookConst =1\n"), std::string::npos) << "stdout=" << r.stdout_text;
  EXPECT_NE(r.stdout_text.find("Sheet1!SheetLocal =Sheet1!$A$1\n"), std::string::npos) << "stdout=" << r.stdout_text;
}

TEST(FormulonCli, DumpMetadataRecoversFromOutOfRangeDefinedNameScope) {
  std::string in = temp_path("dump_metadata_out_of_range_scope.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_out_of_range_defined_names_fixture(in));

  CliRun r = run_cli({"dump", "--view", "metadata", in});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_TRUE(r.stderr_text.empty());
  EXPECT_EQ(r.stdout_text,
            "[definedNames]\n"
            "Before =1\n"
            "#99!Bad =Sheet1!$A$1\n"
            "After =2\n"
            "[tables]\n"
            "[passthroughParts]\n");
}

TEST(FormulonCli, PaginateOptionTerminatorAcceptsDashLeadingRelativePath) {
  const std::string input = "-fm_cli_paginate_dash_input.xlsx";
  PathGuard guard(input);
  ASSERT_TRUE(write_fixture_workbook(input));

  const CliRun without_terminator = run_cli({"paginate", "--sheet", "0", input});
  EXPECT_EQ(without_terminator.exit_code, 64);

  const CliRun with_terminator = run_cli({"paginate", "--sheet", "0", "--", input});
  EXPECT_EQ(with_terminator.exit_code, 0) << with_terminator.stderr_text;
  EXPECT_EQ(with_terminator.stdout_text,
            "name=Sheet1\npageCount=1\nprintArea=\nrowBreaks=\ncolumnBreaks=\npages=A1:B1\n");
}

TEST(FormulonCli, PaginatePostTerminatorHelpIsAnInputPath) {
  const std::string input = "-h";
  PathGuard guard(input);
  std::remove(input.c_str());

  const CliRun result = run_cli({"paginate", "--", input});
  EXPECT_EQ(result.exit_code, 66);
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

TEST(FormulonCli, DumpRangeNarrowsTheCellViews) {
  const std::string in = temp_path("dump_range.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in, "text"));  // A1 = 7, B1 = =A1+1, A2 = text
  const CliRun values = run_cli({"dump", "--view", "values", "--range", "B1:B1", in});
  EXPECT_EQ(values.exit_code, 0) << values.stderr_text;
  EXPECT_EQ(values.stdout_text, "Sheet1!B1 8\n");
  const CliRun column = run_cli({"dump", "--view", "values", "--range=A1:A2", in});
  EXPECT_EQ(column.stdout_text, "Sheet1!A1 7\nSheet1!A2 text\n");
  const CliRun formulas = run_cli({"dump", "--range", "A1:B2", in});
  EXPECT_EQ(formulas.stdout_text, "Sheet1!B1 =A1+1\n");
  const CliRun json = run_cli({"dump", "--json", "--sheet", "0", "--range", "A2:A2", in});
  EXPECT_EQ(json.exit_code, 0) << json.stderr_text;
  EXPECT_NE(json.stdout_text.find("\"ref\":\"A2\""), std::string::npos) << json.stdout_text;
  EXPECT_EQ(json.stdout_text.find("\"ref\":\"A1\""), std::string::npos) << json.stdout_text;
}

TEST(FormulonCli, DumpSheetSelectsOneSheetByNameOrIndex) {
  const std::string in = temp_path("dump_sheet.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  EXPECT_EQ(run_cli({"dump", "--sheet", "Sheet1", in}).stdout_text, "Sheet1!B1 =A1+1\n");
  EXPECT_EQ(run_cli({"dump", "--sheet", "0", in}).stdout_text, "Sheet1!B1 =A1+1\n");
  const CliRun missing = run_cli({"dump", "--sheet", "Nope", in});
  EXPECT_EQ(missing.exit_code, 64);
  EXPECT_EQ(missing.stdout_text, "");
}

TEST(FormulonCli, DumpRangeAndSheetRejectBadUseAsUsageErrors) {
  const std::string in = temp_path("dump_range_bad.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  EXPECT_EQ(run_cli({"dump", "--range", "not a range", in}).exit_code, 64);
  EXPECT_EQ(run_cli({"dump", "--range", "Sheet1!A1:B2", in}).exit_code, 64);
  EXPECT_EQ(run_cli({"dump", "--view", "sheets", "--range", "A1:B2", in}).exit_code, 64);
  EXPECT_EQ(run_cli({"dump", "--view", "metadata", "--sheet", "0", in}).exit_code, 64);
  EXPECT_EQ(run_cli({"dump", "--range"}).exit_code, 64);
}

TEST(FormulonCli, DumpViewAcceptsTheEqualsFormAndRejectsUnknownViews) {
  const std::string in = temp_path("dump_view.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  EXPECT_EQ(run_cli({"dump", "--view=sheets", in}).stdout_text, run_cli({"dump", "--view", "sheets", in}).stdout_text);
  EXPECT_EQ(run_cli({"dump", "--view", "bogus", in}).exit_code, 64);
  EXPECT_EQ(run_cli({"dump", "--view"}).exit_code, 64);
}

TEST(FormulonCli, DumpRetiredModeFlagsAreUnknownOptions) {
  const std::string in = temp_path("dump_retired_flags.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  for (const char* flag : {"--formulas", "--values", "--sheets", "--metadata"}) {
    const CliRun r = run_cli({"dump", flag, in});
    EXPECT_EQ(r.exit_code, 64) << flag;
    EXPECT_EQ(r.stdout_text, "") << flag;
  }
}

TEST(FormulonCli, DumpViewDoesNotTakeAnOperandBesideTheInput) {
  const std::string in = temp_path("dump_view_operand.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  const CliRun r = run_cli({"dump", "--view", "sheets", "Data", in});
  EXPECT_EQ(r.exit_code, 64);
  EXPECT_EQ(r.stdout_text, "");
}

TEST(FormulonCli, PaginateSheetTakesANameOrAnIndex) {
  const std::string in = temp_path("paginate_sheet_name.xlsx");
  PathGuard guard(in);
  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_create_default(&wb), 0);
  std::size_t second = 0;
  ASSERT_EQ(fm_workbook_add_sheet(wb, "Sheet2", 1, &second), 0);
  ASSERT_EQ(fm_sheet_set_number(wb, second, 0, 0, 5.0), 0);
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ASSERT_EQ(fm_workbook_save(wb, FM_WORKBOOK_FORMAT_XLSX, &bytes, &len), 0);
  ASSERT_TRUE(write_bytes(in, std::string_view(reinterpret_cast<const char*>(bytes), len)));
  fm_buffer_destroy(bytes);
  fm_workbook_destroy(wb);

  const CliRun by_name = run_cli({"paginate", "--sheet", "Sheet2", in});
  ASSERT_EQ(by_name.exit_code, 0) << by_name.stderr_text;
  EXPECT_EQ(by_name.stdout_text, run_cli({"paginate", "--sheet", "1", in}).stdout_text);
  EXPECT_NE(by_name.stdout_text.find("name=Sheet2\n"), std::string::npos) << by_name.stdout_text;
}

TEST(FormulonCli, PaginateUnknownSheetFailsAsDumpDoes) {
  const std::string in = temp_path("paginate_unknown_sheet.xlsx");
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  const CliRun paginate = run_cli({"paginate", "--sheet", "Nope", in});
  const CliRun dump = run_cli({"dump", "--sheet", "Nope", in});
  EXPECT_EQ(paginate.exit_code, 64);
  EXPECT_EQ(paginate.exit_code, dump.exit_code);
  EXPECT_NE(paginate.stderr_text.find("no sheet named or numbered 'Nope'"), std::string::npos) << paginate.stderr_text;
}
