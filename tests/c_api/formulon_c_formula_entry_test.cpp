// Stable C ABI tests for the formula-text entry check: every public setter
// that stores formula text rejects what the formula parser cannot parse and
// leaves the workbook unchanged, while the file-load path keeps such text.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "defined_name.h"
#include "formulon_c_count_of.h"
#include "formulon_c_out_index.h"
#include "formulon_c_record_arg.h"
#include "formulon_c_setters.h"
#include "gtest/gtest.h"
#include "parser/parse_status.h"
#include "utils/error.h"
#include "workbook.h"

namespace {

constexpr fm_status_t kParserUnexpectedToken =
    static_cast<fm_status_t>(formulon::FormulonErrorCode::kParserUnexpectedToken);

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

// "Sheet1" (from create), "Data" and "S2".
void MakeWorkbook(WorkbookGuard* wb) {
  ASSERT_EQ(fm_workbook_create_default(&wb->handle), 0);
  for (const char* name : {"Data", "S2"}) {
    ASSERT_EQ(fm_workbook_add_sheet(wb->handle, name, append_position(wb->handle), ignored_index()), 0)
        << fm_get_last_error_message();
  }
}

// Checks the last failure is a formula syntax refusal: `rc` is the parser
// status of the recorded code and the recorded range lies inside `formula`.
// Call it right after the failing call.
void ExpectParseRefusal(fm_status_t rc, const std::string& formula) {
  EXPECT_STRNE(fm_get_last_error_message(), "");
  fm_parse_error_info_t err = FM_STRUCT_INIT(fm_parse_error_info_t);
  int32_t present = 0;
  ASSERT_EQ(fm_get_last_parse_error(&present, &err), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(rc, static_cast<fm_status_t>(
                    formulon::parser::parse_error_status(static_cast<formulon::parser::ParseErrorCode>(err.code))));
  EXPECT_GE(rc, 1000);
  EXPECT_LE(rc, 1014);
  EXPECT_LE(err.start, err.end);
  EXPECT_LE(err.end, formula.size());
  EXPECT_GE(err.line, 1U);
  EXPECT_GE(err.column, 1U);
}

const std::vector<std::string>& RejectedFormulas() {
  static const std::vector<std::string> formulas = {
      "=TRUE!A1",        "=SUM(",      "=1+",           "=Data!SUM(A1:A2)",        "=SUM(Data:Sheet1!A1",
      "=(1+2",           "SUM(",       "=1\t+2",        "=SUM(1,\r\n2)",           "=SUM (1)",
      "=SUBTOTAL(9,)",   "=SUMIF(,1)", "=COUNTIF(1,1)", "=OFFSET(A1+0,0,0)",       "=ROWS()",
      "=ABS(1,2)",       "=SUM()",     "=IF(1)",        "=COUNTIFS(A1,1,A1)",      "=1+ROWS()",
      "=LARGE(A1,A1,1)", "=#CALC!+1",  "={#CALC!}",     "=IFERROR(#SPILL!,\"x\")", "=#UNKNOWN!",
  };
  return formulas;
}

const std::vector<std::string>& AcceptedFormulas() {
  static const std::vector<std::string> formulas = {
      "='S2'!A1",
      "='S2'!$A$1",
      "='S2'!A:A",
      "='S2'!1:1",
      "=SUM(Data:Sheet1!A1)",
      "=Data!A1",
      "=Gone!A1",
      "=SUM(1,2)",
      "=ROW()",
      "=COUNTIFS(A1,1,A1,1)",
      "=SORTBY(A1,A1)",
      "=TRANSLATE(\"a\")",
      "=FOO(1,2,3)",
      "=A1B!A1",
      "=1+2",
      "=SUM(Data:Other!A1)",
      "=[1]Sheet1!A1",
      "='[Book.xlsx]Sheet'!A1",
      "SUM(1,2)",
      "=S2!A1",
      "=2024!A1",
      "=R1C1!A1",
      "=XFE1!$A$1",
      "=S2!A:A",
      "=S2!1:1",
      "=S2! A1",
      "=Data:S2!A1",
      "=[Book.xlsx]Sheet!A1",
      "=[Book.xlsx]Sheet!$A$1:B2",
      "=[Book.xlsx]Sheet!A:A",
      "=[Book.xlsx]Sheet!1:1",
      "=[Book.xlsx]S1:S2!A1",
      "='/p/[Book.xlsx]S'!A1",
      "='/p/[Book.xlsx]S1:S2'!A1",
      "=Book.xlsx!Name",
      "='/p/Book.xlsx'!Name",
      "=SUM(,)",
      "=IF(A1,,1)",
      "=AGGREGATE(9,4,)",
      "=SUBTOTAL(9,INDIRECT(\"A1\"))",
      "=SUMIF(Nope,1)",
      "=#N/A",
      "=#DIV/0!+1",
      "=#GETTING_DATA",
  };
  return formulas;
}

std::string GetFormula(const fm_workbook_t* wb, std::uint32_t row, std::uint32_t col) {
  int32_t present = 0;
  const char* text = nullptr;
  EXPECT_EQ(fm_sheet_get_formula(wb, 0, row, col, &present, &text), 0);
  return text != nullptr ? text : "";
}

TEST(FormulonCApiFormulaEntry, SetFormulaRejectsUnparseableTextAndKeepsState) {
  WorkbookGuard wb;
  MakeWorkbook(&wb);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 1, "=1+2"), 0);
  for (const std::string& formula : RejectedFormulas()) {
    SCOPED_TRACE(formula);
    ExpectParseRefusal(fm_sheet_set_formula(wb.handle, 0, 0, 1, formula.c_str()), formula);
    EXPECT_NE(std::string(fm_get_last_error_context()).find("row=0 col=1"), std::string::npos);
    EXPECT_EQ(GetFormula(wb.handle, 0, 1), "=1+2");

    ExpectParseRefusal(fm_sheet_set_formula(wb.handle, 0, 5, 5, formula.c_str()), formula);
    EXPECT_EQ(GetFormula(wb.handle, 5, 5), "");
  }
}

TEST(FormulonCApiFormulaEntry, SetFormulaAcceptedTextAlwaysHasAnR1C1) {
  WorkbookGuard wb;
  MakeWorkbook(&wb);
  for (const std::string& formula : AcceptedFormulas()) {
    SCOPED_TRACE(formula);
    ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 2, 2, formula.c_str()), 0) << fm_get_last_error_message();
    int32_t present = 0;
    const char* r1c1 = nullptr;
    EXPECT_EQ(fm_sheet_get_formula_r1c1(wb.handle, 0, 2, 2, &present, &r1c1), 0) << fm_get_last_error_message();
  }
}

// `[Book]Sheet` with no `!` is accepted as a name: verbatim in A1 and R1C1,
// #NAME? (ISREF FALSE), and no external link.
TEST(FormulonCApiFormulaEntry, BracketNameWithoutBangIsAcceptedAsAName) {
  WorkbookGuard wb;
  MakeWorkbook(&wb);
  constexpr const char* kFormulas[] = {"=[Book.xlsx]Sheet", "=[Book.xlsx]Sheet+1", "=ISREF([Src.xlsx]Data)"};
  for (std::uint32_t row = 0; row < 3U; ++row) {
    ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, row, 0, kFormulas[row]), 0) << fm_get_last_error_message();
  }
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  for (std::uint32_t row = 0; row < 3U; ++row) {
    SCOPED_TRACE(kFormulas[row]);
    EXPECT_EQ(GetFormula(wb.handle, row, 0), kFormulas[row]);
    int32_t present = 0;
    const char* r1c1 = nullptr;
    ASSERT_EQ(fm_sheet_get_formula_r1c1(wb.handle, 0, row, 0, &present, &r1c1), 0) << fm_get_last_error_message();
    EXPECT_STREQ(r1c1, kFormulas[row]);  // these formulas read the same in R1C1
    fm_value_t value{};
    ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, row, 0, &value), 0);
    if (row < 2U) {
      ASSERT_EQ(value.kind, FM_VALUE_KIND_ERROR);
      EXPECT_EQ(value.u.error_code, 4);  // #NAME?
    } else {
      ASSERT_EQ(value.kind, FM_VALUE_KIND_BOOL);
      EXPECT_EQ(value.u.boolean, 0);
    }
  }
  std::size_t links = 1;
  ASSERT_EQ(fm_workbook_get_external_link_count(wb.handle, &links), 0);
  EXPECT_EQ(links, 0U);
}

TEST(FormulonCApiFormulaEntry, DefinedNameRejectsUnparseableTextAndKeepsState) {
  WorkbookGuard wb;
  MakeWorkbook(&wb);
  ASSERT_EQ(SetDefinedName(wb.handle, "Rate", "Sheet1!$A$1"), 0);
  ASSERT_EQ(SetDefinedName(wb.handle, "Local", "Sheet1!$B$1", 0), 0);
  ASSERT_EQ(count_of(fm_workbook_get_defined_name_count, wb.handle), 2U);
  for (const std::string& formula : RejectedFormulas()) {
    SCOPED_TRACE(formula);
    ExpectParseRefusal(SetDefinedName(wb.handle, "Rate", formula.c_str()), formula);
    ExpectParseRefusal(SetDefinedName(wb.handle, "Fresh", formula.c_str()), formula);
    ExpectParseRefusal(SetDefinedName(wb.handle, "Local", formula.c_str(), 0), formula);
    ExpectParseRefusal(SetDefinedName(wb.handle, "Fresh", formula.c_str(), -1), formula);
  }
  ASSERT_EQ(count_of(fm_workbook_get_defined_name_count, wb.handle), 2U);
  fm_defined_name_t entry = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 0, &entry), 0);
  EXPECT_STREQ(entry.formula, "=Sheet1!$A$1");
  ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 1, &entry), 0);
  EXPECT_STREQ(entry.formula, "=Sheet1!$B$1");
}

TEST(FormulonCApiFormulaEntry, DefinedNameAcceptsParseableText) {
  WorkbookGuard wb;
  MakeWorkbook(&wb);
  EXPECT_EQ(SetDefinedName(wb.handle, "Quoted", "'S2'!$A$1"), 0);
  EXPECT_EQ(SetDefinedName(wb.handle, "Missing", "Gone!$A$1"), 0);
  EXPECT_EQ(SetDefinedName(wb.handle, "Const", "1.5"), 0);
  EXPECT_EQ(SetDefinedName(wb.handle, "Local", "'S2'!$A$1", 1), 0);
  ASSERT_EQ(count_of(fm_workbook_get_defined_name_count, wb.handle), 4U);
}

// A workbook of `profile`.
void MakeProfileWorkbook(WorkbookGuard* wb, const char* profile) {
  ASSERT_EQ(fm_workbook_create_for_profile(profile, &wb->handle), 0) << fm_get_last_error_message();
}

TEST(FormulonCApiFormulaEntry, DefinedNameOpeningWithLocaleR1C1TokenIsRefused) {
  struct Case {
    const char* profile;
    const char* name;
  };
  for (const Case& c : {Case{"win-365-de_DE", "z"}, Case{"win-365-de_DE", "S"}, Case{"win-365-de_DE", "Z2S3"},
                        Case{"win-365-es_ES", "f"}, Case{"win-365-es_MX", "F"}, Case{"win-365-nl_NL", "k"},
                        Case{"win-365-nl_NL", "K"}}) {
    SCOPED_TRACE(std::string(c.profile) + " " + c.name);
    WorkbookGuard wb;
    MakeProfileWorkbook(&wb, c.profile);
    EXPECT_EQ(SetDefinedName(wb.handle, c.name, "Sheet1!$A$1"), static_cast<fm_status_t>(2));
    EXPECT_STREQ(fm_get_last_error_message(), "name is not a valid defined name");
    EXPECT_EQ(count_of(fm_workbook_get_defined_name_count, wb.handle), 0U);
  }
}

TEST(FormulonCApiFormulaEntry, NewDefinedNameOpeningWithProfileR1C1ReferenceIsRefusedOnlyUnderThatProfile) {
  WorkbookGuard de;
  MakeProfileWorkbook(&de, "win-365-de_DE");
  WorkbookGuard en;
  MakeProfileWorkbook(&en, "win-365-en_US");
  for (const char* name : {"Z1S1x", "z1s1"}) {
    SCOPED_TRACE(name);
    EXPECT_EQ(SetDefinedName(de.handle, name, "Sheet1!$A$1"), static_cast<fm_status_t>(2));
    EXPECT_EQ(SetDefinedName(en.handle, name, "Sheet1!$A$1"), 0) << fm_get_last_error_message();
  }
  EXPECT_EQ(count_of(fm_workbook_get_defined_name_count, de.handle), 0U);
  for (const char* name : {"Zeile", "Zone"}) {
    SCOPED_TRACE(name);
    EXPECT_EQ(SetDefinedName(de.handle, name, "Sheet1!$A$1"), 0) << fm_get_last_error_message();
  }
}

TEST(FormulonCApiFormulaEntry, RedefiningNameByCaseAfterProfileSwitchReplacesIt) {
  WorkbookGuard wb;
  MakeProfileWorkbook(&wb, "win-365-en_US");
  ASSERT_EQ(SetDefinedName(wb.handle, "Z1S1x", "Sheet1!$A$1"), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "win-365-de_DE"), 0) << fm_get_last_error_message();
  ASSERT_EQ(SetDefinedName(wb.handle, "z1s1X", "Sheet1!$A$2"), 0) << fm_get_last_error_message();
  ASSERT_EQ(count_of(fm_workbook_get_defined_name_count, wb.handle), 1U);
  fm_defined_name_t out = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 0, &out), 0);
  EXPECT_STREQ(out.name, "z1s1X");
}

TEST(FormulonCApiFormulaEntry, ScopedDefinedNameIsCheckedPerScope) {
  WorkbookGuard wb;
  MakeProfileWorkbook(&wb, "win-365-en_US");
  ASSERT_EQ(SetDefinedName(wb.handle, "Z1S1x", "Sheet1!$A$1", 0), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "win-365-de_DE"), 0) << fm_get_last_error_message();
  // The sheet-scoped entry exists, so it may be redefined; the workbook scope has none, so it is a new name.
  EXPECT_EQ(SetDefinedName(wb.handle, "z1s1x", "Sheet1!$A$2", 0), 0) << fm_get_last_error_message();
  EXPECT_EQ(SetDefinedName(wb.handle, "Z1S1x", "Sheet1!$A$1", -1), static_cast<fm_status_t>(2));
  EXPECT_EQ(count_of(fm_workbook_get_defined_name_count, wb.handle), 1U);
}

TEST(FormulonCApiFormulaEntry, TableNamesOpeningWithProfileR1C1ReferenceAreRefused) {
  WorkbookGuard wb;
  MakeProfileWorkbook(&wb, "win-365-de_DE");
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, "Col"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 0, 1.0), 0);
  const char* columns[] = {"Col"};
  std::size_t index = 0;
  EXPECT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{0, 0, 1, 0}), "Z1S1T", columns, 1, "", 1, 0, &index),
            static_cast<fm_status_t>(2));
  ASSERT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{0, 0, 1, 0}), "Sales", columns, 1, "", 1, 0, &index), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_rename_table(wb.handle, 0, index, "Z1S1T"), static_cast<fm_status_t>(2));
  EXPECT_EQ(wb.handle->workbook().tables()[index].name, "Sales");
}

// A table over A1:A3 named `table_name` (header Col, data 1 and 2), created under en-US and then
// read under `profile`, since a new table name opening with the profile's R1C1 letters is refused.
void MakeTableWorkbook(WorkbookGuard* wb, const char* profile, const char* table_name) {
  MakeProfileWorkbook(wb, "win-365-en_US");
  ASSERT_EQ(fm_sheet_set_text(wb->handle, 0, 0, 0, "Col"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb->handle, 0, 1, 0, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_number(wb->handle, 0, 2, 0, 2.0), 0);
  const char* columns[] = {"Col"};
  std::size_t index = 0;
  ASSERT_EQ(AddTable(wb->handle, 0, record_arg(fm_cell_rect_t{0, 0, 2, 0}), table_name, columns, 1, "", 1, 0, &index),
            0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_profile_id(wb->handle, profile), 0) << fm_get_last_error_message();
}

TEST(FormulonCApiFormulaEntry, TableMayBeRenamedByCaseUnderProfileWhoseR1C1LettersItOpensWith) {
  WorkbookGuard wb;
  MakeTableWorkbook(&wb, "win-365-de_DE", "Z1S1T");
  EXPECT_EQ(fm_sheet_rename_table(wb.handle, 0, 0, "z1s1t"), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_rename_table(wb.handle, 0, 0, "Z1S1U"), static_cast<fm_status_t>(2));
}

TEST(FormulonCApiFormulaEntry, DefinedNameLocaleR1C1TokenIsAcceptedUnderOtherProfiles) {
  struct Case {
    const char* profile;
    const char* name;
  };
  for (const Case& c :
       {Case{"win-365-en_US", "z"}, Case{"win-365-en_US", "s"}, Case{"win-365-en_US", "f"}, Case{"win-365-en_US", "k"},
        Case{"win-365-ja_JP", "z"}, Case{"win-365-ja_JP", "f"}, Case{"win-365-ja_JP", "k"}, Case{"win-365-de_DE", "f"},
        Case{"win-365-es_ES", "z"}, Case{"win-365-nl_NL", "s"}}) {
    SCOPED_TRACE(std::string(c.profile) + " " + c.name);
    WorkbookGuard wb;
    MakeProfileWorkbook(&wb, c.profile);
    EXPECT_EQ(SetDefinedName(wb.handle, c.name, "Sheet1!$A$1"), 0) << fm_get_last_error_message();
  }
}

TEST(FormulonCApiFormulaEntry, DefinedNameRefusedEverywhereForRAndC) {
  for (const char* profile : {"win-365-en_US", "win-365-de_DE", "win-365-es_ES", "win-365-nl_NL"}) {
    WorkbookGuard wb;
    MakeProfileWorkbook(&wb, profile);
    EXPECT_NE(SetDefinedName(wb.handle, "r", "Sheet1!$A$1"), 0) << profile;
    EXPECT_NE(SetDefinedName(wb.handle, "c", "Sheet1!$A$1"), 0) << profile;
  }
}

TEST(FormulonCApiFormulaEntry, LoadedDefinedNameOpeningWithLocaleR1C1TokenIsKept) {
  WorkbookGuard source;
  MakeProfileWorkbook(&source, "win-365-de_DE");
  // Workbook::set_defined_names is the readers' ingestion point and stores names as given.
  formulon::DefinedName entry;
  entry.name = "z";
  entry.formula = "Sheet1!$A$1";
  source.handle->workbook().set_defined_names({entry});

  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ASSERT_EQ(fm_workbook_save(source.handle, FM_WORKBOOK_FORMAT_XLSX, &bytes, &len), 0) << fm_get_last_error_message();
  WorkbookGuard loaded;
  const fm_status_t rc = fm_workbook_load(bytes, len, &loaded.handle);
  fm_buffer_destroy(bytes);
  ASSERT_EQ(rc, 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_profile_id(loaded.handle, "win-365-de_DE"), 0) << fm_get_last_error_message();
  ASSERT_EQ(count_of(fm_workbook_get_defined_name_count, loaded.handle), 1U);
  fm_defined_name_t out = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(fm_workbook_get_defined_name_at(loaded.handle, 0, &out), 0);
  EXPECT_STREQ(out.name, "z");
  // The loaded name stays editable under its own spelling; a new one is still refused.
  EXPECT_EQ(SetDefinedName(loaded.handle, "z", "Sheet1!$A$2"), 0) << fm_get_last_error_message();
  EXPECT_NE(SetDefinedName(loaded.handle, "s", "Sheet1!$A$2"), 0);
  EXPECT_EQ(count_of(fm_workbook_get_defined_name_count, loaded.handle), 1U);
}

TEST(FormulonCApiFormulaEntry, TableNamedLikeLocaleR1C1TokenRefusesItsReferenceOnEntry) {
  struct Case {
    const char* profile;
    const char* table;
  };
  for (const Case& c : {Case{"win-365-es_ES", "F"}, Case{"win-365-nl_NL", "K"}}) {
    SCOPED_TRACE(std::string(c.profile) + " " + c.table);
    WorkbookGuard wb;
    MakeTableWorkbook(&wb, c.profile, c.table);
    for (const std::string body : {"SUM(%s[Col])", "SUM(%s[[#Data],[Col]])"}) {
      std::string formula = "=" + body;
      formula.replace(formula.find("%s"), 2, c.table);
      SCOPED_TRACE(formula);
      ExpectParseRefusal(fm_sheet_set_formula(wb.handle, 0, 0, 3, formula.c_str()), formula);
      EXPECT_EQ(GetFormula(wb.handle, 0, 3), "");
    }
  }
}

TEST(FormulonCApiFormulaEntry, TableReferenceInStoredFormulaFieldsIsRefusedUnderItsLocale) {
  WorkbookGuard wb;
  MakeTableWorkbook(&wb, "win-365-es_ES", "F");
  const std::string formula = "=SUM(F[Col])";
  ExpectParseRefusal(SetDefinedName(wb.handle, "Rate", formula.c_str()), formula);
  EXPECT_EQ(count_of(fm_workbook_get_defined_name_count, wb.handle), 0U);
}

TEST(FormulonCApiFormulaEntry, ThisRowReferenceInsideTableNamedLikeLocaleR1C1TokenIsAccepted) {
  WorkbookGuard wb;
  MakeProfileWorkbook(&wb, "win-365-en_US");
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, "Col"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 1, "Other"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 0, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 2, 0, 2.0), 0);
  const char* columns[] = {"Col", "Other"};
  std::size_t index = 0;
  ASSERT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{0, 0, 2, 1}), "F", columns, 2, "", 1, 0, &index), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "win-365-es_ES"), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_set_formula(wb.handle, 0, 1, 1, "=[@Col]"), 0) << fm_get_last_error_message();
}

TEST(FormulonCApiFormulaEntry, TableReferenceOpeningWithR1C1TokenIsAcceptedOutsideItsLocale) {
  for (const auto& [profile, table] : {std::pair<const char*, const char*>{"win-365-en_US", "F"},
                                       {"win-365-ja_JP", "K"},
                                       {"win-365-de_DE", "F"},
                                       {"win-365-de_DE", "K"},
                                       {"win-365-es_ES", "K"},
                                       {"win-365-nl_NL", "F"}}) {
    SCOPED_TRACE(std::string(profile) + " " + table);
    WorkbookGuard wb;
    MakeTableWorkbook(&wb, profile, table);
    const std::string formula = std::string("=SUM(") + table + "[Col])";
    EXPECT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 3, formula.c_str()), 0) << fm_get_last_error_message();
  }
}

TEST(FormulonCApiFormulaEntry, LoadedFormulaReferencingTableNamedLikeLocaleR1C1TokenStillEvaluates) {
  WorkbookGuard source;
  MakeTableWorkbook(&source, "win-365-es_ES", "F");
  ASSERT_TRUE(source.handle->workbook().set_cell_formula(0, 0, 3, "=SUM(F[Col])"));

  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ASSERT_EQ(fm_workbook_save(source.handle, FM_WORKBOOK_FORMAT_XLSX, &bytes, &len), 0) << fm_get_last_error_message();
  WorkbookGuard loaded;
  const fm_status_t rc = fm_workbook_load(bytes, len, &loaded.handle);
  fm_buffer_destroy(bytes);
  ASSERT_EQ(rc, 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);
  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_value(loaded.handle, 0, 0, 3, &value), 0);
  ASSERT_EQ(value.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_EQ(value.u.number, 3.0);
}

fm_cf_rule_t MakeCfRule(const fm_cell_rect_t* sqref, const char* formula1, const char* formula2) {
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.type = 0;  // Expression
  rule.formula1 = formula1;
  rule.formula2 = formula2;
  rule.ranges = sqref;
  rule.range_count = 1;
  return rule;
}

TEST(FormulonCApiFormulaEntry, CfAddRuleRejectsUnparseableFormulasAndKeepsState) {
  WorkbookGuard wb;
  MakeWorkbook(&wb);
  const fm_cell_rect_t sqref{0, 0, 9, 0};
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, record_arg(MakeCfRule(&sqref, "A1>0", nullptr)), ignored_id()), 0);
  for (const std::string& formula : RejectedFormulas()) {
    SCOPED_TRACE(formula);
    ExpectParseRefusal(
        fm_sheet_add_cf_rule(wb.handle, 0, record_arg(MakeCfRule(&sqref, formula.c_str(), nullptr)), ignored_id()),
        formula);
    EXPECT_NE(std::string(fm_get_last_error_context()).find("formula1"), std::string::npos);
    ExpectParseRefusal(
        fm_sheet_add_cf_rule(wb.handle, 0, record_arg(MakeCfRule(&sqref, "A1>0", formula.c_str())), ignored_id()),
        formula);
    EXPECT_NE(std::string(fm_get_last_error_context()).find("formula2"), std::string::npos);
  }
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);
}

TEST(FormulonCApiFormulaEntry, CfAddRuleAcceptsParseableAndEmptyFormulas) {
  WorkbookGuard wb;
  MakeWorkbook(&wb);
  const fm_cell_rect_t sqref{0, 0, 9, 0};
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, record_arg(MakeCfRule(&sqref, "'S2'!$A$1>0", "")), ignored_id()), 0);
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, record_arg(MakeCfRule(&sqref, "Gone!$A$1>0", nullptr)), ignored_id()),
            0);
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, record_arg(MakeCfRule(&sqref, "", "")), ignored_id()), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 3U);
}

fm_data_validation_t MakeValidation(const fm_cell_rect_t* range, std::uint8_t type, const char* formula1,
                                    const char* formula2) {
  fm_data_validation_t v = FM_STRUCT_INIT(fm_data_validation_t);
  v.ranges = range;
  v.range_count = 1;
  v.type = type;
  v.formula1 = formula1;
  v.formula2 = formula2;
  return v;
}

TEST(FormulonCApiFormulaEntry, AddValidationRejectsUnparseableFormulasAndKeepsState) {
  WorkbookGuard wb;
  MakeWorkbook(&wb);
  const fm_cell_rect_t range{0, 0, 9, 0};
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, record_arg(MakeValidation(&range, 1, "0", "10")), ignored_id()), 0);
  for (const std::string& formula : RejectedFormulas()) {
    SCOPED_TRACE(formula);
    ExpectParseRefusal(fm_sheet_add_validation(
                           wb.handle, 0, record_arg(MakeValidation(&range, 1, formula.c_str(), "10")), ignored_id()),
                       formula);
    EXPECT_NE(std::string(fm_get_last_error_context()).find("formula1"), std::string::npos);
    ExpectParseRefusal(fm_sheet_add_validation(
                           wb.handle, 0, record_arg(MakeValidation(&range, 1, "0", formula.c_str())), ignored_id()),
                       formula);
    EXPECT_NE(std::string(fm_get_last_error_context()).find("formula2"), std::string::npos);
  }
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_validation_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);
}

TEST(FormulonCApiFormulaEntry, AddValidationAcceptsEveryListAndScalarShape) {
  WorkbookGuard wb;
  MakeWorkbook(&wb);
  const fm_cell_rect_t range{0, 0, 9, 0};
  constexpr std::uint8_t kList = 3;
  EXPECT_EQ(fm_sheet_add_validation(wb.handle, 0, record_arg(MakeValidation(&range, kList, "\"a,b,c\"", nullptr)),
                                    ignored_id()),
            0);
  EXPECT_EQ(fm_sheet_add_validation(wb.handle, 0, record_arg(MakeValidation(&range, kList, "'S2'!$A$1:$A$5", nullptr)),
                                    ignored_id()),
            0);
  EXPECT_EQ(fm_sheet_add_validation(wb.handle, 0, record_arg(MakeValidation(&range, kList, "Gone!$A$1:$A$5", nullptr)),
                                    ignored_id()),
            0);
  EXPECT_EQ(fm_sheet_add_validation(wb.handle, 0, record_arg(MakeValidation(&range, kList, "ListName", nullptr)),
                                    ignored_id()),
            0);
  EXPECT_EQ(
      fm_sheet_add_validation(wb.handle, 0, record_arg(MakeValidation(&range, 1, "-3.5", "TODAY()")), ignored_id()), 0);
  EXPECT_EQ(
      fm_sheet_add_validation(wb.handle, 0, record_arg(MakeValidation(&range, 7, "ISNUMBER(A1)", "")), ignored_id()),
      0);
  EXPECT_EQ(
      fm_sheet_add_validation(wb.handle, 0, record_arg(MakeValidation(&range, 0, nullptr, nullptr)), ignored_id()), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_validation_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 7U);
}

TEST(FormulonCApiFormulaEntry, EveryFormulaSetterStoresTheEnteredSpelling) {
  WorkbookGuard wb;
  MakeWorkbook(&wb);
  // Cell, defined name, CF rule and DV rule all take the spelling a cell does.
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 1, "=SUM( +1.50 , 1E3 ) "), 0);
  EXPECT_EQ(GetFormula(wb.handle, 0, 1), "=SUM( 1.5, 1000 )");

  ASSERT_EQ(SetDefinedName(wb.handle, "Rate", "=Sheet1!$A$1+-0 "), 0);
  ASSERT_EQ(SetDefinedName(wb.handle, "Local", " 1+ +5", 0), 0);
  fm_defined_name_t entry = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 0, &entry), 0);
  EXPECT_STREQ(entry.formula, "=Sheet1!$A$1+0");
  ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 1, &entry), 0);
  EXPECT_STREQ(entry.formula, "= 1+ 5");

  const fm_cell_rect_t sqref{0, 0, 9, 0};
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, record_arg(MakeCfRule(&sqref, "=$A$1>+1.50", "{1, 2}")), ignored_id()),
            0);
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 0, &rule), 0);
  EXPECT_STREQ(rule.formula1, "=$A$1>1.5");
  EXPECT_STREQ(rule.formula2, "={1,2}");

  const fm_cell_rect_t range{0, 0, 9, 0};
  constexpr std::uint8_t kList = 3;
  constexpr std::uint8_t kCustom = 7;
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, record_arg(MakeValidation(&range, kList, "\"1.50,1E3,+5\"", nullptr)),
                                    ignored_id()),
            0);
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, record_arg(MakeValidation(&range, kList, "+$A$1:$A$3", nullptr)),
                                    ignored_id()),
            0);
  ASSERT_EQ(fm_sheet_add_validation(
                wb.handle, 0, record_arg(MakeValidation(&range, kList, "OFFSET($A$1,+0,-0,1E3/1E3+2,1)", nullptr)),
                ignored_id()),
            0);
  ASSERT_EQ(
      fm_sheet_add_validation(
          wb.handle, 0, record_arg(MakeValidation(&range, kCustom, "SUM(A1 : A2)>-1E-309", nullptr)), ignored_id()),
      0);
  const char* const expected[] = {"=\"1.50,1E3,+5\"", "=+$A$1:$A$3", "=OFFSET($A$1,0,0,1000/1000+2,1)",
                                  "=SUM(A1:A2)>0"};
  for (std::uint32_t i = 0; i < 4U; ++i) {
    fm_data_validation_t v = FM_STRUCT_INIT(fm_data_validation_t);
    ASSERT_EQ(fm_sheet_get_validation_at(wb.handle, 0, i, &v), 0);
    EXPECT_STREQ(v.formula1, expected[i]);
  }
}

// Whitespace Excel keeps is stored as typed; spaces before a line feed go.
TEST(FormulonCApiFormulaEntry, EnteredWhitespaceTakesTheStoredSpelling) {
  WorkbookGuard wb;
  MakeWorkbook(&wb);
  const std::pair<const char*, const char*> cases[] = {
      {"=1+ (2)", "=1+ (2)"},
      {"=SUM(1\n)", "=SUM(1\n)"},
      {"=NOW( )", "=NOW( )"},
      {"=1+ \n2", "=1+\n2"},
      {"=LET(x, 1,\n x)", "=LET(x, 1,\n x)"},
      {"=(A1, A2)", "=(A1, A2)"},
  };
  std::uint32_t row = 10;
  for (const auto& [typed, stored] : cases) {
    SCOPED_TRACE(typed);
    ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, row, 2, typed), 0) << fm_get_last_error_message();
    EXPECT_EQ(GetFormula(wb.handle, row, 2), stored);
    ++row;
  }
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  for (const auto& [r, expected] : {std::pair<std::uint32_t, double>{10, 3.0}, {11, 1.0}, {13, 3.0}, {14, 1.0}}) {
    fm_value_t value{};
    ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, r, 2, &value), 0);
    ASSERT_EQ(value.kind, FM_VALUE_KIND_NUMBER) << "row " << r;
    EXPECT_EQ(value.u.number, expected) << "row " << r;
  }
  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 15, 2, &value), 0);
  ASSERT_EQ(value.kind, FM_VALUE_KIND_ERROR);
  EXPECT_EQ(value.u.error_code, 2);  // #VALUE!
}

TEST(FormulonCApiFormulaEntry, LoadedUnparseableFormulaIsKeptVerbatim) {
  WorkbookGuard source;
  MakeWorkbook(&source);
  // Workbook::set_cell_formula is the reader's ingestion point and stores text as given.
  ASSERT_TRUE(source.handle->workbook().set_cell_formula(0, 0, 1, "=TRUE!A1"));
  ASSERT_TRUE(source.handle->workbook().set_cell_formula(0, 1, 1, "=SUM("));

  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ASSERT_EQ(fm_workbook_save(source.handle, FM_WORKBOOK_FORMAT_XLSX, &bytes, &len), 0) << fm_get_last_error_message();
  WorkbookGuard loaded;
  const fm_status_t rc = fm_workbook_load(bytes, len, &loaded.handle);
  fm_buffer_destroy(bytes);
  ASSERT_EQ(rc, 0) << fm_get_last_error_message();

  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);
  for (const auto& [row, text] : {std::pair<std::uint32_t, const char*>{0, "=TRUE!A1"}, {1, "=SUM("}}) {
    SCOPED_TRACE(text);
    EXPECT_EQ(GetFormula(loaded.handle, row, 1), text);
    fm_value_t value{};
    ASSERT_EQ(fm_sheet_get_value(loaded.handle, 0, row, 1, &value), 0);
    ASSERT_EQ(value.kind, FM_VALUE_KIND_ERROR);
    EXPECT_EQ(value.u.error_code, 4);  // #NAME?
    int32_t present = 0;
    const char* r1c1 = nullptr;
    EXPECT_EQ(fm_sheet_get_formula_r1c1(loaded.handle, 0, row, 1, &present, &r1c1), kParserUnexpectedToken);
  }
}

TEST(FormulonCApiFormulaEntry, LoadedWrongArityFormulaKeepsEvaluatingButIsRefusedOnEntry) {
  WorkbookGuard source;
  MakeWorkbook(&source);
  ASSERT_TRUE(source.handle->workbook().set_cell_formula(0, 0, 1, "=ROWS()"));

  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ASSERT_EQ(fm_workbook_save(source.handle, FM_WORKBOOK_FORMAT_XLSX, &bytes, &len), 0) << fm_get_last_error_message();
  WorkbookGuard loaded;
  const fm_status_t rc = fm_workbook_load(bytes, len, &loaded.handle);
  fm_buffer_destroy(bytes);
  ASSERT_EQ(rc, 0) << fm_get_last_error_message();

  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);
  EXPECT_EQ(GetFormula(loaded.handle, 0, 1), "=ROWS()");
  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_value(loaded.handle, 0, 0, 1, &value), 0);
  EXPECT_EQ(value.kind, FM_VALUE_KIND_ERROR);
  EXPECT_EQ(fm_sheet_set_formula(loaded.handle, 0, 0, 1, "=ROWS()"), kParserUnexpectedToken);
  EXPECT_EQ(GetFormula(loaded.handle, 0, 1), "=ROWS()");
}

TEST(FormulonCApiFormulaEntry, NewerErrorLiteralsAreRefusedOnEntryButLoadAndStoreAsValues) {
  constexpr std::int32_t kCalc = static_cast<std::int32_t>(formulon::ErrorCode::Calc);
  WorkbookGuard source;
  MakeWorkbook(&source);
  ASSERT_TRUE(source.handle->workbook().set_cell_formula(0, 0, 1, "=#CALC!+1"));
  ASSERT_EQ(fm_sheet_set_error(source.handle, 0, 1, 1, kCalc), 0);
  fm_value_t stored{};
  ASSERT_EQ(fm_sheet_get_value(source.handle, 0, 1, 1, &stored), 0);
  ASSERT_EQ(stored.kind, FM_VALUE_KIND_ERROR);
  EXPECT_EQ(stored.u.error_code, kCalc);

  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ASSERT_EQ(fm_workbook_save(source.handle, FM_WORKBOOK_FORMAT_XLSX, &bytes, &len), 0) << fm_get_last_error_message();
  WorkbookGuard loaded;
  const fm_status_t rc = fm_workbook_load(bytes, len, &loaded.handle);
  fm_buffer_destroy(bytes);
  ASSERT_EQ(rc, 0) << fm_get_last_error_message();

  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);
  EXPECT_EQ(GetFormula(loaded.handle, 0, 1), "=#CALC!+1");
  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_value(loaded.handle, 0, 0, 1, &value), 0);
  ASSERT_EQ(value.kind, FM_VALUE_KIND_ERROR);
  EXPECT_EQ(value.u.error_code, kCalc);
  EXPECT_EQ(fm_sheet_set_formula(loaded.handle, 0, 0, 1, "=#CALC!+1"), kParserUnexpectedToken);
  EXPECT_EQ(GetFormula(loaded.handle, 0, 1), "=#CALC!+1");
}

}  // namespace
