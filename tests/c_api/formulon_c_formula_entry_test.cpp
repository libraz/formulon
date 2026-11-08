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
    ASSERT_EQ(fm_workbook_add_sheet(wb->handle, name, ignored_index()), 0) << fm_get_last_error_message();
  }
}

// Checks the last failure is a formula syntax refusal: `rc` is the parser
// status of the recorded code and the recorded range lies inside `formula`.
// Call it right after the failing call.
void ExpectParseRefusal(fm_status_t rc, const std::string& formula) {
  EXPECT_STRNE(fm_get_last_error_message(), "");
  fm_parse_error_t err = FM_STRUCT_INIT(fm_parse_error_t);
  ASSERT_EQ(fm_get_last_parse_error(&err), 0);
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
      "=TRUE!A1",        "=SUM(",      "=1+",           "=Data!SUM(A1:A2)",   "=SUM(Data:Sheet1!A1",
      "=(1+2",           "SUM(",       "=1\t+2",        "=SUM(1,\r\n2)",      "=SUM (1)",
      "=SUBTOTAL(9,)",   "=SUMIF(,1)", "=COUNTIF(1,1)", "=OFFSET(A1+0,0,0)",  "=ROWS()",
      "=ABS(1,2)",       "=SUM()",     "=IF(1)",        "=COUNTIFS(A1,1,A1)", "=1+ROWS()",
      "=LARGE(A1,A1,1)",
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
  };
  return formulas;
}

std::string GetFormula(const fm_workbook_t* wb, std::uint32_t row, std::uint32_t col) {
  const char* text = nullptr;
  EXPECT_EQ(fm_sheet_get_formula(wb, 0, row, col, &text), 0);
  return text != nullptr ? text : "<null>";
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
    const char* r1c1 = nullptr;
    EXPECT_EQ(fm_sheet_get_formula_r1c1(wb.handle, 0, 2, 2, &r1c1), 0) << fm_get_last_error_message();
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
    const char* r1c1 = nullptr;
    ASSERT_EQ(fm_sheet_get_formula_r1c1(wb.handle, 0, row, 0, &r1c1), 0) << fm_get_last_error_message();
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
  std::size_t index = 0;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, record_arg(MakeCfRule(&sqref, "A1>0", nullptr)), &index), 0);
  for (const std::string& formula : RejectedFormulas()) {
    SCOPED_TRACE(formula);
    ExpectParseRefusal(
        fm_sheet_add_cf_rule(wb.handle, 0, record_arg(MakeCfRule(&sqref, formula.c_str(), nullptr)), &index), formula);
    EXPECT_NE(std::string(fm_get_last_error_context()).find("formula1"), std::string::npos);
    ExpectParseRefusal(
        fm_sheet_add_cf_rule(wb.handle, 0, record_arg(MakeCfRule(&sqref, "A1>0", formula.c_str())), &index), formula);
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
  std::size_t index = 0;
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, record_arg(MakeCfRule(&sqref, "'S2'!$A$1>0", "")), &index), 0);
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, record_arg(MakeCfRule(&sqref, "Gone!$A$1>0", nullptr)), &index), 0);
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, record_arg(MakeCfRule(&sqref, "", "")), &index), 0);
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
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, record_arg(MakeValidation(&range, 1, "0", "10")), ignored_index()),
            0);
  for (const std::string& formula : RejectedFormulas()) {
    SCOPED_TRACE(formula);
    ExpectParseRefusal(fm_sheet_add_validation(
                           wb.handle, 0, record_arg(MakeValidation(&range, 1, formula.c_str(), "10")), ignored_index()),
                       formula);
    EXPECT_NE(std::string(fm_get_last_error_context()).find("formula1"), std::string::npos);
    ExpectParseRefusal(fm_sheet_add_validation(
                           wb.handle, 0, record_arg(MakeValidation(&range, 1, "0", formula.c_str())), ignored_index()),
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
                                    ignored_index()),
            0);
  EXPECT_EQ(fm_sheet_add_validation(wb.handle, 0, record_arg(MakeValidation(&range, kList, "'S2'!$A$1:$A$5", nullptr)),
                                    ignored_index()),
            0);
  EXPECT_EQ(fm_sheet_add_validation(wb.handle, 0, record_arg(MakeValidation(&range, kList, "Gone!$A$1:$A$5", nullptr)),
                                    ignored_index()),
            0);
  EXPECT_EQ(fm_sheet_add_validation(wb.handle, 0, record_arg(MakeValidation(&range, kList, "ListName", nullptr)),
                                    ignored_index()),
            0);
  EXPECT_EQ(
      fm_sheet_add_validation(wb.handle, 0, record_arg(MakeValidation(&range, 1, "-3.5", "TODAY()")), ignored_index()),
      0);
  EXPECT_EQ(
      fm_sheet_add_validation(wb.handle, 0, record_arg(MakeValidation(&range, 7, "ISNUMBER(A1)", "")), ignored_index()),
      0);
  EXPECT_EQ(
      fm_sheet_add_validation(wb.handle, 0, record_arg(MakeValidation(&range, 0, nullptr, nullptr)), ignored_index()),
      0);
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
  std::size_t index = 0;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, record_arg(MakeCfRule(&sqref, "=$A$1>+1.50", "{1, 2}")), &index), 0);
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule(wb.handle, 0, index, &rule), 0);
  EXPECT_STREQ(rule.formula1, "=$A$1>1.5");
  EXPECT_STREQ(rule.formula2, "={1,2}");

  const fm_cell_rect_t range{0, 0, 9, 0};
  constexpr std::uint8_t kList = 3;
  constexpr std::uint8_t kCustom = 7;
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, record_arg(MakeValidation(&range, kList, "\"1.50,1E3,+5\"", nullptr)),
                                    ignored_index()),
            0);
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, record_arg(MakeValidation(&range, kList, "+$A$1:$A$3", nullptr)),
                                    ignored_index()),
            0);
  ASSERT_EQ(fm_sheet_add_validation(
                wb.handle, 0, record_arg(MakeValidation(&range, kList, "OFFSET($A$1,+0,-0,1E3/1E3+2,1)", nullptr)),
                ignored_index()),
            0);
  ASSERT_EQ(
      fm_sheet_add_validation(
          wb.handle, 0, record_arg(MakeValidation(&range, kCustom, "SUM(A1 : A2)>-1E-309", nullptr)), ignored_index()),
      0);
  const char* const expected[] = {"=\"1.50,1E3,+5\"", "=+$A$1:$A$3", "=OFFSET($A$1,0,0,1000/1000+2,1)",
                                  "=SUM(A1:A2)>0"};
  for (std::uint32_t i = 0; i < 4U; ++i) {
    fm_data_validation_t v = FM_STRUCT_INIT(fm_data_validation_t);
    ASSERT_EQ(fm_sheet_get_validation(wb.handle, 0, i, &v), 0);
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
  ASSERT_EQ(fm_workbook_save(source.handle, &bytes, &len), 0) << fm_get_last_error_message();
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
    const char* r1c1 = nullptr;
    EXPECT_EQ(fm_sheet_get_formula_r1c1(loaded.handle, 0, row, 1, &r1c1), kParserUnexpectedToken);
  }
}

TEST(FormulonCApiFormulaEntry, LoadedWrongArityFormulaKeepsEvaluatingButIsRefusedOnEntry) {
  WorkbookGuard source;
  MakeWorkbook(&source);
  ASSERT_TRUE(source.handle->workbook().set_cell_formula(0, 0, 1, "=ROWS()"));

  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ASSERT_EQ(fm_workbook_save(source.handle, &bytes, &len), 0) << fm_get_last_error_message();
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

}  // namespace
