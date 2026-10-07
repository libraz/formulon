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
#include "gtest/gtest.h"
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
  ASSERT_EQ(fm_workbook_create(&wb->handle), 0);
  for (const char* name : {"Data", "S2"}) {
    ASSERT_EQ(fm_workbook_add_sheet(wb->handle, name), 0) << fm_last_error_message();
  }
}

const std::vector<std::string>& RejectedFormulas() {
  static const std::vector<std::string> formulas = {
      "=TRUE!A1", "=SUM(", "=1+", "=Data!SUM(A1:A2)", "=SUM(Data:Sheet1!A1", "=(1+2", "SUM(",
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
  };
  return formulas;
}

std::string GetFormula(const fm_workbook_t* wb, std::uint32_t row, std::uint32_t col) {
  const char* text = nullptr;
  EXPECT_EQ(fm_workbook_get_formula(wb, 0, row, col, &text), 0);
  return text != nullptr ? text : "<null>";
}

TEST(FormulonCApiFormulaEntry, SetFormulaRejectsUnparseableTextAndKeepsState) {
  WorkbookGuard wb;
  MakeWorkbook(&wb);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 1, "=1+2"), 0);
  for (const std::string& formula : RejectedFormulas()) {
    SCOPED_TRACE(formula);
    EXPECT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 1, formula.c_str()), kParserUnexpectedToken);
    EXPECT_NE(std::string(fm_last_error_message()).find("fm_workbook_set_formula: formula does not parse"),
              std::string::npos);
    EXPECT_NE(std::string(fm_last_error_context()).find("row=0 col=1"), std::string::npos);
    EXPECT_EQ(GetFormula(wb.handle, 0, 1), "=1+2");

    EXPECT_EQ(fm_workbook_set_formula(wb.handle, 0, 5, 5, formula.c_str()), kParserUnexpectedToken);
    EXPECT_EQ(GetFormula(wb.handle, 5, 5), "");
  }
}

TEST(FormulonCApiFormulaEntry, SetFormulaAcceptedTextAlwaysHasAnR1C1) {
  WorkbookGuard wb;
  MakeWorkbook(&wb);
  for (const std::string& formula : AcceptedFormulas()) {
    SCOPED_TRACE(formula);
    ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 2, 2, formula.c_str()), 0) << fm_last_error_message();
    const char* r1c1 = nullptr;
    EXPECT_EQ(fm_workbook_get_formula_r1c1(wb.handle, 0, 2, 2, &r1c1), 0) << fm_last_error_message();
  }
}

TEST(FormulonCApiFormulaEntry, DefinedNameRejectsUnparseableTextAndKeepsState) {
  WorkbookGuard wb;
  MakeWorkbook(&wb);
  ASSERT_EQ(fm_workbook_set_defined_name(wb.handle, "Rate", "Sheet1!$A$1"), 0);
  ASSERT_EQ(fm_workbook_set_defined_name_scoped(wb.handle, "Local", "Sheet1!$B$1", 0), 0);
  ASSERT_EQ(fm_workbook_defined_name_count(wb.handle), 2U);
  for (const std::string& formula : RejectedFormulas()) {
    SCOPED_TRACE(formula);
    EXPECT_EQ(fm_workbook_set_defined_name(wb.handle, "Rate", formula.c_str()), kParserUnexpectedToken);
    EXPECT_NE(std::string(fm_last_error_message()).find("fm_workbook_set_defined_name: formula does not parse"),
              std::string::npos);
    EXPECT_EQ(fm_workbook_set_defined_name(wb.handle, "Fresh", formula.c_str()), kParserUnexpectedToken);
    EXPECT_EQ(fm_workbook_set_defined_name_scoped(wb.handle, "Local", formula.c_str(), 0), kParserUnexpectedToken);
    EXPECT_NE(std::string(fm_last_error_message()).find("fm_workbook_set_defined_name_scoped: formula does not parse"),
              std::string::npos);
    EXPECT_EQ(fm_workbook_set_defined_name_scoped(wb.handle, "Fresh", formula.c_str(), -1), kParserUnexpectedToken);
  }
  ASSERT_EQ(fm_workbook_defined_name_count(wb.handle), 2U);
  const char* name = nullptr;
  const char* formula = nullptr;
  int32_t scope = 0;
  ASSERT_EQ(fm_workbook_defined_name_at(wb.handle, 0, &name, &formula, &scope), 0);
  EXPECT_STREQ(formula, "Sheet1!$A$1");
  ASSERT_EQ(fm_workbook_defined_name_at(wb.handle, 1, &name, &formula, &scope), 0);
  EXPECT_STREQ(formula, "Sheet1!$B$1");
}

TEST(FormulonCApiFormulaEntry, DefinedNameAcceptsParseableAndEmptyText) {
  WorkbookGuard wb;
  MakeWorkbook(&wb);
  EXPECT_EQ(fm_workbook_set_defined_name(wb.handle, "Quoted", "'S2'!$A$1"), 0);
  EXPECT_EQ(fm_workbook_set_defined_name(wb.handle, "Missing", "Gone!$A$1"), 0);
  EXPECT_EQ(fm_workbook_set_defined_name(wb.handle, "Const", "1.5"), 0);
  EXPECT_EQ(fm_workbook_set_defined_name_scoped(wb.handle, "Local", "'S2'!$A$1", 1), 0);
  ASSERT_EQ(fm_workbook_defined_name_count(wb.handle), 4U);
  // An empty formula removes the entry rather than being parsed.
  EXPECT_EQ(fm_workbook_set_defined_name(wb.handle, "Const", ""), 0);
  EXPECT_EQ(fm_workbook_set_defined_name_scoped(wb.handle, "Local", "", 1), 0);
  EXPECT_EQ(fm_workbook_defined_name_count(wb.handle), 2U);
}

fm_cf_rule_t MakeCfRule(const fm_cf_cell_range_t* sqref, const char* formula1, const char* formula2) {
  fm_cf_rule_t rule{};
  rule.type = 0;  // Expression
  rule.formula1 = formula1;
  rule.formula2 = formula2;
  rule.sqref = sqref;
  rule.sqref_count = 1;
  return rule;
}

TEST(FormulonCApiFormulaEntry, CfAddRuleRejectsUnparseableFormulasAndKeepsState) {
  WorkbookGuard wb;
  MakeWorkbook(&wb);
  const fm_cf_cell_range_t sqref{0, 0, 9, 0};
  std::size_t index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, MakeCfRule(&sqref, "A1>0", nullptr), &index), 0);
  for (const std::string& formula : RejectedFormulas()) {
    SCOPED_TRACE(formula);
    EXPECT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, MakeCfRule(&sqref, formula.c_str(), nullptr), &index),
              kParserUnexpectedToken);
    EXPECT_NE(std::string(fm_last_error_message()).find("fm_sheet_cf_add_rule: formula does not parse"),
              std::string::npos);
    EXPECT_NE(std::string(fm_last_error_context()).find("formula1"), std::string::npos);
    EXPECT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, MakeCfRule(&sqref, "A1>0", formula.c_str()), &index),
              kParserUnexpectedToken);
    EXPECT_NE(std::string(fm_last_error_context()).find("formula2"), std::string::npos);
  }
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_cf_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);
}

TEST(FormulonCApiFormulaEntry, CfAddRuleAcceptsParseableAndEmptyFormulas) {
  WorkbookGuard wb;
  MakeWorkbook(&wb);
  const fm_cf_cell_range_t sqref{0, 0, 9, 0};
  std::size_t index = 0;
  EXPECT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, MakeCfRule(&sqref, "'S2'!$A$1>0", ""), &index), 0);
  EXPECT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, MakeCfRule(&sqref, "Gone!$A$1>0", nullptr), &index), 0);
  EXPECT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, MakeCfRule(&sqref, "", ""), &index), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_cf_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 3U);
}

fm_data_validation MakeValidation(const fm_merge_range* range, std::uint8_t type, const char* formula1,
                                  const char* formula2) {
  fm_data_validation v{};
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
  const fm_merge_range range{0, 0, 9, 0};
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, MakeValidation(&range, 1, "0", "10")), 0);
  for (const std::string& formula : RejectedFormulas()) {
    SCOPED_TRACE(formula);
    EXPECT_EQ(fm_sheet_add_validation(wb.handle, 0, MakeValidation(&range, 1, formula.c_str(), "10")),
              kParserUnexpectedToken);
    EXPECT_NE(std::string(fm_last_error_message()).find("fm_sheet_add_validation: formula does not parse"),
              std::string::npos);
    EXPECT_NE(std::string(fm_last_error_context()).find("formula1"), std::string::npos);
    EXPECT_EQ(fm_sheet_add_validation(wb.handle, 0, MakeValidation(&range, 1, "0", formula.c_str())),
              kParserUnexpectedToken);
    EXPECT_NE(std::string(fm_last_error_context()).find("formula2"), std::string::npos);
  }
  std::uint32_t count = 0;
  ASSERT_EQ(fm_sheet_get_validation_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);
}

TEST(FormulonCApiFormulaEntry, AddValidationAcceptsEveryListAndScalarShape) {
  WorkbookGuard wb;
  MakeWorkbook(&wb);
  const fm_merge_range range{0, 0, 9, 0};
  constexpr std::uint8_t kList = 3;
  EXPECT_EQ(fm_sheet_add_validation(wb.handle, 0, MakeValidation(&range, kList, "\"a,b,c\"", nullptr)), 0);
  EXPECT_EQ(fm_sheet_add_validation(wb.handle, 0, MakeValidation(&range, kList, "'S2'!$A$1:$A$5", nullptr)), 0);
  EXPECT_EQ(fm_sheet_add_validation(wb.handle, 0, MakeValidation(&range, kList, "Gone!$A$1:$A$5", nullptr)), 0);
  EXPECT_EQ(fm_sheet_add_validation(wb.handle, 0, MakeValidation(&range, kList, "ListName", nullptr)), 0);
  EXPECT_EQ(fm_sheet_add_validation(wb.handle, 0, MakeValidation(&range, 1, "-3.5", "TODAY()")), 0);
  EXPECT_EQ(fm_sheet_add_validation(wb.handle, 0, MakeValidation(&range, 7, "ISNUMBER(A1)", "")), 0);
  EXPECT_EQ(fm_sheet_add_validation(wb.handle, 0, MakeValidation(&range, 0, nullptr, nullptr)), 0);
  std::uint32_t count = 0;
  ASSERT_EQ(fm_sheet_get_validation_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 7U);
}

TEST(FormulonCApiFormulaEntry, LoadedUnparseableFormulaIsKeptVerbatim) {
  WorkbookGuard source;
  MakeWorkbook(&source);
  // Workbook::set_cell_formula is the reader's ingestion point and stores text as given.
  ASSERT_TRUE(source.handle->workbook().set_cell_formula(0, 0, 1, "=TRUE!A1"));
  ASSERT_TRUE(source.handle->workbook().set_cell_formula(0, 1, 1, "=SUM("));

  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ASSERT_EQ(fm_workbook_save(source.handle, &bytes, &len), 0) << fm_last_error_message();
  WorkbookGuard loaded;
  const fm_status_t rc = fm_workbook_load(bytes, len, &loaded.handle);
  fm_buffer_free(bytes);
  ASSERT_EQ(rc, 0) << fm_last_error_message();

  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0);
  for (const auto& [row, text] : {std::pair<std::uint32_t, const char*>{0, "=TRUE!A1"}, {1, "=SUM("}}) {
    SCOPED_TRACE(text);
    EXPECT_EQ(GetFormula(loaded.handle, row, 1), text);
    fm_value_t value{};
    ASSERT_EQ(fm_workbook_get_value(loaded.handle, 0, row, 1, &value), 0);
    ASSERT_EQ(value.kind, FM_VAL_ERROR);
    EXPECT_EQ(value.u.error_code, 4);  // #NAME?
    const char* r1c1 = nullptr;
    EXPECT_EQ(fm_workbook_get_formula_r1c1(loaded.handle, 0, row, 1, &r1c1), kParserUnexpectedToken);
  }
}

}  // namespace
