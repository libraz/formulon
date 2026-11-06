//
// Stable C ABI tests for the external-links surface
// (`fm_workbook_get_external_link_count` / `fm_workbook_get_external_link`).
// The reader/writer round-trip lives in the integration suite; these
// tests exercise the boundary's null / range guards and the empty-
// workbook behaviour.

#include <cstddef>
#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "formulon_c_out_index.h"
#include "formulon_c_record_arg.h"
#include "formulon_c_setters.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

uint32_t LinkCount(fm_workbook_t* wb) {
  std::size_t count = 99;
  EXPECT_EQ(fm_workbook_get_external_link_count(wb, &count), 0);
  return static_cast<uint32_t>(count);
}

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);

std::string FileName(fm_workbook_t* wb) {
  const char* name = nullptr;
  EXPECT_EQ(fm_workbook_get_file_name(wb, &name), 0);
  return name != nullptr ? name : "<null>";
}

std::string Formula(fm_workbook_t* wb, uint32_t row, uint32_t col) {
  const char* text = nullptr;
  EXPECT_EQ(fm_sheet_get_formula(wb, 0, row, col, &text), 0) << fm_get_last_error_message();
  return text != nullptr ? text : "<null>";
}

std::string Str(const char* text) {
  return text != nullptr ? text : "<null>";
}

// A workbook with the workbook-scope name `LN` = 10.
void CreateWithName(WorkbookGuard* wb) {
  ASSERT_EQ(fm_workbook_create_default(&wb->handle), 0);
  ASSERT_EQ(SetDefinedName(wb->handle, "LN", "10"), 0) << fm_get_last_error_message();
}

}  // namespace

TEST(FormulonCApiExternalLinks, FreshWorkbookHasZeroLinks) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t count = 99;
  EXPECT_EQ(fm_workbook_get_external_link_count(wb.handle, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiExternalLinks, IndexOutOfRangeReturnsInvalidArgument) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_external_link_record_t rec{};
  EXPECT_EQ(fm_workbook_get_external_link(wb.handle, 0, &rec),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kIndexOutOfRange));
}

TEST(FormulonCApiExternalLinks, NullArgsReturnBindingNullPointer) {
  fm_external_link_record_t rec{};
  std::size_t count = 0;
  EXPECT_EQ(fm_workbook_get_external_link_count(nullptr, &count),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_get_external_link(nullptr, 0, &rec),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_workbook_get_external_link_count(wb.handle, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_get_external_link(wb.handle, 0, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(FormulonCApiExternalLinks, RoundTripPreservesZeroCount) {
  // No setter on the C ABI today; this guards the empty-default
  // round-trip the way the named-styles tests do. Reader-side records
  // surface through the integration suite.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  uint8_t* saved_data = nullptr;
  size_t saved_len = 0;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved_data, &saved_len), 0);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved_data, saved_len, &wb2.handle), 0);
  fm_buffer_destroy(saved_data);

  std::size_t count = 99;
  EXPECT_EQ(fm_workbook_get_external_link_count(wb2.handle, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiExternalLinks, SetFormulaBindsANewBookOnce) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=[Book.xlsx]Sheet1!A1"), 0) << fm_get_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 1U);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 1, 0, "=[BOOK.XLSX]Sheet1!B2"), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 2, 0, "=[book.xlsx]Other!A1"), 0) << fm_get_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 1U);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 3, 0, "=[Second.xlsx]Sheet1!A1"), 0) << fm_get_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 2U);
}

TEST(FormulonCApiExternalLinks, CfRuleNamingANewBookBindsIt) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t sqref{0, 0, 9, 0};
  fm_cf_rule_t rule{};
  rule.type = 0;  // Expression
  rule.formula1 = "[Book.xlsx]Sheet1!$A$1>0";
  rule.ranges = &sqref;
  rule.range_count = 1;
  std::size_t index = 0;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &index), 0) << fm_get_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 1U);
  rule.formula1 = "[BOOK.xlsx]Sheet1!$A$2>0";
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &index), 0) << fm_get_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 1U);
  rule.formula1 = "A1>0";
  rule.formula2 = "[Other.xlsx]Sheet1!$A$1";
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &index), 0) << fm_get_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 2U);
}

TEST(FormulonCApiExternalLinks, ValidationNamingANewBookBindsIt) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t range{0, 0, 9, 0};
  fm_data_validation_t v{};
  v.ranges = &range;
  v.range_count = 1;
  v.type = 3;  // list
  v.formula1 = "[Book.xlsx]Sheet1!$A$1:$A$5";
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, &v, ignored_index()), 0) << fm_get_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 1U);
  v.formula1 = "[book.XLSX]Sheet1!$B$1:$B$5";
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, &v, ignored_index()), 0) << fm_get_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 1U);
  v.type = 1;
  v.formula1 = "0";
  v.formula2 = "[Other.xlsx]Sheet1!$A$1";
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, &v, ignored_index()), 0) << fm_get_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 2U);
}

TEST(FormulonCApiExternalLinks, FileNameNullArgumentsAreRefused) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const char* name = nullptr;
  EXPECT_EQ(fm_workbook_get_file_name(nullptr, &name), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_get_file_name(wb.handle, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_set_file_name(nullptr, "M5.xlsx"), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_set_file_name(wb.handle, nullptr), kBindingNullPointer);
}

TEST(FormulonCApiExternalLinks, FileNameIsUnsetByDefaultAndByEmptyText) {
  WorkbookGuard wb;
  CreateWithName(&wb);
  EXPECT_EQ(FileName(wb.handle), "");
  ASSERT_EQ(fm_workbook_set_file_name(wb.handle, "M5.xlsx"), 0) << fm_get_last_error_message();
  EXPECT_EQ(FileName(wb.handle), "M5.xlsx");
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=M5.xlsx!LN+1"), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_file_name(wb.handle, ""), 0);
  EXPECT_EQ(FileName(wb.handle), "");
  EXPECT_EQ(Formula(wb.handle, 0, 0), "=[0]!LN+1");
  EXPECT_EQ(LinkCount(wb.handle), 0U);
}

TEST(FormulonCApiExternalLinks, FileNameRefusesInvalidNames) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_file_name(wb.handle, "Keep.xlsx"), 0);
  const std::string too_long = std::string(214, 'a') + ".xlsx";
  for (const char* name : {"Book1", ".xlsx", "Book1.txt", "a[b].xlsx", "a]b.xlsx", "a:b.xlsx", "a/b.xlsx", "a\\b.xlsx",
                           "a?b.xlsx", "a*b.xlsx", "\xFF.xlsx", "\xED\xA0\x80.xlsx", too_long.c_str()}) {
    EXPECT_EQ(fm_workbook_set_file_name(wb.handle, name), kInvalidArgument) << name;
  }
  EXPECT_EQ(FileName(wb.handle), "Keep.xlsx");
}

TEST(FormulonCApiExternalLinks, FileNameAcceptsTwoHundredEighteenBytes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::string name = std::string(213, 'a') + ".xlsx";
  ASSERT_EQ(name.size(), 218U);
  EXPECT_EQ(fm_workbook_set_file_name(wb.handle, name.c_str()), 0) << fm_get_last_error_message();
  EXPECT_EQ(FileName(wb.handle), name);
}

TEST(FormulonCApiExternalLinks, FileNameRefusesAnExternalLinkBookName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=[Ext.xlsx]Sheet1!A1"), 0) << fm_get_last_error_message();
  ASSERT_EQ(LinkCount(wb.handle), 1U);
  EXPECT_EQ(fm_workbook_set_file_name(wb.handle, "EXT.XLSX"), kInvalidArgument);
  EXPECT_EQ(FileName(wb.handle), "");
  EXPECT_EQ(fm_workbook_set_file_name(wb.handle, "Other.xlsx"), 0) << fm_get_last_error_message();
}

TEST(FormulonCApiExternalLinks, FileNameSpellsTheWorkbooksOwnNames) {
  WorkbookGuard wb;
  CreateWithName(&wb);
  ASSERT_EQ(fm_workbook_set_file_name(wb.handle, "M5.xlsx"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=M5.xlsx!LN+1"), 0) << fm_get_last_error_message();
  EXPECT_EQ(Formula(wb.handle, 0, 0), "=M5.xlsx!LN+1");
  const struct {
    const char* name;
    const char* shown;
  } kCases[] = {
      {"M5.XLSX", "=M5.XLSX!LN+1"},        {"M5 b-1.xlsx", "='M5 b-1.xlsx'!LN+1"}, {"1M5.xlsx", "='1M5.xlsx'!LN+1"},
      {"It's.xlsx", "='It''s.xlsx'!LN+1"}, {"M5x.xlsb", "=M5x.xlsb!LN+1"},
  };
  for (const auto& c : kCases) {
    ASSERT_EQ(fm_workbook_set_file_name(wb.handle, c.name), 0) << c.name;
    EXPECT_EQ(Formula(wb.handle, 0, 0), c.shown) << c.name;
  }
  EXPECT_EQ(LinkCount(wb.handle), 0U);
}

TEST(FormulonCApiExternalLinks, FileNameQuotedEntriesBindToTheWorkbook) {
  WorkbookGuard wb;
  CreateWithName(&wb);
  ASSERT_EQ(fm_workbook_set_file_name(wb.handle, "M5 b-1.xlsx"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "='M5 b-1.xlsx'!LN+1"), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_file_name(wb.handle, "1M5.xlsx"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 1, 0, "='1M5.xlsx'!LN+1"), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_file_name(wb.handle, "M5.xlsx"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 2, 0, "=m5.XLSX!LN+1"), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_file_name(wb.handle, ""), 0);
  EXPECT_EQ(Formula(wb.handle, 0, 0), "=[0]!LN+1");
  EXPECT_EQ(Formula(wb.handle, 1, 0), "=[0]!LN+1");
  EXPECT_EQ(Formula(wb.handle, 2, 0), "=[0]!LN+1");
  EXPECT_EQ(LinkCount(wb.handle), 0U);
}

TEST(FormulonCApiExternalLinks, FileNameBracketedSelfBookDropsToTheSheet) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_file_name(wb.handle, "M5.xlsx"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 7.0), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 1, 0, "=[M5.xlsx]Sheet1!A1"), 0) << fm_get_last_error_message();
  EXPECT_EQ(Formula(wb.handle, 1, 0), "=Sheet1!A1");
  EXPECT_EQ(LinkCount(wb.handle), 0U);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 1, 0, &v), 0);
  ASSERT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_EQ(v.u.number, 7.0);
}

TEST(FormulonCApiExternalLinks, FileNameNonAsciiNameRoundTrips) {
  WorkbookGuard wb;
  CreateWithName(&wb);
  ASSERT_EQ(fm_workbook_set_file_name(wb.handle, "\xE3\x83\x96\xE3\x83\x83\xE3\x82\xAF.xlsx"), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=\xE3\x83\x96\xE3\x83\x83\xE3\x82\xAF.xlsx!LN+1"), 0)
      << fm_get_last_error_message();
  const std::string shown = Formula(wb.handle, 0, 0);
  EXPECT_EQ(shown.find("[0]!"), std::string::npos) << shown;
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 1, 0, shown.c_str()), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_file_name(wb.handle, ""), 0);
  EXPECT_EQ(Formula(wb.handle, 0, 0), "=[0]!LN+1");
  EXPECT_EQ(Formula(wb.handle, 1, 0), "=[0]!LN+1");
  EXPECT_EQ(LinkCount(wb.handle), 0U);
}

TEST(FormulonCApiExternalLinks, FileNameCellGettersSpellTheName) {
  WorkbookGuard wb;
  CreateWithName(&wb);
  ASSERT_EQ(fm_workbook_set_file_name(wb.handle, "M5.xlsx"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=M5.xlsx!LN+1"), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_set_array_formula(wb.handle, 0, record_arg(fm_cell_rect_t{2, 0, 2U, 0U}), "=M5.xlsx!LN*2"), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 0U);

  for (const bool named : {true, false}) {
    ASSERT_EQ(fm_workbook_set_file_name(wb.handle, named ? "M5.xlsx" : ""), 0);
    const std::string book = named ? "M5.xlsx!" : "[0]!";
    EXPECT_EQ(Formula(wb.handle, 0, 0), "=" + book + "LN+1");

    const char* r1c1 = nullptr;
    ASSERT_EQ(fm_sheet_get_formula_r1c1(wb.handle, 0, 0, 0, &r1c1), 0) << fm_get_last_error_message();
    EXPECT_EQ(Str(r1c1), book + "LN+1");

    uint32_t row = 99;
    uint32_t col = 99;
    const char* formula = nullptr;
    fm_value_t v{};
    ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 0, 0, &row, &col, &formula, &v), 0) << fm_get_last_error_message();
    EXPECT_EQ(row, 0U);
    EXPECT_EQ(Str(formula), "=" + book + "LN+1");

    fm_cell_range_t* range = nullptr;
    ASSERT_EQ(fm_sheet_get_cells_in_range(wb.handle, 0, record_arg(fm_cell_rect_t{0, 0, 9, 0}), 0, 0, &range), 0)
        << fm_get_last_error_message();
    ASSERT_EQ(fm_cell_range_get_cell_at(range, 0, &row, &col, &formula, &v), 0);
    EXPECT_EQ(Str(formula), "=" + book + "LN+1");
    fm_cell_range_destroy(range);

    fm_array_formula_t array{};
    ASSERT_EQ(fm_sheet_get_array_formula(wb.handle, 0, 2, 0, &array), 0) << fm_get_last_error_message();
    ASSERT_EQ(array.engaged, 1);
    EXPECT_EQ(Str(array.formula), "=" + book + "LN*2");
  }
}

TEST(FormulonCApiExternalLinks, FileNameDefinedNamesStoreAndSpellTheName) {
  WorkbookGuard wb;
  CreateWithName(&wb);
  ASSERT_EQ(fm_workbook_set_file_name(wb.handle, "M5.xlsx"), 0);
  ASSERT_EQ(SetDefinedName(wb.handle, "LP", "M5.xlsx!LN*2"), 0) << fm_get_last_error_message();
  ASSERT_EQ(SetDefinedName(wb.handle, "LQ", "=M5.xlsx!LN*3", 0), 0) << fm_get_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 0U);
  for (const bool named : {true, false}) {
    ASSERT_EQ(fm_workbook_set_file_name(wb.handle, named ? "M5.xlsx" : ""), 0);
    const std::string book = named ? "M5.xlsx!" : "[0]!";
    fm_defined_name_t dn{};
    dn.local_sheet_id = -2;
    ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 1, &dn), 0) << fm_get_last_error_message();
    EXPECT_EQ(Str(dn.name), "LP");
    EXPECT_EQ(Str(dn.formula), book + "LN*2");
    ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 2, &dn), 0) << fm_get_last_error_message();
    EXPECT_EQ(Str(dn.name), "LQ");
    EXPECT_EQ(dn.local_sheet_id, 0);
    EXPECT_EQ(Str(dn.formula), book + "LN*3");
  }
}

TEST(FormulonCApiExternalLinks, FileNameValidationsStoreAndSpellTheName) {
  WorkbookGuard wb;
  CreateWithName(&wb);
  ASSERT_EQ(fm_workbook_set_file_name(wb.handle, "M5.xlsx"), 0);
  const fm_cell_rect_t range{0, 0, 9, 0};
  fm_data_validation_t v{};
  v.ranges = &range;
  v.range_count = 1;
  v.type = 2;  // decimal, between
  v.formula1 = "M5.xlsx!LN";
  v.formula2 = "=M5.xlsx!LN*2";
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, &v, ignored_index()), 0) << fm_get_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 0U);
  for (const bool named : {true, false}) {
    ASSERT_EQ(fm_workbook_set_file_name(wb.handle, named ? "M5.xlsx" : ""), 0);
    const std::string book = named ? "M5.xlsx!" : "[0]!";
    fm_data_validation_t out{};
    ASSERT_EQ(fm_sheet_get_validation(wb.handle, 0, 0, &out), 0) << fm_get_last_error_message();
    EXPECT_EQ(Str(out.formula1), book + "LN");
    EXPECT_EQ(Str(out.formula2), book + "LN*2");
  }
}

TEST(FormulonCApiExternalLinks, FileNameConditionalFormatsStoreAndSpellTheName) {
  WorkbookGuard wb;
  CreateWithName(&wb);
  ASSERT_EQ(fm_workbook_set_file_name(wb.handle, "M5.xlsx"), 0);
  const fm_cell_rect_t sqref{0, 0, 9, 0};
  fm_cf_rule_t rule{};
  rule.type = 0;  // Expression
  rule.formula1 = "M5.xlsx!LN>0";
  rule.ranges = &sqref;
  rule.range_count = 1;
  std::size_t index = 0;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &index), 0) << fm_get_last_error_message();
  fm_cfvo_t thresholds[2]{};
  thresholds[0].type = 5;  // Formula
  thresholds[0].value = "M5.xlsx!LN";
  thresholds[0].gte = 1;
  thresholds[1].type = 4;  // Max
  thresholds[1].gte = 1;
  fm_cf_color_t colors[2]{{255, 0, 0, 255}, {0, 255, 0, 255}};
  fm_cf_rule_t scale{};
  scale.type = 2;  // ColorScale
  scale.ranges = &sqref;
  scale.range_count = 1;
  scale.color_scale_thresholds = thresholds;
  scale.color_scale_colors = colors;
  scale.color_scale_count = 2;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &scale, &index), 0) << fm_get_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 0U);
  for (const bool named : {true, false}) {
    ASSERT_EQ(fm_workbook_set_file_name(wb.handle, named ? "M5.xlsx" : ""), 0);
    const std::string book = named ? "M5.xlsx!" : "[0]!";
    fm_cf_rule_t out{};
    ASSERT_EQ(fm_sheet_get_cf_rule(wb.handle, 0, 0, &out), 0) << fm_get_last_error_message();
    EXPECT_EQ(Str(out.formula1), book + "LN>0");
    ASSERT_EQ(fm_sheet_get_cf_rule(wb.handle, 0, 1, &out), 0) << fm_get_last_error_message();
    ASSERT_EQ(out.color_scale_count, 2U);
    EXPECT_EQ(Str(out.color_scale_thresholds[0].value), book + "LN");
  }
}

TEST(FormulonCApiExternalLinks, FileNameTableTotalsStoreAndSpellTheName) {
  WorkbookGuard wb;
  CreateWithName(&wb);
  ASSERT_EQ(fm_workbook_set_file_name(wb.handle, "M5.xlsx"), 0);
  const char* names[] = {"a", "b"};
  std::size_t table = 99;
  ASSERT_EQ(fm_workbook_add_table(wb.handle, 0, record_arg(fm_cell_rect_t{0, 3, 3, 4}), "T1", "T1", names, 2, nullptr,
                                  1, 1, &table),
            0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_table_column_totals(wb.handle, table, 0, "custom", nullptr, "M5.xlsx!LN+1"), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 0U);
  for (const bool named : {true, false}) {
    ASSERT_EQ(fm_workbook_set_file_name(wb.handle, named ? "M5.xlsx" : ""), 0);
    const std::string book = named ? "M5.xlsx!" : "[0]!";
    fm_table_column_info_t column{};
    ASSERT_EQ(fm_workbook_get_table_column(wb.handle, table, 0, &column), 0) << fm_get_last_error_message();
    EXPECT_EQ(Str(column.totals_row_formula), book + "LN+1");
    EXPECT_EQ(Formula(wb.handle, 3, 3), "=" + book + "LN+1");
  }
}

TEST(FormulonCApiExternalLinks, FileNameChartSeriesStoreAndSpellTheName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 1, "S1"), 0);
  for (uint32_t r = 1; r <= 4; ++r) {
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, r, 0, "c"), 0);
    ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, r, 1, static_cast<double>(r)), 0);
    ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, r, 2, static_cast<double>(r * 2)), 0);
  }
  ASSERT_EQ(SetDefinedName(wb.handle, "LR", "Sheet1!$C$2:$C$5"), 0) << fm_get_last_error_message();
  fm_chart_insert_t opts{};
  opts.kind = FM_CHART_INSERT_COLUMN;
  opts.source = "Sheet1!$A$1:$B$5";
  opts.legend = FM_CHART_LEGEND_RIGHT;
  opts.placement.anchor.anchor_kind = FM_ANCHOR_KIND_TWO_CELL;
  opts.placement.anchor.edit_as = FM_ANCHOR_EDIT_AS_TWO_CELL;
  opts.placement.anchor.row = 6;
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_chart(wb.handle, 0, &opts, &id), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_file_name(wb.handle, "M5.xlsx"), 0);
  fm_chart_series_patch_t patch{};
  patch.values = "M5.xlsx!LR";
  ASSERT_EQ(fm_sheet_set_chart_series(wb.handle, 0, id, 0, 0, &patch), 0) << fm_get_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 0U);
  for (const bool named : {true, false}) {
    ASSERT_EQ(fm_workbook_set_file_name(wb.handle, named ? "M5.xlsx" : ""), 0);
    fm_chart_t* chart = nullptr;
    int32_t present = 0;
    ASSERT_EQ(fm_sheet_get_chart(wb.handle, 0, id, &present, &chart), 0) << fm_get_last_error_message();
    ASSERT_EQ(present, 1);
    fm_chart_series_t series{};
    EXPECT_EQ(fm_chart_get_series_at(chart, 0, 0, &series), 0) << fm_get_last_error_message();
    EXPECT_EQ(Str(series.values.formula), std::string(named ? "M5.xlsx!" : "[0]!") + "LR");
    fm_chart_destroy(chart);
  }
}

TEST(FormulonCApiExternalLinks, FileNameAdHocEvaluationReadsTheName) {
  WorkbookGuard wb;
  CreateWithName(&wb);
  ASSERT_EQ(fm_workbook_set_file_name(wb.handle, "M5.xlsx"), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_sheet_evaluate_formula(wb.handle, 0, 0, 0, "=M5.xlsx!LN*3", &v), 0) << fm_get_last_error_message();
  ASSERT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_EQ(v.u.number, 30.0);
  ASSERT_EQ(fm_sheet_evaluate_cf_formula(wb.handle, 0, 0, 0, 0, 0, "M5.xlsx!LN>5", &v), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(v.kind, FM_VAL_BOOL);
  EXPECT_EQ(v.u.boolean, 1);
  EXPECT_EQ(LinkCount(wb.handle), 0U);
}

TEST(FormulonCApiExternalLinks, FileNameFormulaTextSpellsTheName) {
  WorkbookGuard wb;
  CreateWithName(&wb);
  ASSERT_EQ(fm_workbook_set_file_name(wb.handle, "M10a.xlsx"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=M10a.xlsx!LN+1"), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 1, "=FORMULATEXT(A1)"), 0) << fm_get_last_error_message();
  for (const bool named : {true, false}) {
    ASSERT_EQ(fm_workbook_set_file_name(wb.handle, named ? "M10a.xlsx" : ""), 0);
    ASSERT_EQ(fm_workbook_recalc(wb.handle), 0) << fm_get_last_error_message();
    fm_value_t v{};
    ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 0, 1, &v), 0);
    ASSERT_EQ(v.kind, FM_VAL_TEXT);
    EXPECT_EQ(Str(v.u.text), std::string(named ? "=M10a.xlsx!" : "=[0]!") + "LN+1");
  }
}
