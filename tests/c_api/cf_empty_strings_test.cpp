// Empty-value rules of the conditional-format rule record (api-conventions §5 rule 30): `""` is refused for the id, the
// two formulas and a threshold value, and is a value for the text of a text rule, as Excel stores it.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "io/zip_reader.h"
#include "utils/error.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

using Bytes = std::vector<std::uint8_t>;

struct CfResultsGuard {
  fm_cf_results_t* handle = nullptr;
  ~CfResultsGuard() { fm_cf_results_destroy(handle); }
  CfResultsGuard() = default;
  CfResultsGuard(const CfResultsGuard&) = delete;
  CfResultsGuard& operator=(const CfResultsGuard&) = delete;
};

const fm_cell_rect_t kRange{0, 0, 3, 0};

fm_cf_rule_t Rule(int32_t type) {
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.type = type;
  rule.ranges = &kRange;
  rule.range_count = 1;
  return rule;
}

std::size_t RuleCount(const WorkbookGuard& wb) {
  std::size_t count = 99;
  EXPECT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  return count;
}

fm_cf_rule_t ReadAt(const WorkbookGuard& wb, std::size_t index) {
  fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
  EXPECT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, index, &out), 0) << fm_get_last_error_message();
  return out;
}

Bytes Save(const WorkbookGuard& wb, int32_t format) {
  BufferGuard buffer;
  EXPECT_EQ(fm_workbook_save(wb.handle, format, &buffer.data, &buffer.len), 0);
  return Bytes(buffer.data, buffer.data + buffer.len);
}

void Load(const Bytes& bytes, WorkbookGuard& out) {
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), nullptr, &out.handle), 0) << fm_get_last_error_message();
}

std::string SheetXml(const Bytes& package) {
  formulon::io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(formulon::io::ByteSpan{package.data(), package.size()})));
  auto body = zip.read_entry("xl/worksheets/sheet1.xml");
  EXPECT_TRUE(static_cast<bool>(body));
  return body ? std::string(body.value().begin(), body.value().end()) : std::string();
}

// A1 "abc", A2 5, A3 empty, A4 an error: every kind of cell a text rule meets.
void FillCells(WorkbookGuard& wb) {
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, "abc"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 0, 5.0), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 3, 0, "=1/0"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
}

std::vector<std::uint32_t> MatchedRows(const WorkbookGuard& wb) {
  CfResultsGuard results;
  const fm_cell_rect_t range{0, 0, 3, 0};
  EXPECT_EQ(fm_sheet_evaluate_cf_range(wb.handle, 0, &range, nullptr, &results.handle), 0);
  std::size_t count = 0;
  EXPECT_EQ(fm_cf_results_get_cell_count(results.handle, &count), 0);
  std::vector<std::uint32_t> rows;
  for (std::size_t i = 0; i < count; ++i) {
    std::uint32_t row = 0;
    std::uint32_t col = 0;
    std::size_t matches = 0;
    EXPECT_EQ(fm_cf_results_get_cell_at(results.handle, i, &row, &col, &matches), 0);
    rows.push_back(row);
  }
  return rows;
}

// ---------------------------------------------------------------------------
// id, formula1, formula2: "" is refused, the absent value keeps its meaning
// ---------------------------------------------------------------------------

TEST(CfEmptyStrings, IdEmptyIsRefusedAndAbsentIsGenerated) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cf_rule_t rule = Rule(FM_CF_RULE_TYPE_EXPRESSION);
  rule.formula1 = "TRUE";
  const char* id = nullptr;

  rule.id = "";
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &id), kInvalidArgument);
  EXPECT_EQ(RuleCount(wb), 0U);

  rule.id = nullptr;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &id), 0) << fm_get_last_error_message();
  ASSERT_NE(id, nullptr);
  const std::string generated(id);
  EXPECT_EQ(generated.size(), 38U);
  EXPECT_EQ(generated.front(), '{');
  EXPECT_EQ(RuleCount(wb), 1U);
}

TEST(CfEmptyStrings, Formula1EmptyIsRefusedAndAbsentIsFilledOrMissing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const char* id = nullptr;

  fm_cf_rule_t expression = Rule(FM_CF_RULE_TYPE_EXPRESSION);
  expression.formula1 = "";
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &expression, &id), kInvalidArgument);

  fm_cf_rule_t cell_is = Rule(FM_CF_RULE_TYPE_CELL_IS);
  cell_is.has_op = 1;
  cell_is.op = FM_COMPARISON_OPERATOR_GREATER_THAN;
  cell_is.formula1 = "";
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &cell_is, &id), kInvalidArgument);

  fm_cf_rule_t text = Rule(FM_CF_RULE_TYPE_CONTAINS_TEXT);
  text.text = "ab";
  text.formula1 = "";
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &text, &id), kInvalidArgument);
  EXPECT_EQ(RuleCount(wb), 0U);

  // Absent, a text rule's formula is filled in as before.
  text.formula1 = nullptr;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &text, &id), 0) << fm_get_last_error_message();
  EXPECT_STREQ(ReadAt(wb, 0).formula1, "=NOT(ISERROR(SEARCH(\"ab\",A1)))");
}

TEST(CfEmptyStrings, Formula2EmptyIsRefusedAndAbsentIsAbsent) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cf_rule_t rule = Rule(FM_CF_RULE_TYPE_CELL_IS);
  rule.has_op = 1;
  rule.op = FM_COMPARISON_OPERATOR_BETWEEN;
  rule.formula1 = "1";
  rule.formula2 = "";
  const char* id = nullptr;
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &id), kInvalidArgument);
  EXPECT_EQ(RuleCount(wb), 0U);

  rule.has_op = 1;
  rule.op = FM_COMPARISON_OPERATOR_GREATER_THAN;
  rule.formula2 = nullptr;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &id), 0) << fm_get_last_error_message();
  EXPECT_EQ(ReadAt(wb, 0).formula2, nullptr);
}

TEST(CfEmptyStrings, UpdateRefusesEmptyFormulaAndChangesNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cf_rule_t rule = Rule(FM_CF_RULE_TYPE_CELL_IS);
  rule.has_op = 1;
  rule.op = FM_COMPARISON_OPERATOR_BETWEEN;
  rule.formula1 = "1";
  rule.formula2 = "9";
  const char* id = nullptr;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &id), 0) << fm_get_last_error_message();
  const std::string rule_id(id);

  fm_cf_rule_patch_t patch = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  patch.formula1_op = FM_PATCH_OP_SET;
  patch.formula1 = "";
  EXPECT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, rule_id.c_str(), &patch), kInvalidArgument);

  patch = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  patch.formula2_op = FM_PATCH_OP_SET;
  patch.formula2 = "";
  EXPECT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, rule_id.c_str(), &patch), kInvalidArgument);

  const fm_cf_rule_t out = ReadAt(wb, 0);
  EXPECT_STREQ(out.formula1, "=1");
  EXPECT_STREQ(out.formula2, "=9");
}

// ---------------------------------------------------------------------------
// Threshold value: "" is refused, NULL is the type that reads none
// ---------------------------------------------------------------------------

TEST(CfEmptyStrings, ThresholdValueEmptyIsRefused) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_color_t colors[2]{};
  colors[0].kind = FM_COLOR_KIND_RGB;
  colors[0].argb = 0xFFFF0000U;
  colors[1].kind = FM_COLOR_KIND_RGB;
  colors[1].argb = 0xFF00FF00U;
  fm_cfvo_t thresholds[2]{};
  thresholds[0].type = FM_CF_VALUE_OBJECT_TYPE_MIN;
  thresholds[1].type = FM_CF_VALUE_OBJECT_TYPE_NUMBER;
  thresholds[1].value = "";
  fm_cf_rule_t rule = Rule(FM_CF_RULE_TYPE_COLOR_SCALE);
  rule.color_scale_thresholds = thresholds;
  rule.color_scale_colors = colors;
  rule.color_scale_count = 2;
  const char* id = nullptr;
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &id), kInvalidArgument);
  EXPECT_EQ(RuleCount(wb), 0U);

  thresholds[1].value = "10";
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &id), 0) << fm_get_last_error_message();
  EXPECT_EQ(ReadAt(wb, 0).color_scale_thresholds[0].value, nullptr);

  fm_cf_rule_t bar = Rule(FM_CF_RULE_TYPE_DATA_BAR);
  bar.has_data_bar = 1;
  bar.data_bar_min.type = FM_CF_VALUE_OBJECT_TYPE_MIN;
  bar.data_bar_max.type = FM_CF_VALUE_OBJECT_TYPE_NUMBER;
  bar.data_bar_max.value = "";
  bar.data_bar_fill = colors[0];
  bar.data_bar_show_value = 1;
  bar.data_bar_min_length_pct = 10;
  bar.data_bar_max_length_pct = 90;
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &bar, &id), kInvalidArgument);
  EXPECT_EQ(RuleCount(wb), 1U);
}

// ---------------------------------------------------------------------------
// text: "" is a value
// ---------------------------------------------------------------------------

TEST(CfEmptyStrings, TextEmptyIsKeptAndAbsentStaysAbsent) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cf_rule_t rule = Rule(FM_CF_RULE_TYPE_CONTAINS_TEXT);
  rule.text = "";
  const char* id = nullptr;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &id), 0) << fm_get_last_error_message();
  {
    const fm_cf_rule_t out = ReadAt(wb, 0);
    ASSERT_NE(out.text, nullptr);
    EXPECT_STREQ(out.text, "");
    EXPECT_STREQ(out.formula1, "=NOT(ISERROR(SEARCH(\"\",A1)))");
  }

  rule.text = nullptr;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &id), 0) << fm_get_last_error_message();
  EXPECT_EQ(ReadAt(wb, 1).text, nullptr);
}

TEST(CfEmptyStrings, TextUpdateToEmptyIsKept) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cf_rule_t rule = Rule(FM_CF_RULE_TYPE_BEGINS_WITH);
  rule.text = "ab";
  const char* id = nullptr;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &id), 0) << fm_get_last_error_message();
  const std::string rule_id(id);

  fm_cf_rule_patch_t patch = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  patch.text_op = FM_PATCH_OP_SET;
  patch.text = "";
  ASSERT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, rule_id.c_str(), &patch), 0) << fm_get_last_error_message();
  const fm_cf_rule_t out = ReadAt(wb, 0);
  ASSERT_NE(out.text, nullptr);
  EXPECT_STREQ(out.text, "");
}

TEST(CfEmptyStrings, TextEmptyIsSavedAsAnAttributeAndRoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cf_rule_t rule = Rule(FM_CF_RULE_TYPE_CONTAINS_TEXT);
  rule.text = "";
  const char* id = nullptr;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &id), 0) << fm_get_last_error_message();

  const Bytes xlsx = Save(wb, FM_WORKBOOK_FORMAT_XLSX);
  const std::string xml = SheetXml(xlsx);
  EXPECT_NE(xml.find("type=\"containsText\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("operator=\"containsText\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("text=\"\""), std::string::npos) << xml;

  for (const Bytes& bytes : {xlsx, Save(wb, FM_WORKBOOK_FORMAT_XLSB)}) {
    WorkbookGuard reloaded;
    Load(bytes, reloaded);
    ASSERT_EQ(RuleCount(reloaded), 1U);
    const fm_cf_rule_t out = ReadAt(reloaded, 0);
    ASSERT_NE(out.text, nullptr);
    EXPECT_STREQ(out.text, "");
  }
}

TEST(CfEmptyStrings, ContainsTextEmptyMatchesEveryCellThatIsNotAnError) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillCells(wb);
  fm_cf_rule_t rule = Rule(FM_CF_RULE_TYPE_CONTAINS_TEXT);
  rule.text = "";
  const char* id = nullptr;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &id), 0) << fm_get_last_error_message();
  // SEARCH("", x) is 1 for text, number and empty cells; the error cell propagates its error.
  EXPECT_EQ(MatchedRows(wb), (std::vector<std::uint32_t>{0, 1, 2}));

  // The same rule after a save and reload evaluates the same way.
  WorkbookGuard reloaded;
  Load(Save(wb, FM_WORKBOOK_FORMAT_XLSX), reloaded);
  ASSERT_EQ(fm_workbook_recalc(reloaded.handle), 0);
  EXPECT_EQ(MatchedRows(reloaded), (std::vector<std::uint32_t>{0, 1, 2}));
}

}  // namespace
