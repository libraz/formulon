//
// C ABI contracts shared by the bindings' usability surface: the dirty-cell
// query, the process-scope image probe, NotFound for an absent drawing or
// chart id, the duplicate CF id refusal, and the leading `=` on formula text
// read back from cells, rules, tables and pivot caches.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_setters.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"

namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr const char* kRuleId = "{0A1B2C3D-0000-4000-8000-00000000AB01}";

Bytes Png(std::uint32_t w, std::uint32_t h) {
  Bytes b = {0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, 'I', 'H', 'D', 'R'};
  for (std::uint32_t v : {w, h}) {
    for (int shift = 24; shift >= 0; shift -= 8) {
      b.push_back(static_cast<std::uint8_t>(v >> shift));
    }
  }
  b.insert(b.end(), {8, 6, 0, 0, 0, 0x1F, 0x15, 0xC4, 0x89});
  return b;
}

std::int32_t Dirty(const fm_workbook_t* wb) {
  std::int32_t dirty = -1;
  EXPECT_EQ(fm_workbook_has_dirty_cells(wb, &dirty), 0) << fm_get_last_error_message();
  return dirty;
}

void SaveAndReload(const fm_workbook_t* wb, WorkbookGuard* out) {
  BufferGuard buf;
  ASSERT_EQ(fm_workbook_save(wb, &buf.data, &buf.len), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_load(buf.data, buf.len, &out->handle), 0) << fm_get_last_error_message();
}

std::string Str(const char* text) {
  return text != nullptr ? std::string(text) : std::string("<null>");
}

fm_cf_rule_t ExpressionRule(const fm_cell_rect_t* sqref, const char* formula, const char* id = nullptr) {
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.id = id;
  rule.type = 0;  // Expression
  rule.formula1 = formula;
  rule.ranges = sqref;
  rule.range_count = 1;
  return rule;
}

std::size_t CfRuleCount(const fm_workbook_t* wb) {
  std::size_t count = 99;
  EXPECT_EQ(fm_sheet_get_cf_rule_count(wb, 0, &count), 0);
  return count;
}

fm_data_validation_t WholeBetween(const fm_cell_rect_t* sqref, const char* low, const char* high) {
  fm_data_validation_t v = FM_STRUCT_INIT(fm_data_validation_t);
  v.ranges = sqref;
  v.range_count = 1;
  v.type = 1;  // whole
  v.op = 0;    // between
  v.formula1 = low;
  v.formula2 = high;
  return v;
}

// A progress callback that queries the dirty set of the workbook it reports on.
fm_status_t g_reentrant_status = 0;

extern "C" int32_t query_dirty_from_progress(uint32_t /*iteration*/, double /*max_residual*/,
                                             uint32_t /*max_iterations*/, void* user_data) {
  std::int32_t dirty = -1;
  g_reentrant_status = fm_workbook_has_dirty_cells(static_cast<const fm_workbook_t*>(user_data), &dirty);
  return 0;
}

}  // namespace

// ---------------------------------------------------------------------------
// fm_workbook_has_dirty_cells
// ---------------------------------------------------------------------------

TEST(FormulonCApiUsability, HasDirtyCellsIsFalseOnANewWorkbookAndAfterRecalc) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(Dirty(wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 5.0), 0);
  EXPECT_EQ(Dirty(wb.handle), 1);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_EQ(Dirty(wb.handle), 0);
}

TEST(FormulonCApiUsability, HasDirtyCellsRejectsNullArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::int32_t dirty = 0;
  EXPECT_EQ(fm_workbook_has_dirty_cells(nullptr, &dirty), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_workbook_has_dirty_cells(wb.handle, nullptr), FM_STATUS_BINDING_NULL_POINTER);
}

// A loaded workbook holds the file's cached results; its formulas stay pending
// until the first full recalc re-evaluates them.
TEST(FormulonCApiUsability, HasDirtyCellsAfterLoadReflectsThePendingFormulas) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 2.0), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  WorkbookGuard literals;
  SaveAndReload(wb.handle, &literals);
  EXPECT_EQ(Dirty(literals.handle), 0);

  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 1, "=A1*2"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  WorkbookGuard formulas;
  SaveAndReload(wb.handle, &formulas);
  EXPECT_EQ(Dirty(formulas.handle), 1);
  ASSERT_EQ(fm_workbook_recalc(formulas.handle), 0);
  EXPECT_EQ(Dirty(formulas.handle), 0);
}

TEST(FormulonCApiUsability, HasDirtyCellsAfterAReferencedDefinedNameChanges) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetDefinedName(wb.handle, "Rate", "3"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=Rate*2"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(Dirty(wb.handle), 0);
  ASSERT_EQ(SetDefinedName(wb.handle, "Rate", "4"), 0);
  EXPECT_EQ(Dirty(wb.handle), 1);
}

TEST(FormulonCApiUsability, HasDirtyCellsAfterInsertRows) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 1, 0, "=SUM(A1:A1)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(Dirty(wb.handle), 0);
  ASSERT_EQ(fm_sheet_insert_rows(wb.handle, 0, 0, 1), 0);
  EXPECT_EQ(Dirty(wb.handle), 1);
}

TEST(FormulonCApiUsability, HasDirtyCellsAfterRemovingAReferencedSheet) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t second = 0;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Data", &second), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, second, 0, 0, 7.0), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=Data!A1+1"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(Dirty(wb.handle), 0);
  ASSERT_EQ(fm_workbook_remove_sheet(wb.handle, second), 0);
  EXPECT_EQ(Dirty(wb.handle), 1);
}

TEST(FormulonCApiUsability, HasDirtyCellsFollowsAProfileChangeOnlyWhenItChanges) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=LEN(\"abc\")"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  const char* current = nullptr;
  ASSERT_EQ(fm_workbook_get_profile_id(wb.handle, &current), 0);
  const std::string same = Str(current);
  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, same.c_str()), 0);
  EXPECT_EQ(Dirty(wb.handle), 0);
  const char* other = same == "win-365-ja_JP" ? "win-365-en_US" : "win-365-ja_JP";
  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, other), 0);
  EXPECT_EQ(Dirty(wb.handle), 1);
}

TEST(FormulonCApiUsability, HasDirtyCellsLeavesVolatileCellsOut) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=NOW()"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_EQ(Dirty(wb.handle), 0);
}

TEST(FormulonCApiUsability, HasDirtyCellsStaysTrueAfterAViewportRecalc) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 1, "=A1+1"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 500, 1, "=A1+2"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 2.0), 0);
  const fm_cell_rect_t viewport{0, 0, 9, 9};
  ASSERT_EQ(fm_sheet_recalc_viewport(wb.handle, 0, &viewport, nullptr), 0);
  EXPECT_EQ(Dirty(wb.handle), 1);
}

TEST(FormulonCApiUsability, HasDirtyCellsFromTheProgressCallbackIsReentrant) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetIterative(wb.handle, 1, 100, 0.001), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=(A1+10)/2"), 0);
  g_reentrant_status = 0;
  ASSERT_EQ(fm_workbook_set_iterative_progress(wb.handle, &query_dirty_from_progress, wb.handle), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_EQ(g_reentrant_status, FM_STATUS_GRAPH_RECALC_REENTRANT);
}

// ---------------------------------------------------------------------------
// fm_get_image_info
// ---------------------------------------------------------------------------

TEST(FormulonCApiUsability, ImageInfoNeedsNoWorkbook) {
  const Bytes png = Png(16, 8);
  fm_image_info_t info = FM_STRUCT_INIT(fm_image_info_t);
  ASSERT_EQ(fm_get_image_info(png.data(), png.size(), &info), 0) << fm_get_last_error_message();
  EXPECT_EQ(info.format, FM_IMAGE_FORMAT_PNG);
  EXPECT_EQ(info.width_px, 16U);
  EXPECT_EQ(info.height_px, 8U);

  const Bytes junk = {'n', 'o', 't', ' ', 'a', 'n', ' ', 'i', 'm', 'a', 'g', 'e'};
  EXPECT_EQ(fm_get_image_info(junk.data(), junk.size(), &info), FM_STATUS_IO_IMAGE_UNSUPPORTED);
  EXPECT_EQ(fm_get_image_info(nullptr, 0, &info), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_get_image_info(png.data(), png.size(), nullptr), FM_STATUS_BINDING_NULL_POINTER);
}

// ---------------------------------------------------------------------------
// Absent drawing and chart ids
// ---------------------------------------------------------------------------

TEST(FormulonCApiUsability, AbsentDrawingAndChartIdsAreNotFound) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_chart_data_t* data = nullptr;
  const fm_chart_preview_overrides_t overrides = FM_STRUCT_INIT(fm_chart_preview_overrides_t);
  // A sheet with no drawing at all.
  EXPECT_EQ(fm_sheet_remove_drawing(wb.handle, 0, 42U), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_get_chart_data(wb.handle, 0, 42U, &data), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(data, nullptr);
  EXPECT_EQ(fm_sheet_get_chart_preview(wb.handle, 0, 42U, &overrides, &data), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(data, nullptr);

  // A drawing whose only object is a picture: its id names no chart.
  const Bytes png = Png(4, 4);
  fm_picture_insert_spec_t opts{};
  opts.anchor.anchor_kind = FM_ANCHOR_KIND_ONE_CELL;
  opts.anchor.edit_as = FM_ANCHOR_EDIT_AS_TWO_CELL;
  std::uint32_t picture = 0;
  ASSERT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), &opts, &picture), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_get_chart_data(wb.handle, 0, picture, &data), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_get_chart_preview(wb.handle, 0, picture, &overrides, &data), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_remove_drawing(wb.handle, 0, picture + 100U), FM_STATUS_NOT_FOUND);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_drawing_object_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);

  ASSERT_EQ(fm_sheet_remove_drawing(wb.handle, 0, picture), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_remove_drawing(wb.handle, 0, picture), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_remove_drawing(wb.handle, 9, picture), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

// ---------------------------------------------------------------------------
// Conditional-format rule ids
// ---------------------------------------------------------------------------

// A read rule re-added to its own sheet collides with itself; the copy takes a fresh GUID.
std::string CfRuleId(const fm_workbook_t* wb, std::size_t index) {
  fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
  EXPECT_EQ(fm_sheet_get_cf_rule(wb, 0, index, &out), 0);
  return out.id != nullptr ? std::string(out.id) : std::string();
}

TEST(FormulonCApiUsability, CfRuleWithAnExistingIdGetsAFreshId) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t sqref{0, 0, 9, 0};
  std::size_t index = 99;
  fm_cf_rule_t rule = ExpressionRule(&sqref, "A1>0", kRuleId);
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &index), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &index), 0) << fm_get_last_error_message();
  std::string lower(kRuleId);
  for (char& c : lower) {
    if (c >= 'A' && c <= 'F') {
      c = static_cast<char>(c - 'A' + 'a');
    }
  }
  rule.id = lower.c_str();
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &index), 0) << fm_get_last_error_message();
  ASSERT_EQ(CfRuleCount(wb.handle), 3U);
  const std::string first = CfRuleId(wb.handle, 0);
  const std::string second = CfRuleId(wb.handle, 1);
  const std::string third = CfRuleId(wb.handle, 2);
  EXPECT_EQ(first, kRuleId);
  EXPECT_FALSE(second.empty());
  EXPECT_FALSE(third.empty());
  EXPECT_NE(second, first);
  EXPECT_NE(third, first);
  EXPECT_NE(third, second);

  // The same id on another sheet keeps its id.
  std::size_t other = 0;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Other", &other), 0);
  rule.id = kRuleId;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, other, &rule, &index), 0) << fm_get_last_error_message();
  fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule(wb.handle, other, index, &out), 0);
  ASSERT_NE(out.id, nullptr);
  EXPECT_STREQ(out.id, kRuleId);
}

TEST(FormulonCApiUsability, CfRuleReadBackFeedsAddWithoutDoublingTheEquals) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t sqref{0, 0, 9, 0};
  std::size_t index = 99;
  fm_cf_rule_t rule = ExpressionRule(&sqref, "=A1>0", kRuleId);
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &index), 0) << fm_get_last_error_message();

  fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(Str(out.formula1), "=A1>0");
  // The read record goes back in once its own rule is gone.
  ASSERT_EQ(fm_sheet_remove_cf_rule(wb.handle, 0, 0), 0);
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &out, &index), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_cf_rule(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(Str(out.formula1), "=A1>0");
  EXPECT_EQ(Str(out.id), kRuleId);
}

// ---------------------------------------------------------------------------
// Formula text: the leading `=`
// ---------------------------------------------------------------------------

TEST(FormulonCApiUsability, CellFormulaTextCarriesTheEqualsInBothNotations) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "1+1"), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 1, 1, "=A1*2"), 0) << fm_get_last_error_message();
  const char* text = nullptr;
  ASSERT_EQ(fm_sheet_get_formula(wb.handle, 0, 0, 0, &text), 0);
  EXPECT_EQ(Str(text), "=1+1");
  ASSERT_EQ(fm_sheet_get_formula_r1c1(wb.handle, 0, 0, 0, &text), 0);
  EXPECT_EQ(Str(text), "=1+1");
  ASSERT_EQ(fm_sheet_get_formula_r1c1(wb.handle, 0, 1, 1, &text), 0);
  EXPECT_EQ(Str(text), "=R[-1]C[-1]*2");
  // A cell without a formula stays empty.
  ASSERT_EQ(fm_sheet_get_formula_r1c1(wb.handle, 0, 5, 5, &text), 0);
  EXPECT_EQ(Str(text), "");
}

TEST(FormulonCApiUsability, ArrayFormulaTextCarriesTheEqualsWhicheverWayItWasEntered) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  for (std::uint32_t r = 0; r < 3; ++r) {
    ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, r, 0, r + 1.0), 0);
  }
  const fm_cell_rect_t block{0, 2, 2, 2};
  ASSERT_EQ(fm_sheet_set_array_formula(wb.handle, 0, &block, "A1:A3*2"), 0) << fm_get_last_error_message();
  fm_array_formula_t info = FM_STRUCT_INIT(fm_array_formula_t);
  ASSERT_EQ(fm_sheet_get_array_formula(wb.handle, 0, 1, 2, &info), 0);
  EXPECT_EQ(Str(info.formula), "=A1:A3*2");
  ASSERT_EQ(fm_sheet_set_array_formula(wb.handle, 0, &block, "=A1:A3*3"), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_array_formula(wb.handle, 0, 1, 2, &info), 0);
  EXPECT_EQ(Str(info.formula), "=A1:A3*3");
}

TEST(FormulonCApiUsability, CfFormulasCarryTheEqualsWhicheverWayTheyWereEntered) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t sqref{0, 0, 9, 0};
  std::size_t index = 99;
  fm_cf_rule_t bare = FM_STRUCT_INIT(fm_cf_rule_t);
  bare.type = 1;  // CellIs
  bare.op_engaged = 1;
  bare.op = 0;  // Between
  bare.formula1 = "1";
  bare.formula2 = "=10";
  bare.ranges = &sqref;
  bare.range_count = 1;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &bare, &index), 0) << fm_get_last_error_message();
  fm_cf_rule_t out = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(Str(out.formula1), "=1");
  EXPECT_EQ(Str(out.formula2), "=10");

  // A derived formula reads back with the `=` as well.
  fm_cf_rule_t text = FM_STRUCT_INIT(fm_cf_rule_t);
  text.type = 7;  // ContainsText
  text.text = "ab";
  text.ranges = &sqref;
  text.range_count = 1;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &text, &index), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_cf_rule(wb.handle, 0, index, &out), 0);
  EXPECT_EQ(Str(out.formula1), "=NOT(ISERROR(SEARCH(\"ab\",A1)))");

  // A doubled `=` is still refused.
  fm_cf_rule_t doubled = ExpressionRule(&sqref, "==A1>0");
  EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &doubled, &index), FM_STATUS_PARSER_UNEXPECTED_TOKEN);
}

TEST(FormulonCApiUsability, ValidationFormulasCarryTheEqualsWhicheverWayTheyWereEntered) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t sqref{0, 0, 9, 0};
  const fm_data_validation_t v = WholeBetween(&sqref, "1", "=B1*2");
  std::size_t index = 99;
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, &v, &index), 0) << fm_get_last_error_message();
  fm_data_validation_t out = FM_STRUCT_INIT(fm_data_validation_t);
  ASSERT_EQ(fm_sheet_get_validation(wb.handle, 0, index, &out), 0);
  EXPECT_EQ(Str(out.formula1), "=1");
  EXPECT_EQ(Str(out.formula2), "=B1*2");

  // An absent formula stays empty.
  const fm_data_validation_t one = WholeBetween(&sqref, "5", nullptr);
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, &one, &index), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_validation(wb.handle, 0, index, &out), 0);
  EXPECT_EQ(Str(out.formula1), "=5");
  EXPECT_EQ(Str(out.formula2), "");
}

TEST(FormulonCApiUsability, TableTotalsFormulaCarriesTheEqualsWhicheverWayItWasEntered) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, "Product"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 1, "Amount"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 1, 10.0), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 2, 1, 20.0), 0);
  const fm_cell_rect_t ref{0, 0, 3, 1};
  const char* columns[] = {"Product", "Amount"};
  std::size_t table = 99;
  ASSERT_EQ(fm_sheet_add_table(wb.handle, 0, &ref, "Sales", columns, 2, "TableStyleMedium2", 1, 1, &table), 0)
      << fm_get_last_error_message();
  fm_table_column_info_t column = FM_STRUCT_INIT(fm_table_column_info_t);
  ASSERT_EQ(fm_sheet_get_table_column(wb.handle, 0, table, 1, &column), 0);
  EXPECT_EQ(Str(column.totals_row_formula), "");

  ASSERT_EQ(fm_sheet_set_table_column_totals(wb.handle, 0, table, 1, "custom", nullptr, "SUM(Sales[Amount])*2"), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_table_column(wb.handle, 0, table, 1, &column), 0);
  EXPECT_EQ(Str(column.totals_row_formula), "=SUM([Amount])*2");
  ASSERT_EQ(fm_sheet_set_table_column_totals(wb.handle, 0, table, 1, "custom", nullptr, "=SUM(Sales[Amount])*3"), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_table_column(wb.handle, 0, table, 1, &column), 0);
  EXPECT_EQ(Str(column.totals_row_formula), "=SUM([Amount])*3");
}

TEST(FormulonCApiUsability, PivotCacheFieldFormulaCarriesTheEqualsWhicheverWayItWasEntered) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0, &cache), 0) << fm_get_last_error_message();
  for (const char* name : {"Amt", "Qty"}) {
    std::size_t field = 99;
    ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb.handle, cache, name, &field), 0) << fm_get_last_error_message();
  }
  std::size_t price = 99;
  ASSERT_EQ(fm_workbook_add_pivot_cache_calculated_field(wb.handle, cache, "Price", "Amt/Qty", &price), 0)
      << fm_get_last_error_message();
  const char* formula = nullptr;
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_formula(wb.handle, cache, price, &formula), 0);
  EXPECT_EQ(Str(formula), "=Amt/Qty");
  ASSERT_EQ(fm_workbook_set_pivot_cache_calculated_field_formula(wb.handle, cache, price, "=Amt*2"), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_formula(wb.handle, cache, price, &formula), 0);
  EXPECT_EQ(Str(formula), "=Amt*2");
  // A database field has no formula.
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_formula(wb.handle, cache, 0, &formula), 0);
  EXPECT_EQ(Str(formula), "");
}

TEST(FormulonCApiUsability, SaveAndReloadDoesNotDoubleTheEquals) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 1, "A1+1"), 0);
  const fm_cell_rect_t sqref{0, 0, 9, 0};
  std::size_t index = 99;
  const fm_cf_rule_t rule = ExpressionRule(&sqref, "=A1>0");
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &index), 0) << fm_get_last_error_message();
  const fm_data_validation_t v = WholeBetween(&sqref, "=1", "10");
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, &v, &index), 0) << fm_get_last_error_message();
  ASSERT_EQ(SetDefinedName(wb.handle, "Rate", "=0.5"), 0);

  WorkbookGuard reloaded;
  SaveAndReload(wb.handle, &reloaded);
  const char* text = nullptr;
  ASSERT_EQ(fm_sheet_get_formula(reloaded.handle, 0, 0, 1, &text), 0);
  EXPECT_EQ(Str(text), "=A1+1");
  ASSERT_EQ(fm_sheet_get_formula_r1c1(reloaded.handle, 0, 0, 1, &text), 0);
  EXPECT_EQ(Str(text), "=RC[-1]+1");
  fm_cf_rule_t cf = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule(reloaded.handle, 0, 0, &cf), 0);
  EXPECT_EQ(Str(cf.formula1), "=A1>0");
  fm_data_validation_t dv = FM_STRUCT_INIT(fm_data_validation_t);
  ASSERT_EQ(fm_sheet_get_validation(reloaded.handle, 0, 0, &dv), 0);
  EXPECT_EQ(Str(dv.formula1), "=1");
  EXPECT_EQ(Str(dv.formula2), "=10");
  fm_defined_name_t name = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(fm_workbook_get_defined_name_at(reloaded.handle, 0, &name), 0);
  EXPECT_EQ(Str(name.formula), "=0.5");
}
