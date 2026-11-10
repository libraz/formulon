//
// C ABI: the getters that read a natural-key collection by its key, the
// defined-name remove, and the index each index-keyed add returns.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>

#include "c_api/formulon_c.h"
#include "formulon_c_out_index.h"
#include "formulon_c_setters.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "utils/error.h"
#include "value.h"

namespace {

constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);

constexpr const char* kAlice = "{0A1B2C3D-0000-4000-8000-000000000001}";
constexpr const char* kThread = "{6F1E0D2C-1111-4111-8111-000000000001}";

void CreateDefault(WorkbookGuard& wb) {
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
}

TEST(FormulonCApiNaturalKey, DefinedNameByNameAndScope) {
  WorkbookGuard wb;
  CreateDefault(wb);
  const fm_defined_name_t rate{sizeof(fm_defined_name_t),      "Rate",     "Sheet1!$A$1",
                               FM_DEFINED_NAME_SCOPE_WORKBOOK, "the rate", 1};
  ASSERT_EQ(fm_workbook_set_defined_name(wb.handle, &rate), 0);
  ASSERT_EQ(SetDefinedName(wb.handle, "Rate", "Sheet1!$B$1", 0), 0);

  int32_t present = 0;
  fm_defined_name_t found = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(fm_workbook_get_defined_name(wb.handle, "rate", FM_DEFINED_NAME_SCOPE_WORKBOOK, &present, &found), 0);
  ASSERT_EQ(present, 1);
  EXPECT_STREQ(found.name, "Rate");
  EXPECT_STREQ(found.formula, "=Sheet1!$A$1");
  EXPECT_EQ(found.local_sheet_id, FM_DEFINED_NAME_SCOPE_WORKBOOK);
  EXPECT_STREQ(found.comment, "the rate");
  EXPECT_EQ(found.hidden, 1);
  ASSERT_EQ(fm_workbook_get_defined_name(wb.handle, "Rate", 0, &present, &found), 0);
  ASSERT_EQ(present, 1);
  EXPECT_STREQ(found.formula, "=Sheet1!$B$1");
  EXPECT_EQ(found.local_sheet_id, 0);
  EXPECT_EQ(found.hidden, 0);

  ASSERT_EQ(fm_workbook_get_defined_name(wb.handle, "Missing", FM_DEFINED_NAME_SCOPE_WORKBOOK, &present, &found), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(found.name, nullptr);
  EXPECT_EQ(found.formula, nullptr);
  EXPECT_EQ(fm_workbook_get_defined_name(wb.handle, "Rate", FM_DEFINED_NAME_SCOPE_WORKBOOK, nullptr, &found),
            kBindingNullPointer);
  EXPECT_EQ(fm_workbook_get_defined_name(wb.handle, "Rate", FM_DEFINED_NAME_SCOPE_WORKBOOK, &present, nullptr),
            kBindingNullPointer);

  ASSERT_EQ(fm_workbook_remove_defined_name(wb.handle, "RATE", FM_DEFINED_NAME_SCOPE_WORKBOOK), 0);
  ASSERT_EQ(fm_workbook_get_defined_name(wb.handle, "Rate", FM_DEFINED_NAME_SCOPE_WORKBOOK, &present, &found), 0);
  EXPECT_EQ(present, 0);
  ASSERT_EQ(fm_workbook_get_defined_name(wb.handle, "Rate", 0, &present, &found), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(fm_workbook_remove_defined_name(wb.handle, "Rate", FM_DEFINED_NAME_SCOPE_WORKBOOK), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_workbook_remove_defined_name(wb.handle, nullptr, FM_DEFINED_NAME_SCOPE_WORKBOOK), kBindingNullPointer);
}

TEST(FormulonCApiNaturalKey, DefinedNameScopeRefusalNamesTheScopeInContext) {
  WorkbookGuard wb;
  CreateDefault(wb);
  EXPECT_EQ(fm_workbook_remove_defined_name(wb.handle, "Missing", FM_DEFINED_NAME_SCOPE_WORKBOOK), FM_STATUS_NOT_FOUND);
  EXPECT_NE(std::string(fm_get_last_error_context()).find("scope=workbook"), std::string::npos);

  EXPECT_EQ(fm_workbook_remove_defined_name(wb.handle, "Missing", 7),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kSheetIndexOutOfRange));
  EXPECT_NE(std::string(fm_get_last_error_context()).find("scope=7"), std::string::npos);
}

TEST(FormulonCApiNaturalKey, RemovingADefinedNameRefreshesItsReaders) {
  WorkbookGuard wb;
  CreateDefault(wb);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 5.0), 0);
  ASSERT_EQ(SetDefinedName(wb.handle, "Five", "Sheet1!$A$1"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 1, "=Five"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(fm_workbook_remove_defined_name(wb.handle, "Five", FM_DEFINED_NAME_SCOPE_WORKBOOK), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 0, 1, &v), 0);
  ASSERT_EQ(v.kind, FM_VALUE_KIND_ERROR);
  EXPECT_EQ(v.u.error_code, static_cast<int32_t>(formulon::ErrorCode::Name));
}

TEST(FormulonCApiNaturalKey, TableAndCellStyleByName) {
  WorkbookGuard wb;
  CreateDefault(wb);
  const fm_table_style_element_t element{FM_TABLE_STYLE_ELEMENT_TYPE_WHOLE_TABLE, 1U, 0, 0U};
  const fm_table_style_t style{sizeof(fm_table_style_t), "Corporate", 0, 1, &element, 1U};
  ASSERT_EQ(fm_workbook_set_table_style(wb.handle, &style), 0);

  int32_t present = 0;
  fm_table_style_t table_style = FM_STRUCT_INIT(fm_table_style_t);
  ASSERT_EQ(fm_workbook_get_table_style(wb.handle, "corporate", &present, &table_style), 0);
  ASSERT_EQ(present, 1);
  EXPECT_STREQ(table_style.name, "Corporate");
  EXPECT_EQ(table_style.element_count, 1U);
  ASSERT_EQ(fm_workbook_get_table_style(wb.handle, "Other", &present, &table_style), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(table_style.name, nullptr);

  fm_cell_style_t cell_style{};
  ASSERT_EQ(fm_workbook_get_cell_style(wb.handle, "Normal", &present, &cell_style), 0);
  ASSERT_EQ(present, 1);
  EXPECT_STREQ(cell_style.name, "Normal");
  EXPECT_EQ(cell_style.builtin_id, 0U);
  ASSERT_EQ(fm_workbook_get_cell_style(wb.handle, "normal", &present, &cell_style), 0);
  EXPECT_EQ(present, 1);
  EXPECT_STREQ(cell_style.name, "Normal");
  ASSERT_EQ(fm_workbook_get_cell_style(wb.handle, "Other", &present, &cell_style), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(fm_workbook_get_cell_style(wb.handle, nullptr, &present, &cell_style), kBindingNullPointer);
}

TEST(FormulonCApiNaturalKey, HyperlinkByAnchorCell) {
  WorkbookGuard wb;
  CreateDefault(wb);
  fm_hyperlink_t link = FM_STRUCT_INIT(fm_hyperlink_t);
  link.rect = fm_cell_rect_t{1U, 1U, 2U, 3U};
  link.target = "https://example.com/";
  ASSERT_EQ(fm_sheet_set_hyperlink(wb.handle, 0, 1U, 1U, &link), 0);

  int32_t present = 0;
  fm_hyperlink_t out = FM_STRUCT_INIT(fm_hyperlink_t);
  ASSERT_EQ(fm_sheet_get_hyperlink(wb.handle, 0, 1U, 1U, &present, &out), 0);
  ASSERT_EQ(present, 1);
  EXPECT_STREQ(out.target, "https://example.com/");
  EXPECT_EQ(out.rect.last_col, 3U);
  ASSERT_EQ(fm_sheet_get_hyperlink(wb.handle, 0, 2U, 3U, &present, &out), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(fm_sheet_get_hyperlink(wb.handle, 9, 1U, 1U, &present, &out), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

TEST(FormulonCApiNaturalKey, PersonAndThreadedCommentById) {
  WorkbookGuard wb;
  CreateDefault(wb);
  const fm_person_t alice{sizeof(fm_person_t), kAlice, "Alice", "alice@example.com", "None"};
  const char* added_person_id = nullptr;
  ASSERT_EQ(fm_workbook_add_person(wb.handle, &alice, &added_person_id), 0) << fm_get_last_error_message();
  EXPECT_STREQ(added_person_id, kAlice);
  fm_threaded_comment_t comment = FM_STRUCT_INIT(fm_threaded_comment_t);
  comment.id = kThread;
  comment.row = 4U;
  comment.col = 2U;
  comment.person_id = kAlice;
  comment.created = "2026-01-02T03:04:05.67";
  comment.text = "Check this";
  ASSERT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, FM_TEXT_OFFSET_UNIT_UTF16, &comment, ignored_id()), 0)
      << fm_get_last_error_message();

  int32_t present = 0;
  fm_person_t person = FM_STRUCT_INIT(fm_person_t);
  ASSERT_EQ(fm_workbook_get_person(wb.handle, kAlice, &present, &person), 0);
  ASSERT_EQ(present, 1);
  EXPECT_STREQ(person.display_name, "Alice");
  ASSERT_EQ(fm_workbook_get_person(wb.handle, "{0A1B2C3D-0000-4000-8000-000000000009}", &present, &person), 0);
  EXPECT_EQ(present, 0);

  fm_threaded_comment_t thread = FM_STRUCT_INIT(fm_threaded_comment_t);
  ASSERT_EQ(fm_sheet_get_threaded_comment(wb.handle, 0, kThread, &present, &thread), 0);
  ASSERT_EQ(present, 1);
  EXPECT_STREQ(thread.text, "Check this");
  EXPECT_EQ(thread.row, 4U);
  ASSERT_EQ(fm_sheet_get_threaded_comment(wb.handle, 0, kAlice, &present, &thread), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(fm_sheet_get_threaded_comment(wb.handle, 0, nullptr, &present, &thread), kBindingNullPointer);
}

TEST(FormulonCApiNaturalKey, ShapeAndDrawingObjectByObjectId) {
  WorkbookGuard wb;
  CreateDefault(wb);
  fm_shape_spec_t spec = FM_STRUCT_INIT(fm_shape_spec_t);
  spec.placement = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  spec.preset_geometry = "rect";
  spec.placement.anchor.kind = FM_ANCHOR_KIND_ONE_CELL;
  spec.placement.anchor.from_row = 2;
  spec.placement.anchor.width_emu = 914400;
  spec.placement.anchor.height_emu = 457200;
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &spec, &id), 0) << fm_get_last_error_message();

  int32_t present = 0;
  fm_shape_info_t shape = FM_STRUCT_INIT(fm_shape_info_t);
  ASSERT_EQ(fm_sheet_get_shape(wb.handle, 0, id, &present, &shape), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(shape.object_id, id);
  EXPECT_STREQ(shape.preset_geometry, "rect");
  ASSERT_EQ(fm_sheet_get_shape(wb.handle, 0, id + 1U, &present, &shape), 0);
  EXPECT_EQ(present, 0);

  fm_drawing_object_info_t object = FM_STRUCT_INIT(fm_drawing_object_info_t);
  ASSERT_EQ(fm_sheet_get_drawing_object(wb.handle, 0, id, &present, &object), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(object.object_id, id);
  EXPECT_EQ(object.anchor.from_row, 2U);
  ASSERT_EQ(fm_sheet_get_drawing_object(wb.handle, 0, id + 1U, &present, &object), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(fm_sheet_get_drawing_object(wb.handle, 0, id, nullptr, &object), kBindingNullPointer);
}

TEST(FormulonCApiNaturalKey, ThemeScriptFontAndPassthroughByKey) {
  WorkbookGuard wb;
  CreateDefault(wb);
  int32_t present = 0;
  const char* typeface = nullptr;
  ASSERT_EQ(fm_workbook_get_theme_script_font(wb.handle, FM_FONT_SCHEME_MAJOR, "Jpan", &present, &typeface), 0);
  ASSERT_EQ(present, 1);
  ASSERT_NE(typeface, nullptr);
  EXPECT_GT(std::strlen(typeface), 0U);
  ASSERT_EQ(fm_workbook_has_passthrough(wb.handle, "xl/theme/theme1.xml", &present), 0);
  EXPECT_EQ(present, 0);

  ASSERT_EQ(fm_workbook_set_theme_script_font(wb.handle, FM_FONT_SCHEME_MINOR, "Zzzz", "Test Face"), 0);
  ASSERT_EQ(fm_workbook_get_theme_script_font(wb.handle, FM_FONT_SCHEME_MINOR, "Zzzz", &present, &typeface), 0);
  ASSERT_EQ(present, 1);
  EXPECT_STREQ(typeface, "Test Face");
  ASSERT_EQ(fm_workbook_get_theme_script_font(wb.handle, FM_FONT_SCHEME_MAJOR, "Zzzz", &present, &typeface), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(typeface, nullptr);

  // The edited theme is carried as a passthrough part.
  std::size_t parts = 0;
  ASSERT_EQ(fm_workbook_get_passthrough_count(wb.handle, &parts), 0);
  ASSERT_GT(parts, 0U);
  const char* path = nullptr;
  ASSERT_EQ(fm_workbook_get_passthrough_at(wb.handle, 0, &path), 0);
  const std::string first(path);
  ASSERT_EQ(fm_workbook_has_passthrough(wb.handle, first.c_str(), &present), 0);
  EXPECT_EQ(present, 1);
  ASSERT_EQ(fm_workbook_has_passthrough(wb.handle, ("/" + first).c_str(), &present), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(fm_workbook_has_passthrough(wb.handle, nullptr, &present), kBindingNullPointer);
}

TEST(FormulonCApiNaturalKey, IndexKeyedAddsReturnTheirIndex) {
  WorkbookGuard wb;
  CreateDefault(wb);
  std::size_t sheet = 99;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Second", append_position(wb.handle), &sheet), 0);
  EXPECT_EQ(sheet, 1U);
  EXPECT_EQ(fm_workbook_add_sheet(wb.handle, "Third", append_position(wb.handle), nullptr), kBindingNullPointer);

  fm_cell_rect_t range{0U, 0U, 3U, 0U};
  fm_data_validation_t validation = FM_STRUCT_INIT(fm_data_validation_t);
  validation.ranges = &range;
  validation.range_count = 1;
  validation.type = 1;
  validation.op = 0;
  validation.formula1 = "1";
  validation.formula2 = "9";
  const char* added = nullptr;
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, sheet, &validation, &added), 0) << fm_get_last_error_message();
  ASSERT_NE(added, nullptr);
  const std::string first = added;
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, sheet, &validation, &added), 0);
  EXPECT_NE(first, added);
  fm_data_validation_t read = FM_STRUCT_INIT(fm_data_validation_t);
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_validation(wb.handle, sheet, first.c_str(), &present, &read), 0);
  EXPECT_EQ(present, 1);

  fm_protected_range_t protected_range = FM_STRUCT_INIT(fm_protected_range_t);
  protected_range.name = "Inputs";
  protected_range.ranges = &range;
  protected_range.range_count = 1;
  const char* protected_name = nullptr;
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, sheet, &protected_range, &protected_name), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, sheet, nullptr, &protected_name), kBindingNullPointer);
}

TEST(FormulonCApiNaturalKey, ARejectedAddLeavesItsOutIndexUntouched) {
  WorkbookGuard wb;
  CreateDefault(wb);
  std::size_t out = 99;
  EXPECT_NE(fm_sheet_add_pivot_field_item_by_cache_index(wb.handle, 0, 0, 0, 0, 1, &out), 0);
  EXPECT_EQ(out, 99U);
  EXPECT_NE(fm_sheet_add_pivot_field_item(wb.handle, 0, 0, 0, "East", 1, &out), 0);
  EXPECT_EQ(out, 99U);
  EXPECT_NE(fm_workbook_add_sheet(wb.handle, "Sheet1", append_position(wb.handle), &out), 0);
  EXPECT_EQ(out, 99U);
  fm_data_validation_t validation = FM_STRUCT_INIT(fm_data_validation_t);
  const char* out_id = nullptr;
  EXPECT_NE(fm_sheet_add_validation(wb.handle, 0, &validation, &out_id), 0);
  EXPECT_EQ(out_id, nullptr);
}

}  // namespace
