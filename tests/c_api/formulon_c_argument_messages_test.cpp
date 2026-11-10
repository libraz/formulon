// Argument errors a binding can surface: the message leads with the C parameter or field, in the
// `<field> <constraint>` shape, and carries no C-only vocabulary.

#include <cstdint>
#include <string>
#include <string_view>

#include "c_api/binding_message.h"
#include "formulon_c_pivot_test_helpers.h"
#include "formulon_c_setters.h"
#include "utils/error.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kSheetIndexOutOfRange =
    static_cast<fm_status_t>(formulon::FormulonErrorCode::kSheetIndexOutOfRange);
constexpr fm_status_t kIndexOutOfRange = static_cast<fm_status_t>(formulon::FormulonErrorCode::kIndexOutOfRange);

std::string LastError() {
  const char* message = fm_get_last_error_message();
  return message != nullptr ? message : "";
}

/// Words a binding must never see in a message built from a C body.
void ExpectNoCOnlyVocabulary(const std::string& message) {
  for (const char* banned : {"argument ", "`", "FM_", "fm_", "NULL", "spec->", "_op", "SET", "KEEP", "CLEAR",
                             "4294967295", "uint32", "int32", "2^63"}) {
    EXPECT_EQ(message.find(banned), std::string::npos) << "'" << banned << "' in: " << message;
  }
}

/// The first whitespace-delimited token of `message`.
std::string FirstToken(const std::string& message) {
  return message.substr(0, message.find(' '));
}

}  // namespace

TEST(FormulonCApiArgumentMessages, BindingReachableBodiesLeadWithTheField) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  // Empty table style name.
  const fm_cell_rect_t ref = {0, 0, 1, 0};
  const char* columns[] = {"Product"};
  std::size_t table = 99U;
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, "Product"), 0);
  ASSERT_EQ(AddTable(wb.handle, 0, &ref, "Sales", columns, 1, "TableStyleMedium2", 1, 0, &table), 0) << LastError();
  fm_table_style_options_patch_t style = FM_STRUCT_INIT(fm_table_style_options_patch_t);
  style.style_name_op = FM_PATCH_OP_SET;
  style.style_name = "";
  EXPECT_EQ(fm_sheet_update_table_style_options(wb.handle, 0, table, &style), kInvalidArgument);
  EXPECT_EQ(LastError(), "style_name must not be empty");
  ExpectNoCOnlyVocabulary(LastError());

  // Empty shape name.
  fm_shape_spec_t shape = FM_STRUCT_INIT(fm_shape_spec_t);
  shape.placement = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  shape.placement.anchor.kind = FM_ANCHOR_KIND_TWO_CELL;
  shape.placement.anchor.from_row = 2;
  shape.placement.anchor.from_col = 1;
  shape.placement.anchor.to_row = 6;
  shape.placement.anchor.to_col = 4;
  std::uint32_t shape_id = 0;
  ASSERT_EQ(fm_sheet_add_shape(wb.handle, 0, &shape, &shape_id), 0) << LastError();
  fm_drawing_patch_t common = FM_STRUCT_INIT(fm_drawing_patch_t);
  common.name_op = FM_PATCH_OP_SET;
  common.name = "";
  EXPECT_EQ(fm_sheet_update_shape(wb.handle, 0, shape_id, &common, nullptr), kInvalidArgument);
  EXPECT_EQ(LastError(), "name must not be empty");
  ExpectNoCOnlyVocabulary(LastError());

  // Empty sparkline items.
  fm_sparkline_group_t group = FM_STRUCT_INIT(fm_sparkline_group_t);
  group.type = FM_SPARKLINE_TYPE_COLUMN;
  group.line_width_pt = 0.75;
  group.empty_cells_as = FM_BLANKS_AS_GAP;
  const fm_sparkline_t items[] = {{"Sheet1!A1:D1", "F1"}};
  const char* group_id = nullptr;
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &group, items, 1, &group_id), 0) << LastError();
  ASSERT_NE(group_id, nullptr);
  const std::string id = group_id;
  fm_sparkline_group_patch_t sparklines = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  sparklines.items_op = FM_PATCH_OP_SET;
  sparklines.items = nullptr;
  sparklines.items_count = 0;
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, id.c_str(), &sparklines), kInvalidArgument);
  EXPECT_EQ(LastError(), "items needs at least one sparkline");
  ExpectNoCOnlyVocabulary(LastError());

  // Bad theme scheme.
  std::size_t script_fonts = 0;
  EXPECT_EQ(fm_workbook_get_theme_script_font_count(wb.handle, 7, &script_fonts), kInvalidArgument);
  EXPECT_EQ(LastError(), "scheme 7 is not a FontScheme");
  ExpectNoCOnlyVocabulary(LastError());

  // Data field index 99.
  std::uint32_t cache_id = 0U;
  std::size_t pivot = 0U;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot), 0) << LastError();
  std::size_t data_fields = 0U;
  ASSERT_EQ(fm_sheet_get_pivot_data_field_count(wb.handle, 0, pivot, &data_fields), 0);
  fm_pivot_data_field_t field = FM_STRUCT_INIT(fm_pivot_data_field_t);
  EXPECT_EQ(fm_sheet_get_pivot_data_field_at(wb.handle, 0, pivot, 99U, &field), kIndexOutOfRange);
  EXPECT_EQ(LastError(), "data_field_index 99 is out of range (" + std::to_string(data_fields) +
                             (data_fields == 1 ? " item)" : " items)"));
  ExpectNoCOnlyVocabulary(LastError());

  // max_iterations above the Excel limit.
  fm_iterative_settings_patch_t iterative = FM_STRUCT_INIT(fm_iterative_settings_patch_t);
  iterative.max_iterations_op = FM_PATCH_OP_SET;
  iterative.max_iterations = 40000U;
  EXPECT_EQ(fm_workbook_update_iterative(wb.handle, &iterative), kInvalidArgument);
  EXPECT_EQ(LastError(), "max_iterations must be in 1..32767");
  ExpectNoCOnlyVocabulary(LastError());
}

TEST(FormulonCApiArgumentMessages, SheetIndexNamesTheFieldAndTheSheetCount) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t cells = 0;
  EXPECT_EQ(fm_sheet_get_cell_count(wb.handle, 99, &cells), kSheetIndexOutOfRange);
  EXPECT_EQ(LastError(), "sheet_index 99 is out of range (1 sheet)");
  ExpectNoCOnlyVocabulary(LastError());
  EXPECT_EQ(FirstToken(LastError()), "sheet_index");
}

TEST(FormulonCApiArgumentMessages, BindingMessageBuildersComposeThePath) {
  using formulon::binding::argument_message;
  using formulon::binding::element_path;
  using formulon::binding::field_path;
  using formulon::binding::record_param;

  EXPECT_EQ(argument_message("Sheet.updateView", "patch.zoomPct", "must be an integer"),
            "Sheet.updateView patch.zoomPct must be an integer");
  EXPECT_EQ(argument_message("Sheet.addTable", "", "takes at most 1 argument"),
            "Sheet.addTable takes at most 1 argument");

  EXPECT_EQ(field_path("patch", "zoomPct"), "patch.zoomPct");
  EXPECT_EQ(field_path("", "zoomPct"), "zoomPct");

  EXPECT_EQ(element_path("patch.sparklines", 2), "patch.sparklines[2]");
  EXPECT_EQ(field_path(element_path("patch.sparklines", 2), "location"), "patch.sparklines[2].location");

  EXPECT_EQ(record_param("Sheet.updateView"), "patch");
  EXPECT_EQ(record_param("WorkbookOoxml.setFont"), "record");
  EXPECT_EQ(record_param("Sheet.noSuchCall"), "");
  EXPECT_EQ(record_param(""), "");
}
