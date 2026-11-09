// C ABI pivot patch records: the style options patch, the field and item
// patches (validated as a whole, applied together), and the date-group spec.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <vector>

#include "formulon_c_pivot_test_helpers.h"

namespace {

constexpr std::size_t kPivot = 0U;
constexpr std::size_t kRegionField = 0U;
constexpr std::size_t kAmountField = 1U;

/// A saved and reloaded scratch pivot: Region (row field, items North and
/// South) and Amount summed as "Sum of Amount".
void LoadScratchPivot(WorkbookGuard& out) {
  WorkbookGuard built;
  ASSERT_EQ(fm_workbook_create_default(&built.handle), FM_STATUS_OK);
  ASSERT_EQ(BuildScratchPivot(built.handle, nullptr, nullptr), FM_STATUS_OK) << fm_get_last_error_message();
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(built.handle, &saved.data, &saved.len), FM_STATUS_OK) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &out.handle), FM_STATUS_OK) << fm_get_last_error_message();
}

struct StyleSnapshot {
  std::string style_name;
  int32_t present;
  int32_t row_headers;
  int32_t column_headers;
  int32_t row_stripes;
  int32_t column_stripes;
  int32_t last_column;

  bool operator==(const StyleSnapshot& other) const {
    return style_name == other.style_name && present == other.present && row_headers == other.row_headers &&
           column_headers == other.column_headers && row_stripes == other.row_stripes &&
           column_stripes == other.column_stripes && last_column == other.last_column;
  }
};

StyleSnapshot Style(const fm_workbook_t* wb) {
  fm_pivot_style_options_t style = FM_STRUCT_INIT(fm_pivot_style_options_t);
  int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_pivot_style_options(wb, 0, kPivot, &present, &style), FM_STATUS_OK);
  return StyleSnapshot{style.style_name != nullptr ? style.style_name : "",
                       present,
                       style.show_row_headers,
                       style.show_column_headers,
                       style.show_row_stripes,
                       style.show_column_stripes,
                       style.show_last_column};
}

struct FieldSnapshot {
  std::string custom_name;
  std::string number_format;
  std::string sort_by_field;
  int32_t sort_order;
  int32_t default_subtotal;
  int32_t insert_blank_row;
  int32_t fill_down_labels;
  int32_t show_all;
  int32_t subtotal_top;
  int32_t show_detail;

  bool operator==(const FieldSnapshot& other) const {
    return custom_name == other.custom_name && number_format == other.number_format &&
           sort_by_field == other.sort_by_field && sort_order == other.sort_order &&
           default_subtotal == other.default_subtotal && insert_blank_row == other.insert_blank_row &&
           fill_down_labels == other.fill_down_labels && show_all == other.show_all &&
           subtotal_top == other.subtotal_top && show_detail == other.show_detail;
  }
};

FieldSnapshot Field(const fm_workbook_t* wb, std::size_t field = kRegionField) {
  fm_pivot_field_info_t info = FM_STRUCT_INIT(fm_pivot_field_info_t);
  EXPECT_EQ(fm_sheet_get_pivot_field_at(wb, 0, kPivot, field, &info), FM_STATUS_OK);
  return FieldSnapshot{info.custom_name != nullptr ? info.custom_name : "",
                       info.number_format != nullptr ? info.number_format : "",
                       info.sort_by_field != nullptr ? info.sort_by_field : "",
                       info.sort_order,
                       info.default_subtotal,
                       info.insert_blank_row,
                       info.fill_down_labels,
                       info.show_all,
                       info.subtotal_top,
                       info.show_detail};
}

struct ItemSnapshot {
  int32_t visible;
  int32_t show_detail;
  std::string caption;

  bool operator==(const ItemSnapshot& other) const {
    return visible == other.visible && show_detail == other.show_detail && caption == other.caption;
  }
};

ItemSnapshot Item(const fm_workbook_t* wb, std::size_t item) {
  fm_pivot_item_info_t info = FM_STRUCT_INIT(fm_pivot_item_info_t);
  EXPECT_EQ(fm_sheet_get_pivot_field_item_at(wb, 0, kPivot, kRegionField, item, &info), FM_STATUS_OK);
  return ItemSnapshot{info.visible, info.show_detail, info.caption != nullptr ? info.caption : ""};
}

// ---------------------------------------------------------------------------
// Style options patch
// ---------------------------------------------------------------------------

TEST(FormulonCApiPivotPatch, StylePatchChangesOnlyTheMembersItSets) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  const StyleSnapshot before = Style(wb.handle);
  ASSERT_EQ(before.present, 1);

  fm_pivot_style_options_patch_t patch = FM_STRUCT_INIT(fm_pivot_style_options_patch_t);
  patch.show_row_stripes_op = FM_PATCH_OP_SET;
  patch.show_row_stripes = before.row_stripes == 0 ? 1 : 0;
  ASSERT_EQ(fm_sheet_update_pivot_style_options(wb.handle, 0, kPivot, &patch), FM_STATUS_OK)
      << fm_get_last_error_message();
  StyleSnapshot expected = before;
  expected.row_stripes = patch.show_row_stripes;
  EXPECT_EQ(Style(wb.handle), expected);

  // A style name and a header toggle together.
  patch = FM_STRUCT_INIT(fm_pivot_style_options_patch_t);
  patch.style_name_op = FM_PATCH_OP_SET;
  patch.style_name = "PivotStyleMedium9";
  patch.show_row_headers_op = FM_PATCH_OP_SET;
  patch.show_row_headers = 0;
  ASSERT_EQ(fm_sheet_update_pivot_style_options(wb.handle, 0, kPivot, &patch), FM_STATUS_OK)
      << fm_get_last_error_message();
  expected.style_name = "PivotStyleMedium9";
  expected.row_headers = 0;
  EXPECT_EQ(Style(wb.handle), expected);
}

TEST(FormulonCApiPivotPatch, StylePatchClearOnTheNameKeepsTheElementWithoutAStyle) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  fm_pivot_style_options_patch_t patch = FM_STRUCT_INIT(fm_pivot_style_options_patch_t);
  patch.style_name_op = FM_PATCH_OP_SET;
  patch.style_name = "PivotStyleMedium9";
  ASSERT_EQ(fm_sheet_update_pivot_style_options(wb.handle, 0, kPivot, &patch), FM_STATUS_OK)
      << fm_get_last_error_message();
  const StyleSnapshot named = Style(wb.handle);
  ASSERT_EQ(named.style_name, "PivotStyleMedium9");

  patch = FM_STRUCT_INIT(fm_pivot_style_options_patch_t);
  patch.style_name_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_pivot_style_options(wb.handle, 0, kPivot, &patch), FM_STATUS_OK)
      << fm_get_last_error_message();
  StyleSnapshot expected = named;
  expected.style_name = "";
  EXPECT_EQ(Style(wb.handle), expected);
  EXPECT_EQ(Style(wb.handle).present, 1);
}

TEST(FormulonCApiPivotPatch, AllKeepStylePatchChangesNothing) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  const StyleSnapshot before = Style(wb.handle);
  const fm_pivot_style_options_patch_t keep = FM_STRUCT_INIT(fm_pivot_style_options_patch_t);
  ASSERT_EQ(fm_sheet_update_pivot_style_options(wb.handle, 0, kPivot, &keep), FM_STATUS_OK)
      << fm_get_last_error_message();
  EXPECT_EQ(Style(wb.handle), before);
  EXPECT_EQ(fm_sheet_update_pivot_style_options(wb.handle, 0, kPivot, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_update_pivot_style_options(nullptr, 0, kPivot, &keep), FM_STATUS_BINDING_NULL_POINTER);
}

TEST(FormulonCApiPivotPatch, StylePatchWithABadMemberChangesNothing) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  const StyleSnapshot before = Style(wb.handle);
  const auto set_style = [] {
    fm_pivot_style_options_patch_t patch = FM_STRUCT_INIT(fm_pivot_style_options_patch_t);
    patch.style_name_op = FM_PATCH_OP_SET;
    patch.style_name = "PivotStyleLight16";
    return patch;
  };

  // CLEAR on a toggle.
  fm_pivot_style_options_patch_t patch = set_style();
  patch.show_column_stripes_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_pivot_style_options(wb.handle, 0, kPivot, &patch), FM_STATUS_INVALID_ARGUMENT);
  // A toggle outside 0 / 1.
  patch = set_style();
  patch.show_last_column_op = FM_PATCH_OP_SET;
  patch.show_last_column = 2;
  EXPECT_EQ(fm_sheet_update_pivot_style_options(wb.handle, 0, kPivot, &patch), FM_STATUS_INVALID_ARGUMENT);
  // SET with no style name.
  patch = FM_STRUCT_INIT(fm_pivot_style_options_patch_t);
  patch.style_name_op = FM_PATCH_OP_SET;
  patch.show_row_headers_op = FM_PATCH_OP_SET;
  patch.show_row_headers = 0;
  EXPECT_EQ(fm_sheet_update_pivot_style_options(wb.handle, 0, kPivot, &patch), FM_STATUS_INVALID_ARGUMENT);
  // An op outside the enumeration.
  patch = set_style();
  patch.show_row_stripes_op = static_cast<fm_patch_op_t>(9);
  EXPECT_EQ(fm_sheet_update_pivot_style_options(wb.handle, 0, kPivot, &patch), FM_STATUS_INVALID_ARGUMENT);

  EXPECT_EQ(Style(wb.handle), before);
}

// ---------------------------------------------------------------------------
// Field patch
// ---------------------------------------------------------------------------

TEST(FormulonCApiPivotPatch, FieldPatchWithABadMemberChangesNothingEvenWhenAnEarlierMemberIsValid) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  const FieldSnapshot before = Field(wb.handle);

  // A custom name another field already carries, beside a valid show_all.
  fm_pivot_field_patch_t patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  patch.show_all_op = FM_PATCH_OP_SET;
  patch.show_all = before.show_all == 0 ? 1 : 0;
  patch.custom_name_op = FM_PATCH_OP_SET;
  patch.custom_name = "Amount";
  EXPECT_EQ(fm_sheet_update_pivot_field(wb.handle, 0, kPivot, kRegionField, &patch), FM_STATUS_ALREADY_EXISTS);
  EXPECT_EQ(Field(wb.handle), before);

  // CLEAR on a flag.
  patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  patch.insert_blank_row_op = FM_PATCH_OP_SET;
  patch.insert_blank_row = 1;
  patch.default_subtotal_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_pivot_field(wb.handle, 0, kPivot, kRegionField, &patch), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(Field(wb.handle), before);

  // A flag outside 0 / 1.
  patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  patch.show_all_op = FM_PATCH_OP_SET;
  patch.show_all = 2;
  EXPECT_EQ(fm_sheet_update_pivot_field(wb.handle, 0, kPivot, kRegionField, &patch), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(Field(wb.handle), before);

  // A sort order outside the enumeration, and MANUAL with a by-field key.
  patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  patch.fill_down_labels_op = FM_PATCH_OP_SET;
  patch.fill_down_labels = 1;
  patch.sort_order_op = FM_PATCH_OP_SET;
  patch.sort_order = 99;
  EXPECT_EQ(fm_sheet_update_pivot_field(wb.handle, 0, kPivot, kRegionField, &patch), FM_STATUS_INVALID_ARGUMENT);
  patch.sort_order = FM_PIVOT_SORT_ORDER_MANUAL;
  patch.sort_by_field_op = FM_PATCH_OP_SET;
  patch.sort_by_field = "Sum of Amount";
  EXPECT_EQ(fm_sheet_update_pivot_field(wb.handle, 0, kPivot, kRegionField, &patch), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(Field(wb.handle), before);

  // SET with a string member left NULL.
  patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  patch.show_all_op = FM_PATCH_OP_SET;
  patch.show_all = 1;
  patch.custom_name_op = FM_PATCH_OP_SET;
  patch.custom_name = nullptr;
  EXPECT_EQ(fm_sheet_update_pivot_field(wb.handle, 0, kPivot, kRegionField, &patch), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(Field(wb.handle), before);
}

TEST(FormulonCApiPivotPatch, FieldPatchAppliesEveryMemberTogether) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  fm_pivot_field_patch_t patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  patch.custom_name_op = FM_PATCH_OP_SET;
  patch.custom_name = "Area";
  patch.number_format_op = FM_PATCH_OP_SET;
  patch.number_format = "0.00";
  patch.default_subtotal_op = FM_PATCH_OP_SET;
  patch.default_subtotal = 0;
  patch.insert_blank_row_op = FM_PATCH_OP_SET;
  patch.insert_blank_row = 1;
  patch.fill_down_labels_op = FM_PATCH_OP_SET;
  patch.fill_down_labels = 1;
  patch.show_all_op = FM_PATCH_OP_SET;
  patch.show_all = 1;
  patch.subtotal_top_op = FM_PATCH_OP_SET;
  patch.subtotal_top = 1;
  ASSERT_EQ(fm_sheet_update_pivot_field(wb.handle, 0, kPivot, kRegionField, &patch), FM_STATUS_OK)
      << fm_get_last_error_message();
  const FieldSnapshot after = Field(wb.handle);
  EXPECT_EQ(after.custom_name, "Area");
  EXPECT_EQ(after.number_format, "0.00");
  EXPECT_EQ(after.default_subtotal, 0);
  EXPECT_EQ(after.insert_blank_row, 1);
  EXPECT_EQ(after.fill_down_labels, 1);
  EXPECT_EQ(after.show_all, 1);
  EXPECT_EQ(after.subtotal_top, FM_TRISTATE_TRUE);

  // CLEAR restores the source name and removes the format; the flags stay.
  patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  patch.custom_name_op = FM_PATCH_OP_CLEAR;
  patch.number_format_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_pivot_field(wb.handle, 0, kPivot, kRegionField, &patch), FM_STATUS_OK)
      << fm_get_last_error_message();
  FieldSnapshot cleared = Field(wb.handle);
  EXPECT_EQ(cleared.custom_name, "");
  EXPECT_EQ(cleared.number_format, "");
  EXPECT_EQ(cleared.insert_blank_row, 1);
  EXPECT_EQ(cleared.show_all, 1);
}

TEST(FormulonCApiPivotPatch, FieldPatchSortAndShowDetailRoundTrip) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  fm_pivot_field_patch_t patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  patch.sort_order_op = FM_PATCH_OP_SET;
  patch.sort_order = FM_PIVOT_SORT_ORDER_DESCENDING;
  patch.sort_by_field_op = FM_PATCH_OP_SET;
  patch.sort_by_field = "Sum of Amount";
  ASSERT_EQ(fm_sheet_update_pivot_field(wb.handle, 0, kPivot, kRegionField, &patch), FM_STATUS_OK)
      << fm_get_last_error_message();
  FieldSnapshot field = Field(wb.handle);
  EXPECT_EQ(field.sort_order, FM_PIVOT_SORT_ORDER_DESCENDING);
  EXPECT_EQ(field.sort_by_field, "Sum of Amount");

  // CLEAR on the by-field key sorts by label and keeps the direction.
  patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  patch.sort_by_field_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_pivot_field(wb.handle, 0, kPivot, kRegionField, &patch), FM_STATUS_OK)
      << fm_get_last_error_message();
  field = Field(wb.handle);
  EXPECT_EQ(field.sort_by_field, "");
  EXPECT_EQ(field.sort_order, FM_PIVOT_SORT_ORDER_DESCENDING);

  // MANUAL takes no by-field key.
  patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  patch.sort_order_op = FM_PATCH_OP_SET;
  patch.sort_order = FM_PIVOT_SORT_ORDER_MANUAL;
  ASSERT_EQ(fm_sheet_update_pivot_field(wb.handle, 0, kPivot, kRegionField, &patch), FM_STATUS_OK)
      << fm_get_last_error_message();
  EXPECT_EQ(Field(wb.handle).sort_order, FM_PIVOT_SORT_ORDER_MANUAL);

  // Show detail collapses every item and reads back on the field.
  patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  patch.show_detail_op = FM_PATCH_OP_SET;
  patch.show_detail = 0;
  ASSERT_EQ(fm_sheet_update_pivot_field(wb.handle, 0, kPivot, kRegionField, &patch), FM_STATUS_OK)
      << fm_get_last_error_message();
  EXPECT_EQ(Field(wb.handle).show_detail, 0);
  EXPECT_EQ(Item(wb.handle, 0).show_detail, 0);
  EXPECT_EQ(Item(wb.handle, 1).show_detail, 0);
  patch.show_detail = 1;
  ASSERT_EQ(fm_sheet_update_pivot_field(wb.handle, 0, kPivot, kRegionField, &patch), FM_STATUS_OK)
      << fm_get_last_error_message();
  EXPECT_EQ(Field(wb.handle).show_detail, 1);
  EXPECT_EQ(Item(wb.handle, 0).show_detail, 1);
  EXPECT_EQ(Item(wb.handle, 1).show_detail, 1);
}

TEST(FormulonCApiPivotPatch, AllKeepFieldPatchChangesNothingAndNullIsRefused) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  const FieldSnapshot region = Field(wb.handle, kRegionField);
  const FieldSnapshot amount = Field(wb.handle, kAmountField);
  const fm_pivot_field_patch_t keep = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  ASSERT_EQ(fm_sheet_update_pivot_field(wb.handle, 0, kPivot, kRegionField, &keep), FM_STATUS_OK)
      << fm_get_last_error_message();
  EXPECT_EQ(Field(wb.handle, kRegionField), region);
  EXPECT_EQ(Field(wb.handle, kAmountField), amount);
  EXPECT_EQ(fm_sheet_update_pivot_field(wb.handle, 0, kPivot, kRegionField, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_update_pivot_field(nullptr, 0, kPivot, kRegionField, &keep), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_update_pivot_field(wb.handle, 9, kPivot, kRegionField, &keep), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

// ---------------------------------------------------------------------------
// Item patch
// ---------------------------------------------------------------------------

TEST(FormulonCApiPivotPatch, ItemPatchWithABadMemberChangesNothingEvenWhenAnEarlierMemberIsValid) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  const ItemSnapshot before = Item(wb.handle, 0);

  // A caption the other item already shows, beside a valid visibility change.
  fm_pivot_item_patch_t patch = FM_STRUCT_INIT(fm_pivot_item_patch_t);
  patch.visible_op = FM_PATCH_OP_SET;
  patch.visible = 0;
  patch.caption_op = FM_PATCH_OP_SET;
  patch.caption = "South";
  EXPECT_EQ(fm_sheet_update_pivot_field_item(wb.handle, 0, kPivot, kRegionField, 0, &patch), FM_STATUS_ALREADY_EXISTS);
  EXPECT_EQ(Item(wb.handle, 0), before);

  // CLEAR on a flag, and a flag outside 0 / 1.
  patch = FM_STRUCT_INIT(fm_pivot_item_patch_t);
  patch.caption_op = FM_PATCH_OP_SET;
  patch.caption = "Nord";
  patch.visible_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_pivot_field_item(wb.handle, 0, kPivot, kRegionField, 0, &patch),
            FM_STATUS_INVALID_ARGUMENT);
  patch.visible_op = FM_PATCH_OP_SET;
  patch.visible = 2;
  EXPECT_EQ(fm_sheet_update_pivot_field_item(wb.handle, 0, kPivot, kRegionField, 0, &patch),
            FM_STATUS_INVALID_ARGUMENT);
  // SET with a NULL caption.
  patch = FM_STRUCT_INIT(fm_pivot_item_patch_t);
  patch.visible_op = FM_PATCH_OP_SET;
  patch.visible = 0;
  patch.caption_op = FM_PATCH_OP_SET;
  patch.caption = nullptr;
  EXPECT_EQ(fm_sheet_update_pivot_field_item(wb.handle, 0, kPivot, kRegionField, 0, &patch),
            FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(Item(wb.handle, 0), before);
}

TEST(FormulonCApiPivotPatch, ItemPatchSetsVisibilityDetailAndCaptionTogether) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  fm_pivot_item_patch_t patch = FM_STRUCT_INIT(fm_pivot_item_patch_t);
  patch.visible_op = FM_PATCH_OP_SET;
  patch.visible = 0;
  patch.show_detail_op = FM_PATCH_OP_SET;
  patch.show_detail = 0;
  patch.caption_op = FM_PATCH_OP_SET;
  patch.caption = "Nord";
  ASSERT_EQ(fm_sheet_update_pivot_field_item(wb.handle, 0, kPivot, kRegionField, 0, &patch), FM_STATUS_OK)
      << fm_get_last_error_message();
  ItemSnapshot item = Item(wb.handle, 0);
  EXPECT_EQ(item.visible, 0);
  EXPECT_EQ(item.show_detail, 0);
  EXPECT_EQ(item.caption, "Nord");
  // The other item is untouched.
  EXPECT_EQ(Item(wb.handle, 1).visible, 1);
  EXPECT_EQ(Item(wb.handle, 1).caption, "");

  // CLEAR removes the caption and leaves the flags.
  patch = FM_STRUCT_INIT(fm_pivot_item_patch_t);
  patch.caption_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_pivot_field_item(wb.handle, 0, kPivot, kRegionField, 0, &patch), FM_STATUS_OK)
      << fm_get_last_error_message();
  item = Item(wb.handle, 0);
  EXPECT_EQ(item.caption, "");
  EXPECT_EQ(item.visible, 0);
  EXPECT_EQ(item.show_detail, 0);

  // An all-KEEP patch changes nothing; a NULL one is refused.
  const fm_pivot_item_patch_t keep = FM_STRUCT_INIT(fm_pivot_item_patch_t);
  ASSERT_EQ(fm_sheet_update_pivot_field_item(wb.handle, 0, kPivot, kRegionField, 0, &keep), FM_STATUS_OK)
      << fm_get_last_error_message();
  EXPECT_EQ(Item(wb.handle, 0), item);
  EXPECT_EQ(fm_sheet_update_pivot_field_item(wb.handle, 0, kPivot, kRegionField, 0, nullptr),
            FM_STATUS_BINDING_NULL_POINTER);
}

// ---------------------------------------------------------------------------
// Date group spec
// ---------------------------------------------------------------------------

constexpr double kFirstDay = 45000.0;
constexpr int kDayCount = 40;

/// A pivot over a cache of `kDayCount` consecutive days with a Date row field
/// (field 0) and a summed Amount.
fm_status_t BuildDatePivot(fm_workbook_t* wb) {
  std::uint32_t cache_id = 0;
  fm_status_t st = fm_workbook_add_pivot_cache(wb, 0, &cache_id);
  if (st != 0) {
    return st;
  }
  st = SetPivotCacheWorksheetSource(wb, cache_id, /*present=*/1, "A1:B41", "Sheet1", nullptr);
  if (st != 0) {
    return st;
  }
  std::size_t date_index = 99;
  std::size_t amount_index = 99;
  if ((st = fm_workbook_add_pivot_cache_field(wb, cache_id, "Date", &date_index)) != 0 ||
      (st = fm_workbook_add_pivot_cache_field(wb, cache_id, "Amount", &amount_index)) != 0) {
    return st;
  }
  for (int day = 0; day < kDayCount; ++day) {
    std::size_t record = 99;
    if ((st = fm_workbook_add_pivot_cache_record(wb, cache_id, &record)) != 0 ||
        (st = fm_workbook_set_pivot_cache_record_number(wb, cache_id, record, date_index, kFirstDay + day)) != 0 ||
        (st = fm_workbook_set_pivot_cache_record_number(wb, cache_id, record, amount_index, 1.0)) != 0) {
      return st;
    }
  }
  std::size_t pivot_index = 99;
  if ((st = AddPivot(wb, 0, "PT", cache_id, 0U, 3U, &pivot_index)) != 0) {
    return st;
  }
  fm_pivot_field_spec_t date_spec = FM_STRUCT_INIT(fm_pivot_field_spec_t);
  date_spec.source_name = "Date";
  date_spec.custom_name = "";
  date_spec.axis = FM_PIVOT_AXIS_ROW;
  date_spec.subtotal_top = FM_TRISTATE_FALSE;
  date_spec.number_format = "";
  std::size_t date_field = 99;
  if ((st = fm_sheet_add_pivot_field(wb, 0, pivot_index, &date_spec, &date_field)) != 0) {
    return st;
  }
  fm_pivot_field_spec_t amount_spec = FM_STRUCT_INIT(fm_pivot_field_spec_t);
  amount_spec.source_name = "Amount";
  amount_spec.custom_name = "";
  amount_spec.axis = FM_PIVOT_AXIS_VALUE;
  amount_spec.subtotal_top = FM_TRISTATE_FALSE;
  amount_spec.number_format = "";
  std::size_t amount_field = 99;
  if ((st = fm_sheet_add_pivot_field(wb, 0, pivot_index, &amount_spec, &amount_field)) != 0) {
    return st;
  }
  const std::uint32_t row_order[] = {static_cast<std::uint32_t>(date_field)};
  if ((st = fm_sheet_set_pivot_row_field_order(wb, 0, pivot_index, row_order, 1U)) != 0) {
    return st;
  }
  fm_pivot_data_field_t data_spec = FM_STRUCT_INIT(fm_pivot_data_field_t);
  data_spec.name = "Sum of Amount";
  data_spec.field_index = static_cast<std::uint32_t>(amount_field);
  data_spec.aggregation = FM_PIVOT_AGGREGATION_SUM;
  data_spec.number_format = "";
  data_spec.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;
  data_spec.has_show_as_base_field = 0;
  data_spec.has_show_as_base_item = 0;
  std::size_t data_index = 99;
  return fm_sheet_add_pivot_data_field(wb, 0, pivot_index, &data_spec, &data_index);
}

fm_pivot_date_group_spec_t WeekGroup() {
  fm_pivot_date_group_spec_t group = FM_STRUCT_INIT(fm_pivot_date_group_spec_t);
  group.periods = FM_PIVOT_PERIOD_DAYS;
  group.calendar = FM_PIVOT_CALENDAR_GREGORIAN;
  group.interval_days = 7;
  return group;
}

std::vector<std::string> RowLabels(const fm_workbook_t* wb) {
  PivotCellsGuard projected;
  EXPECT_EQ(fm_sheet_get_pivot_layout(wb, 0, kPivot, &projected.handle), FM_STATUS_OK) << fm_get_last_error_message();
  std::vector<std::string> labels;
  for (const fm_pivot_cell_info_t& cell : CollectCells(projected.handle)) {
    if (cell.kind == FM_PIVOT_CELL_KIND_ROW_LABEL && cell.value.kind == FM_VALUE_KIND_TEXT) {
      labels.emplace_back(cell.value.u.text);
    }
  }
  return labels;
}

std::uint32_t DatePeriods(const fm_workbook_t* wb) {
  fm_pivot_field_info_t info = FM_STRUCT_INIT(fm_pivot_field_info_t);
  EXPECT_EQ(fm_sheet_get_pivot_field_at(wb, 0, kPivot, 0, &info), FM_STATUS_OK);
  return info.date_periods;
}

TEST(FormulonCApiPivotPatch, DateGroupWithoutABoundStartsAtTheFirstValue) {
  WorkbookGuard open_start;
  ASSERT_EQ(fm_workbook_create_default(&open_start.handle), FM_STATUS_OK);
  ASSERT_EQ(BuildDatePivot(open_start.handle), FM_STATUS_OK) << fm_get_last_error_message();
  fm_pivot_date_group_spec_t group = WeekGroup();
  ASSERT_EQ(fm_sheet_set_pivot_field_date_group(open_start.handle, 0, kPivot, 0, &group), FM_STATUS_OK)
      << fm_get_last_error_message();
  EXPECT_EQ(DatePeriods(open_start.handle), FM_PIVOT_PERIOD_DAYS);

  // Stating the first value as the start groups the same way.
  WorkbookGuard stated_start;
  ASSERT_EQ(fm_workbook_create_default(&stated_start.handle), FM_STATUS_OK);
  ASSERT_EQ(BuildDatePivot(stated_start.handle), FM_STATUS_OK) << fm_get_last_error_message();
  group.has_start = 1;
  group.start = kFirstDay;
  ASSERT_EQ(fm_sheet_set_pivot_field_date_group(stated_start.handle, 0, kPivot, 0, &group), FM_STATUS_OK)
      << fm_get_last_error_message();
  const std::vector<std::string> open_labels = RowLabels(open_start.handle);
  EXPECT_FALSE(open_labels.empty());
  EXPECT_EQ(open_labels, RowLabels(stated_start.handle));

  // A later stated start moves the first group.
  WorkbookGuard later_start;
  ASSERT_EQ(fm_workbook_create_default(&later_start.handle), FM_STATUS_OK);
  ASSERT_EQ(BuildDatePivot(later_start.handle), FM_STATUS_OK) << fm_get_last_error_message();
  group.start = kFirstDay + 3.0;
  ASSERT_EQ(fm_sheet_set_pivot_field_date_group(later_start.handle, 0, kPivot, 0, &group), FM_STATUS_OK)
      << fm_get_last_error_message();
  EXPECT_NE(RowLabels(later_start.handle), open_labels);
}

TEST(FormulonCApiPivotPatch, DateGroupRefusesABadBoundFlagOrNaNAndGroupsNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  ASSERT_EQ(BuildDatePivot(wb.handle), FM_STATUS_OK) << fm_get_last_error_message();
  const double nan = std::numeric_limits<double>::quiet_NaN();
  const double inf = std::numeric_limits<double>::infinity();

  fm_pivot_date_group_spec_t group = WeekGroup();
  group.has_start = 1;
  group.start = nan;
  EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, kPivot, 0, &group), FM_STATUS_INVALID_ARGUMENT);
  group = WeekGroup();
  group.has_end = 1;
  group.end = nan;
  EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, kPivot, 0, &group), FM_STATUS_INVALID_ARGUMENT);
  group = WeekGroup();
  group.has_end = 1;
  group.end = inf;
  EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, kPivot, 0, &group), FM_STATUS_INVALID_ARGUMENT);
  group = WeekGroup();
  group.has_start = 2;
  EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, kPivot, 0, &group), FM_STATUS_INVALID_ARGUMENT);
  group = WeekGroup();
  group.has_end = -1;
  EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, kPivot, 0, &group), FM_STATUS_INVALID_ARGUMENT);
  group = WeekGroup();
  group.has_start = 1;
  group.start = kFirstDay + 10.0;
  group.has_end = 1;
  group.end = kFirstDay;
  EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, kPivot, 0, &group), FM_STATUS_INVALID_ARGUMENT);
  group = WeekGroup();
  group.interval_days = 0;
  EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, kPivot, 0, &group), FM_STATUS_INVALID_ARGUMENT);
  group = WeekGroup();
  group.periods = 0;
  EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, kPivot, 0, &group), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, kPivot, 0, nullptr), FM_STATUS_OK);

  EXPECT_EQ(DatePeriods(wb.handle), 0U);
}

TEST(FormulonCApiPivotPatch, DateGroupOfAPivotOrFieldPastTheEndIsOutOfRange) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  ASSERT_EQ(BuildDatePivot(wb.handle), FM_STATUS_OK) << fm_get_last_error_message();
  const fm_pivot_date_group_spec_t group = WeekGroup();
  EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, kPivot + 1U, 0, &group), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, kPivot, 99U, &group), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, kPivot + 1U, 0, nullptr), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, kPivot, 99U, nullptr), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 9, kPivot, 0, &group), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(DatePeriods(wb.handle), 0U);
}

}  // namespace
