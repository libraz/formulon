// Stable C ABI pivot report options: grand total caption, autofit flag, field
// display flags, item captions and expansion, the Values level in the field
// orders, Show Details, the pivot and cache options structs, field names and
// showAll, and default data-field names. Every setter must leave the table
// layout-dirty and reject out-of-range indices and malformed text.

#include <cstddef>
#include <cstdint>
#include <functional>
#include <limits>
#include <string>
#include <vector>

#include "formulon_c_out_index.h"
#include "formulon_c_pivot_test_helpers.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "workbook.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kBindingUtf8EncodingError =
    static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingUtf8EncodingError);
constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr fm_status_t kPreconditionFailed = static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed);

constexpr std::size_t kPivot = 0U;
constexpr std::size_t kRegionField = 0U;
constexpr std::size_t kAmountField = 1U;

formulon::pivot::PivotTable& Table(WorkbookGuard& wb, std::size_t sheet = 0U) {
  return *wb.handle->wb->sheet(sheet).mutable_pivot_tables()[kPivot];
}

// The scratch pivot saved and loaded again, so its table starts layout-clean.
void LoadScratchPivot(WorkbookGuard& out) {
  WorkbookGuard built;
  ASSERT_EQ(fm_workbook_create_default(&built.handle), 0);
  ASSERT_EQ(BuildScratchPivot(built.handle, nullptr, nullptr), 0) << fm_get_last_error_message();
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(built.handle, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &out.handle), 0) << fm_get_last_error_message();
  ASSERT_EQ(out.handle->wb->sheet(0U).pivot_tables().size(), 1U);
}

// Adds a second data field so the Values level exists.
void AddSecondDataField(fm_workbook_t* wb) {
  fm_pivot_data_field_t spec = FM_STRUCT_INIT(fm_pivot_data_field_t);
  spec.name = "Count of Amount";
  spec.field_index = static_cast<std::uint32_t>(kAmountField);
  spec.aggregation = FM_PIVOT_AGGREGATION_COUNT;
  spec.number_format = "";
  spec.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;
  spec.has_show_as_base_field = 0;
  spec.has_show_as_base_item = 0;
  std::size_t index = 99;
  ASSERT_EQ(fm_sheet_add_pivot_data_field(wb, 0U, kPivot, &spec, &index), 0) << fm_get_last_error_message();
}

// Runs `mutation` on a layout-clean pivot and expects it to succeed and dirty the table.
void ExpectMarksDirty(const std::function<fm_status_t(fm_workbook_t*)>& mutation) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  ASSERT_FALSE(Table(wb).layout_dirty());
  ASSERT_EQ(mutation(wb.handle), 0) << fm_get_last_error_message();
  EXPECT_TRUE(Table(wb).layout_dirty());
}

std::vector<std::uint32_t> ReadOrder(const fm_workbook_t* wb, bool row) {
  const auto read = [&](std::uint32_t* out, std::size_t capacity, std::size_t* count) {
    return row ? fm_sheet_get_pivot_row_field_order(wb, 0U, kPivot, out, capacity, count)
               : fm_sheet_get_pivot_column_field_order(wb, 0U, kPivot, out, capacity, count);
  };
  std::size_t count = 0;
  fm_status_t status = read(nullptr, 0U, &count);
  EXPECT_EQ(status, 0) << fm_get_last_error_message();
  std::vector<std::uint32_t> indices(count);
  if (status == 0 && count != 0U) {
    status = read(indices.data(), indices.size(), &count);
    EXPECT_EQ(status, 0) << fm_get_last_error_message();
  }
  return status == 0 ? indices : std::vector<std::uint32_t>{};
}

fm_status_t SetOrder(fm_workbook_t* wb, bool row, std::vector<std::uint32_t> order) {
  return row ? fm_sheet_set_pivot_row_field_order(wb, 0U, kPivot, order.data(), order.size())
             : fm_sheet_set_pivot_column_field_order(wb, 0U, kPivot, order.data(), order.size());
}

std::string Repeat(const std::string& unit, std::size_t times) {
  std::string out;
  for (std::size_t i = 0; i < times; ++i) {
    out += unit;
  }
  return out;
}

const char* const kEmoji = "\xF0\x9F\x98\x80";  // two UTF-16 code units

/// `fm_sheet_update_pivot_options` writing every member of `record`: a NULL
/// string clears its option, anything else is SET. A NULL `record` passes a
/// NULL patch.
fm_status_t UpdateOptionsFrom(fm_workbook_t* wb, std::size_t sheet, std::size_t pivot,
                              const fm_pivot_options_t* record) {
  if (record == nullptr) {
    return fm_sheet_update_pivot_options(wb, sheet, pivot, nullptr);
  }
  const fm_pivot_options_t& o = *record;
  fm_pivot_options_patch_t p = FM_STRUCT_INIT(fm_pivot_options_patch_t);
  const auto text = [](const char* value, fm_patch_op_t* op, const char** out) {
    *op = value != nullptr ? FM_PATCH_OP_SET : FM_PATCH_OP_CLEAR;
    *out = value;
  };
#define FORMULON_SET_OPTION(member) \
  p.member##_op = FM_PATCH_OP_SET;  \
  p.member = o.member
  FORMULON_SET_OPTION(show_error);
  FORMULON_SET_OPTION(show_missing);
  FORMULON_SET_OPTION(show_headers);
  FORMULON_SET_OPTION(show_drill);
  FORMULON_SET_OPTION(show_data_tips);
  FORMULON_SET_OPTION(grid_drop_zones);
  FORMULON_SET_OPTION(show_empty_row);
  FORMULON_SET_OPTION(show_empty_column);
  FORMULON_SET_OPTION(show_data_drop_down);
  FORMULON_SET_OPTION(print_drill);
  FORMULON_SET_OPTION(field_print_titles);
  FORMULON_SET_OPTION(item_print_titles);
  FORMULON_SET_OPTION(merge_item);
  FORMULON_SET_OPTION(indent);
  FORMULON_SET_OPTION(page_over_then_down);
  FORMULON_SET_OPTION(page_wrap);
  FORMULON_SET_OPTION(multiple_field_filters);
  FORMULON_SET_OPTION(preserve_formatting);
  FORMULON_SET_OPTION(custom_list_sort);
  FORMULON_SET_OPTION(auto_formatting);
  FORMULON_SET_OPTION(enable_drill);
#undef FORMULON_SET_OPTION
  text(o.error_caption, &p.error_caption_op, &p.error_caption);
  text(o.missing_caption, &p.missing_caption_op, &p.missing_caption);
  text(o.alt_text, &p.alt_text_op, &p.alt_text);
  text(o.alt_text_summary, &p.alt_text_summary_op, &p.alt_text_summary);
  text(o.grand_total_caption, &p.grand_total_caption_op, &p.grand_total_caption);
  return fm_sheet_update_pivot_options(wb, sheet, pivot, &p);
}

/// `fm_workbook_update_pivot_cache_options` writing every member of `record`;
/// a record without `has_missing_items_limit` clears the limit.
fm_status_t UpdateCacheOptionsFrom(fm_workbook_t* wb, std::uint32_t cache_id, const fm_pivot_cache_options_t* record) {
  if (record == nullptr) {
    return fm_workbook_update_pivot_cache_options(wb, cache_id, nullptr);
  }
  fm_pivot_cache_options_patch_t p = FM_STRUCT_INIT(fm_pivot_cache_options_patch_t);
  p.save_data_op = FM_PATCH_OP_SET;
  p.save_data = record->save_data;
  p.missing_items_limit_op = record->has_missing_items_limit == 0 ? FM_PATCH_OP_CLEAR : FM_PATCH_OP_SET;
  p.missing_items_limit = record->missing_items_limit;
  p.refresh_on_load_op = FM_PATCH_OP_SET;
  p.refresh_on_load = record->refresh_on_load;
  return fm_workbook_update_pivot_cache_options(wb, cache_id, &p);
}

// Read options with every string copied into owned storage. The getter's
// strings are read-scratch views that the next successful read may invalidate.
class OwnedPivotOptions {
 public:
  explicit OwnedPivotOptions(const fm_pivot_options_t& read) : options_(read) {
    const char* fm_pivot_options_t::*const members[] = {
        &fm_pivot_options_t::error_caption, &fm_pivot_options_t::missing_caption, &fm_pivot_options_t::alt_text,
        &fm_pivot_options_t::alt_text_summary, &fm_pivot_options_t::grand_total_caption};
    static_assert(sizeof(members) / sizeof(members[0]) == kStrings, "one owned copy per string member");
    for (std::size_t i = 0; i < kStrings; ++i) {
      if (read.*members[i] != nullptr) {
        texts_[i] = read.*members[i];
        options_.*members[i] = texts_[i].c_str();
      }
    }
  }
  OwnedPivotOptions(const OwnedPivotOptions&) = delete;
  OwnedPivotOptions& operator=(const OwnedPivotOptions&) = delete;

  const fm_pivot_options_t* get() const { return &options_; }

 private:
  static constexpr std::size_t kStrings = 5;
  fm_pivot_options_t options_;
  std::string texts_[kStrings];
};

}  // namespace

// ---------------------------------------------------------------------------
// Table level
// ---------------------------------------------------------------------------

// Read-modify-write of one options member through the struct API.
fm_status_t SetGrandTotalCaption(fm_workbook_t* wb, std::size_t sheet, std::size_t pivot, const char* caption) {
  fm_pivot_options_t o = FM_STRUCT_INIT(fm_pivot_options_t);
  const fm_status_t rc = fm_sheet_get_pivot_options(wb, sheet, pivot, &o);
  if (rc != 0) {
    return rc;
  }
  o.grand_total_caption = caption;
  return UpdateOptionsFrom(wb, sheet, pivot, &o);
}

fm_status_t SetAutoFormatting(fm_workbook_t* wb, std::size_t sheet, std::size_t pivot, std::int32_t on) {
  fm_pivot_options_t o = FM_STRUCT_INIT(fm_pivot_options_t);
  const fm_status_t rc = fm_sheet_get_pivot_options(wb, sheet, pivot, &o);
  if (rc != 0) {
    return rc;
  }
  o.auto_formatting = on;
  return UpdateOptionsFrom(wb, sheet, pivot, &o);
}

TEST(FormulonCApiPivotOptions, GrandTotalCaptionSetsGetsAndClears) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  fm_pivot_options_t o = FM_STRUCT_INIT(fm_pivot_options_t);
  ASSERT_EQ(fm_sheet_get_pivot_options(wb.handle, 0U, kPivot, &o), 0);
  EXPECT_EQ(o.grand_total_caption, nullptr);

  ASSERT_EQ(SetGrandTotalCaption(wb.handle, 0U, kPivot, "Total \xE5\x90\x88\xE8\xA8\x88"), 0);
  ASSERT_EQ(fm_sheet_get_pivot_options(wb.handle, 0U, kPivot, &o), 0);
  EXPECT_STREQ(o.grand_total_caption, "Total \xE5\x90\x88\xE8\xA8\x88");
  EXPECT_TRUE(LayoutHasText(wb.handle, kPivot, "Total \xE5\x90\x88\xE8\xA8\x88"));

  ASSERT_EQ(SetGrandTotalCaption(wb.handle, 0U, kPivot, nullptr), 0);
  ASSERT_EQ(fm_sheet_get_pivot_options(wb.handle, 0U, kPivot, &o), 0);
  EXPECT_EQ(o.grand_total_caption, nullptr);

  ASSERT_EQ(SetGrandTotalCaption(wb.handle, 0U, kPivot, "All"), 0);
  ASSERT_EQ(SetGrandTotalCaption(wb.handle, 0U, kPivot, ""), 0);
  ASSERT_EQ(fm_sheet_get_pivot_options(wb.handle, 0U, kPivot, &o), 0);
  EXPECT_EQ(o.grand_total_caption, nullptr);
}

TEST(FormulonCApiPivotOptions, GrandTotalCaptionBoundaries) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  // 255 UTF-16 units fit; 256 do not, and a surrogate pair counts as two.
  const std::string at_limit = Repeat(kEmoji, 127) + "a";
  ASSERT_EQ(SetGrandTotalCaption(wb.handle, 0U, kPivot, at_limit.c_str()), 0) << fm_get_last_error_message();
  fm_pivot_options_t o = FM_STRUCT_INIT(fm_pivot_options_t);
  ASSERT_EQ(fm_sheet_get_pivot_options(wb.handle, 0U, kPivot, &o), 0);
  EXPECT_EQ(std::string(o.grand_total_caption), at_limit);
  const OwnedPivotOptions keep(o);

  EXPECT_EQ(SetGrandTotalCaption(wb.handle, 0U, kPivot, Repeat(kEmoji, 128).c_str()), kInvalidArgument);
  EXPECT_EQ(SetGrandTotalCaption(wb.handle, 0U, kPivot, Repeat("a", 256).c_str()), kInvalidArgument);
  EXPECT_EQ(SetGrandTotalCaption(wb.handle, 0U, kPivot, "\xC3"), kBindingUtf8EncodingError);
  EXPECT_EQ(UpdateOptionsFrom(wb.handle, 9U, kPivot, keep.get()), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(UpdateOptionsFrom(wb.handle, 0U, 9U, keep.get()), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(UpdateOptionsFrom(wb.handle, 0U, std::numeric_limits<std::size_t>::max(), keep.get()),
            FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_pivot_options(wb.handle, 0U, std::numeric_limits<std::size_t>::max(), &o),
            FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(UpdateOptionsFrom(nullptr, 0U, kPivot, keep.get()), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_pivot_options(wb.handle, 0U, kPivot, nullptr), kBindingNullPointer);
  // The rejected calls left the accepted caption alone.
  ASSERT_EQ(fm_sheet_get_pivot_options(wb.handle, 0U, kPivot, &o), 0);
  EXPECT_EQ(std::string(o.grand_total_caption), at_limit);
}

TEST(FormulonCApiPivotOptions, GrandTotalCaptionMarksLayoutDirty) {
  ExpectMarksDirty([](fm_workbook_t* wb) { return SetGrandTotalCaption(wb, 0U, kPivot, "All"); });
}

TEST(FormulonCApiPivotOptions, AutoFormattingSetsAndGets) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  fm_pivot_options_t o = FM_STRUCT_INIT(fm_pivot_options_t);
  ASSERT_EQ(SetAutoFormatting(wb.handle, 0U, kPivot, 0), 0);
  ASSERT_EQ(fm_sheet_get_pivot_options(wb.handle, 0U, kPivot, &o), 0);
  EXPECT_EQ(o.auto_formatting, 0);
  ASSERT_EQ(SetAutoFormatting(wb.handle, 0U, kPivot, 7), 0);
  ASSERT_EQ(fm_sheet_get_pivot_options(wb.handle, 0U, kPivot, &o), 0);
  EXPECT_EQ(o.auto_formatting, 1);

  EXPECT_EQ(SetAutoFormatting(wb.handle, 1U, kPivot, 1), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(SetAutoFormatting(wb.handle, 0U, 1U, 1), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(SetAutoFormatting(nullptr, 0U, kPivot, 1), kBindingNullPointer);
}

TEST(FormulonCApiPivotOptions, AutoFormattingMarksLayoutDirty) {
  ExpectMarksDirty([](fm_workbook_t* wb) { return SetAutoFormatting(wb, 0U, kPivot, 0); });
}

// ---------------------------------------------------------------------------
// Field flags
// ---------------------------------------------------------------------------

TEST(FormulonCApiPivotOptions, FieldFlagsSetAndReadBack) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  fm_pivot_field_info_t info = FM_STRUCT_INIT(fm_pivot_field_info_t);
  ASSERT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0U, kPivot, kRegionField, &info), 0);
  EXPECT_EQ(info.default_subtotal, 1);
  EXPECT_EQ(info.insert_blank_row, 0);
  EXPECT_EQ(info.fill_down_labels, 0);

  ASSERT_EQ(SetPivotFieldDefaultSubtotal(wb.handle, 0U, kPivot, kRegionField, 0), 0);
  ASSERT_EQ(SetPivotFieldInsertBlankRow(wb.handle, 0U, kPivot, kRegionField, 1), 0);
  ASSERT_EQ(SetPivotFieldFillDownLabels(wb.handle, 0U, kPivot, kRegionField, 1), 0);
  ASSERT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0U, kPivot, kRegionField, &info), 0);
  EXPECT_EQ(info.default_subtotal, 0);
  EXPECT_EQ(info.insert_blank_row, 1);
  EXPECT_EQ(info.fill_down_labels, 1);

  ASSERT_EQ(SetPivotFieldDefaultSubtotal(wb.handle, 0U, kPivot, kRegionField, 1), 0);
  ASSERT_EQ(SetPivotFieldInsertBlankRow(wb.handle, 0U, kPivot, kRegionField, 0), 0);
  ASSERT_EQ(SetPivotFieldFillDownLabels(wb.handle, 0U, kPivot, kRegionField, 0), 0);
  ASSERT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0U, kPivot, kRegionField, &info), 0);
  EXPECT_EQ(info.default_subtotal, 1);
  EXPECT_EQ(info.insert_blank_row, 0);
  EXPECT_EQ(info.fill_down_labels, 0);
}

TEST(FormulonCApiPivotOptions, FieldFlagsRejectBadIndices) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  using Setter = fm_status_t (*)(fm_workbook_t*, std::size_t, std::size_t, std::size_t, std::int32_t);
  for (const Setter setter : {&SetPivotFieldDefaultSubtotal, &SetPivotFieldInsertBlankRow, &SetPivotFieldFillDownLabels,
                              &SetPivotFieldShowDetail}) {
    EXPECT_EQ(setter(wb.handle, 9U, kPivot, kRegionField, 1), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
    EXPECT_EQ(setter(wb.handle, 0U, 9U, kRegionField, 1), FM_STATUS_INDEX_OUT_OF_RANGE);
    EXPECT_EQ(setter(wb.handle, 0U, kPivot, 9U, 1), FM_STATUS_INDEX_OUT_OF_RANGE);
    EXPECT_EQ(setter(wb.handle, 0U, kPivot, std::numeric_limits<std::size_t>::max(), 1), FM_STATUS_INDEX_OUT_OF_RANGE);
    if constexpr (sizeof(std::size_t) > 4U) {
      EXPECT_EQ(setter(wb.handle, 0U, kPivot, static_cast<std::size_t>(1) << 32U, 1), FM_STATUS_INDEX_OUT_OF_RANGE);
    }
    EXPECT_EQ(setter(nullptr, 0U, kPivot, kRegionField, 1), kBindingNullPointer);
  }
}

TEST(FormulonCApiPivotOptions, FieldFlagsMarkLayoutDirty) {
  ExpectMarksDirty([](fm_workbook_t* wb) { return SetPivotFieldDefaultSubtotal(wb, 0U, kPivot, kRegionField, 0); });
  ExpectMarksDirty([](fm_workbook_t* wb) { return SetPivotFieldInsertBlankRow(wb, 0U, kPivot, kRegionField, 1); });
  ExpectMarksDirty([](fm_workbook_t* wb) { return SetPivotFieldFillDownLabels(wb, 0U, kPivot, kRegionField, 1); });
}

// ---------------------------------------------------------------------------
// Expansion
// ---------------------------------------------------------------------------

TEST(FormulonCApiPivotOptions, ShowDetailCollapsesFieldAndItems) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  fm_pivot_item_info_t item = FM_STRUCT_INIT(fm_pivot_item_info_t);
  for (std::size_t i = 0; i < 2U; ++i) {
    ASSERT_EQ(fm_sheet_get_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, i, &item), 0);
    EXPECT_EQ(item.show_detail, 1);
  }

  ASSERT_EQ(SetPivotFieldShowDetail(wb.handle, 0U, kPivot, kRegionField, 0), 0);
  for (std::size_t i = 0; i < 2U; ++i) {
    ASSERT_EQ(fm_sheet_get_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, i, &item), 0);
    EXPECT_EQ(item.show_detail, 0);
  }
  for (const formulon::pivot::PivotMarkerItem& marker : Table(wb).fields()[kRegionField].marker_items) {
    if (marker.type == "default") {
      EXPECT_FALSE(marker.show_detail);
    }
  }

  ASSERT_EQ(SetPivotFieldItemShowDetail(wb.handle, 0U, kPivot, kRegionField, 1, 1), 0);
  ASSERT_EQ(fm_sheet_get_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, 0, &item), 0);
  EXPECT_EQ(item.show_detail, 0);
  ASSERT_EQ(fm_sheet_get_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, 1, &item), 0);
  EXPECT_EQ(item.show_detail, 1);

  ASSERT_EQ(SetPivotFieldShowDetail(wb.handle, 0U, kPivot, kRegionField, 1), 0);
  for (std::size_t i = 0; i < 2U; ++i) {
    ASSERT_EQ(fm_sheet_get_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, i, &item), 0);
    EXPECT_EQ(item.show_detail, 1);
  }
}

TEST(FormulonCApiPivotOptions, ItemShowDetailRejectsBadIndices) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  EXPECT_EQ(SetPivotFieldItemShowDetail(wb.handle, 9U, kPivot, kRegionField, 0, 0), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(SetPivotFieldItemShowDetail(wb.handle, 0U, 9U, kRegionField, 0, 0), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(SetPivotFieldItemShowDetail(wb.handle, 0U, kPivot, 9U, 0, 0), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(SetPivotFieldItemShowDetail(wb.handle, 0U, kPivot, kRegionField, 2U, 0), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(
      SetPivotFieldItemShowDetail(wb.handle, 0U, kPivot, kRegionField, std::numeric_limits<std::size_t>::max(), 0),
      FM_STATUS_INDEX_OUT_OF_RANGE);
  if constexpr (sizeof(std::size_t) > 4U) {
    EXPECT_EQ(SetPivotFieldItemShowDetail(wb.handle, 0U, kPivot, kRegionField, static_cast<std::size_t>(1) << 32U, 0),
              FM_STATUS_INDEX_OUT_OF_RANGE);
  }
  EXPECT_EQ(SetPivotFieldItemShowDetail(nullptr, 0U, kPivot, kRegionField, 0, 0), kBindingNullPointer);
}

TEST(FormulonCApiPivotOptions, ShowDetailSettersMarkLayoutDirty) {
  ExpectMarksDirty([](fm_workbook_t* wb) { return SetPivotFieldShowDetail(wb, 0U, kPivot, kRegionField, 0); });
  ExpectMarksDirty([](fm_workbook_t* wb) { return SetPivotFieldItemShowDetail(wb, 0U, kPivot, kRegionField, 0, 0); });
}

// ---------------------------------------------------------------------------
// Item captions
// ---------------------------------------------------------------------------

TEST(FormulonCApiPivotOptions, ItemCaptionSetsKeepsValueAndClears) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  fm_pivot_item_info_t item = FM_STRUCT_INIT(fm_pivot_item_info_t);
  ASSERT_EQ(fm_sheet_get_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, 0, &item), 0);
  EXPECT_EQ(item.caption, nullptr);

  ASSERT_EQ(SetPivotFieldItemCaption(wb.handle, 0U, kPivot, kRegionField, 0, "Nord \xE5\x8C\x97"), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, 0, &item), 0);
  EXPECT_STREQ(item.caption, "Nord \xE5\x8C\x97");
  EXPECT_STREQ(item.name, "North");
  EXPECT_TRUE(LayoutHasText(wb.handle, kPivot, "Nord \xE5\x8C\x97"));
  EXPECT_NE(SavedPivotXml(wb.handle).find("n=\"Nord \xE5\x8C\x97\""), std::string::npos);

  // NULL and empty clear.
  ASSERT_EQ(SetPivotFieldItemCaption(wb.handle, 0U, kPivot, kRegionField, 0, nullptr), 0);
  ASSERT_EQ(fm_sheet_get_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, 0, &item), 0);
  EXPECT_EQ(item.caption, nullptr);
  ASSERT_EQ(SetPivotFieldItemCaption(wb.handle, 0U, kPivot, kRegionField, 0, "Nord"), 0);
  ASSERT_EQ(SetPivotFieldItemCaption(wb.handle, 0U, kPivot, kRegionField, 0, ""), 0);
  ASSERT_EQ(fm_sheet_get_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, 0, &item), 0);
  EXPECT_EQ(item.caption, nullptr);

  // A caption equal to the item's own value clears, whatever its case.
  ASSERT_EQ(SetPivotFieldItemCaption(wb.handle, 0U, kPivot, kRegionField, 0, "Nord"), 0);
  ASSERT_EQ(SetPivotFieldItemCaption(wb.handle, 0U, kPivot, kRegionField, 0, "NORTH"), 0);
  ASSERT_EQ(fm_sheet_get_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, 0, &item), 0);
  EXPECT_EQ(item.caption, nullptr);
}

TEST(FormulonCApiPivotOptions, ItemCaptionRejectsDuplicatesOfShownLabels) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  // Another item's value, in any case.
  EXPECT_EQ(SetPivotFieldItemCaption(wb.handle, 0U, kPivot, kRegionField, 0, "South"), FM_STATUS_ALREADY_EXISTS);
  EXPECT_EQ(SetPivotFieldItemCaption(wb.handle, 0U, kPivot, kRegionField, 0, "south"), FM_STATUS_ALREADY_EXISTS);
  // Another item's caption.
  ASSERT_EQ(SetPivotFieldItemCaption(wb.handle, 0U, kPivot, kRegionField, 0, "Nord"), 0);
  EXPECT_EQ(SetPivotFieldItemCaption(wb.handle, 0U, kPivot, kRegionField, 1, "NORD"), FM_STATUS_ALREADY_EXISTS);
  // The value a caption replaced is free to use again.
  EXPECT_EQ(SetPivotFieldItemCaption(wb.handle, 0U, kPivot, kRegionField, 1, "North"), 0);
  // The same caption on the same item is not a duplicate.
  EXPECT_EQ(SetPivotFieldItemCaption(wb.handle, 0U, kPivot, kRegionField, 0, "Nord"), 0);

  fm_pivot_item_info_t item = FM_STRUCT_INIT(fm_pivot_item_info_t);
  ASSERT_EQ(fm_sheet_get_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, 1, &item), 0);
  EXPECT_STREQ(item.caption, "North");
}

TEST(FormulonCApiPivotOptions, ItemCaptionBoundaries) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  const std::string at_limit = Repeat(kEmoji, 127) + "a";
  ASSERT_EQ(SetPivotFieldItemCaption(wb.handle, 0U, kPivot, kRegionField, 0, at_limit.c_str()), 0)
      << fm_get_last_error_message();
  fm_pivot_item_info_t item = FM_STRUCT_INIT(fm_pivot_item_info_t);
  ASSERT_EQ(fm_sheet_get_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, 0, &item), 0);
  EXPECT_EQ(std::string(item.caption), at_limit);

  EXPECT_EQ(SetPivotFieldItemCaption(wb.handle, 0U, kPivot, kRegionField, 0, Repeat(kEmoji, 128).c_str()),
            kInvalidArgument);
  EXPECT_EQ(SetPivotFieldItemCaption(wb.handle, 0U, kPivot, kRegionField, 0, Repeat("a", 256).c_str()),
            kInvalidArgument);
  EXPECT_EQ(SetPivotFieldItemCaption(wb.handle, 0U, kPivot, kRegionField, 0, "\xC3"), kBindingUtf8EncodingError);
  EXPECT_EQ(SetPivotFieldItemCaption(wb.handle, 9U, kPivot, kRegionField, 0, "x"), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(SetPivotFieldItemCaption(wb.handle, 0U, 9U, kRegionField, 0, "x"), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(SetPivotFieldItemCaption(wb.handle, 0U, kPivot, 9U, 0, "x"), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(SetPivotFieldItemCaption(wb.handle, 0U, kPivot, kRegionField, 2U, "x"), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(SetPivotFieldItemCaption(wb.handle, 0U, kPivot, kRegionField, std::numeric_limits<std::size_t>::max(), "x"),
            FM_STATUS_INDEX_OUT_OF_RANGE);
  if constexpr (sizeof(std::size_t) > 4U) {
    EXPECT_EQ(SetPivotFieldItemCaption(wb.handle, 0U, kPivot, kRegionField, static_cast<std::size_t>(1) << 32U, "x"),
              FM_STATUS_INDEX_OUT_OF_RANGE);
  }
  EXPECT_EQ(SetPivotFieldItemCaption(nullptr, 0U, kPivot, kRegionField, 0, "x"), kBindingNullPointer);
  // The rejected calls left the accepted caption alone.
  ASSERT_EQ(fm_sheet_get_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, 0, &item), 0);
  EXPECT_EQ(std::string(item.caption), at_limit);
}

TEST(FormulonCApiPivotOptions, ItemCaptionMarksLayoutDirty) {
  ExpectMarksDirty([](fm_workbook_t* wb) { return SetPivotFieldItemCaption(wb, 0U, kPivot, kRegionField, 0, "Nord"); });
}

// ---------------------------------------------------------------------------
// The Values level in the field orders
// ---------------------------------------------------------------------------

TEST(FormulonCApiPivotOptions, ValuesFieldConstantMatchesTheOoxmlMarker) {
  EXPECT_EQ(FM_PIVOT_VALUES_FIELD, static_cast<std::uint32_t>(-2));
}

TEST(FormulonCApiPivotOptions, ValuesFieldNeedsTwoDataFields) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, nullptr), 0) << fm_get_last_error_message();
  EXPECT_EQ(ReadOrder(wb.handle, true), (std::vector<std::uint32_t>{0U}));
  EXPECT_TRUE(ReadOrder(wb.handle, false).empty());
  EXPECT_EQ(SetOrder(wb.handle, true, {0U, FM_PIVOT_VALUES_FIELD}), kInvalidArgument);
  EXPECT_EQ(SetOrder(wb.handle, false, {FM_PIVOT_VALUES_FIELD}), kInvalidArgument);
  EXPECT_EQ(ReadOrder(wb.handle, true), (std::vector<std::uint32_t>{0U}));
  EXPECT_FALSE(Table(wb).row_values_position().has_value());
  EXPECT_FALSE(Table(wb).col_values_position().has_value());
}

TEST(FormulonCApiPivotOptions, ValuesFieldDefaultsToTheColumnEnd) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, nullptr), 0) << fm_get_last_error_message();
  AddSecondDataField(wb.handle);
  EXPECT_EQ(ReadOrder(wb.handle, true), (std::vector<std::uint32_t>{0U}));
  EXPECT_EQ(ReadOrder(wb.handle, false), (std::vector<std::uint32_t>{FM_PIVOT_VALUES_FIELD}));
}

TEST(FormulonCApiPivotOptions, ValuesFieldMovesBetweenAxesAndPositions) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, nullptr), 0) << fm_get_last_error_message();
  AddSecondDataField(wb.handle);

  ASSERT_EQ(SetOrder(wb.handle, true, {0U, FM_PIVOT_VALUES_FIELD}), 0) << fm_get_last_error_message();
  EXPECT_EQ(ReadOrder(wb.handle, true), (std::vector<std::uint32_t>{0U, FM_PIVOT_VALUES_FIELD}));
  EXPECT_TRUE(ReadOrder(wb.handle, false).empty());
  EXPECT_EQ(Table(wb).row_values_position(), 1U);
  EXPECT_FALSE(Table(wb).col_values_position().has_value());

  ASSERT_EQ(SetOrder(wb.handle, true, {FM_PIVOT_VALUES_FIELD, 0U}), 0) << fm_get_last_error_message();
  EXPECT_EQ(ReadOrder(wb.handle, true), (std::vector<std::uint32_t>{FM_PIVOT_VALUES_FIELD, 0U}));
  EXPECT_EQ(Table(wb).row_values_position(), 0U);

  // The getter's own listing is a valid input.
  std::uint32_t listed[8] = {};
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_pivot_row_field_order(wb.handle, 0U, kPivot, listed, 8U, &count), 0);
  ASSERT_EQ(fm_sheet_set_pivot_row_field_order(wb.handle, 0U, kPivot, listed, count), 0) << fm_get_last_error_message();
  EXPECT_EQ(ReadOrder(wb.handle, true), (std::vector<std::uint32_t>{FM_PIVOT_VALUES_FIELD, 0U}));

  // Back to the columns: the rows give it up, and the default layout returns.
  ASSERT_EQ(SetOrder(wb.handle, false, {FM_PIVOT_VALUES_FIELD}), 0) << fm_get_last_error_message();
  EXPECT_EQ(ReadOrder(wb.handle, false), (std::vector<std::uint32_t>{FM_PIVOT_VALUES_FIELD}));
  EXPECT_EQ(ReadOrder(wb.handle, true), (std::vector<std::uint32_t>{0U}));
  EXPECT_FALSE(Table(wb).row_values_position().has_value());
  EXPECT_EQ(Table(wb).col_values_position(), 0U);
}

TEST(FormulonCApiPivotOptions, ValuesFieldMustStayListedOnItsAxis) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, nullptr), 0) << fm_get_last_error_message();
  AddSecondDataField(wb.handle);

  // The columns hold it by default.
  EXPECT_EQ(SetOrder(wb.handle, false, {}), kInvalidArgument);
  EXPECT_EQ(ReadOrder(wb.handle, false), (std::vector<std::uint32_t>{FM_PIVOT_VALUES_FIELD}));

  ASSERT_EQ(SetOrder(wb.handle, true, {0U, FM_PIVOT_VALUES_FIELD}), 0);
  EXPECT_EQ(SetOrder(wb.handle, true, {0U}), kInvalidArgument);
  EXPECT_EQ(SetOrder(wb.handle, true, {}), kInvalidArgument);
  EXPECT_EQ(ReadOrder(wb.handle, true), (std::vector<std::uint32_t>{0U, FM_PIVOT_VALUES_FIELD}));
  // The other axis may be rewritten without it.
  EXPECT_EQ(SetOrder(wb.handle, false, {}), 0) << fm_get_last_error_message();
  EXPECT_EQ(ReadOrder(wb.handle, true), (std::vector<std::uint32_t>{0U, FM_PIVOT_VALUES_FIELD}));
}

TEST(FormulonCApiPivotOptions, ValuesFieldRejectsDuplicatesAndStrayIndices) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, nullptr), 0) << fm_get_last_error_message();
  AddSecondDataField(wb.handle);

  EXPECT_EQ(SetOrder(wb.handle, true, {FM_PIVOT_VALUES_FIELD, 0U, FM_PIVOT_VALUES_FIELD}), kInvalidArgument);
  EXPECT_EQ(SetOrder(wb.handle, false, {FM_PIVOT_VALUES_FIELD, FM_PIVOT_VALUES_FIELD}), kInvalidArgument);
  EXPECT_EQ(SetOrder(wb.handle, true, {0U, FM_PIVOT_VALUES_FIELD - 1U}), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(SetOrder(wb.handle, true, {0U, 0xFFFFFFFFU}), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(ReadOrder(wb.handle, true), (std::vector<std::uint32_t>{0U}));
  EXPECT_EQ(ReadOrder(wb.handle, false), (std::vector<std::uint32_t>{FM_PIVOT_VALUES_FIELD}));
}

TEST(FormulonCApiPivotOptions, ValuesFieldOrderMarksLayoutDirty) {
  WorkbookGuard built;
  ASSERT_EQ(fm_workbook_create_default(&built.handle), 0);
  ASSERT_EQ(BuildScratchPivot(built.handle, nullptr, nullptr), 0);
  AddSecondDataField(built.handle);
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(built.handle, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  WorkbookGuard clean;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &clean.handle), 0) << fm_get_last_error_message();
  ASSERT_FALSE(Table(clean).layout_dirty());
  ASSERT_EQ(SetOrder(clean.handle, true, {0U, FM_PIVOT_VALUES_FIELD}), 0) << fm_get_last_error_message();
  EXPECT_TRUE(Table(clean).layout_dirty());
}

// ---------------------------------------------------------------------------
// Show Details
// ---------------------------------------------------------------------------

namespace {

// First projected cell of `kind`, as absolute coordinates.
bool FindCell(fm_workbook_t* wb, fm_pivot_cell_kind_t kind, std::uint32_t* row, std::uint32_t* col) {
  PivotCellsGuard projected;
  EXPECT_EQ(fm_sheet_get_pivot_layout(wb, 0U, kPivot, &projected.handle), 0) << fm_get_last_error_message();
  for (const fm_pivot_cell_info_t& cell : CollectCells(projected.handle)) {
    if (cell.kind == kind) {
      *row = cell.row;
      *col = cell.col;
      return true;
    }
  }
  return false;
}

}  // namespace

TEST(FormulonCApiPivotOptions, ShowDetailsAddsASheetBeforeThePivotAndActivatesIt) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, nullptr), 0) << fm_get_last_error_message();
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  ASSERT_TRUE(FindCell(wb.handle, FM_PIVOT_CELL_KIND_DATA, &row, &col));

  const std::size_t before = count_of(fm_workbook_get_sheet_count, wb.handle);
  std::size_t created = 99;
  ASSERT_EQ(fm_sheet_show_pivot_details(wb.handle, 0U, kPivot, row, col, &created), 0) << fm_get_last_error_message();
  EXPECT_EQ(created, 0U);
  EXPECT_EQ(count_of(fm_workbook_get_sheet_count, wb.handle), before + 1U);
  std::size_t active = 99;
  ASSERT_EQ(fm_workbook_get_active_sheet(wb.handle, &active), 0);
  EXPECT_EQ(active, 0U);
  // The pivot moved with its sheet.
  EXPECT_EQ(wb.handle->wb->sheet(1U).pivot_tables().size(), 1U);
}

TEST(FormulonCApiPivotOptions, ShowDetailsRejectsBadCellsAndArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, nullptr), 0) << fm_get_last_error_message();
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  ASSERT_TRUE(FindCell(wb.handle, FM_PIVOT_CELL_KIND_DATA, &row, &col));
  std::uint32_t label_row = 0;
  std::uint32_t label_col = 0;
  ASSERT_TRUE(FindCell(wb.handle, FM_PIVOT_CELL_KIND_ROW_LABEL, &label_row, &label_col));

  std::size_t created = 99;
  EXPECT_EQ(fm_sheet_show_pivot_details(wb.handle, 0U, kPivot, label_row, label_col, &created), kInvalidArgument);
  EXPECT_EQ(fm_sheet_show_pivot_details(wb.handle, 0U, kPivot, 5000U, 5000U, &created), kInvalidArgument);
  EXPECT_EQ(fm_sheet_show_pivot_details(wb.handle, 0U, kPivot, 0xFFFFFFFFU, 0xFFFFFFFFU, &created), kInvalidArgument);
  EXPECT_EQ(fm_sheet_show_pivot_details(wb.handle, 9U, kPivot, row, col, &created), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_show_pivot_details(wb.handle, 0U, 9U, row, col, &created), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_show_pivot_details(wb.handle, 0U, kPivot, row, col, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_show_pivot_details(nullptr, 0U, kPivot, row, col, &created), kBindingNullPointer);
  EXPECT_EQ(created, 99U);
  EXPECT_EQ(count_of(fm_workbook_get_sheet_count, wb.handle), 1U);

  Table(wb).set_enable_drill(false);
  EXPECT_EQ(fm_sheet_show_pivot_details(wb.handle, 0U, kPivot, row, col, &created), kPreconditionFailed);
}

// ---------------------------------------------------------------------------
// Cache edits drop the encoding flags of the cells they overwrite
// ---------------------------------------------------------------------------

TEST(FormulonCApiPivotOptions, RecordEditsResetDateAndIndexFlags) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, nullptr), 0) << fm_get_last_error_message();
  formulon::pivot::PivotCache* cache = wb.handle->wb->mutable_pivot_caches().front().get();
  formulon::pivot::PivotCacheRecord& record = cache->mutable_records()[0];
  record.cell_is_index.assign(2U, false);
  record.cell_is_index[0] = true;
  record.cell_is_date.assign(2U, false);
  record.cell_is_date[1] = true;

  // A plain number over a date cell is an inline number again.
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, 0U, 1U, 5.0), 0)
      << fm_get_last_error_message();
  EXPECT_FALSE(record.cell_is_date[1]);
  EXPECT_FALSE(record.cell_is_index[1]);
  const std::string records = ExtractZipEntry(
      [&] {
        BufferGuard saved;
        EXPECT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_get_last_error_message();
        return std::vector<std::uint8_t>(saved.data, saved.data + saved.len);
      }(),
      "xl/pivotCache/pivotCacheRecords1.xml");
  EXPECT_NE(records.find("<n v=\"5\"/>"), std::string::npos);
  EXPECT_EQ(records.find("<d "), std::string::npos);

  // Text and blank over an index cell are no longer an index.
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_text(wb.handle, cache_id, 0U, 0U, "Elsewhere"), 0);
  EXPECT_FALSE(record.cell_is_index[0]);
  record.cell_is_index[0] = true;
  record.cell_is_date[0] = true;
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_blank(wb.handle, cache_id, 0U, 0U), 0);
  EXPECT_FALSE(record.cell_is_index[0]);
  EXPECT_FALSE(record.cell_is_date[0]);

  // Clearing the shared items forgets which of them were dates.
  cache->mutable_fields()[0].shared_item_is_date.assign(2U, true);
  ASSERT_EQ(fm_workbook_clear_pivot_cache_field_shared_items(wb.handle, cache_id, 0U), 0);
  EXPECT_TRUE(cache->fields()[0].shared_item_is_date.empty());
}

// ---------------------------------------------------------------------------
// PivotTable options struct
// ---------------------------------------------------------------------------

namespace {

// Every member away from the value a created pivot starts with.
fm_pivot_options_t ChangedOptions() {
  fm_pivot_options_t o = FM_STRUCT_INIT(fm_pivot_options_t);
  o.show_error = 1;
  o.error_caption = "ERR";
  o.show_missing = 0;
  o.missing_caption = "N/A";
  o.show_headers = 0;
  o.show_drill = 0;
  o.show_data_tips = 0;
  o.grid_drop_zones = 1;
  o.show_empty_row = 1;
  o.show_empty_column = 1;
  o.show_data_drop_down = 0;
  o.print_drill = 1;
  o.field_print_titles = 1;
  o.item_print_titles = 0;
  o.merge_item = 1;
  o.indent = 3U;
  o.page_over_then_down = 1;
  o.page_wrap = 2U;
  o.multiple_field_filters = 1;
  o.preserve_formatting = 0;
  o.custom_list_sort = 0;
  o.alt_text = "Alt title";
  o.alt_text_summary = "Alt summary text";
  o.grand_total_caption = "All of it";
  o.auto_formatting = 0;
  o.enable_drill = 0;
  return o;
}

std::string Text(const char* text) {
  return text != nullptr ? std::string(text) : std::string("<null>");
}

void ExpectSameOptions(const fm_pivot_options_t& a, const fm_pivot_options_t& b) {
  EXPECT_EQ(a.show_error, b.show_error);
  EXPECT_EQ(Text(a.error_caption), Text(b.error_caption));
  EXPECT_EQ(a.show_missing, b.show_missing);
  EXPECT_EQ(Text(a.missing_caption), Text(b.missing_caption));
  EXPECT_EQ(a.show_headers, b.show_headers);
  EXPECT_EQ(a.show_drill, b.show_drill);
  EXPECT_EQ(a.show_data_tips, b.show_data_tips);
  EXPECT_EQ(a.grid_drop_zones, b.grid_drop_zones);
  EXPECT_EQ(a.show_empty_row, b.show_empty_row);
  EXPECT_EQ(a.show_empty_column, b.show_empty_column);
  EXPECT_EQ(a.show_data_drop_down, b.show_data_drop_down);
  EXPECT_EQ(a.print_drill, b.print_drill);
  EXPECT_EQ(a.field_print_titles, b.field_print_titles);
  EXPECT_EQ(a.item_print_titles, b.item_print_titles);
  EXPECT_EQ(a.merge_item, b.merge_item);
  EXPECT_EQ(a.indent, b.indent);
  EXPECT_EQ(a.page_over_then_down, b.page_over_then_down);
  EXPECT_EQ(a.page_wrap, b.page_wrap);
  EXPECT_EQ(a.multiple_field_filters, b.multiple_field_filters);
  EXPECT_EQ(a.preserve_formatting, b.preserve_formatting);
  EXPECT_EQ(a.custom_list_sort, b.custom_list_sort);
  EXPECT_EQ(Text(a.alt_text), Text(b.alt_text));
  EXPECT_EQ(Text(a.alt_text_summary), Text(b.alt_text_summary));
  EXPECT_EQ(Text(a.grand_total_caption), Text(b.grand_total_caption));
  EXPECT_EQ(a.auto_formatting, b.auto_formatting);
  EXPECT_EQ(a.enable_drill, b.enable_drill);
}

fm_pivot_options_t GetOptions(const fm_workbook_t* wb) {
  fm_pivot_options_t options = FM_STRUCT_INIT(fm_pivot_options_t);
  EXPECT_EQ(fm_sheet_get_pivot_options(wb, 0U, kPivot, &options), 0) << fm_get_last_error_message();
  return options;
}

std::vector<std::uint8_t> SavedPackage(fm_workbook_t* wb) {
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save(wb, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  return std::vector<std::uint8_t>(saved.data, saved.data + saved.len);
}

bool HasZipEntry(const std::vector<std::uint8_t>& package, const char* path) {
  mz_zip_archive reader{};
  if (mz_zip_reader_init_mem(&reader, package.data(), package.size(), 0) == MZ_FALSE) {
    ADD_FAILURE() << "mz_zip_reader_init_mem failed";
    return false;
  }
  const bool found = mz_zip_reader_locate_file(&reader, path, nullptr, 0) >= 0;
  mz_zip_reader_end(&reader);
  return found;
}

// The `<pivotField ...>` start tag of the first field on axis `axis`.
std::string FieldTag(const std::string& pivot_xml, const std::string& axis) {
  for (std::size_t at = pivot_xml.find("<pivotField "); at != std::string::npos;
       at = pivot_xml.find("<pivotField ", at + 1U)) {
    const std::string tag = pivot_xml.substr(at, pivot_xml.find('>', at) - at);
    if (tag.find("axis=\"" + axis + "\"") != std::string::npos) {
      return tag;
    }
  }
  ADD_FAILURE() << "no pivotField on " << axis << "\n" << pivot_xml;
  return {};
}

void CreateJaScratch(WorkbookGuard& wb) {
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "win-365-ja_JP"), 0);
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, nullptr), 0) << fm_get_last_error_message();
}

std::string DefaultName(const fm_workbook_t* wb, std::size_t field, fm_pivot_aggregation_t aggregation,
                        std::size_t exclude = SIZE_MAX) {
  const char* name = nullptr;
  EXPECT_EQ(fm_sheet_get_pivot_data_field_default_name(wb, 0U, kPivot, field, aggregation, exclude, &name), 0)
      << fm_get_last_error_message();
  return name != nullptr ? std::string(name) : std::string();
}

// Adds a data field with an empty name and returns the name it got.
std::string AddUnnamed(fm_workbook_t* wb, std::size_t field, fm_pivot_aggregation_t aggregation) {
  fm_pivot_data_field_t spec = FM_STRUCT_INIT(fm_pivot_data_field_t);
  spec.name = "";
  spec.field_index = static_cast<std::uint32_t>(field);
  spec.aggregation = aggregation;
  spec.number_format = "";
  spec.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;
  spec.has_show_as_base_field = 0;
  spec.has_show_as_base_item = 0;
  std::size_t index = 99;
  EXPECT_EQ(fm_sheet_add_pivot_data_field(wb, 0U, kPivot, &spec, &index), 0) << fm_get_last_error_message();
  fm_pivot_data_field_t read = FM_STRUCT_INIT(fm_pivot_data_field_t);
  EXPECT_EQ(fm_sheet_get_pivot_data_field_at(wb, 0U, kPivot, index, &read), 0);
  return read.name != nullptr ? std::string(read.name) : std::string();
}

const char* const kSumOfAmount = "\xE5\x90\x88\xE8\xA8\x88 / Amount";  // 合計 / Amount

}  // namespace

TEST(FormulonCApiPivotOptions, CreatedPivotStartsWithExcelsNewPivotOptions) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, nullptr), 0) << fm_get_last_error_message();
  const fm_pivot_options_t o = GetOptions(wb.handle);
  EXPECT_EQ(o.item_print_titles, 1);
  EXPECT_EQ(o.indent, 0U);
  EXPECT_EQ(o.multiple_field_filters, 0);
  EXPECT_EQ(o.show_error, 0);
  EXPECT_EQ(o.error_caption, nullptr);
  EXPECT_EQ(o.show_missing, 1);
  EXPECT_EQ(o.missing_caption, nullptr);
  EXPECT_EQ(o.show_headers, 1);
  EXPECT_EQ(o.show_drill, 1);
  EXPECT_EQ(o.show_data_tips, 1);
  EXPECT_EQ(o.grid_drop_zones, 0);
  EXPECT_EQ(o.show_data_drop_down, 1);
  EXPECT_EQ(o.page_wrap, 0U);
  EXPECT_EQ(o.preserve_formatting, 1);
  EXPECT_EQ(o.custom_list_sort, 1);
  EXPECT_EQ(o.alt_text, nullptr);
  EXPECT_EQ(o.alt_text_summary, nullptr);
  EXPECT_EQ(o.grand_total_caption, nullptr);
  EXPECT_EQ(o.auto_formatting, 1);
  EXPECT_EQ(o.enable_drill, 1);
}

TEST(FormulonCApiPivotOptions, OptionsRoundTripThroughSaveInExcelsSpelling) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, nullptr), 0) << fm_get_last_error_message();
  const fm_pivot_options_t changed = ChangedOptions();
  ASSERT_EQ(UpdateOptionsFrom(wb.handle, 0U, kPivot, &changed), 0) << fm_get_last_error_message();
  ExpectSameOptions(GetOptions(wb.handle), changed);

  const std::vector<std::uint8_t> package = SavedPackage(wb.handle);
  const std::string xml = ExtractZipEntry(package, "xl/pivotTables/pivotTable1.xml");
  for (const char* attribute : {"showError=\"1\"",
                                "errorCaption=\"ERR\"",
                                "showMissing=\"0\"",
                                "missingCaption=\"N/A\"",
                                "showHeaders=\"0\"",
                                "showDrill=\"0\"",
                                "showDataTips=\"0\"",
                                "gridDropZones=\"1\"",
                                "showEmptyRow=\"1\"",
                                "showEmptyCol=\"1\"",
                                "showDataDropDown=\"0\"",
                                "printDrill=\"1\"",
                                "fieldPrintTitles=\"1\"",
                                "mergeItem=\"1\"",
                                "indent=\"3\"",
                                "pageOverThenDown=\"1\"",
                                "pageWrap=\"2\"",
                                "preserveFormatting=\"0\"",
                                "customListSort=\"0\"",
                                "enableDrill=\"0\"",
                                "grandTotalCaption=\"All of it\"",
                                "altText=\"Alt title\"",
                                "altTextSummary=\"Alt summary text\""}) {
    EXPECT_NE(xml.find(attribute), std::string::npos) << attribute << "\n" << xml;
  }
  // Schema defaults are not written, and neither is the OLAP-only attribute.
  for (const char* attribute : {"itemPrintTitles=", "multipleFieldFilters=", "useAutoFormatting=", "asteriskTotals="}) {
    EXPECT_EQ(xml.find(attribute), std::string::npos) << attribute << "\n" << xml;
  }

  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(package.data(), package.size(), &loaded.handle), 0) << fm_get_last_error_message();
  ExpectSameOptions(GetOptions(loaded.handle), changed);
}

TEST(FormulonCApiPivotOptions, OptionsUnsetStringsAndFlags) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  fm_pivot_options_t o = ChangedOptions();
  ASSERT_EQ(UpdateOptionsFrom(wb.handle, 0U, kPivot, &o), 0) << fm_get_last_error_message();
  // NULL and empty strings unset; any non-zero flag reads 1.
  o.error_caption = nullptr;
  o.missing_caption = "";
  o.alt_text = nullptr;
  o.alt_text_summary = "";
  o.grand_total_caption = nullptr;
  o.show_error = 7;
  o.merge_item = -1;
  ASSERT_EQ(UpdateOptionsFrom(wb.handle, 0U, kPivot, &o), 0) << fm_get_last_error_message();
  const fm_pivot_options_t read = GetOptions(wb.handle);
  EXPECT_EQ(read.error_caption, nullptr);
  EXPECT_EQ(read.missing_caption, nullptr);
  EXPECT_EQ(read.alt_text, nullptr);
  EXPECT_EQ(read.alt_text_summary, nullptr);
  EXPECT_EQ(read.grand_total_caption, nullptr);
  EXPECT_EQ(read.show_error, 1);
  EXPECT_EQ(read.merge_item, 1);
  EXPECT_TRUE(Table(wb).layout_dirty());
}

TEST(FormulonCApiPivotOptions, ErrorCaptionIsKeptWhileShowErrorIsOff) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  fm_pivot_options_t o = GetOptions(wb.handle);
  o.show_error = 0;
  o.error_caption = "ERR";
  ASSERT_EQ(UpdateOptionsFrom(wb.handle, 0U, kPivot, &o), 0) << fm_get_last_error_message();
  EXPECT_EQ(Text(GetOptions(wb.handle).error_caption), "ERR");
  const std::string xml = SavedPivotXml(wb.handle);
  EXPECT_NE(xml.find("errorCaption=\"ERR\""), std::string::npos) << xml;
  EXPECT_EQ(xml.find("showError="), std::string::npos) << xml;
}

TEST(FormulonCApiPivotOptions, OptionsArgumentBoundaries) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  const fm_pivot_options_t before = GetOptions(wb.handle);
  fm_pivot_options_t o = before;
  fm_pivot_options_t out = FM_STRUCT_INIT(fm_pivot_options_t);
  EXPECT_EQ(fm_sheet_get_pivot_options(nullptr, 0U, kPivot, &out), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_pivot_options(wb.handle, 0U, kPivot, nullptr), kBindingNullPointer);
  EXPECT_EQ(UpdateOptionsFrom(nullptr, 0U, kPivot, &o), kBindingNullPointer);
  EXPECT_EQ(UpdateOptionsFrom(wb.handle, 0U, kPivot, nullptr), kBindingNullPointer);
  for (const std::size_t index : {std::size_t{1}, std::size_t{0x80000000U}, static_cast<std::size_t>(0x100000000ULL)}) {
    EXPECT_EQ(fm_sheet_get_pivot_options(wb.handle, index, kPivot, &out), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE) << index;
    EXPECT_EQ(fm_sheet_get_pivot_options(wb.handle, 0U, index, &out), FM_STATUS_INDEX_OUT_OF_RANGE) << index;
    EXPECT_EQ(UpdateOptionsFrom(wb.handle, index, kPivot, &o), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE) << index;
    EXPECT_EQ(UpdateOptionsFrom(wb.handle, 0U, index, &o), FM_STATUS_INDEX_OUT_OF_RANGE) << index;
  }

  // indent 0..126, page_wrap 0..255.
  o = before;
  o.indent = 126U;
  o.page_wrap = 255U;
  ASSERT_EQ(UpdateOptionsFrom(wb.handle, 0U, kPivot, &o), 0) << fm_get_last_error_message();
  o.indent = 127U;
  EXPECT_EQ(UpdateOptionsFrom(wb.handle, 0U, kPivot, &o), kInvalidArgument);
  o.indent = 0xFFFFFFFFU;
  EXPECT_EQ(UpdateOptionsFrom(wb.handle, 0U, kPivot, &o), kInvalidArgument);
  o.indent = 0U;
  o.page_wrap = 256U;
  EXPECT_EQ(UpdateOptionsFrom(wb.handle, 0U, kPivot, &o), kInvalidArgument);
  EXPECT_EQ(GetOptions(wb.handle).indent, 126U);

  // 255 UTF-16 units per caption and title, a surrogate pair counting two; the
  // description has no limit.
  const std::string at_limit = Repeat(kEmoji, 127) + "a";
  const std::string over_limit = Repeat(kEmoji, 128);
  const std::string long_summary = Repeat("x", 4000);
  o = before;
  o.error_caption = at_limit.c_str();
  o.missing_caption = at_limit.c_str();
  o.alt_text = at_limit.c_str();
  o.grand_total_caption = at_limit.c_str();
  o.alt_text_summary = long_summary.c_str();
  ASSERT_EQ(UpdateOptionsFrom(wb.handle, 0U, kPivot, &o), 0) << fm_get_last_error_message();
  EXPECT_EQ(Text(GetOptions(wb.handle).alt_text_summary), long_summary);
  for (const char* fm_pivot_options_t::*member :
       {&fm_pivot_options_t::error_caption, &fm_pivot_options_t::missing_caption, &fm_pivot_options_t::alt_text,
        &fm_pivot_options_t::grand_total_caption}) {
    fm_pivot_options_t bad = before;
    bad.*member = over_limit.c_str();
    EXPECT_EQ(UpdateOptionsFrom(wb.handle, 0U, kPivot, &bad), kInvalidArgument);
    bad.*member = "\xC3";
    EXPECT_EQ(UpdateOptionsFrom(wb.handle, 0U, kPivot, &bad), kBindingUtf8EncodingError);
  }
  o = before;
  o.alt_text_summary = "\xC3";
  EXPECT_EQ(UpdateOptionsFrom(wb.handle, 0U, kPivot, &o), kBindingUtf8EncodingError);
  // Rejected calls change nothing.
  EXPECT_EQ(Text(GetOptions(wb.handle).alt_text), at_limit);
}

TEST(FormulonCApiPivotOptions, OptionsPatchOfOnlyKeepsChangesNothing) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  const OwnedPivotOptions before(GetOptions(wb.handle));
  ASSERT_FALSE(Table(wb).layout_dirty());
  const fm_pivot_options_patch_t keep = FM_STRUCT_INIT(fm_pivot_options_patch_t);
  ASSERT_EQ(fm_sheet_update_pivot_options(wb.handle, 0U, kPivot, &keep), 0) << fm_get_last_error_message();
  EXPECT_FALSE(Table(wb).layout_dirty()) << "an all-KEEP patch does not touch the table";
  ExpectSameOptions(GetOptions(wb.handle), *before.get());

  fm_pivot_options_patch_t clear_flag = FM_STRUCT_INIT(fm_pivot_options_patch_t);
  clear_flag.merge_item_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_pivot_options(wb.handle, 0U, kPivot, &clear_flag), kInvalidArgument);
  fm_pivot_options_patch_t bad_op = FM_STRUCT_INIT(fm_pivot_options_patch_t);
  bad_op.error_caption_op = static_cast<fm_patch_op_t>(3);
  bad_op.error_caption = "x";
  EXPECT_EQ(fm_sheet_update_pivot_options(wb.handle, 0U, kPivot, &bad_op), kInvalidArgument);
  fm_pivot_options_patch_t unsized = FM_STRUCT_INIT(fm_pivot_options_patch_t);
  unsized.struct_size = 0U;
  EXPECT_EQ(fm_sheet_update_pivot_options(wb.handle, 0U, kPivot, &unsized), kInvalidArgument);
  EXPECT_FALSE(Table(wb).layout_dirty());
}

TEST(FormulonCApiPivotOptions, CacheOptionsPatchOfOnlyKeepsLeavesThePivotsProjected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, nullptr), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_project_pivot(wb.handle, 0U, kPivot), 0) << fm_get_last_error_message();
  const fm_pivot_cache_options_patch_t keep = FM_STRUCT_INIT(fm_pivot_cache_options_patch_t);
  ASSERT_EQ(fm_workbook_update_pivot_cache_options(wb.handle, cache_id, &keep), 0) << fm_get_last_error_message();
  std::int32_t pending = -1;
  ASSERT_EQ(fm_sheet_is_pivot_projection_pending(wb.handle, 0U, kPivot, &pending), 0);
  EXPECT_EQ(pending, 0);
}

TEST(FormulonCApiPivotOptions, OptionsKeepALoadedOutOfRangeValue) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  Table(wb).set_indent(200U);
  Table(wb).set_page_wrap(300U);
  fm_pivot_options_t o = GetOptions(wb.handle);
  EXPECT_EQ(o.indent, 200U);
  EXPECT_EQ(o.page_wrap, 300U);
  o.show_headers = 0;
  ASSERT_EQ(UpdateOptionsFrom(wb.handle, 0U, kPivot, &o), 0) << fm_get_last_error_message();
  EXPECT_EQ(GetOptions(wb.handle).indent, 200U);
  o.indent = 201U;
  EXPECT_EQ(UpdateOptionsFrom(wb.handle, 0U, kPivot, &o), kInvalidArgument);
}

// ---------------------------------------------------------------------------
// PivotCache options
// ---------------------------------------------------------------------------

TEST(FormulonCApiPivotOptions, CacheOptionsRoundTripThroughSave) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, nullptr), 0) << fm_get_last_error_message();
  fm_pivot_cache_options_t o = FM_STRUCT_INIT(fm_pivot_cache_options_t);
  ASSERT_EQ(fm_workbook_get_pivot_cache_options(wb.handle, cache_id, &o), 0);
  EXPECT_EQ(o.save_data, 1);
  EXPECT_EQ(o.has_missing_items_limit, 0);
  EXPECT_EQ(o.refresh_on_load, 0);

  o.save_data = 0;
  o.has_missing_items_limit = 1;
  o.missing_items_limit = 0;
  o.refresh_on_load = 1;
  ASSERT_EQ(UpdateCacheOptionsFrom(wb.handle, cache_id, &o), 0) << fm_get_last_error_message();
  std::int32_t pending = -1;
  ASSERT_EQ(fm_sheet_is_pivot_projection_pending(wb.handle, 0U, kPivot, &pending), 0);
  EXPECT_EQ(pending, 1);

  const std::vector<std::uint8_t> package = SavedPackage(wb.handle);
  const std::string def = ExtractZipEntry(package, "xl/pivotCache/pivotCacheDefinition1.xml");
  EXPECT_NE(def.find("saveData=\"0\""), std::string::npos) << def;
  EXPECT_NE(def.find("missingItemsLimit=\"0\""), std::string::npos) << def;
  EXPECT_NE(def.find("refreshOnLoad=\"1\""), std::string::npos) << def;
  // Without records: no part, no content-type override, no relationship.
  EXPECT_FALSE(HasZipEntry(package, "xl/pivotCache/pivotCacheRecords1.xml"));
  EXPECT_EQ(ExtractZipEntry(package, "[Content_Types].xml").find("pivotCacheRecords"), std::string::npos);
  EXPECT_EQ(def.find("r:id="), std::string::npos) << def;
  if (HasZipEntry(package, "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels")) {
    EXPECT_EQ(ExtractZipEntry(package, "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels").find("pivotCacheRecords"),
              std::string::npos);
  }

  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(package.data(), package.size(), &loaded.handle), 0) << fm_get_last_error_message();
  fm_pivot_cache_options_t read = FM_STRUCT_INIT(fm_pivot_cache_options_t);
  ASSERT_EQ(fm_workbook_get_pivot_cache_options(loaded.handle, cache_id, &read), 0);
  EXPECT_EQ(read.save_data, 0);
  EXPECT_EQ(read.has_missing_items_limit, 1);
  EXPECT_EQ(read.missing_items_limit, 0U);
  EXPECT_EQ(read.refresh_on_load, 1);
}

TEST(FormulonCApiPivotOptions, CacheMissingItemsLimitMaxAndBoundaries) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, nullptr), 0) << fm_get_last_error_message();
  fm_pivot_cache_options_t o = FM_STRUCT_INIT(fm_pivot_cache_options_t);
  ASSERT_EQ(fm_workbook_get_pivot_cache_options(wb.handle, cache_id, &o), 0);
  o.has_missing_items_limit = 1;
  o.missing_items_limit = 32500;
  ASSERT_EQ(UpdateCacheOptionsFrom(wb.handle, cache_id, &o), 0) << fm_get_last_error_message();
  const std::string def = ExtractZipEntry(SavedPackage(wb.handle), "xl/pivotCache/pivotCacheDefinition1.xml");
  EXPECT_NE(def.find("missingItemsLimit=\"32500\""), std::string::npos) << def;

  for (const std::uint32_t limit : {0U, 1U, 1048576U}) {
    o.has_missing_items_limit = 1;
    o.missing_items_limit = limit;
    EXPECT_EQ(UpdateCacheOptionsFrom(wb.handle, cache_id, &o), 0) << limit;
    fm_pivot_cache_options_t read = FM_STRUCT_INIT(fm_pivot_cache_options_t);
    ASSERT_EQ(fm_workbook_get_pivot_cache_options(wb.handle, cache_id, &read), 0);
    EXPECT_EQ(read.has_missing_items_limit, 1);
    EXPECT_EQ(read.missing_items_limit, limit);
  }
  o.has_missing_items_limit = 0;
  EXPECT_EQ(UpdateCacheOptionsFrom(wb.handle, cache_id, &o), 0);
  fm_pivot_cache_options_t automatic = FM_STRUCT_INIT(fm_pivot_cache_options_t);
  ASSERT_EQ(fm_workbook_get_pivot_cache_options(wb.handle, cache_id, &automatic), 0);
  EXPECT_EQ(automatic.has_missing_items_limit, 0);
  for (const std::uint32_t limit : {1048577U, UINT32_MAX}) {
    o.has_missing_items_limit = 1;
    o.missing_items_limit = limit;
    EXPECT_EQ(UpdateCacheOptionsFrom(wb.handle, cache_id, &o), kInvalidArgument) << limit;
  }
  // A non-zero flag reads 1.
  o.has_missing_items_limit = 0;
  o.save_data = 5;
  o.refresh_on_load = -3;
  ASSERT_EQ(UpdateCacheOptionsFrom(wb.handle, cache_id, &o), 0);
  fm_pivot_cache_options_t read = FM_STRUCT_INIT(fm_pivot_cache_options_t);
  ASSERT_EQ(fm_workbook_get_pivot_cache_options(wb.handle, cache_id, &read), 0);
  EXPECT_EQ(read.save_data, 1);
  EXPECT_EQ(read.refresh_on_load, 1);

  EXPECT_EQ(fm_workbook_get_pivot_cache_options(nullptr, cache_id, &read), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_get_pivot_cache_options(wb.handle, cache_id, nullptr), kBindingNullPointer);
  EXPECT_EQ(UpdateCacheOptionsFrom(nullptr, cache_id, &o), kBindingNullPointer);
  EXPECT_EQ(UpdateCacheOptionsFrom(wb.handle, cache_id, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_get_pivot_cache_options(wb.handle, cache_id + 1U, &read), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(UpdateCacheOptionsFrom(wb.handle, cache_id + 1U, &o), FM_STATUS_NOT_FOUND);
}

// ---------------------------------------------------------------------------
// Field names and showAll
// ---------------------------------------------------------------------------

TEST(FormulonCApiPivotOptions, CustomNameShowsInTheHeaderAndClears) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  ASSERT_EQ(fm_sheet_set_pivot_report_layout(wb.handle, 0U, kPivot, FM_PIVOT_REPORT_LAYOUT_TABULAR), 0);
  ASSERT_EQ(SetPivotFieldCustomName(wb.handle, 0U, kPivot, kRegionField, "Area"), 0) << fm_get_last_error_message();
  fm_pivot_field_info_t info = FM_STRUCT_INIT(fm_pivot_field_info_t);
  ASSERT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0U, kPivot, kRegionField, &info), 0);
  EXPECT_STREQ(info.custom_name, "Area");
  EXPECT_STREQ(info.source_name, "Region");
  EXPECT_TRUE(LayoutHasText(wb.handle, kPivot, "Area"));
  EXPECT_NE(FieldTag(SavedPivotXml(wb.handle), "axisRow").find("name=\"Area\""), std::string::npos);

  for (const char* cleared : {static_cast<const char*>(nullptr), ""}) {
    ASSERT_EQ(SetPivotFieldCustomName(wb.handle, 0U, kPivot, kRegionField, "Area"), 0);
    ASSERT_EQ(SetPivotFieldCustomName(wb.handle, 0U, kPivot, kRegionField, cleared), 0);
    ASSERT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0U, kPivot, kRegionField, &info), 0);
    EXPECT_EQ(info.custom_name, nullptr);
    EXPECT_FALSE(LayoutHasText(wb.handle, kPivot, "Area"));
    EXPECT_TRUE(LayoutHasText(wb.handle, kPivot, "Region"));
  }
}

TEST(FormulonCApiPivotOptions, CustomNameMayNotRepeatAnotherName) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  // Another field's source name, or a data field's, under the collation.
  for (const char* name : {"Amount", "AMOUNT", "Sum of Amount", "sum of amount"}) {
    EXPECT_EQ(SetPivotFieldCustomName(wb.handle, 0U, kPivot, kRegionField, name), FM_STATUS_ALREADY_EXISTS) << name;
  }
  // Another field's custom name.
  ASSERT_EQ(SetPivotFieldCustomName(wb.handle, 0U, kPivot, kAmountField, "Value"), 0);
  EXPECT_EQ(SetPivotFieldCustomName(wb.handle, 0U, kPivot, kRegionField, "value"), FM_STATUS_ALREADY_EXISTS);
  // The field's own source name is allowed.
  EXPECT_EQ(SetPivotFieldCustomName(wb.handle, 0U, kPivot, kRegionField, "Region"), 0) << fm_get_last_error_message();
}

TEST(FormulonCApiPivotOptions, CustomNameBoundaries) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  EXPECT_EQ(SetPivotFieldCustomName(wb.handle, 0U, kPivot, kRegionField, (Repeat(kEmoji, 127) + "a").c_str()), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(SetPivotFieldCustomName(wb.handle, 0U, kPivot, kRegionField, Repeat(kEmoji, 128).c_str()),
            kInvalidArgument);
  EXPECT_EQ(SetPivotFieldCustomName(wb.handle, 0U, kPivot, kRegionField, "\xC3"), kBindingUtf8EncodingError);
  EXPECT_EQ(SetPivotFieldCustomName(nullptr, 0U, kPivot, kRegionField, "x"), kBindingNullPointer);
  for (const std::size_t index : {std::size_t{9}, std::size_t{0x80000000U}, static_cast<std::size_t>(0x100000000ULL)}) {
    EXPECT_EQ(SetPivotFieldCustomName(wb.handle, 0U, kPivot, index, "x"), FM_STATUS_INDEX_OUT_OF_RANGE) << index;
    EXPECT_EQ(SetPivotFieldCustomName(wb.handle, 0U, index, kRegionField, "x"), FM_STATUS_INDEX_OUT_OF_RANGE);
    EXPECT_EQ(SetPivotFieldShowAll(wb.handle, 0U, kPivot, index, 1), FM_STATUS_INDEX_OUT_OF_RANGE) << index;
    EXPECT_EQ(SetPivotFieldShowAll(wb.handle, index, kPivot, kRegionField, 1), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  }
  EXPECT_EQ(SetPivotFieldShowAll(nullptr, 0U, kPivot, kRegionField, 1), kBindingNullPointer);
}

TEST(FormulonCApiPivotOptions, ShowAllListsItemsWithoutData) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, nullptr), 0) << fm_get_last_error_message();
  // East is an item of the field with no records.
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(wb.handle, cache_id, 0U, "East"), 0);
  ASSERT_EQ(fm_sheet_add_pivot_field_item(wb.handle, 0U, kPivot, kRegionField, "East", 1, ignored_index()), 0);
  fm_pivot_field_info_t info = FM_STRUCT_INIT(fm_pivot_field_info_t);
  ASSERT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0U, kPivot, kRegionField, &info), 0);
  EXPECT_EQ(info.show_all, 0);
  EXPECT_FALSE(LayoutHasText(wb.handle, kPivot, "East"));
  EXPECT_NE(FieldTag(SavedPivotXml(wb.handle), "axisRow").find("showAll=\"0\""), std::string::npos);

  ASSERT_EQ(SetPivotFieldShowAll(wb.handle, 0U, kPivot, kRegionField, 1), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0U, kPivot, kRegionField, &info), 0);
  EXPECT_EQ(info.show_all, 1);
  EXPECT_TRUE(LayoutHasText(wb.handle, kPivot, "East"));

  // The empty value cell shows the missing caption; a lone "-" is the number 0.
  fm_pivot_options_t o = GetOptions(wb.handle);
  o.missing_caption = "N/A";
  ASSERT_EQ(UpdateOptionsFrom(wb.handle, 0U, kPivot, &o), 0) << fm_get_last_error_message();
  EXPECT_TRUE(LayoutHasText(wb.handle, kPivot, "N/A"));
  o.missing_caption = "-";
  ASSERT_EQ(UpdateOptionsFrom(wb.handle, 0U, kPivot, &o), 0) << fm_get_last_error_message();
  EXPECT_FALSE(LayoutHasText(wb.handle, kPivot, "-"));

  const std::vector<std::uint8_t> package = SavedPackage(wb.handle);
  EXPECT_EQ(FieldTag(ExtractZipEntry(package, "xl/pivotTables/pivotTable1.xml"), "axisRow").find("showAll"),
            std::string::npos);
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(package.data(), package.size(), &loaded.handle), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_pivot_field_at(loaded.handle, 0U, kPivot, kRegionField, &info), 0);
  EXPECT_EQ(info.show_all, 1);
  ASSERT_EQ(SetPivotFieldShowAll(loaded.handle, 0U, kPivot, kRegionField, 0), 0);
  ASSERT_EQ(fm_sheet_get_pivot_field_at(loaded.handle, 0U, kPivot, kRegionField, &info), 0);
  EXPECT_EQ(info.show_all, 0);
  EXPECT_FALSE(LayoutHasText(loaded.handle, kPivot, "East"));
}

// ---------------------------------------------------------------------------
// Default data-field names
// ---------------------------------------------------------------------------

TEST(FormulonCApiPivotOptions, DefaultNameIgnoresAnUnrelatedName) {
  WorkbookGuard wb;
  CreateJaScratch(wb);
  // The existing data field is "Sum of Amount"; the new one is the plain name.
  EXPECT_EQ(DefaultName(wb.handle, kAmountField, FM_PIVOT_AGGREGATION_SUM), kSumOfAmount);
  EXPECT_EQ(AddUnnamed(wb.handle, kAmountField, FM_PIVOT_AGGREGATION_SUM), kSumOfAmount);
  EXPECT_EQ(DefaultName(wb.handle, kAmountField, FM_PIVOT_AGGREGATION_COUNT), "\xE5\x80\x8B\xE6\x95\xB0 / Amount");
}

TEST(FormulonCApiPivotOptions, DefaultNameNumbersPastTakenDataFieldNames) {
  WorkbookGuard wb;
  CreateJaScratch(wb);
  fm_pivot_data_field_t renamed = FM_STRUCT_INIT(fm_pivot_data_field_t);
  ASSERT_EQ(fm_sheet_get_pivot_data_field_at(wb.handle, 0U, kPivot, 0U, &renamed), 0);
  renamed.name = kSumOfAmount;
  renamed.number_format = "";
  ASSERT_EQ(fm_sheet_set_pivot_data_field(wb.handle, 0U, kPivot, 0U, &renamed), 0) << fm_get_last_error_message();
  // DN4: 合計 / Amount2, then 合計 / Amount3.
  EXPECT_EQ(DefaultName(wb.handle, kAmountField, FM_PIVOT_AGGREGATION_SUM), std::string(kSumOfAmount) + "2");
  EXPECT_EQ(AddUnnamed(wb.handle, kAmountField, FM_PIVOT_AGGREGATION_SUM), std::string(kSumOfAmount) + "2");
  EXPECT_EQ(DefaultName(wb.handle, kAmountField, FM_PIVOT_AGGREGATION_SUM), std::string(kSumOfAmount) + "3");
  EXPECT_EQ(AddUnnamed(wb.handle, kAmountField, FM_PIVOT_AGGREGATION_SUM), std::string(kSumOfAmount) + "3");
  // The entry a set replaces does not count against itself.
  EXPECT_EQ(DefaultName(wb.handle, kAmountField, FM_PIVOT_AGGREGATION_SUM, 0U), kSumOfAmount);
  fm_pivot_data_field_t unnamed = renamed;
  unnamed.name = nullptr;
  unnamed.number_format = "";
  ASSERT_EQ(fm_sheet_set_pivot_data_field(wb.handle, 0U, kPivot, 0U, &unnamed), 0) << fm_get_last_error_message();
  fm_pivot_data_field_t read = FM_STRUCT_INIT(fm_pivot_data_field_t);
  ASSERT_EQ(fm_sheet_get_pivot_data_field_at(wb.handle, 0U, kPivot, 0U, &read), 0);
  EXPECT_STREQ(read.name, kSumOfAmount);
}

TEST(FormulonCApiPivotOptions, DefaultNameNumbersPastAFieldName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "win-365-ja_JP"), 0);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0, &cache_id), 0);
  for (const char* name : {"Region", "Amount", kSumOfAmount}) {
    std::size_t index = 99;
    ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb.handle, cache_id, name, &index), 0);
  }
  std::size_t record = 99;
  ASSERT_EQ(fm_workbook_add_pivot_cache_record(wb.handle, cache_id, &record), 0);
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_text(wb.handle, cache_id, record, 0U, "North"), 0);
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, record, 1U, 10.0), 0);
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, record, 2U, 3.0), 0);
  std::size_t pivot = 99;
  ASSERT_EQ(AddPivot(wb.handle, 0U, "PT", cache_id, 0U, 4U, &pivot), 0) << fm_get_last_error_message();
  fm_pivot_field_spec_t spec = FM_STRUCT_INIT(fm_pivot_field_spec_t);
  spec.source_name = "Region";
  spec.custom_name = "";
  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.subtotal_top = FM_TRISTATE_FALSE;
  spec.number_format = "";
  std::size_t field = 99;
  ASSERT_EQ(fm_sheet_add_pivot_field(wb.handle, 0U, pivot, &spec, &field), 0) << fm_get_last_error_message();
  spec.axis = FM_PIVOT_AXIS_VALUE;
  for (const char* name : {"Amount", kSumOfAmount}) {
    spec.source_name = name;
    ASSERT_EQ(fm_sheet_add_pivot_field(wb.handle, 0U, pivot, &spec, &field), 0) << fm_get_last_error_message();
  }
  ASSERT_EQ(field, 2U);
  // DN3: a field named like the default caption pushes Amount to 2, and that
  // field's own Sum is named after it.
  EXPECT_EQ(DefaultName(wb.handle, 1U, FM_PIVOT_AGGREGATION_SUM), std::string(kSumOfAmount) + "2");
  EXPECT_EQ(AddUnnamed(wb.handle, 1U, FM_PIVOT_AGGREGATION_SUM), std::string(kSumOfAmount) + "2");
  const std::string nested = "\xE5\x90\x88\xE8\xA8\x88 / " + std::string(kSumOfAmount);
  EXPECT_EQ(DefaultName(wb.handle, 2U, FM_PIVOT_AGGREGATION_SUM), nested);
  EXPECT_EQ(AddUnnamed(wb.handle, 2U, FM_PIVOT_AGGREGATION_SUM), nested);
}

TEST(FormulonCApiPivotOptions, DefaultNameWritesNothing) {
  WorkbookGuard wb;
  CreateJaScratch(wb);
  const std::string before = SavedPivotXml(wb.handle);
  fm_value_t cell{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0U, 0U, 4U, &cell), 0);
  const std::string header = cell.kind == FM_VALUE_KIND_TEXT ? std::string(cell.u.text) : std::string();
  EXPECT_EQ(DefaultName(wb.handle, kAmountField, FM_PIVOT_AGGREGATION_SUM), kSumOfAmount);
  EXPECT_EQ(DefaultName(wb.handle, kRegionField, FM_PIVOT_AGGREGATION_COUNT, 0U), "\xE5\x80\x8B\xE6\x95\xB0 / Region");
  EXPECT_EQ(SavedPivotXml(wb.handle), before);
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0U, 0U, 4U, &cell), 0);
  EXPECT_EQ(cell.kind == FM_VALUE_KIND_TEXT ? std::string(cell.u.text) : std::string(), header);
  std::size_t data_fields = 0;
  ASSERT_EQ(fm_sheet_get_pivot_data_field_count(wb.handle, 0U, kPivot, &data_fields), 0);
  EXPECT_EQ(data_fields, 1U);
}

TEST(FormulonCApiPivotOptions, DefaultNameArgumentBoundaries) {
  WorkbookGuard wb;
  CreateJaScratch(wb);
  const char* name = nullptr;
  EXPECT_EQ(
      fm_sheet_get_pivot_data_field_default_name(nullptr, 0U, kPivot, 1U, FM_PIVOT_AGGREGATION_SUM, SIZE_MAX, &name),
      kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_pivot_data_field_default_name(wb.handle, 0U, kPivot, 1U, FM_PIVOT_AGGREGATION_SUM, SIZE_MAX,
                                                       nullptr),
            kBindingNullPointer);
  for (const std::int32_t aggregation : {-1, 11, INT32_MAX}) {
    EXPECT_EQ(fm_sheet_get_pivot_data_field_default_name(wb.handle, 0U, kPivot, 1U, aggregation, SIZE_MAX, &name),
              kInvalidArgument)
        << aggregation;
  }
  EXPECT_EQ(
      fm_sheet_get_pivot_data_field_default_name(wb.handle, 0U, kPivot, 1U, FM_PIVOT_AGGREGATION_VARP, SIZE_MAX, &name),
      0);
  for (const std::size_t index : {std::size_t{2}, std::size_t{0x80000000U}, static_cast<std::size_t>(0x100000000ULL)}) {
    EXPECT_EQ(fm_sheet_get_pivot_data_field_default_name(wb.handle, 0U, kPivot, index, FM_PIVOT_AGGREGATION_SUM,
                                                         SIZE_MAX, &name),
              FM_STATUS_INDEX_OUT_OF_RANGE)
        << index;
    EXPECT_EQ(fm_sheet_get_pivot_data_field_default_name(wb.handle, 0U, kPivot, 1U, FM_PIVOT_AGGREGATION_SUM,
                                                         index - 1U, &name),
              FM_STATUS_INDEX_OUT_OF_RANGE)
        << index;
    EXPECT_EQ(fm_sheet_get_pivot_data_field_default_name(wb.handle, index, kPivot, 1U, FM_PIVOT_AGGREGATION_SUM,
                                                         SIZE_MAX, &name),
              FM_STATUS_SHEET_INDEX_OUT_OF_RANGE)
        << index;
    EXPECT_EQ(
        fm_sheet_get_pivot_data_field_default_name(wb.handle, 0U, index, 1U, FM_PIVOT_AGGREGATION_SUM, SIZE_MAX, &name),
        FM_STATUS_INDEX_OUT_OF_RANGE)
        << index;
  }
}
