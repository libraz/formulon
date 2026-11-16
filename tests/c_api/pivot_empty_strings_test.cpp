// Empty strings in the PIVOT setters: where an empty caption is a stored state (item captions, data field names) it
// is a value and the absent value clears or derives; where the file holds one state for both (field names and
// number formats, option captions) they mean the same; a worksheet source attribute may not be empty.

#include <cstddef>
#include <cstdint>
#include <string>

#include "formulon_c_pivot_test_helpers.h"

namespace {

constexpr std::size_t kPivot = 0U;
constexpr std::size_t kRegionField = 0U;
constexpr std::size_t kAmountField = 1U;

// A scratch pivot whose cache is refreshed and whose report is projected, so every save format can write it.
void BuildProjectedPivot(WorkbookGuard& out, std::uint32_t* cache_id = nullptr) {
  ASSERT_EQ(fm_workbook_create_default(&out.handle), 0);
  std::uint32_t cache = 0;
  ASSERT_EQ(BuildScratchPivot(out.handle, &cache, nullptr), 0) << fm_get_last_error_message();
  FillScratchSource(out.handle);
  ASSERT_EQ(fm_workbook_refresh_pivot_cache(out.handle, cache), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_project_pivot(out.handle, 0U, kPivot), 0) << fm_get_last_error_message();
  if (cache_id != nullptr) {
    *cache_id = cache;
  }
}

void SaveAndReload(fm_workbook_t* wb, std::int32_t format, WorkbookGuard& out) {
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb, format, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, nullptr, &out.handle), 0) << fm_get_last_error_message();
}

fm_status_t PatchItemCaption(fm_workbook_t* wb, std::size_t item, fm_patch_op_t op, const char* caption) {
  fm_pivot_item_patch_t patch = FM_STRUCT_INIT(fm_pivot_item_patch_t);
  patch.caption_op = op;
  patch.caption = caption;
  return fm_sheet_update_pivot_field_item(wb, 0U, kPivot, kRegionField, item, &patch);
}

// The item's caption; `present` tells an empty label from none.
std::string ItemCaption(const fm_workbook_t* wb, std::size_t item, bool* present) {
  fm_pivot_item_info_t info = FM_STRUCT_INIT(fm_pivot_item_info_t);
  EXPECT_EQ(fm_sheet_get_pivot_field_item_at(wb, 0U, kPivot, kRegionField, item, &info), 0);
  *present = info.caption != nullptr;
  return info.caption != nullptr ? info.caption : "";
}

fm_status_t SetDataFieldName(fm_workbook_t* wb, const char* name, const char* number_format = nullptr) {
  fm_pivot_data_field_t spec = FM_STRUCT_INIT(fm_pivot_data_field_t);
  spec.name = name;
  spec.field_index = static_cast<std::uint32_t>(kAmountField);
  spec.aggregation = FM_PIVOT_AGGREGATION_SUM;
  spec.number_format = number_format;
  spec.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;
  return fm_sheet_set_pivot_data_field(wb, 0U, kPivot, 0U, &spec);
}

fm_pivot_data_field_t DataField(const fm_workbook_t* wb) {
  fm_pivot_data_field_t out = FM_STRUCT_INIT(fm_pivot_data_field_t);
  EXPECT_EQ(fm_sheet_get_pivot_data_field_at(wb, 0U, kPivot, 0U, &out), 0);
  return out;
}

// ---------------------------------------------------------------------------
// Field custom name and number format: "" and the clear op are one state
// ---------------------------------------------------------------------------

TEST(PivotEmptyStrings, FieldCustomNameEmptyRestoresTheSourceName) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  for (const fm_patch_op_t op : {FM_PATCH_OP_SET, FM_PATCH_OP_CLEAR}) {
    fm_pivot_field_patch_t patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
    patch.custom_name_op = FM_PATCH_OP_SET;
    patch.custom_name = "Area";
    ASSERT_EQ(fm_sheet_update_pivot_field(wb.handle, 0U, kPivot, kRegionField, &patch), 0)
        << fm_get_last_error_message();
    fm_pivot_field_info_t info = FM_STRUCT_INIT(fm_pivot_field_info_t);
    ASSERT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0U, kPivot, kRegionField, &info), 0);
    ASSERT_NE(info.custom_name, nullptr);
    EXPECT_STREQ(info.custom_name, "Area");

    patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
    patch.custom_name_op = op;
    patch.custom_name = op == FM_PATCH_OP_SET ? "" : nullptr;
    ASSERT_EQ(fm_sheet_update_pivot_field(wb.handle, 0U, kPivot, kRegionField, &patch), 0)
        << fm_get_last_error_message();
    info = FM_STRUCT_INIT(fm_pivot_field_info_t);
    ASSERT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0U, kPivot, kRegionField, &info), 0);
    // No report name is left, so the report shows the source name.
    EXPECT_EQ(info.custom_name, nullptr) << "op=" << op;
    EXPECT_STREQ(info.source_name, "Region");
  }
}

TEST(PivotEmptyStrings, FieldNumberFormatEmptyMeansNone) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  for (const fm_patch_op_t op : {FM_PATCH_OP_SET, FM_PATCH_OP_CLEAR}) {
    fm_pivot_field_patch_t patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
    patch.number_format_op = FM_PATCH_OP_SET;
    patch.number_format = "0.00";
    ASSERT_EQ(fm_sheet_update_pivot_field(wb.handle, 0U, kPivot, kAmountField, &patch), 0)
        << fm_get_last_error_message();
    fm_pivot_field_info_t info = FM_STRUCT_INIT(fm_pivot_field_info_t);
    ASSERT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0U, kPivot, kAmountField, &info), 0);
    ASSERT_NE(info.number_format, nullptr);

    patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
    patch.number_format_op = op;
    patch.number_format = op == FM_PATCH_OP_SET ? "" : nullptr;
    ASSERT_EQ(fm_sheet_update_pivot_field(wb.handle, 0U, kPivot, kAmountField, &patch), 0)
        << fm_get_last_error_message();
    info = FM_STRUCT_INIT(fm_pivot_field_info_t);
    ASSERT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0U, kPivot, kAmountField, &info), 0);
    EXPECT_EQ(info.number_format, nullptr) << "op=" << op;
  }
}

// ---------------------------------------------------------------------------
// Item caption: "" is an empty label, the clear op un-renames
// ---------------------------------------------------------------------------

TEST(PivotEmptyStrings, ItemCaptionEmptyIsALabelAndClearUnrenames) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  bool present = true;
  ItemCaption(wb.handle, 0U, &present);
  EXPECT_FALSE(present);

  ASSERT_EQ(PatchItemCaption(wb.handle, 0U, FM_PATCH_OP_SET, ""), 0) << fm_get_last_error_message();
  EXPECT_EQ(ItemCaption(wb.handle, 0U, &present), "");
  EXPECT_TRUE(present) << "an empty label reads as an empty string, not as no caption";

  ASSERT_EQ(PatchItemCaption(wb.handle, 0U, FM_PATCH_OP_CLEAR, nullptr), 0) << fm_get_last_error_message();
  ItemCaption(wb.handle, 0U, &present);
  EXPECT_FALSE(present);

  // The label sticks until it is replaced.
  ASSERT_EQ(PatchItemCaption(wb.handle, 1U, FM_PATCH_OP_SET, ""), 0) << fm_get_last_error_message();
  ASSERT_EQ(PatchItemCaption(wb.handle, 1U, FM_PATCH_OP_SET, "Sud"), 0) << fm_get_last_error_message();
  EXPECT_EQ(ItemCaption(wb.handle, 1U, &present), "Sud");
  ASSERT_EQ(PatchItemCaption(wb.handle, 1U, FM_PATCH_OP_SET, ""), 0) << fm_get_last_error_message();
  EXPECT_EQ(ItemCaption(wb.handle, 1U, &present), "");
  EXPECT_TRUE(present);
}

TEST(PivotEmptyStrings, ItemCaptionEmptyLabelIsStoredAsAnEmptyNameAttribute) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  ASSERT_EQ(PatchItemCaption(wb.handle, 0U, FM_PATCH_OP_SET, ""), 0) << fm_get_last_error_message();
  const std::string xml = SavedPivotXml(wb.handle);
  EXPECT_NE(xml.find(" n=\"\""), std::string::npos) << xml;

  ASSERT_EQ(PatchItemCaption(wb.handle, 0U, FM_PATCH_OP_CLEAR, nullptr), 0) << fm_get_last_error_message();
  EXPECT_EQ(SavedPivotXml(wb.handle).find(" n=\""), std::string::npos);
}

class PivotEmptyStringsFormat : public ::testing::TestWithParam<std::int32_t> {};

TEST_P(PivotEmptyStringsFormat, ItemCaptionEmptyLabelSurvivesSaveAndReload) {
  WorkbookGuard wb;
  BuildProjectedPivot(wb);
  ASSERT_EQ(PatchItemCaption(wb.handle, 0U, FM_PATCH_OP_SET, ""), 0) << fm_get_last_error_message();
  ASSERT_EQ(PatchItemCaption(wb.handle, 1U, FM_PATCH_OP_SET, "Sud"), 0) << fm_get_last_error_message();
  WorkbookGuard reloaded;
  SaveAndReload(wb.handle, GetParam(), reloaded);
  ASSERT_NE(reloaded.handle, nullptr);
  bool present = false;
  EXPECT_EQ(ItemCaption(reloaded.handle, 0U, &present), "");
  EXPECT_TRUE(present);
  EXPECT_EQ(ItemCaption(reloaded.handle, 1U, &present), "Sud");

  // Un-renamed, the item has no caption after a round trip either.
  ASSERT_EQ(PatchItemCaption(wb.handle, 0U, FM_PATCH_OP_CLEAR, nullptr), 0) << fm_get_last_error_message();
  WorkbookGuard again;
  SaveAndReload(wb.handle, GetParam(), again);
  ASSERT_NE(again.handle, nullptr);
  ItemCaption(again.handle, 0U, &present);
  EXPECT_FALSE(present);
}

// ---------------------------------------------------------------------------
// Data field name: "" is an empty caption, NULL derives
// ---------------------------------------------------------------------------

TEST(PivotEmptyStrings, DataFieldNameEmptyIsKeptAndNullDerives) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  const char* derived = nullptr;
  ASSERT_EQ(fm_sheet_get_pivot_data_field_default_name(wb.handle, 0U, kPivot, kAmountField, FM_PIVOT_AGGREGATION_SUM,
                                                       0U, &derived),
            0);
  ASSERT_NE(derived, nullptr);
  const std::string default_name = derived;
  ASSERT_FALSE(default_name.empty());

  ASSERT_EQ(SetDataFieldName(wb.handle, ""), 0) << fm_get_last_error_message();
  ASSERT_NE(DataField(wb.handle).name, nullptr);
  EXPECT_STREQ(DataField(wb.handle).name, "");

  ASSERT_EQ(SetDataFieldName(wb.handle, nullptr), 0) << fm_get_last_error_message();
  EXPECT_EQ(std::string(DataField(wb.handle).name), default_name);

  // An add with "" keeps the empty caption beside the derived one.
  fm_pivot_data_field_t spec = FM_STRUCT_INIT(fm_pivot_data_field_t);
  spec.name = "";
  spec.field_index = static_cast<std::uint32_t>(kAmountField);
  spec.aggregation = FM_PIVOT_AGGREGATION_COUNT;
  spec.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;
  std::size_t index = 99;
  ASSERT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0U, kPivot, &spec, &index), 0) << fm_get_last_error_message();
  fm_pivot_data_field_t added = FM_STRUCT_INIT(fm_pivot_data_field_t);
  ASSERT_EQ(fm_sheet_get_pivot_data_field_at(wb.handle, 0U, kPivot, index, &added), 0);
  EXPECT_STREQ(added.name, "");
}

TEST(PivotEmptyStrings, DataFieldEmptyNameIsStoredAsAnEmptyNameAttribute) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  ASSERT_EQ(SetDataFieldName(wb.handle, ""), 0) << fm_get_last_error_message();
  const std::string xml = SavedPivotXml(wb.handle);
  EXPECT_NE(xml.find("<dataField name=\"\""), std::string::npos) << xml;
}

TEST_P(PivotEmptyStringsFormat, DataFieldEmptyNameSurvivesSaveAndReload) {
  WorkbookGuard wb;
  BuildProjectedPivot(wb);
  ASSERT_EQ(SetDataFieldName(wb.handle, ""), 0) << fm_get_last_error_message();
  WorkbookGuard reloaded;
  SaveAndReload(wb.handle, GetParam(), reloaded);
  ASSERT_NE(reloaded.handle, nullptr);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_pivot_data_field_count(reloaded.handle, 0U, kPivot, &count), 0);
  ASSERT_EQ(count, 1U);
  ASSERT_NE(DataField(reloaded.handle).name, nullptr);
  EXPECT_STREQ(DataField(reloaded.handle).name, "");
}

INSTANTIATE_TEST_SUITE_P(Formats, PivotEmptyStringsFormat, ::testing::Values(FM_WORKBOOK_FORMAT_XLSX));

TEST(PivotEmptyStrings, DataFieldNumberFormatEmptyMeansNone) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  ASSERT_EQ(SetDataFieldName(wb.handle, "Total", "0.00"), 0) << fm_get_last_error_message();
  EXPECT_NE(DataField(wb.handle).number_format, nullptr);
  ASSERT_EQ(SetDataFieldName(wb.handle, "Total", ""), 0) << fm_get_last_error_message();
  EXPECT_EQ(DataField(wb.handle).number_format, nullptr);
  ASSERT_EQ(SetDataFieldName(wb.handle, "Total", "0.00"), 0) << fm_get_last_error_message();
  ASSERT_EQ(SetDataFieldName(wb.handle, "Total", nullptr), 0) << fm_get_last_error_message();
  EXPECT_EQ(DataField(wb.handle).number_format, nullptr);
}

// ---------------------------------------------------------------------------
// Option captions: "" and the clear op unset
// ---------------------------------------------------------------------------

TEST(PivotEmptyStrings, OptionCaptionsEmptyUnsetLikeClear) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  for (const fm_patch_op_t op : {FM_PATCH_OP_SET, FM_PATCH_OP_CLEAR}) {
    fm_pivot_options_patch_t set = FM_STRUCT_INIT(fm_pivot_options_patch_t);
    set.error_caption_op = set.missing_caption_op = set.alt_text_op = set.alt_text_summary_op =
        set.grand_total_caption_op = FM_PATCH_OP_SET;
    set.error_caption = "E";
    set.missing_caption = "M";
    set.alt_text = "T";
    set.alt_text_summary = "S";
    set.grand_total_caption = "G";
    ASSERT_EQ(fm_sheet_update_pivot_options(wb.handle, 0U, kPivot, &set), 0) << fm_get_last_error_message();
    fm_pivot_options_t read = FM_STRUCT_INIT(fm_pivot_options_t);
    ASSERT_EQ(fm_sheet_get_pivot_options(wb.handle, 0U, kPivot, &read), 0);
    ASSERT_NE(read.error_caption, nullptr);
    ASSERT_NE(read.grand_total_caption, nullptr);

    fm_pivot_options_patch_t unset = FM_STRUCT_INIT(fm_pivot_options_patch_t);
    unset.error_caption_op = unset.missing_caption_op = unset.alt_text_op = unset.alt_text_summary_op =
        unset.grand_total_caption_op = op;
    if (op == FM_PATCH_OP_SET) {
      unset.error_caption = unset.missing_caption = unset.alt_text = unset.alt_text_summary =
          unset.grand_total_caption = "";
    }
    ASSERT_EQ(fm_sheet_update_pivot_options(wb.handle, 0U, kPivot, &unset), 0) << fm_get_last_error_message();
    read = FM_STRUCT_INIT(fm_pivot_options_t);
    ASSERT_EQ(fm_sheet_get_pivot_options(wb.handle, 0U, kPivot, &read), 0);
    EXPECT_EQ(read.error_caption, nullptr) << "op=" << op;
    EXPECT_EQ(read.missing_caption, nullptr) << "op=" << op;
    EXPECT_EQ(read.alt_text, nullptr) << "op=" << op;
    EXPECT_EQ(read.alt_text_summary, nullptr) << "op=" << op;
    // Nothing is stored, so the locale's grand total word shows again.
    EXPECT_EQ(read.grand_total_caption, nullptr) << "op=" << op;
  }
}

TEST(PivotEmptyStrings, GrandTotalCaptionEmptyWritesNothing) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  fm_pivot_options_patch_t patch = FM_STRUCT_INIT(fm_pivot_options_patch_t);
  patch.grand_total_caption_op = FM_PATCH_OP_SET;
  patch.grand_total_caption = "TOTALX";
  ASSERT_EQ(fm_sheet_update_pivot_options(wb.handle, 0U, kPivot, &patch), 0) << fm_get_last_error_message();
  EXPECT_NE(SavedPivotXml(wb.handle).find("grandTotalCaption=\"TOTALX\""), std::string::npos);
  patch.grand_total_caption = "";
  ASSERT_EQ(fm_sheet_update_pivot_options(wb.handle, 0U, kPivot, &patch), 0) << fm_get_last_error_message();
  EXPECT_EQ(SavedPivotXml(wb.handle).find("grandTotalCaption"), std::string::npos);
}

// ---------------------------------------------------------------------------
// Worksheet source: "" is refused, NULL is absent
// ---------------------------------------------------------------------------

TEST(PivotEmptyStrings, WorksheetSourceAttributeEmptyIsRefused) {
  WorkbookGuard wb;
  std::uint32_t cache = 0;
  BuildProjectedPivot(wb, &cache);
  struct Case {
    const char* ref;
    const char* sheet;
    const char* name;
  };
  for (const Case& c : {Case{"", "Sheet1", nullptr}, Case{"A1:B5", "", nullptr}, Case{nullptr, nullptr, ""}}) {
    EXPECT_EQ(SetPivotCacheWorksheetSource(wb.handle, cache, 1, c.ref, c.sheet, c.name), FM_STATUS_INVALID_ARGUMENT);
  }
  // The refused calls changed nothing.
  int32_t present = 0;
  fm_pivot_worksheet_source_t source = FM_STRUCT_INIT(fm_pivot_worksheet_source_t);
  ASSERT_EQ(fm_workbook_get_pivot_cache_worksheet_source(wb.handle, cache, &present, &source), 0);
  EXPECT_EQ(present, 1);
  EXPECT_STREQ(source.ref, "A1:B5");
  EXPECT_STREQ(source.sheet, "Sheet1");

  // NULL is an absent attribute.
  EXPECT_EQ(SetPivotCacheWorksheetSource(wb.handle, cache, 1, "A1:B5", nullptr, nullptr), 0);
  ASSERT_EQ(fm_workbook_get_pivot_cache_worksheet_source(wb.handle, cache, &present, &source), 0);
  EXPECT_EQ(source.sheet, nullptr);
}

}  // namespace
