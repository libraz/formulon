// Stable C ABI pivot report options: grand total caption, autofit flag, field
// display flags, item captions and expansion, the Values level in the field
// orders, and Show Details. Every setter must leave the table layout-dirty and
// reject out-of-range indices and malformed text.

#include <cstddef>
#include <cstdint>
#include <functional>
#include <limits>
#include <string>
#include <vector>

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
  ASSERT_EQ(fm_workbook_create(&built.handle), 0);
  ASSERT_EQ(BuildScratchPivot(built.handle, nullptr, nullptr), 0) << fm_last_error_message();
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(built.handle, &saved.data, &saved.len), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &out.handle), 0) << fm_last_error_message();
  ASSERT_EQ(out.handle->wb->sheet(0U).pivot_tables().size(), 1U);
}

// Adds a second data field so the Values level exists.
void AddSecondDataField(fm_workbook_t* wb) {
  fm_pivot_data_field_spec_t spec{};
  spec.name = "Count of Amount";
  spec.field_index = static_cast<std::uint32_t>(kAmountField);
  spec.aggregation = FM_PIVOT_AGG_COUNT;
  spec.number_format = "";
  spec.show_as = FM_PIVOT_SHOW_AS_NORMAL;
  spec.show_as_base_field = -1;
  spec.show_as_base_item = -1;
  std::size_t index = 99;
  ASSERT_EQ(fm_workbook_pivot_data_field_add(wb, 0U, kPivot, &spec, &index), 0) << fm_last_error_message();
}

// Runs `mutation` on a layout-clean pivot and expects it to succeed and dirty the table.
void ExpectMarksDirty(const std::function<fm_status_t(fm_workbook_t*)>& mutation) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  ASSERT_FALSE(Table(wb).layout_dirty());
  ASSERT_EQ(mutation(wb.handle), 0) << fm_last_error_message();
  EXPECT_TRUE(Table(wb).layout_dirty());
}

std::vector<std::uint32_t> ReadOrder(const fm_workbook_t* wb, bool row) {
  const std::uint32_t* indices = nullptr;
  std::size_t count = 0;
  const fm_status_t status = row ? fm_workbook_pivot_get_row_field_order(wb, 0U, kPivot, &indices, &count)
                                 : fm_workbook_pivot_get_col_field_order(wb, 0U, kPivot, &indices, &count);
  EXPECT_EQ(status, 0) << fm_last_error_message();
  if (status != 0 || count == 0U) {
    return {};
  }
  return {indices, indices + count};
}

fm_status_t SetOrder(fm_workbook_t* wb, bool row, std::vector<std::uint32_t> order) {
  return row ? fm_workbook_pivot_set_row_field_order(wb, 0U, kPivot, order.data(), order.size())
             : fm_workbook_pivot_set_col_field_order(wb, 0U, kPivot, order.data(), order.size());
}

std::string Repeat(const std::string& unit, std::size_t times) {
  std::string out;
  for (std::size_t i = 0; i < times; ++i) {
    out += unit;
  }
  return out;
}

const char* const kEmoji = "\xF0\x9F\x98\x80";  // two UTF-16 code units

}  // namespace

// ---------------------------------------------------------------------------
// Table level
// ---------------------------------------------------------------------------

TEST(FormulonCApiPivotOptions, GrandTotalCaptionSetsGetsAndClears) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  const char* caption = "unset";
  ASSERT_EQ(fm_workbook_pivot_get_grand_total_caption(wb.handle, 0U, kPivot, &caption), 0);
  EXPECT_EQ(caption, nullptr);

  ASSERT_EQ(fm_workbook_pivot_set_grand_total_caption(wb.handle, 0U, kPivot, "Total \xE5\x90\x88\xE8\xA8\x88"), 0);
  ASSERT_EQ(fm_workbook_pivot_get_grand_total_caption(wb.handle, 0U, kPivot, &caption), 0);
  EXPECT_STREQ(caption, "Total \xE5\x90\x88\xE8\xA8\x88");
  EXPECT_TRUE(LayoutHasText(wb.handle, kPivot, "Total \xE5\x90\x88\xE8\xA8\x88"));

  ASSERT_EQ(fm_workbook_pivot_set_grand_total_caption(wb.handle, 0U, kPivot, nullptr), 0);
  ASSERT_EQ(fm_workbook_pivot_get_grand_total_caption(wb.handle, 0U, kPivot, &caption), 0);
  EXPECT_EQ(caption, nullptr);

  ASSERT_EQ(fm_workbook_pivot_set_grand_total_caption(wb.handle, 0U, kPivot, "All"), 0);
  ASSERT_EQ(fm_workbook_pivot_set_grand_total_caption(wb.handle, 0U, kPivot, ""), 0);
  ASSERT_EQ(fm_workbook_pivot_get_grand_total_caption(wb.handle, 0U, kPivot, &caption), 0);
  EXPECT_EQ(caption, nullptr);
}

TEST(FormulonCApiPivotOptions, GrandTotalCaptionBoundaries) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  // 255 UTF-16 units fit; 256 do not, and a surrogate pair counts as two.
  const std::string at_limit = Repeat(kEmoji, 127) + "a";
  ASSERT_EQ(fm_workbook_pivot_set_grand_total_caption(wb.handle, 0U, kPivot, at_limit.c_str()), 0)
      << fm_last_error_message();
  const char* caption = nullptr;
  ASSERT_EQ(fm_workbook_pivot_get_grand_total_caption(wb.handle, 0U, kPivot, &caption), 0);
  EXPECT_EQ(std::string(caption), at_limit);

  EXPECT_EQ(fm_workbook_pivot_set_grand_total_caption(wb.handle, 0U, kPivot, Repeat(kEmoji, 128).c_str()),
            kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_set_grand_total_caption(wb.handle, 0U, kPivot, Repeat("a", 256).c_str()),
            kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_set_grand_total_caption(wb.handle, 0U, kPivot, "\xC3"), kBindingUtf8EncodingError);
  EXPECT_EQ(fm_workbook_pivot_set_grand_total_caption(wb.handle, 9U, kPivot, "x"), kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_set_grand_total_caption(wb.handle, 0U, 9U, "x"), kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_set_grand_total_caption(wb.handle, 0U, std::numeric_limits<std::size_t>::max(), "x"),
            kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_set_grand_total_caption(nullptr, 0U, kPivot, "x"), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_pivot_get_grand_total_caption(wb.handle, 0U, kPivot, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_pivot_get_grand_total_caption(wb.handle, 0U, 9U, &caption), kInvalidArgument);
  // The rejected calls left the accepted caption alone.
  ASSERT_EQ(fm_workbook_pivot_get_grand_total_caption(wb.handle, 0U, kPivot, &caption), 0);
  EXPECT_EQ(std::string(caption), at_limit);
}

TEST(FormulonCApiPivotOptions, GrandTotalCaptionMarksLayoutDirty) {
  ExpectMarksDirty([](fm_workbook_t* wb) { return fm_workbook_pivot_set_grand_total_caption(wb, 0U, kPivot, "All"); });
}

TEST(FormulonCApiPivotOptions, AutoFormattingSetsAndGets) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  std::int32_t value = -1;
  ASSERT_EQ(fm_workbook_pivot_set_auto_formatting(wb.handle, 0U, kPivot, 0), 0);
  ASSERT_EQ(fm_workbook_pivot_get_auto_formatting(wb.handle, 0U, kPivot, &value), 0);
  EXPECT_EQ(value, 0);
  ASSERT_EQ(fm_workbook_pivot_set_auto_formatting(wb.handle, 0U, kPivot, 7), 0);
  ASSERT_EQ(fm_workbook_pivot_get_auto_formatting(wb.handle, 0U, kPivot, &value), 0);
  EXPECT_EQ(value, 1);

  EXPECT_EQ(fm_workbook_pivot_set_auto_formatting(wb.handle, 1U, kPivot, 1), kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_set_auto_formatting(wb.handle, 0U, 1U, 1), kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_set_auto_formatting(nullptr, 0U, kPivot, 1), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_pivot_get_auto_formatting(wb.handle, 0U, kPivot, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_pivot_get_auto_formatting(wb.handle, 0U, 1U, &value), kInvalidArgument);
}

TEST(FormulonCApiPivotOptions, AutoFormattingMarksLayoutDirty) {
  ExpectMarksDirty([](fm_workbook_t* wb) { return fm_workbook_pivot_set_auto_formatting(wb, 0U, kPivot, 0); });
}

// ---------------------------------------------------------------------------
// Field flags
// ---------------------------------------------------------------------------

TEST(FormulonCApiPivotOptions, FieldFlagsSetAndReadBack) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  fm_pivot_field_info_t info{};
  ASSERT_EQ(fm_workbook_pivot_field_at(wb.handle, 0U, kPivot, kRegionField, &info), 0);
  EXPECT_EQ(info.default_subtotal, 1);
  EXPECT_EQ(info.insert_blank_row, 0);
  EXPECT_EQ(info.fill_down_labels, 0);

  ASSERT_EQ(fm_workbook_pivot_field_set_default_subtotal(wb.handle, 0U, kPivot, kRegionField, 0), 0);
  ASSERT_EQ(fm_workbook_pivot_field_set_insert_blank_row(wb.handle, 0U, kPivot, kRegionField, 5), 0);
  ASSERT_EQ(fm_workbook_pivot_field_set_fill_down_labels(wb.handle, 0U, kPivot, kRegionField, 1), 0);
  ASSERT_EQ(fm_workbook_pivot_field_at(wb.handle, 0U, kPivot, kRegionField, &info), 0);
  EXPECT_EQ(info.default_subtotal, 0);
  EXPECT_EQ(info.insert_blank_row, 1);
  EXPECT_EQ(info.fill_down_labels, 1);

  ASSERT_EQ(fm_workbook_pivot_field_set_default_subtotal(wb.handle, 0U, kPivot, kRegionField, 1), 0);
  ASSERT_EQ(fm_workbook_pivot_field_set_insert_blank_row(wb.handle, 0U, kPivot, kRegionField, 0), 0);
  ASSERT_EQ(fm_workbook_pivot_field_set_fill_down_labels(wb.handle, 0U, kPivot, kRegionField, 0), 0);
  ASSERT_EQ(fm_workbook_pivot_field_at(wb.handle, 0U, kPivot, kRegionField, &info), 0);
  EXPECT_EQ(info.default_subtotal, 1);
  EXPECT_EQ(info.insert_blank_row, 0);
  EXPECT_EQ(info.fill_down_labels, 0);
}

TEST(FormulonCApiPivotOptions, FieldFlagsRejectBadIndices) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  using Setter = fm_status_t (*)(fm_workbook_t*, std::size_t, std::size_t, std::size_t, std::int32_t);
  for (const Setter setter :
       {&fm_workbook_pivot_field_set_default_subtotal, &fm_workbook_pivot_field_set_insert_blank_row,
        &fm_workbook_pivot_field_set_fill_down_labels, &fm_workbook_pivot_field_set_show_detail}) {
    EXPECT_EQ(setter(wb.handle, 9U, kPivot, kRegionField, 1), kInvalidArgument);
    EXPECT_EQ(setter(wb.handle, 0U, 9U, kRegionField, 1), kInvalidArgument);
    EXPECT_EQ(setter(wb.handle, 0U, kPivot, 9U, 1), kInvalidArgument);
    EXPECT_EQ(setter(wb.handle, 0U, kPivot, std::numeric_limits<std::size_t>::max(), 1), kInvalidArgument);
    if constexpr (sizeof(std::size_t) > 4U) {
      EXPECT_EQ(setter(wb.handle, 0U, kPivot, static_cast<std::size_t>(1) << 32U, 1), kInvalidArgument);
    }
    EXPECT_EQ(setter(nullptr, 0U, kPivot, kRegionField, 1), kBindingNullPointer);
  }
}

TEST(FormulonCApiPivotOptions, FieldFlagsMarkLayoutDirty) {
  ExpectMarksDirty(
      [](fm_workbook_t* wb) { return fm_workbook_pivot_field_set_default_subtotal(wb, 0U, kPivot, kRegionField, 0); });
  ExpectMarksDirty(
      [](fm_workbook_t* wb) { return fm_workbook_pivot_field_set_insert_blank_row(wb, 0U, kPivot, kRegionField, 1); });
  ExpectMarksDirty(
      [](fm_workbook_t* wb) { return fm_workbook_pivot_field_set_fill_down_labels(wb, 0U, kPivot, kRegionField, 1); });
}

// ---------------------------------------------------------------------------
// Expansion
// ---------------------------------------------------------------------------

TEST(FormulonCApiPivotOptions, ShowDetailCollapsesFieldAndItems) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  fm_pivot_item_info_t item{};
  for (std::size_t i = 0; i < 2U; ++i) {
    ASSERT_EQ(fm_workbook_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, i, &item), 0);
    EXPECT_EQ(item.show_detail, 1);
  }

  ASSERT_EQ(fm_workbook_pivot_field_set_show_detail(wb.handle, 0U, kPivot, kRegionField, 0), 0);
  for (std::size_t i = 0; i < 2U; ++i) {
    ASSERT_EQ(fm_workbook_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, i, &item), 0);
    EXPECT_EQ(item.show_detail, 0);
  }
  for (const formulon::pivot::PivotMarkerItem& marker : Table(wb).fields()[kRegionField].marker_items) {
    if (marker.type == "default") {
      EXPECT_FALSE(marker.show_detail);
    }
  }

  ASSERT_EQ(fm_workbook_pivot_field_set_item_show_detail(wb.handle, 0U, kPivot, kRegionField, 1, 1), 0);
  ASSERT_EQ(fm_workbook_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, 0, &item), 0);
  EXPECT_EQ(item.show_detail, 0);
  ASSERT_EQ(fm_workbook_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, 1, &item), 0);
  EXPECT_EQ(item.show_detail, 1);

  ASSERT_EQ(fm_workbook_pivot_field_set_show_detail(wb.handle, 0U, kPivot, kRegionField, 1), 0);
  for (std::size_t i = 0; i < 2U; ++i) {
    ASSERT_EQ(fm_workbook_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, i, &item), 0);
    EXPECT_EQ(item.show_detail, 1);
  }
}

TEST(FormulonCApiPivotOptions, ItemShowDetailRejectsBadIndices) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  EXPECT_EQ(fm_workbook_pivot_field_set_item_show_detail(wb.handle, 9U, kPivot, kRegionField, 0, 0), kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_field_set_item_show_detail(wb.handle, 0U, 9U, kRegionField, 0, 0), kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_field_set_item_show_detail(wb.handle, 0U, kPivot, 9U, 0, 0), kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_field_set_item_show_detail(wb.handle, 0U, kPivot, kRegionField, 2U, 0), kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_field_set_item_show_detail(wb.handle, 0U, kPivot, kRegionField,
                                                         std::numeric_limits<std::size_t>::max(), 0),
            kInvalidArgument);
  if constexpr (sizeof(std::size_t) > 4U) {
    EXPECT_EQ(fm_workbook_pivot_field_set_item_show_detail(wb.handle, 0U, kPivot, kRegionField,
                                                           static_cast<std::size_t>(1) << 32U, 0),
              kInvalidArgument);
  }
  EXPECT_EQ(fm_workbook_pivot_field_set_item_show_detail(nullptr, 0U, kPivot, kRegionField, 0, 0), kBindingNullPointer);
}

TEST(FormulonCApiPivotOptions, ShowDetailSettersMarkLayoutDirty) {
  ExpectMarksDirty(
      [](fm_workbook_t* wb) { return fm_workbook_pivot_field_set_show_detail(wb, 0U, kPivot, kRegionField, 0); });
  ExpectMarksDirty([](fm_workbook_t* wb) {
    return fm_workbook_pivot_field_set_item_show_detail(wb, 0U, kPivot, kRegionField, 0, 0);
  });
}

// ---------------------------------------------------------------------------
// Item captions
// ---------------------------------------------------------------------------

TEST(FormulonCApiPivotOptions, ItemCaptionSetsKeepsValueAndClears) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  fm_pivot_item_info_t item{};
  ASSERT_EQ(fm_workbook_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, 0, &item), 0);
  EXPECT_EQ(item.caption, nullptr);

  ASSERT_EQ(fm_workbook_pivot_field_set_item_caption(wb.handle, 0U, kPivot, kRegionField, 0, "Nord \xE5\x8C\x97"), 0)
      << fm_last_error_message();
  ASSERT_EQ(fm_workbook_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, 0, &item), 0);
  EXPECT_STREQ(item.caption, "Nord \xE5\x8C\x97");
  EXPECT_STREQ(item.name, "North");
  EXPECT_TRUE(LayoutHasText(wb.handle, kPivot, "Nord \xE5\x8C\x97"));
  EXPECT_NE(SavedPivotXml(wb.handle).find("n=\"Nord \xE5\x8C\x97\""), std::string::npos);

  // NULL and empty clear.
  ASSERT_EQ(fm_workbook_pivot_field_set_item_caption(wb.handle, 0U, kPivot, kRegionField, 0, nullptr), 0);
  ASSERT_EQ(fm_workbook_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, 0, &item), 0);
  EXPECT_EQ(item.caption, nullptr);
  ASSERT_EQ(fm_workbook_pivot_field_set_item_caption(wb.handle, 0U, kPivot, kRegionField, 0, "Nord"), 0);
  ASSERT_EQ(fm_workbook_pivot_field_set_item_caption(wb.handle, 0U, kPivot, kRegionField, 0, ""), 0);
  ASSERT_EQ(fm_workbook_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, 0, &item), 0);
  EXPECT_EQ(item.caption, nullptr);

  // A caption equal to the item's own value clears, whatever its case.
  ASSERT_EQ(fm_workbook_pivot_field_set_item_caption(wb.handle, 0U, kPivot, kRegionField, 0, "Nord"), 0);
  ASSERT_EQ(fm_workbook_pivot_field_set_item_caption(wb.handle, 0U, kPivot, kRegionField, 0, "NORTH"), 0);
  ASSERT_EQ(fm_workbook_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, 0, &item), 0);
  EXPECT_EQ(item.caption, nullptr);
}

TEST(FormulonCApiPivotOptions, ItemCaptionRejectsDuplicatesOfShownLabels) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  // Another item's value, in any case.
  EXPECT_EQ(fm_workbook_pivot_field_set_item_caption(wb.handle, 0U, kPivot, kRegionField, 0, "South"),
            kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_field_set_item_caption(wb.handle, 0U, kPivot, kRegionField, 0, "south"),
            kInvalidArgument);
  // Another item's caption.
  ASSERT_EQ(fm_workbook_pivot_field_set_item_caption(wb.handle, 0U, kPivot, kRegionField, 0, "Nord"), 0);
  EXPECT_EQ(fm_workbook_pivot_field_set_item_caption(wb.handle, 0U, kPivot, kRegionField, 1, "NORD"), kInvalidArgument);
  // The value a caption replaced is free to use again.
  EXPECT_EQ(fm_workbook_pivot_field_set_item_caption(wb.handle, 0U, kPivot, kRegionField, 1, "North"), 0);
  // The same caption on the same item is not a duplicate.
  EXPECT_EQ(fm_workbook_pivot_field_set_item_caption(wb.handle, 0U, kPivot, kRegionField, 0, "Nord"), 0);

  fm_pivot_item_info_t item{};
  ASSERT_EQ(fm_workbook_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, 1, &item), 0);
  EXPECT_STREQ(item.caption, "North");
}

TEST(FormulonCApiPivotOptions, ItemCaptionBoundaries) {
  WorkbookGuard wb;
  LoadScratchPivot(wb);
  const std::string at_limit = Repeat(kEmoji, 127) + "a";
  ASSERT_EQ(fm_workbook_pivot_field_set_item_caption(wb.handle, 0U, kPivot, kRegionField, 0, at_limit.c_str()), 0)
      << fm_last_error_message();
  fm_pivot_item_info_t item{};
  ASSERT_EQ(fm_workbook_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, 0, &item), 0);
  EXPECT_EQ(std::string(item.caption), at_limit);

  EXPECT_EQ(
      fm_workbook_pivot_field_set_item_caption(wb.handle, 0U, kPivot, kRegionField, 0, Repeat(kEmoji, 128).c_str()),
      kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_field_set_item_caption(wb.handle, 0U, kPivot, kRegionField, 0, Repeat("a", 256).c_str()),
            kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_field_set_item_caption(wb.handle, 0U, kPivot, kRegionField, 0, "\xC3"),
            kBindingUtf8EncodingError);
  EXPECT_EQ(fm_workbook_pivot_field_set_item_caption(wb.handle, 9U, kPivot, kRegionField, 0, "x"), kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_field_set_item_caption(wb.handle, 0U, 9U, kRegionField, 0, "x"), kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_field_set_item_caption(wb.handle, 0U, kPivot, 9U, 0, "x"), kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_field_set_item_caption(wb.handle, 0U, kPivot, kRegionField, 2U, "x"), kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_field_set_item_caption(wb.handle, 0U, kPivot, kRegionField,
                                                     std::numeric_limits<std::size_t>::max(), "x"),
            kInvalidArgument);
  if constexpr (sizeof(std::size_t) > 4U) {
    EXPECT_EQ(fm_workbook_pivot_field_set_item_caption(wb.handle, 0U, kPivot, kRegionField,
                                                       static_cast<std::size_t>(1) << 32U, "x"),
              kInvalidArgument);
  }
  EXPECT_EQ(fm_workbook_pivot_field_set_item_caption(nullptr, 0U, kPivot, kRegionField, 0, "x"), kBindingNullPointer);
  // The rejected calls left the accepted caption alone.
  ASSERT_EQ(fm_workbook_pivot_field_item_at(wb.handle, 0U, kPivot, kRegionField, 0, &item), 0);
  EXPECT_EQ(std::string(item.caption), at_limit);
}

TEST(FormulonCApiPivotOptions, ItemCaptionMarksLayoutDirty) {
  ExpectMarksDirty([](fm_workbook_t* wb) {
    return fm_workbook_pivot_field_set_item_caption(wb, 0U, kPivot, kRegionField, 0, "Nord");
  });
}

// ---------------------------------------------------------------------------
// The Values level in the field orders
// ---------------------------------------------------------------------------

TEST(FormulonCApiPivotOptions, ValuesFieldConstantMatchesTheOoxmlMarker) {
  EXPECT_EQ(FM_PIVOT_VALUES_FIELD, static_cast<std::uint32_t>(-2));
}

TEST(FormulonCApiPivotOptions, ValuesFieldNeedsTwoDataFields) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, nullptr), 0) << fm_last_error_message();
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
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, nullptr), 0) << fm_last_error_message();
  AddSecondDataField(wb.handle);
  EXPECT_EQ(ReadOrder(wb.handle, true), (std::vector<std::uint32_t>{0U}));
  EXPECT_EQ(ReadOrder(wb.handle, false), (std::vector<std::uint32_t>{FM_PIVOT_VALUES_FIELD}));
}

TEST(FormulonCApiPivotOptions, ValuesFieldMovesBetweenAxesAndPositions) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, nullptr), 0) << fm_last_error_message();
  AddSecondDataField(wb.handle);

  ASSERT_EQ(SetOrder(wb.handle, true, {0U, FM_PIVOT_VALUES_FIELD}), 0) << fm_last_error_message();
  EXPECT_EQ(ReadOrder(wb.handle, true), (std::vector<std::uint32_t>{0U, FM_PIVOT_VALUES_FIELD}));
  EXPECT_TRUE(ReadOrder(wb.handle, false).empty());
  EXPECT_EQ(Table(wb).row_values_position(), 1U);
  EXPECT_FALSE(Table(wb).col_values_position().has_value());

  ASSERT_EQ(SetOrder(wb.handle, true, {FM_PIVOT_VALUES_FIELD, 0U}), 0) << fm_last_error_message();
  EXPECT_EQ(ReadOrder(wb.handle, true), (std::vector<std::uint32_t>{FM_PIVOT_VALUES_FIELD, 0U}));
  EXPECT_EQ(Table(wb).row_values_position(), 0U);

  // The getter's own listing is a valid input.
  const std::uint32_t* listed = nullptr;
  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_pivot_get_row_field_order(wb.handle, 0U, kPivot, &listed, &count), 0);
  ASSERT_EQ(fm_workbook_pivot_set_row_field_order(wb.handle, 0U, kPivot, listed, count), 0) << fm_last_error_message();
  EXPECT_EQ(ReadOrder(wb.handle, true), (std::vector<std::uint32_t>{FM_PIVOT_VALUES_FIELD, 0U}));

  // Back to the columns: the rows give it up, and the default layout returns.
  ASSERT_EQ(SetOrder(wb.handle, false, {FM_PIVOT_VALUES_FIELD}), 0) << fm_last_error_message();
  EXPECT_EQ(ReadOrder(wb.handle, false), (std::vector<std::uint32_t>{FM_PIVOT_VALUES_FIELD}));
  EXPECT_EQ(ReadOrder(wb.handle, true), (std::vector<std::uint32_t>{0U}));
  EXPECT_FALSE(Table(wb).row_values_position().has_value());
  EXPECT_EQ(Table(wb).col_values_position(), 0U);
}

TEST(FormulonCApiPivotOptions, ValuesFieldMustStayListedOnItsAxis) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, nullptr), 0) << fm_last_error_message();
  AddSecondDataField(wb.handle);

  // The columns hold it by default.
  EXPECT_EQ(SetOrder(wb.handle, false, {}), kInvalidArgument);
  EXPECT_EQ(ReadOrder(wb.handle, false), (std::vector<std::uint32_t>{FM_PIVOT_VALUES_FIELD}));

  ASSERT_EQ(SetOrder(wb.handle, true, {0U, FM_PIVOT_VALUES_FIELD}), 0);
  EXPECT_EQ(SetOrder(wb.handle, true, {0U}), kInvalidArgument);
  EXPECT_EQ(SetOrder(wb.handle, true, {}), kInvalidArgument);
  EXPECT_EQ(ReadOrder(wb.handle, true), (std::vector<std::uint32_t>{0U, FM_PIVOT_VALUES_FIELD}));
  // The other axis may be rewritten without it.
  EXPECT_EQ(SetOrder(wb.handle, false, {}), 0) << fm_last_error_message();
  EXPECT_EQ(ReadOrder(wb.handle, true), (std::vector<std::uint32_t>{0U, FM_PIVOT_VALUES_FIELD}));
}

TEST(FormulonCApiPivotOptions, ValuesFieldRejectsDuplicatesAndStrayIndices) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, nullptr), 0) << fm_last_error_message();
  AddSecondDataField(wb.handle);

  EXPECT_EQ(SetOrder(wb.handle, true, {FM_PIVOT_VALUES_FIELD, 0U, FM_PIVOT_VALUES_FIELD}), kInvalidArgument);
  EXPECT_EQ(SetOrder(wb.handle, false, {FM_PIVOT_VALUES_FIELD, FM_PIVOT_VALUES_FIELD}), kInvalidArgument);
  EXPECT_EQ(SetOrder(wb.handle, true, {0U, FM_PIVOT_VALUES_FIELD - 1U}), kInvalidArgument);
  EXPECT_EQ(SetOrder(wb.handle, true, {0U, 0xFFFFFFFFU}), kInvalidArgument);
  EXPECT_EQ(ReadOrder(wb.handle, true), (std::vector<std::uint32_t>{0U}));
  EXPECT_EQ(ReadOrder(wb.handle, false), (std::vector<std::uint32_t>{FM_PIVOT_VALUES_FIELD}));
}

TEST(FormulonCApiPivotOptions, ValuesFieldOrderMarksLayoutDirty) {
  WorkbookGuard built;
  ASSERT_EQ(fm_workbook_create(&built.handle), 0);
  ASSERT_EQ(BuildScratchPivot(built.handle, nullptr, nullptr), 0);
  AddSecondDataField(built.handle);
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(built.handle, &saved.data, &saved.len), 0) << fm_last_error_message();
  WorkbookGuard clean;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &clean.handle), 0) << fm_last_error_message();
  ASSERT_FALSE(Table(clean).layout_dirty());
  ASSERT_EQ(SetOrder(clean.handle, true, {0U, FM_PIVOT_VALUES_FIELD}), 0) << fm_last_error_message();
  EXPECT_TRUE(Table(clean).layout_dirty());
}

// ---------------------------------------------------------------------------
// Show Details
// ---------------------------------------------------------------------------

namespace {

// First projected cell of `kind`, as absolute coordinates.
bool FindCell(fm_workbook_t* wb, fm_pivot_cell_kind_t kind, std::uint32_t* row, std::uint32_t* col) {
  PivotCellsGuard projected;
  EXPECT_EQ(fm_workbook_pivot_layout(wb, 0U, kPivot, &projected.handle), 0) << fm_last_error_message();
  for (const fm_pivot_cell_t& cell : CollectCells(projected.handle)) {
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
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, nullptr), 0) << fm_last_error_message();
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  ASSERT_TRUE(FindCell(wb.handle, FM_PIVOT_CELL_DATA, &row, &col));

  const std::size_t before = fm_workbook_sheet_count(wb.handle);
  std::size_t created = 99;
  ASSERT_EQ(fm_workbook_pivot_show_details(wb.handle, 0U, kPivot, row, col, &created), 0) << fm_last_error_message();
  EXPECT_EQ(created, 0U);
  EXPECT_EQ(fm_workbook_sheet_count(wb.handle), before + 1U);
  std::uint32_t active = 99;
  ASSERT_EQ(fm_workbook_get_active_sheet(wb.handle, &active), 0);
  EXPECT_EQ(active, 0U);
  // The pivot moved with its sheet.
  EXPECT_EQ(wb.handle->wb->sheet(1U).pivot_tables().size(), 1U);
}

TEST(FormulonCApiPivotOptions, ShowDetailsRejectsBadCellsAndArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, nullptr), 0) << fm_last_error_message();
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  ASSERT_TRUE(FindCell(wb.handle, FM_PIVOT_CELL_DATA, &row, &col));
  std::uint32_t label_row = 0;
  std::uint32_t label_col = 0;
  ASSERT_TRUE(FindCell(wb.handle, FM_PIVOT_CELL_ROW_LABEL, &label_row, &label_col));

  std::size_t created = 99;
  EXPECT_EQ(fm_workbook_pivot_show_details(wb.handle, 0U, kPivot, label_row, label_col, &created), kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_show_details(wb.handle, 0U, kPivot, 5000U, 5000U, &created), kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_show_details(wb.handle, 0U, kPivot, 0xFFFFFFFFU, 0xFFFFFFFFU, &created),
            kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_show_details(wb.handle, 9U, kPivot, row, col, &created), kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_show_details(wb.handle, 0U, 9U, row, col, &created), kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_show_details(wb.handle, 0U, kPivot, row, col, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_pivot_show_details(nullptr, 0U, kPivot, row, col, &created), kBindingNullPointer);
  EXPECT_EQ(created, 99U);
  EXPECT_EQ(fm_workbook_sheet_count(wb.handle), 1U);

  Table(wb).set_enable_drill(false);
  EXPECT_EQ(fm_workbook_pivot_show_details(wb.handle, 0U, kPivot, row, col, &created), kPreconditionFailed);
}

// ---------------------------------------------------------------------------
// Cache edits drop the encoding flags of the cells they overwrite
// ---------------------------------------------------------------------------

TEST(FormulonCApiPivotOptions, RecordEditsResetDateAndIndexFlags) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, nullptr), 0) << fm_last_error_message();
  formulon::pivot::PivotCache* cache = wb.handle->wb->mutable_pivot_caches().front().get();
  formulon::pivot::PivotCacheRecord& record = cache->mutable_records()[0];
  record.cell_is_index.assign(2U, false);
  record.cell_is_index[0] = true;
  record.cell_is_date.assign(2U, false);
  record.cell_is_date[1] = true;

  // A plain number over a date cell is an inline number again.
  ASSERT_EQ(fm_workbook_pivot_cache_record_set_number(wb.handle, cache_id, 0U, 1U, 5.0), 0) << fm_last_error_message();
  EXPECT_FALSE(record.cell_is_date[1]);
  EXPECT_FALSE(record.cell_is_index[1]);
  const std::string records = ExtractZipEntry(
      [&] {
        BufferGuard saved;
        EXPECT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_last_error_message();
        return std::vector<std::uint8_t>(saved.data, saved.data + saved.len);
      }(),
      "xl/pivotCache/pivotCacheRecords1.xml");
  EXPECT_NE(records.find("<n v=\"5\"/>"), std::string::npos);
  EXPECT_EQ(records.find("<d "), std::string::npos);

  // Text and blank over an index cell are no longer an index.
  ASSERT_EQ(fm_workbook_pivot_cache_record_set_text(wb.handle, cache_id, 0U, 0U, "Elsewhere"), 0);
  EXPECT_FALSE(record.cell_is_index[0]);
  record.cell_is_index[0] = true;
  record.cell_is_date[0] = true;
  ASSERT_EQ(fm_workbook_pivot_cache_record_set_blank(wb.handle, cache_id, 0U, 0U), 0);
  EXPECT_FALSE(record.cell_is_index[0]);
  EXPECT_FALSE(record.cell_is_date[0]);

  // Clearing the shared items forgets which of them were dates.
  cache->mutable_fields()[0].shared_item_is_date.assign(2U, true);
  ASSERT_EQ(fm_workbook_pivot_cache_field_clear_shared_items(wb.handle, cache_id, 0U), 0);
  EXPECT_TRUE(cache->fields()[0].shared_item_is_date.empty());
}
