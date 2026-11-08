#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <utility>

#include "eval/eval_profile_scope.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "pivot/field_patch.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

using pivot::FieldSortOrder;
using pivot::PivotFieldPatch;
using pivot::PivotItemPatch;
using pivot::PivotTable;

constexpr std::size_t kRegion = 0U;
constexpr std::size_t kProduct = 1U;
constexpr std::size_t kAmt = 2U;

class PivotFieldPatchTest : public ::testing::Test {
 protected:
  PivotFieldPatchTest() : wb_(Workbook::create_for_profile(mac_365_ja_jp_profile())) {
    auto cache = std::make_unique<pivot::PivotCache>();
    cache->set_cache_id(1U);
    cache->mutable_fields().push_back(
        pivot::PivotCacheField{"Region", {Value::text("North"), Value::text("South"), Value::text("East")}});
    cache->mutable_fields().push_back(pivot::PivotCacheField{"Product", {Value::text("A")}});
    cache->mutable_fields().push_back(pivot::PivotCacheField{"Amt", {}});
    wb_.add_pivot_cache(std::move(cache));

    table_.set_pivot_cache_id(1U);
    for (const char* name : {"Region", "Product", "Amt"}) {
      pivot::PivotField field;
      field.source_name = name;
      field.axis = pivot::PivotAxis::None;
      table_.mutable_fields().push_back(std::move(field));
    }
    table_.mutable_fields()[kRegion].axis = pivot::PivotAxis::Row;
    for (const std::uint32_t x : {0U, 1U, 2U}) {
      pivot::PivotItem item;
      item.has_cache_index = true;
      item.cache_index = x;
      table_.mutable_fields()[kRegion].items.push_back(item);
    }
    pivot::PivotMarkerItem marker;
    marker.type = "default";
    marker.after_items = 3U;
    table_.mutable_fields()[kRegion].marker_items.push_back(marker);
    table_.mutable_fields()[kAmt].axis = pivot::PivotAxis::Value;
    pivot::PivotDataField sum;
    sum.name = "Sum of Amt";
    sum.field_index = static_cast<std::uint32_t>(kAmt);
    table_.mutable_data_fields().push_back(std::move(sum));
  }

  const pivot::PivotField& Field(std::size_t i) const { return table_.fields()[i]; }

  template <typename E>
  static void ExpectInvalid(const E& got) {
    ASSERT_FALSE(static_cast<bool>(got));
    EXPECT_EQ(got.error().code, FormulonErrorCode::kInvalidArgument);
  }

  const eval::ScopedEvalProfile scope_{mac_365_ja_jp_profile()};
  Workbook wb_;
  PivotTable table_;
};

TEST_F(PivotFieldPatchTest, RefusedFieldPatchWritesNoKey) {
  PivotFieldPatch taken;
  taken.show_all = true;
  taken.insert_blank_row = true;
  taken.custom_name = "Product";
  ExpectInvalid(pivot::apply_field_patch(wb_, table_, kRegion, taken));
  EXPECT_FALSE(Field(kRegion).show_all);
  EXPECT_FALSE(Field(kRegion).insert_blank_row);
  EXPECT_TRUE(Field(kRegion).custom_name.empty());

  taken.custom_name = "sum of amt";
  ExpectInvalid(pivot::apply_field_patch(wb_, table_, kRegion, taken));
  EXPECT_FALSE(Field(kRegion).show_all);

  // The field's own source name is not taken.
  taken.custom_name = "Region";
  ASSERT_TRUE(static_cast<bool>(pivot::apply_field_patch(wb_, table_, kRegion, taken)));
  EXPECT_TRUE(Field(kRegion).show_all);
  EXPECT_TRUE(Field(kRegion).insert_blank_row);
  EXPECT_EQ(Field(kRegion).custom_name, "Region");

  taken.custom_name = std::string(256, 'n');
  ExpectInvalid(pivot::check_field_patch(wb_, table_, kRegion, taken));
  ExpectInvalid(pivot::check_field_patch(wb_, table_, 3U, PivotFieldPatch{}));
}

TEST_F(PivotFieldPatchTest, EmptyFieldPatchChangesNothing) {
  const pivot::PivotField before = Field(kRegion);
  ASSERT_TRUE(static_cast<bool>(pivot::apply_field_patch(wb_, table_, kRegion, PivotFieldPatch{})));
  const pivot::PivotField& after = Field(kRegion);
  EXPECT_EQ(after.custom_name, before.custom_name);
  EXPECT_EQ(after.number_format, before.number_format);
  EXPECT_EQ(after.default_subtotal, before.default_subtotal);
  EXPECT_EQ(after.subtotal_top, before.subtotal_top);
  EXPECT_EQ(after.sort.stated, before.sort.stated);
  EXPECT_EQ(after.sort.manual, before.sort.manual);
}

TEST_F(PivotFieldPatchTest, SortOrderAndByFieldAreCheckedTogether) {
  PivotFieldPatch manual_by;
  manual_by.sort_order = FieldSortOrder::kManual;
  manual_by.sort_by_field = "Sum of Amt";
  ExpectInvalid(pivot::apply_field_patch(wb_, table_, kRegion, manual_by));
  EXPECT_FALSE(Field(kRegion).sort.stated);

  PivotFieldPatch descending;
  descending.sort_order = FieldSortOrder::kDescending;
  descending.sort_by_field = "Sum of Amt";
  ASSERT_TRUE(static_cast<bool>(pivot::apply_field_patch(wb_, table_, kRegion, descending)));
  EXPECT_FALSE(Field(kRegion).sort.manual);
  EXPECT_FALSE(Field(kRegion).sort.ascending);
  EXPECT_TRUE(Field(kRegion).sort.stated);
  EXPECT_EQ(Field(kRegion).sort.by_field, "Sum of Amt");

  // Manual alone is refused while the field keeps its by-field; clearing the by-field with it passes.
  PivotFieldPatch manual;
  manual.sort_order = FieldSortOrder::kManual;
  ExpectInvalid(pivot::apply_field_patch(wb_, table_, kRegion, manual));
  manual.sort_by_field = std::string();
  ASSERT_TRUE(static_cast<bool>(pivot::apply_field_patch(wb_, table_, kRegion, manual)));
  EXPECT_TRUE(Field(kRegion).sort.manual);
  EXPECT_TRUE(Field(kRegion).sort.by_field.empty());
}

TEST_F(PivotFieldPatchTest, ShowDetailCoversEveryItemAndTheDefaultMarker) {
  EXPECT_TRUE(pivot::field_show_detail(Field(kRegion)));
  EXPECT_TRUE(pivot::field_show_detail(Field(kProduct)));
  PivotFieldPatch collapse;
  collapse.show_detail = false;
  ASSERT_TRUE(static_cast<bool>(pivot::apply_field_patch(wb_, table_, kRegion, collapse)));
  for (const pivot::PivotItem& item : Field(kRegion).items) {
    EXPECT_FALSE(item.show_detail);
  }
  EXPECT_FALSE(Field(kRegion).marker_items[0].show_detail);
  EXPECT_FALSE(pivot::field_show_detail(Field(kRegion)));

  // One expanded item leaves the field mixed, which reads as collapsed.
  PivotItemPatch expand;
  expand.show_detail = true;
  ASSERT_TRUE(static_cast<bool>(pivot::apply_item_patch(wb_, table_, kRegion, 1U, expand)));
  EXPECT_FALSE(pivot::field_show_detail(Field(kRegion)));
  collapse.show_detail = true;
  ASSERT_TRUE(static_cast<bool>(pivot::apply_field_patch(wb_, table_, kRegion, collapse)));
  EXPECT_TRUE(pivot::field_show_detail(Field(kRegion)));
}

TEST_F(PivotFieldPatchTest, RefusedItemPatchWritesNoKey) {
  PivotItemPatch p;
  p.visible = false;
  p.caption = "south";
  ExpectInvalid(pivot::apply_item_patch(wb_, table_, kRegion, 0U, p));
  EXPECT_TRUE(Field(kRegion).items[0].visible);
  EXPECT_TRUE(Field(kRegion).items[0].caption.empty());

  p.caption = "Up";
  ASSERT_TRUE(static_cast<bool>(pivot::apply_item_patch(wb_, table_, kRegion, 0U, p)));
  EXPECT_FALSE(Field(kRegion).items[0].visible);
  EXPECT_EQ(Field(kRegion).items[0].caption, "Up");

  // Another item may not show a caption already shown; the item's own label removes its caption.
  PivotItemPatch clash;
  clash.caption = "up";
  ExpectInvalid(pivot::apply_item_patch(wb_, table_, kRegion, 1U, clash));
  PivotItemPatch own;
  own.caption = "North";
  ASSERT_TRUE(static_cast<bool>(pivot::apply_item_patch(wb_, table_, kRegion, 0U, own)));
  EXPECT_TRUE(Field(kRegion).items[0].caption.empty());

  ExpectInvalid(pivot::check_item_patch(wb_, table_, kRegion, 3U, PivotItemPatch{}));
  ExpectInvalid(pivot::check_item_patch(wb_, table_, 9U, 0U, PivotItemPatch{}));
}

}  // namespace
}  // namespace formulon
