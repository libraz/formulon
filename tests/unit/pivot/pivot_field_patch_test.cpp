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

  template <typename E>
  static void ExpectTaken(const E& got) {
    ASSERT_FALSE(static_cast<bool>(got));
    EXPECT_EQ(got.error().code, FormulonErrorCode::kAlreadyExists);
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
  ExpectTaken(pivot::apply_field_patch(wb_, table_, kRegion, taken));
  EXPECT_FALSE(Field(kRegion).show_all);
  EXPECT_FALSE(Field(kRegion).insert_blank_row);
  EXPECT_TRUE(Field(kRegion).custom_name.empty());

  taken.custom_name = "sum of amt";
  ExpectTaken(pivot::apply_field_patch(wb_, table_, kRegion, taken));
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
  ExpectTaken(pivot::apply_item_patch(wb_, table_, kRegion, 0U, p));
  EXPECT_TRUE(Field(kRegion).items[0].visible);
  EXPECT_TRUE(Field(kRegion).items[0].caption.empty());

  p.caption = "Up";
  ASSERT_TRUE(static_cast<bool>(pivot::apply_item_patch(wb_, table_, kRegion, 0U, p)));
  EXPECT_FALSE(Field(kRegion).items[0].visible);
  EXPECT_EQ(Field(kRegion).items[0].caption, "Up");

  // Another item may not show a caption already shown; the item's own label removes its caption.
  PivotItemPatch clash;
  clash.caption = "up";
  ExpectTaken(pivot::apply_item_patch(wb_, table_, kRegion, 1U, clash));
  PivotItemPatch own;
  own.caption = "North";
  ASSERT_TRUE(static_cast<bool>(pivot::apply_item_patch(wb_, table_, kRegion, 0U, own)));
  EXPECT_TRUE(Field(kRegion).items[0].caption.empty());

  ExpectInvalid(pivot::check_item_patch(wb_, table_, kRegion, 3U, PivotItemPatch{}));
  ExpectInvalid(pivot::check_item_patch(wb_, table_, 9U, 0U, PivotItemPatch{}));
}

// A workbook with cache 1 over Region / Product / Amt and no pivots yet.
class PivotAuthoringTest : public ::testing::Test {
 protected:
  PivotAuthoringTest() : wb_(Workbook::create_for_profile(mac_365_ja_jp_profile())) {
    auto cache = std::make_unique<pivot::PivotCache>();
    cache->set_cache_id(1U);
    cache->mutable_fields().push_back(pivot::PivotCacheField{"Region", {Value::text("North")}});
    cache->mutable_fields().push_back(pivot::PivotCacheField{"Product", {Value::text("A")}});
    cache->mutable_fields().push_back(pivot::PivotCacheField{"Amt", {}});
    wb_.add_pivot_cache(std::move(cache));
  }

  static std::unique_ptr<PivotTable> NewPivot(const char* name, std::uint32_t row) {
    auto table = std::make_unique<PivotTable>();
    table->set_name(name);
    table->set_pivot_cache_id(1U);
    table->set_anchor(row, 5U, 1U, 1U);
    return table;
  }

  const PivotTable& Pivot(std::size_t i) const { return *wb_.sheet(0).pivot_tables()[i]; }

  const eval::ScopedEvalProfile scope_{mac_365_ja_jp_profile()};
  Workbook wb_;
};

using PivotFieldsFromCache = PivotAuthoringTest;

TEST_F(PivotFieldsFromCache, NewPivotHasOneFieldPerCacheField) {
  ASSERT_TRUE(static_cast<bool>(wb_.add_pivot(0, NewPivot("PivotTable1", 0U))));
  const PivotTable& table = Pivot(0);
  ASSERT_EQ(table.fields().size(), wb_.find_pivot_cache(1U)->fields().size());
  EXPECT_EQ(table.fields()[1].source_name, "Product");
  EXPECT_EQ(table.fields()[1].axis, pivot::PivotAxis::None);

  auto unbound = NewPivot("PivotTable2", 20U);
  unbound->set_pivot_cache_id(9U);
  auto refused = wb_.add_pivot(0, std::move(unbound));
  ASSERT_FALSE(static_cast<bool>(refused));
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kNotFound);
}

TEST_F(PivotFieldsFromCache, CacheAddFieldAppendsToPivots) {
  ASSERT_TRUE(static_cast<bool>(wb_.add_pivot(0, NewPivot("PivotTable1", 0U))));
  ASSERT_TRUE(static_cast<bool>(wb_.add_pivot(0, NewPivot("PivotTable2", 20U))));
  auto index = wb_.add_pivot_cache_field(1U, pivot::PivotCacheField{"Qty", {}});
  ASSERT_TRUE(static_cast<bool>(index));
  EXPECT_EQ(index.value(), 3U);
  for (std::size_t p = 0; p < 2U; ++p) {
    ASSERT_EQ(Pivot(p).fields().size(), 4U) << p;
    EXPECT_EQ(Pivot(p).fields()[3].source_name, "Qty");
    EXPECT_EQ(Pivot(p).fields()[3].axis, pivot::PivotAxis::None);
  }
  auto missing = wb_.add_pivot_cache_field(9U, pivot::PivotCacheField{"Qty", {}});
  ASSERT_FALSE(static_cast<bool>(missing));
  EXPECT_EQ(missing.error().code, FormulonErrorCode::kNotFound);
}

using PivotRenameValidation = PivotAuthoringTest;

TEST_F(PivotRenameValidation, EmptyNameInvalidArgument) {
  ASSERT_TRUE(static_cast<bool>(wb_.add_pivot(0, NewPivot("PivotTable1", 0U))));
  auto empty = wb_.rename_pivot(0, 0, "");
  ASSERT_FALSE(static_cast<bool>(empty));
  EXPECT_EQ(empty.error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(Pivot(0).name(), "PivotTable1");
  ASSERT_TRUE(static_cast<bool>(wb_.rename_pivot(0, 0, "Sales")));
  EXPECT_EQ(Pivot(0).name(), "Sales");

  auto unnamed = wb_.add_pivot(0, NewPivot("", 20U));
  ASSERT_FALSE(static_cast<bool>(unnamed));
  EXPECT_EQ(unnamed.error().code, FormulonErrorCode::kInvalidArgument);
}

TEST_F(PivotRenameValidation, DuplicateNameAlreadyExists) {
  ASSERT_TRUE(static_cast<bool>(wb_.add_pivot(0, NewPivot("PivotTable1", 0U))));
  ASSERT_TRUE(static_cast<bool>(wb_.add_pivot(0, NewPivot("PivotTable2", 20U))));
  auto taken = wb_.rename_pivot(0, 1, "pivottable1");
  ASSERT_FALSE(static_cast<bool>(taken));
  EXPECT_EQ(taken.error().code, FormulonErrorCode::kAlreadyExists);
  EXPECT_EQ(Pivot(1).name(), "PivotTable2");
  // A pivot may take its own name in another letter case.
  ASSERT_TRUE(static_cast<bool>(wb_.rename_pivot(0, 1, "PIVOTTABLE2")));
  EXPECT_EQ(Pivot(1).name(), "PIVOTTABLE2");

  auto duplicate = wb_.add_pivot(0, NewPivot("PIVOTTABLE1", 40U));
  ASSERT_FALSE(static_cast<bool>(duplicate));
  EXPECT_EQ(duplicate.error().code, FormulonErrorCode::kAlreadyExists);
}

}  // namespace
}  // namespace formulon
