// C ABI argument validation: iterative max_change, pivot cache numbers and
// the out-pointer contract of fm_workbook_create_for_profile.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>

#include "c_api/formulon_c.h"
#include "formulon_c_out_index.h"
#include "formulon_c_setters.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr double kNaN = std::numeric_limits<double>::quiet_NaN();
constexpr double kInf = std::numeric_limits<double>::infinity();

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

fm_iterative_settings_info_t ReadIterative(const fm_workbook_t* wb) {
  fm_iterative_settings_info_t info = FM_STRUCT_INIT(fm_iterative_settings_info_t);
  EXPECT_EQ(fm_workbook_get_iterative(wb, &info), 0);
  return info;
}

void ExpectIterative(const fm_workbook_t* wb, int32_t enabled, uint32_t max_iterations, double max_change) {
  const fm_iterative_settings_info_t info = ReadIterative(wb);
  EXPECT_EQ(info.enabled, enabled);
  EXPECT_EQ(info.max_iterations, max_iterations);
  EXPECT_DOUBLE_EQ(info.max_change, max_change);
}

TEST(CApiSetIterative, RejectsNonFiniteMaxChangeAndKeepsOptions) {
  WorkbookGuard g;
  ASSERT_EQ(fm_workbook_create_default(&g.handle), 0);
  ASSERT_EQ(SetIterative(g.handle, 1, 50, 0.25), 0);
  for (const double bad : {kNaN, kInf, -kInf}) {
    EXPECT_EQ(SetIterative(g.handle, 0, 7, bad), kInvalidArgument);
  }
  ExpectIterative(g.handle, 1, 50U, 0.25);
}

TEST(CApiSetIterative, AcceptsZeroMaxChange) {
  WorkbookGuard g;
  ASSERT_EQ(fm_workbook_create_default(&g.handle), 0);
  EXPECT_EQ(SetIterative(g.handle, 1, 10, 0.0), 0);
}

// A KEEP field leaves that setting, so each one can change alone and an
// all-KEEP patch changes nothing.
TEST(CApiSetIterative, KeepFieldsLeaveTheStoredSettings) {
  WorkbookGuard g;
  ASSERT_EQ(fm_workbook_create_default(&g.handle), 0);
  ASSERT_EQ(SetIterative(g.handle, 1, 50, 0.25), 0);

  const fm_iterative_settings_patch_t keep = FM_STRUCT_INIT(fm_iterative_settings_patch_t);
  ASSERT_EQ(fm_workbook_update_iterative(g.handle, &keep), 0);
  ExpectIterative(g.handle, 1, 50U, 0.25);

  ASSERT_EQ(SetIterative(g.handle, 0, -1, -1.0), 0);
  ASSERT_EQ(SetIterative(g.handle, -1, 9, -1.0), 0);
  ASSERT_EQ(SetIterative(g.handle, -1, -1, 0.5), 0);
  ExpectIterative(g.handle, 0, 9U, 0.5);

  EXPECT_NE(fm_workbook_update_iterative(g.handle, nullptr), 0);
}

// SET of a value outside a field's domain is refused before anything is
// written: a negative max_change, and a max_iterations of 0.
TEST(CApiSetIterative, SetOutsideTheDomainIsRejectedAndChangesNothing) {
  WorkbookGuard g;
  ASSERT_EQ(fm_workbook_create_default(&g.handle), 0);
  ASSERT_EQ(SetIterative(g.handle, 1, 50, 0.25), 0);

  fm_iterative_settings_patch_t negative = FM_STRUCT_INIT(fm_iterative_settings_patch_t);
  negative.enabled_op = FM_PATCH_OP_SET;
  negative.enabled = 0;
  negative.max_change_op = FM_PATCH_OP_SET;
  negative.max_change = -1.0;
  EXPECT_EQ(fm_workbook_update_iterative(g.handle, &negative), kInvalidArgument);
  ExpectIterative(g.handle, 1, 50U, 0.25);

  fm_iterative_settings_patch_t zero = FM_STRUCT_INIT(fm_iterative_settings_patch_t);
  zero.enabled_op = FM_PATCH_OP_SET;
  zero.enabled = 0;
  zero.max_iterations_op = FM_PATCH_OP_SET;
  zero.max_iterations = 0U;
  EXPECT_EQ(fm_workbook_update_iterative(g.handle, &zero), kInvalidArgument);
  ExpectIterative(g.handle, 1, 50U, 0.25);
}

// No iterative field can be cleared, and an op outside 0..2 is refused;
// either refusal leaves every setting as it was.
TEST(CApiSetIterative, ClearAndUnknownOpsAreRejected) {
  WorkbookGuard g;
  ASSERT_EQ(fm_workbook_create_default(&g.handle), 0);
  ASSERT_EQ(SetIterative(g.handle, 1, 50, 0.25), 0);
  for (const int32_t raw_op : {static_cast<int32_t>(FM_PATCH_OP_CLEAR), 3}) {
    for (int field = 0; field < 3; ++field) {
      fm_iterative_settings_patch_t patch = FM_STRUCT_INIT(fm_iterative_settings_patch_t);
      patch.enabled_op = FM_PATCH_OP_SET;
      patch.enabled = 0;
      const auto op = static_cast<fm_patch_op_t>(raw_op);
      if (field == 0) {
        patch.max_iterations_op = op;
        patch.max_iterations = 5U;
      } else if (field == 1) {
        patch.max_change_op = op;
        patch.max_change = 0.5;
      } else {
        patch.enabled_op = op;
      }
      EXPECT_EQ(fm_workbook_update_iterative(g.handle, &patch), kInvalidArgument) << raw_op << " " << field;
      ExpectIterative(g.handle, 1, 50U, 0.25);
    }
  }
}

TEST(CApiSetIterative, GetterRejectsAShortRecord) {
  WorkbookGuard g;
  ASSERT_EQ(fm_workbook_create_default(&g.handle), 0);
  fm_iterative_settings_info_t info = FM_STRUCT_INIT(fm_iterative_settings_info_t);
  info.struct_size = 4U;
  EXPECT_EQ(fm_workbook_get_iterative(g.handle, &info), kInvalidArgument);
  EXPECT_NE(fm_workbook_get_iterative(g.handle, nullptr), 0);
  EXPECT_NE(fm_workbook_get_iterative(nullptr, &info), 0);
}

// A negative font size is refused with the record field named as `size_pt`.
TEST(CApiFontRecord, NegativeSizeNamesTheSizePtField) {
  WorkbookGuard g;
  ASSERT_EQ(fm_workbook_create_default(&g.handle), 0);
  fm_font_t font{};
  ASSERT_EQ(fm_workbook_get_font_at(g.handle, 0, &font), 0);
  font.size_pt = -1.0;
  std::uint32_t index = 0;
  EXPECT_EQ(fm_workbook_add_font(g.handle, &font, &index), kInvalidArgument);
  EXPECT_NE(std::string(fm_get_last_error_context()).find("font.size_pt="), std::string::npos)
      << fm_get_last_error_context();
}

TEST(CApiPivotCacheNumbers, RejectNonFiniteValues) {
  WorkbookGuard g;
  ASSERT_EQ(fm_workbook_create_default(&g.handle), 0);
  uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(g.handle, 0, &cache_id), 0);
  size_t field_index = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache_field(g.handle, cache_id, "f", &field_index), 0);
  size_t record_index = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache_record(g.handle, cache_id, &record_index), 0);
  for (const double bad : {kNaN, kInf, -kInf}) {
    EXPECT_EQ(
        fm_workbook_add_pivot_cache_field_shared_item_number(g.handle, cache_id, field_index, bad, ignored_index()),
        kInvalidArgument);
    EXPECT_EQ(fm_workbook_set_pivot_cache_record_number(g.handle, cache_id, record_index, field_index, bad),
              kInvalidArgument);
    EXPECT_EQ(fm_workbook_add_pivot_cache_field_shared_item_date(g.handle, cache_id, field_index, bad, ignored_index()),
              kInvalidArgument);
    EXPECT_EQ(fm_workbook_set_pivot_cache_record_date(g.handle, cache_id, record_index, field_index, bad),
              kInvalidArgument);
  }
  size_t count = 99;
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_shared_item_count(g.handle, cache_id, field_index, &count), 0);
  EXPECT_EQ(count, 0U);
  EXPECT_EQ(fm_workbook_add_pivot_cache_field_shared_item_number(g.handle, cache_id, field_index, 1.5, ignored_index()),
            0);
  EXPECT_EQ(fm_workbook_set_pivot_cache_record_number(g.handle, cache_id, record_index, field_index, 1.5), 0);
}

TEST(CApiCreateForProfile, UnknownProfileNullsOutPointer) {
  auto* sentinel = reinterpret_cast<fm_workbook_t*>(static_cast<std::uintptr_t>(0x10));
  fm_workbook_t* out = sentinel;
  EXPECT_EQ(fm_workbook_create_for_profile("bogus", &out), kInvalidArgument);
  EXPECT_EQ(out, nullptr);
}

}  // namespace
