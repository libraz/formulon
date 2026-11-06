// C ABI argument validation: iterative max_change, pivot cache numbers and
// the out-pointer contract of fm_workbook_create_for_profile.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>

#include "c_api/formulon_c.h"
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

TEST(CApiSetIterative, RejectsNonFiniteMaxChangeAndKeepsOptions) {
  WorkbookGuard g;
  ASSERT_EQ(fm_workbook_create_default(&g.handle), 0);
  ASSERT_EQ(SetIterative(g.handle, 1, 50, 0.25), 0);
  for (const double bad : {kNaN, kInf, -kInf}) {
    EXPECT_EQ(SetIterative(g.handle, 0, 7, bad), kInvalidArgument);
  }
  int32_t enabled = 0;
  uint32_t iterations = 0;
  double max_change = 0.0;
  ASSERT_EQ(fm_workbook_get_iterative(g.handle, &enabled, &iterations, &max_change), 0);
  EXPECT_EQ(enabled, 1);
  EXPECT_EQ(iterations, 50U);
  EXPECT_DOUBLE_EQ(max_change, 0.25);
}

TEST(CApiSetIterative, AcceptsZeroMaxChange) {
  WorkbookGuard g;
  ASSERT_EQ(fm_workbook_create_default(&g.handle), 0);
  EXPECT_EQ(SetIterative(g.handle, 1, 10, 0.0), 0);
}

// A negative field of the patch keeps that setting, so each one can change
// alone and an all-negative patch changes nothing.
TEST(CApiSetIterative, NegativePatchFieldsKeepTheStoredSettings) {
  WorkbookGuard g;
  ASSERT_EQ(fm_workbook_create_default(&g.handle), 0);
  ASSERT_EQ(SetIterative(g.handle, 1, 50, 0.25), 0);
  int32_t enabled = 0;
  uint32_t iterations = 0;
  double max_change = 0.0;

  ASSERT_EQ(SetIterative(g.handle, -1, -1, -1.0), 0);
  ASSERT_EQ(fm_workbook_get_iterative(g.handle, &enabled, &iterations, &max_change), 0);
  EXPECT_EQ(enabled, 1);
  EXPECT_EQ(iterations, 50U);
  EXPECT_DOUBLE_EQ(max_change, 0.25);

  ASSERT_EQ(SetIterative(g.handle, 0, -1, -1.0), 0);
  ASSERT_EQ(SetIterative(g.handle, -1, 9, -1.0), 0);
  ASSERT_EQ(SetIterative(g.handle, -1, -1, 0.5), 0);
  ASSERT_EQ(fm_workbook_get_iterative(g.handle, &enabled, &iterations, &max_change), 0);
  EXPECT_EQ(enabled, 0);
  EXPECT_EQ(iterations, 9U);
  EXPECT_DOUBLE_EQ(max_change, 0.5);

  EXPECT_NE(fm_workbook_set_iterative(g.handle, nullptr), 0);
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
    EXPECT_EQ(fm_workbook_add_pivot_cache_field_shared_item_number(g.handle, cache_id, field_index, bad),
              kInvalidArgument);
    EXPECT_EQ(fm_workbook_set_pivot_cache_record_number(g.handle, cache_id, record_index, field_index, bad),
              kInvalidArgument);
    EXPECT_EQ(fm_workbook_add_pivot_cache_field_shared_item_date(g.handle, cache_id, field_index, bad),
              kInvalidArgument);
    EXPECT_EQ(fm_workbook_set_pivot_cache_record_date(g.handle, cache_id, record_index, field_index, bad),
              kInvalidArgument);
  }
  size_t count = 99;
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_shared_item_count(g.handle, cache_id, field_index, &count), 0);
  EXPECT_EQ(count, 0U);
  EXPECT_EQ(fm_workbook_add_pivot_cache_field_shared_item_number(g.handle, cache_id, field_index, 1.5), 0);
  EXPECT_EQ(fm_workbook_set_pivot_cache_record_number(g.handle, cache_id, record_index, field_index, 1.5), 0);
}

TEST(CApiCreateForProfile, UnknownProfileNullsOutPointer) {
  auto* sentinel = reinterpret_cast<fm_workbook_t*>(static_cast<std::uintptr_t>(0x10));
  fm_workbook_t* out = sentinel;
  EXPECT_EQ(fm_workbook_create_for_profile("bogus", &out), kInvalidArgument);
  EXPECT_EQ(out, nullptr);
}

}  // namespace
