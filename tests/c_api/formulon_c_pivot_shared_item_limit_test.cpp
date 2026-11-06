// A pivot-field shared item index wider than the 32-bit cache index is a limit,
// not an index into a counted collection.

#include <cstdint>
#include <limits>

#include "c_api/formulon_c.h"
#include "formulon_c_out_index.h"
#include "formulon_c_pivot_test_helpers.h"

TEST(FormulonCApiPivotSharedItemLimit, AWiderThan32BitCacheIndexIsAnInvalidArgument) {
  if (std::numeric_limits<std::size_t>::max() <= std::numeric_limits<std::uint32_t>::max()) {
    GTEST_SKIP() << "size_t cannot hold a cache index past 32 bits";
  }
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  const std::size_t wide = static_cast<std::size_t>(std::numeric_limits<std::uint32_t>::max()) + 1U;
  EXPECT_EQ(fm_sheet_add_pivot_field_shared_item(wb.handle, 0, pivot_index, 0, wide, 1, ignored_index()),
            FM_STATUS_INVALID_ARGUMENT);
  std::size_t item_count = 99;
  ASSERT_EQ(fm_sheet_get_pivot_field_item_count(wb.handle, 0, pivot_index, 0, &item_count), 0);
  EXPECT_EQ(item_count, 2U);
}
