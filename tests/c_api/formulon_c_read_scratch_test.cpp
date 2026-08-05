//
// Regression test for the per-handle read-path text storage. A long-lived
// handle that loops over text reads must not accumulate one scratch entry
// per call: read-path strings live in `read_scratch`, which is reset at the
// start of each fallible read so it only ever holds the most recent call's
// output. This test reaches into the TU-public handle struct
// (`src/c_api/parts/common.h`) to assert the bound directly.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "gtest/gtest.h"

namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

}  // namespace

// Looping `fm_workbook_get_value` over a text cell must keep the per-handle
// read scratch bounded (one entry, not one-per-iteration).
TEST(FormulonCApiReadScratch, GetValueDoesNotGrowScratch) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "hello"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  constexpr int kIterations = 10000;
  for (int i = 0; i < kIterations; ++i) {
    fm_value_t v{};
    ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
    ASSERT_EQ(v.kind, FM_VAL_TEXT);
    ASSERT_NE(v.u.text, nullptr);
    EXPECT_STREQ(v.u.text, "hello");
  }

  // The scratch must hold at most this call's output, regardless of how
  // many reads ran. Without the per-call reset it would hold `kIterations`.
  EXPECT_LE(wb.handle->read_scratch.size(), 1U);
}

// `fm_workbook_lambda_text_at` shares the same read scratch; repeated calls
// on a non-lambda cell (the error path) and on real reads must not leak.
TEST(FormulonCApiReadScratch, MixedReadsStayBounded) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "alpha"), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 1, 0, 42.0), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  for (int i = 0; i < 5000; ++i) {
    fm_value_t v{};
    ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
    fm_value_t n{};
    ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 1, 0, &n), 0);
    EXPECT_EQ(n.kind, FM_VAL_NUMBER);
  }
  EXPECT_LE(wb.handle->read_scratch.size(), 1U);
}

// Text writes own their current bytes in the cell rather than retaining every
// historical buffer in a handle-global store. Repeated overwrites must leave
// the workbook-level shared-string storage untouched and preserve the final
// text through later reads.
TEST(FormulonCApiReadScratch, RepeatedTextOverwritesDoNotRetainHistoricalBuffers) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  constexpr int kWrites = 1000;
  for (int i = 0; i < kWrites; ++i) {
    const std::string text = "value-" + std::to_string(i);
    ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, text.c_str()), 0);
  }
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  for (int i = 0; i < 1000; ++i) {
    fm_value_t v{};
    ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  }

  EXPECT_TRUE(wb.handle->workbook().text_storage().empty());
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  ASSERT_EQ(v.kind, FM_VAL_TEXT);
  EXPECT_STREQ(v.u.text, "value-999");
}
