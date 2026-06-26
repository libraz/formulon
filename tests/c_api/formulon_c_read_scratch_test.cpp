// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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

// The read scratch must not disturb the intern store: cell text written via
// `fm_workbook_set_text` is interned in `text_store` and stays alive for the
// handle's lifetime even after many read calls reset the scratch.
TEST(FormulonCApiReadScratch, InternStorePersistsAcrossReads) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "persistent"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  const std::size_t interned = wb.handle->text_store.size();
  EXPECT_GE(interned, 1U);

  for (int i = 0; i < 1000; ++i) {
    fm_value_t v{};
    ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  }

  // Reads never touch the intern store, so its size is unchanged and the
  // round-tripped text is still readable.
  EXPECT_EQ(wb.handle->text_store.size(), interned);
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  ASSERT_EQ(v.kind, FM_VAL_TEXT);
  EXPECT_STREQ(v.u.text, "persistent");
}
