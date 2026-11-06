// `fm_workbook_load` input checks: a NULL pointer and an empty buffer are
// different failures, and a short non-archive buffer is a corrupt ZIP.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"

TEST(FormulonCApiEmptyLoad, ANullBufferIsANullPointer) {
  fm_workbook_t* wb = reinterpret_cast<fm_workbook_t*>(1);
  EXPECT_EQ(fm_workbook_load(nullptr, 0, &wb), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(wb, nullptr);
  wb = reinterpret_cast<fm_workbook_t*>(1);
  EXPECT_EQ(fm_workbook_load(nullptr, 16, &wb), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(wb, nullptr);
}

TEST(FormulonCApiEmptyLoad, ANullOutputIsANullPointer) {
  const uint8_t byte = 0;
  EXPECT_EQ(fm_workbook_load(&byte, 1, nullptr), FM_STATUS_BINDING_NULL_POINTER);
}

TEST(FormulonCApiEmptyLoad, AnEmptyBufferIsAnInvalidArgument) {
  const uint8_t byte = 0;
  fm_workbook_t* wb = reinterpret_cast<fm_workbook_t*>(1);
  EXPECT_EQ(fm_workbook_load(&byte, 0, &wb), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(wb, nullptr);
  EXPECT_STREQ(fm_get_last_error_message(), "input is empty");
}

TEST(FormulonCApiEmptyLoad, ShortNonArchiveBuffersAreCorruptZips) {
  const uint8_t one = 'P';
  fm_workbook_t* wb = nullptr;
  EXPECT_EQ(fm_workbook_load(&one, 1, &wb), FM_STATUS_IO_ZIP_CORRUPT);
  EXPECT_EQ(wb, nullptr);

  std::vector<uint8_t> sixteen = {'P', 'K', 0x03, 0x04};
  sixteen.insert(sixteen.end(), 12, 0);
  EXPECT_EQ(fm_workbook_load(sixteen.data(), sixteen.size(), &wb), FM_STATUS_IO_ZIP_CORRUPT);
  EXPECT_EQ(wb, nullptr);
}
