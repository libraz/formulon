// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the Workbook skeleton. Verifies the factory shape, sheet
// accessor mutation, and the basic byte-level shape of the save() output.

#include "workbook.h"

#include <cstddef>
#include <cstdint>
#include <vector>

#include "gtest/gtest.h"
#include "sheet.h"

namespace formulon {
namespace {

TEST(WorkbookTest, CreateYieldsSingleSheetNamedSheet1) {
  Workbook wb = Workbook::create();
  ASSERT_EQ(wb.sheet_count(), 1u);
  EXPECT_EQ(wb.sheet(0).name(), "Sheet1");
}

TEST(WorkbookTest, SheetCountReturnsOne) {
  Workbook wb = Workbook::create();
  EXPECT_EQ(wb.sheet_count(), static_cast<std::size_t>(1));
}

TEST(WorkbookTest, MutateSheetNamePropagates) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_name("Daten");
  EXPECT_EQ(wb.sheet(0).name(), "Daten");

  // The const overload must observe the same state.
  const Workbook& const_wb = wb;
  EXPECT_EQ(const_wb.sheet(0).name(), "Daten");
}

TEST(WorkbookTest, SaveProducesNonEmptyBytes) {
  Workbook wb = Workbook::create();
  auto result = wb.save();
  ASSERT_TRUE(static_cast<bool>(result)) << "save() failed: " << result.error().message;
  const std::vector<std::uint8_t>& bytes = result.value();
  EXPECT_GT(bytes.size(), 0u);
}

TEST(WorkbookTest, SaveIsZipMagicBytes) {
  Workbook wb = Workbook::create();
  auto result = wb.save();
  ASSERT_TRUE(static_cast<bool>(result));
  const std::vector<std::uint8_t>& bytes = result.value();
  ASSERT_GE(bytes.size(), 4u);
  EXPECT_EQ(bytes[0], 0x50u);  // 'P'
  EXPECT_EQ(bytes[1], 0x4Bu);  // 'K'
  EXPECT_EQ(bytes[2], 0x03u);
  EXPECT_EQ(bytes[3], 0x04u);
}

}  // namespace
}  // namespace formulon
