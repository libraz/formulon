#include "utils/a1_column.h"

#include <cstdint>
#include <string>

#include "gtest/gtest.h"

namespace formulon::a1 {
namespace {

TEST(A1Column, EncodesExcelColumnBounds) {
  std::string out;
  EXPECT_TRUE(append_column_letters(out, 0U));
  EXPECT_EQ(out, "A");

  out.clear();
  EXPECT_TRUE(append_column_letters(out, 25U));
  EXPECT_EQ(out, "Z");

  out.clear();
  EXPECT_TRUE(append_column_letters(out, 26U));
  EXPECT_EQ(out, "AA");

  out.clear();
  EXPECT_TRUE(append_column_letters(out, kMaxColumns - 1U));
  EXPECT_EQ(out, "XFD");
}

TEST(A1Column, RejectsColumnsOutsideExcelGridWithoutMutatingOutput) {
  std::string out = "prefix";
  EXPECT_FALSE(append_column_letters(out, kMaxColumns));
  EXPECT_EQ(out, "prefix");
  EXPECT_FALSE(append_column_letters(out, UINT32_MAX));
  EXPECT_EQ(out, "prefix");
}

}  // namespace
}  // namespace formulon::a1
