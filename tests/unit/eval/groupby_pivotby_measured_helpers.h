// Cell-by-cell comparison helper for GROUPBY / PIVOTBY results measured in
// Excel. Expected cells are written as "n:<number>", "t:<text>" or "b:<0|1>";
// an empty-text placeholder is "t:".

#ifndef FORMULON_TESTS_UNIT_EVAL_GROUPBY_PIVOTBY_MEASURED_HELPERS_H_
#define FORMULON_TESTS_UNIT_EVAL_GROUPBY_PIVOTBY_MEASURED_HELPERS_H_

#include <cstdint>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "value.h"

namespace formulon::eval::measured {

inline std::string Render(const Value& v) {
  if (v.is_number()) {
    return "n:" + std::to_string(v.as_number());
  }
  if (v.is_text()) {
    return "t:" + std::string(v.as_text());
  }
  if (v.is_boolean()) {
    return std::string("b:") + (v.as_boolean() ? "1" : "0");
  }
  return "?:" + v.debug_to_string();
}

inline void ExpectCells(const Value& v, std::uint32_t rows, std::uint32_t cols,
                        const std::vector<std::string>& expected) {
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_rows(), rows);
  ASSERT_EQ(v.as_array_cols(), cols);
  ASSERT_EQ(expected.size(), static_cast<std::size_t>(rows) * cols);
  for (std::size_t i = 0; i < expected.size(); ++i) {
    EXPECT_EQ(Render(v.as_array_cells()[i]), expected[i]) << "cell " << i / cols << "," << i % cols;
  }
}

}  // namespace formulon::eval::measured

#endif  // FORMULON_TESTS_UNIT_EVAL_GROUPBY_PIVOTBY_MEASURED_HELPERS_H_
