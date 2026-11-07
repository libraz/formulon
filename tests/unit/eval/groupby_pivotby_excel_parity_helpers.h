// Cell-by-cell comparison helper for GROUPBY / PIVOTBY results from
// Excel. Expected cells are written as "n:<number>", "t:<text>" or "b:<0|1>";
// an empty-text placeholder is "t:".

#ifndef FORMULON_TESTS_UNIT_EVAL_GROUPBY_PIVOTBY_EXCEL_PARITY_HELPERS_H_
#define FORMULON_TESTS_UNIT_EVAL_GROUPBY_PIVOTBY_EXCEL_PARITY_HELPERS_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "value.h"

namespace formulon::eval::excel_parity {

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

/// One expected spill cell: text (the empty string is the placeholder cell)
/// or a number compared bit-exactly.
struct ExactCell {
  explicit ExactCell(const char* t) : text(t) {}
  explicit ExactCell(double n) : is_number(true), number(n) {}
  bool is_number = false;
  double number = 0.0;
  std::string text;
};

/// A formula and the spill Excel produced for it.
struct ExactCase {
  const char* formula;
  std::uint32_t rows;
  std::uint32_t cols;
  std::vector<ExactCell> cells;
};

inline void ExpectExactCells(const Value& v, const ExactCase& expected) {
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_rows(), expected.rows);
  ASSERT_EQ(v.as_array_cols(), expected.cols);
  ASSERT_EQ(expected.cells.size(), static_cast<std::size_t>(expected.rows) * expected.cols);
  for (std::size_t i = 0; i < expected.cells.size(); ++i) {
    const Value& got = v.as_array_cells()[i];
    const ExactCell& want = expected.cells[i];
    if (want.is_number) {
      ASSERT_TRUE(got.is_number()) << "cell " << i / expected.cols << "," << i % expected.cols << ": "
                                   << got.debug_to_string();
      EXPECT_EQ(got.as_number(), want.number) << "cell " << i / expected.cols << "," << i % expected.cols;
    } else {
      ASSERT_TRUE(got.is_text()) << "cell " << i / expected.cols << "," << i % expected.cols << ": "
                                 << got.debug_to_string();
      EXPECT_EQ(std::string(got.as_text()), want.text) << "cell " << i / expected.cols << "," << i % expected.cols;
    }
  }
}

}  // namespace formulon::eval::excel_parity

#endif  // FORMULON_TESTS_UNIT_EVAL_GROUPBY_PIVOTBY_EXCEL_PARITY_HELPERS_H_
