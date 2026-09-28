//
// Unit tests for the shared sort body in `utils/index_sort.h`: the plain
// ascending `uint32_t` sort must agree with `std::sort`, and the element sort
// must keep equal elements in input order.

#include "utils/index_sort.h"

#include <algorithm>
#include <cstdint>
#include <random>
#include <utility>
#include <vector>

#include "gtest/gtest.h"

namespace formulon {
namespace {

TEST(SortAscending, EmptyAndSingle) {
  std::vector<std::uint32_t> empty;
  sort_ascending(empty);
  EXPECT_TRUE(empty.empty());

  std::vector<std::uint32_t> single{42U};
  sort_ascending(single);
  EXPECT_EQ(single, std::vector<std::uint32_t>{42U});
}

TEST(SortAscending, MatchesStdSortWithDuplicatesAndExtremes) {
  std::mt19937 rng(20260929U);
  for (const std::uint32_t modulus : {3U, 1000U, 0xFFFFFFFFU}) {
    std::uniform_int_distribution<std::uint32_t> dist(0U, modulus);
    std::vector<std::uint32_t> values(2000U);
    for (std::uint32_t& value : values) {
      value = dist(rng);
    }
    values.push_back(0U);
    values.push_back(0xFFFFFFFFU);
    std::vector<std::uint32_t> expected = values;
    std::sort(expected.begin(), expected.end());
    sort_ascending(values);
    EXPECT_EQ(values, expected) << "modulus " << modulus;
  }
}

TEST(SortByIndex, KeepsInputOrderForEqualKeys) {
  std::vector<std::pair<double, int>> items{{3.0, 0}, {1.0, 1}, {3.0, 2}, {1.0, 3}, {2.0, 4}, {1.0, 5}};
  sort_by_index(items,
                [](const std::pair<double, int>& a, const std::pair<double, int>& b) { return a.first < b.first; });
  const std::vector<std::pair<double, int>> expected{{1.0, 1}, {1.0, 3}, {1.0, 5}, {2.0, 4}, {3.0, 0}, {3.0, 2}};
  EXPECT_EQ(items, expected);
}

}  // namespace
}  // namespace formulon
