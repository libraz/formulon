#pragma once

#include <gtest/gtest.h>

#include <cstddef>
#include <tuple>
#include <utility>

#include "c_api/formulon_c.h"

namespace count_of_detail {

template <typename... Params, typename... Args, std::size_t... I>
fm_status_t call(fm_status_t (*fn)(Params...), std::index_sequence<I...>, std::size_t* count, Args... args) {
  return fn(static_cast<std::tuple_element_t<I, std::tuple<Params...>>>(args)..., count);
}

}  // namespace count_of_detail

// Count a status + `size_t* out_count` getter reports for `args`; a nonzero
// status fails the calling test.
template <typename... Params, typename... Args>
std::size_t count_of(fm_status_t (*fn)(Params...), Args... args) {
  static_assert(sizeof...(Args) + 1 == sizeof...(Params), "count_of takes every argument but the out-param");
  std::size_t count = 0;
  EXPECT_EQ(count_of_detail::call(fn, std::index_sequence_for<Args...>{}, &count, args...), 0);
  return count;
}
