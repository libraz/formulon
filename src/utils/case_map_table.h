//
// Data shapes for the generated case-mapping tables (`case_map_table.cpp`).

#ifndef FORMULON_UTILS_CASE_MAP_TABLE_H_
#define FORMULON_UTILS_CASE_MAP_TABLE_H_

#include <cstddef>
#include <cstdint>

namespace formulon {
namespace case_map {

/// Every `stride`-th code point in `first..last` maps to code point + `delta`.
struct CaseRun {
  std::uint16_t first;
  std::uint16_t last;
  std::int16_t delta;
  std::uint16_t stride;
};

struct CaseRuns {
  const CaseRun* runs;
  std::size_t count;
};

/// `count` pairs of (first, last) code points, sorted.
struct LetterRanges {
  const std::uint16_t* pairs;
  std::size_t count;
};

CaseRuns upper_runs() noexcept;
CaseRuns lower_runs() noexcept;
CaseRuns proper_start_runs() noexcept;
CaseRuns proper_mid_runs() noexcept;
LetterRanges letter_ranges() noexcept;

}  // namespace case_map
}  // namespace formulon

#endif  // FORMULON_UTILS_CASE_MAP_TABLE_H_
