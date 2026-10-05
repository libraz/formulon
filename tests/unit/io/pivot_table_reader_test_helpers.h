#pragma once

#include <cstddef>
#include <cstdint>
#include <string_view>
#include <vector>

namespace formulon::io::pivot_table_reader_test_support {

inline std::vector<std::uint8_t> Bytes(std::string_view xml) {
  return std::vector<std::uint8_t>(xml.begin(), xml.end());
}

inline std::size_t CountOccurrences(std::string_view haystack, std::string_view needle) {
  std::size_t count = 0;
  for (std::size_t at = haystack.find(needle); at != std::string_view::npos; at = haystack.find(needle, at + 1)) {
    ++count;
  }
  return count;
}

inline constexpr std::string_view kXmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
inline constexpr std::string_view kPivotNs = " xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"";

}  // namespace formulon::io::pivot_table_reader_test_support
