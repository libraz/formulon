#include "io/sqref_text.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "utils/a1_column.h"
#include "utils/a1_ref.h"

namespace formulon::io {

bool parse_sqref_text(std::string_view text, std::vector<MergeRange>& out) {
  std::size_t i = 0;
  while (i < text.size()) {
    if (text[i] == ' ') {
      ++i;
      continue;
    }
    const std::size_t end = std::min(text.find(' ', i), text.size());
    const std::string_view token = text.substr(i, end - i);
    const std::size_t colon = token.find(':');
    MergeRange range;
    if (!a1::parse_a1_ref(token.substr(0, colon), &range.first_row, &range.first_col)) {
      return false;
    }
    range.last_row = range.first_row;
    range.last_col = range.first_col;
    if (colon != std::string_view::npos &&
        !a1::parse_a1_ref(token.substr(colon + 1U), &range.last_row, &range.last_col)) {
      return false;
    }
    out.push_back(range);
    i = end;
  }
  return !out.empty();
}

std::string format_sqref_text(const std::vector<MergeRange>& ranges) {
  std::string out;
  const auto cell = [&out](std::uint32_t row, std::uint32_t col) {
    a1::append_column_letters(out, col);
    out += std::to_string(static_cast<std::uint64_t>(row) + 1U);
  };
  for (const MergeRange& r : ranges) {
    if (!out.empty()) {
      out += ' ';
    }
    cell(r.first_row, r.first_col);
    if (r.first_row != r.last_row || r.first_col != r.last_col) {
      out += ':';
      cell(r.last_row, r.last_col);
    }
  }
  return out;
}

}  // namespace formulon::io
