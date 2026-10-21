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

void parse_selection_sqref_text(std::string_view text, std::vector<MergeRange>& out) {
  for (std::size_t at = 0; at < text.size();) {
    const std::size_t end = std::min(text.find(' ', at), text.size());
    const std::string_view token = text.substr(at, end - at);
    at = end + 1U;
    std::vector<MergeRange> cells;
    if (parse_sqref_text(token, cells)) {
      if (cells.front().first_row <= cells.front().last_row && cells.front().first_col <= cells.front().last_col) {
        out.push_back(cells.front());
      }
      continue;
    }
    const std::size_t colon = token.find(':');
    if (colon == std::string_view::npos) {
      continue;
    }
    const std::string_view head = token.substr(0, colon);
    const std::string_view tail = token.substr(colon + 1U);
    std::size_t head_pos = 0;
    std::size_t tail_pos = 0;
    std::uint32_t first = 0;
    std::uint32_t last = 0;
    if (a1::parse_column_letters(head, &head_pos, &first) && head_pos == head.size() &&
        a1::parse_column_letters(tail, &tail_pos, &last) && tail_pos == tail.size()) {
      if (first <= last && last <= Sheet::kMaxCols) {
        out.push_back({0U, first - 1U, Sheet::kMaxRows - 1U, last - 1U});
      }
    } else if (head_pos = tail_pos = 0; a1::parse_uint(head, &head_pos, &first) && head_pos == head.size() &&
                                        first != 0U && a1::parse_uint(tail, &tail_pos, &last) &&
                                        tail_pos == tail.size() && first <= last && last <= Sheet::kMaxRows) {
      out.push_back({first - 1U, 0U, last - 1U, Sheet::kMaxCols - 1U});
    }
  }
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
