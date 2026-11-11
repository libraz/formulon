#include "utils/name_syntax.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "utils/a1_ref.h"
#include "utils/strings.h"
#include "utils/text_ops.h"
#include "utils/utf8_length.h"

namespace formulon {
namespace {

constexpr std::uint32_t kMaxNameUnits = 255U;

/// Full-width ASCII and U+3000 as the ASCII they stand for, upper-cased;
/// other code points stay as UTF-8.
std::string FoldName(std::string_view name) {
  std::string out;
  out.reserve(name.size());
  for (std::size_t i = 0; i < name.size();) {
    std::size_t bytes = 0;
    const std::uint32_t cp = decode_utf8_step(name, i, &bytes);
    if (cp >= 0xFF01U && cp <= 0xFF5EU) {
      out.push_back(strings::ascii_to_upper(static_cast<char>(cp - 0xFEE0U)));
    } else if (cp == 0x3000U) {
      out.push_back(' ');
    } else if (cp < 0x80U) {
      out.push_back(strings::ascii_to_upper(static_cast<char>(cp)));
    } else {
      out.append(name.substr(i, bytes));
    }
    i += bytes;
  }
  return out;
}

bool IsNameChar(char c, bool first) {
  const auto byte = static_cast<unsigned char>(c);
  if (byte >= 0x80U || strings::is_ascii_alpha(c) || c == '_' || c == '\\') {
    return true;
  }
  return !first && (strings::is_ascii_digit(c) || c == '.' || c == '?');
}

/// Reads the digit run at `*i`; true when it is non-empty and names an index
/// in [1, `limit`].
bool ReadInGridNumber(std::string_view text, std::size_t* i, std::uint32_t limit) {
  std::uint64_t value = 0;
  const std::size_t begin = *i;
  while (*i < text.size() && strings::is_ascii_digit(text[*i])) {
    if (value <= limit) {
      value = value * 10U + static_cast<std::uint64_t>(text[*i] - '0');
    }
    ++*i;
  }
  return *i != begin && value >= 1U && value <= limit;
}

/// True when the upper-cased `name` is an R1C1 reference or opens with one
/// that carries an in-grid number, the row and column letters being `row` and
/// `col` (upper case).
bool OpensWithR1C1(std::string_view name, char row, char col) {
  const auto at = [name](std::size_t i, char upper) { return i < name.size() && name[i] == upper; };
  std::size_t i = 0;
  if (at(0, row)) {
    i = 1;
    if (i < name.size() && strings::is_ascii_digit(name[i])) {
      return ReadInGridNumber(name, &i, a1::kMaxRows);
    }
    if (!at(i, col)) {
      return i == name.size();
    }
  } else if (!at(0, col)) {
    return false;
  }
  ++i;
  if (i < name.size() && strings::is_ascii_digit(name[i])) {
    return ReadInGridNumber(name, &i, a1::kMaxCols);
  }
  return i == name.size();
}

}  // namespace

bool is_valid_table_name(std::string_view name) {
  if (name.empty() || utf16_units_in(name) > kMaxNameUnits) {
    return false;
  }
  const std::string folded = FoldName(name);
  for (std::size_t i = 0; i < folded.size(); ++i) {
    if (!IsNameChar(folded[i], i == 0U)) {
      return false;
    }
  }
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  return folded != "TRUE" && folded != "FALSE" && !a1::parse_a1_ref(folded, &row, &col) &&
         !OpensWithR1C1(folded, 'R', 'C');
}

bool opens_with_r1c1(std::string_view name, char row_letter, char col_letter) {
  return OpensWithR1C1(FoldName(name), strings::ascii_to_upper(row_letter), strings::ascii_to_upper(col_letter));
}

bool is_valid_defined_name(std::string_view name) {
  if (!is_valid_table_name(name)) {
    return false;
  }
  if (name[0] != '\\' || name.size() < 2U) {
    return true;
  }
  std::size_t bytes = 0;
  decode_utf8_step(name, 1U, &bytes);
  return 1U + bytes != name.size();
}

}  // namespace formulon
