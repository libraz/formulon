#include "utils/a1_column.h"

namespace formulon::a1 {

bool append_column_letters(std::string& out, std::uint32_t column) {
  if (column >= kMaxColumns) {
    return false;
  }

  char reverse[3];
  std::uint32_t count = 0;
  std::uint32_t value = column + 1U;
  do {
    reverse[count++] = static_cast<char>('A' + (value - 1U) % 26U);
    value = (value - 1U) / 26U;
  } while (value != 0U);

  while (count != 0U) {
    out.push_back(reverse[--count]);
  }
  return true;
}

}  // namespace formulon::a1
