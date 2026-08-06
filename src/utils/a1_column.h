#ifndef FORMULON_UTILS_A1_COLUMN_H_
#define FORMULON_UTILS_A1_COLUMN_H_

#include <cstdint>
#include <string>

namespace formulon::a1 {

inline constexpr std::uint32_t kMaxColumns = 16384U;

/// Appends Excel's uppercase bijective-base-26 column name for the 0-based
/// `column` (`0 -> A`, `16383 -> XFD`). Returns false and leaves `out`
/// unchanged when the column is outside Excel's grid.
bool append_column_letters(std::string& out, std::uint32_t column);

}  // namespace formulon::a1

#endif  // FORMULON_UTILS_A1_COLUMN_H_
