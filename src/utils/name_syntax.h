//
// Excel's name syntax, shared by defined names and table display names.
//
// A name is 1-255 UTF-16 units. Its first character is a letter, `_` or
// `\`; later ones may also be digits, `.` and `?`. Any non-ASCII code point
// counts as a letter except full-width ASCII and U+3000, which are read as
// their ASCII counterparts first. TRUE / FALSE, an A1 cell inside the grid
// and any name opening with an R1C1 reference (`R`, `C`, `RC`, `R2...`,
// `C3...`, `RC3...`) are refused; a row or column number off the grid
// (`R0`, `C16385` as R1C1) is no reference. A defined name additionally
// refuses `\` followed by exactly one character.

#ifndef FORMULON_UTILS_NAME_SYNTAX_H_
#define FORMULON_UTILS_NAME_SYNTAX_H_

#include <string_view>

namespace formulon {

/// True when `name` is a valid table display name.
bool is_valid_table_name(std::string_view name);

/// True when `name` is a valid defined name.
bool is_valid_defined_name(std::string_view name);

}  // namespace formulon

#endif  // FORMULON_UTILS_NAME_SYNTAX_H_
