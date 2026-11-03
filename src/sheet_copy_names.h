//
// Name generation for a sheet copy: the default copy name of a worksheet and
// the renamed tables it carries.

#ifndef FORMULON_SHEET_COPY_NAMES_H_
#define FORMULON_SHEET_COPY_NAMES_H_

#include <cstddef>
#include <functional>
#include <string>
#include <string_view>
#include <vector>

namespace formulon {

/// Returns the default name of a copy of the worksheet `source_name`.
///
/// A trailing " (<digits>)" is stripped from the source (when something
/// remains) to form the base; the result is the first "base (N)", N >= 2, that
/// matches no entry of `existing_names` under worksheet-name identity. A name
/// over 31 UTF-16 units has its base truncated on a code-point boundary, then
/// trailing spaces and apostrophes are dropped (a base that would become empty
/// keeps its truncated form).
std::string copy_sheet_default_name(std::string_view source_name, const std::vector<std::string_view>& existing_names);

/// Returns the name of the copy of a table named `source_table_name`.
///
/// The candidate is the source name followed by N, where N starts at
/// `tables_in_workbook + 2`; an underscore separates them only when the
/// candidate would read as an A1 or R1C1 cell reference. While
/// `name_taken(candidate)` holds, N advances by one. `name_taken` owns the
/// case-insensitive comparison against table and defined names.
std::string copy_table_name(std::string_view source_table_name, std::size_t tables_in_workbook,
                            const std::function<bool(std::string_view)>& name_taken);

}  // namespace formulon

#endif  // FORMULON_SHEET_COPY_NAMES_H_
