//
// Preconditions and bookkeeping shared by the `Workbook` per-sheet mutators
// that live in separate translation units. Internal to the workbook
// implementation; not part of the public API.

#ifndef FORMULON_WORKBOOK_SHEET_MUTATION_H_
#define FORMULON_WORKBOOK_SHEET_MUTATION_H_

#include <cstddef>
#include <vector>

#include "defined_name.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

Expected<void, Error> check_sheet_index(const char* op, std::size_t sheet_index, std::size_t sheet_count);
bool erase_filter_database_name(std::vector<DefinedName>& names, std::size_t sheet_index);

}  // namespace formulon

#endif  // FORMULON_WORKBOOK_SHEET_MUTATION_H_
