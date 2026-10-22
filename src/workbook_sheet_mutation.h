//
// Preconditions and bookkeeping shared by the `Workbook` per-sheet mutators
// that live in separate translation units. Internal to the workbook
// implementation; not part of the public API.

#ifndef FORMULON_WORKBOOK_SHEET_MUTATION_H_
#define FORMULON_WORKBOOK_SHEET_MUTATION_H_

#include <cstddef>
#include <string>
#include <vector>

#include "defined_name.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

Expected<void, Error> check_sheet_index(const char* op, std::size_t sheet_index, std::size_t sheet_count);
bool erase_filter_database_name(std::vector<DefinedName>& names, std::size_t sheet_index);
bool erase_slicer_cache_names(std::vector<DefinedName>& names, const std::vector<std::string>& cache_names);

}  // namespace formulon

#endif  // FORMULON_WORKBOOK_SHEET_MUTATION_H_
