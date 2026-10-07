//
// OOXML spellings of the `<dataValidation>` `type`, `operator` and `errorStyle`
// attributes, keyed by the stored code. Code 0 (`any`, `between`, `stop`) is the
// omitted default and has no entry.

#ifndef FORMULON_IO_DATA_VALIDATION_ATTR_NAMES_H_
#define FORMULON_IO_DATA_VALIDATION_ATTR_NAMES_H_

#include <array>
#include <cstdint>

#include "io/enum_name_table.h"

namespace formulon::io {

inline constexpr std::array<EnumName, 7> kDataValidationTypeNames = {{
    {1, "whole"},
    {2, "decimal"},
    {3, "list"},
    {4, "date"},
    {5, "time"},
    {6, "textLength"},
    {7, "custom"},
}};

inline constexpr std::array<EnumName, 7> kDataValidationOperatorNames = {{
    {1, "notBetween"},
    {2, "equal"},
    {3, "notEqual"},
    {4, "greaterThan"},
    {5, "lessThan"},
    {6, "greaterThanOrEqual"},
    {7, "lessThanOrEqual"},
}};

inline constexpr std::array<EnumName, 2> kDataValidationErrorStyleNames = {{
    {1, "warning"},
    {2, "information"},
}};

}  // namespace formulon::io

#endif  // FORMULON_IO_DATA_VALIDATION_ATTR_NAMES_H_
