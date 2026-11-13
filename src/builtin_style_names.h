// Excel UI-locale names of the 54 built-in cell styles (builtinId 0..53).

#ifndef FORMULON_BUILTIN_STYLE_NAMES_H_
#define FORMULON_BUILTIN_STYLE_NAMES_H_

#include <cstdint>
#include <string>

#include "excel_profile.h"

namespace formulon {

/// Number of built-in cell styles; valid ids are `0 .. kBuiltinStyleCount - 1`.
constexpr std::uint32_t kBuiltinStyleCount = 54U;

/// Name Excel shows for built-in cell style `builtin_id` in `locale`.
///
/// Ids 1 and 2 (RowLevel_ / ColLevel_) are stems the saved name extends with
/// the decimal of `i_level + 1`; for every other id `i_level` is ignored.
/// Returns an empty string when `builtin_id >= kBuiltinStyleCount`.
std::string builtin_style_name(std::uint32_t builtin_id, std::uint32_t i_level, ExcelLocale locale);

}  // namespace formulon

#endif  // FORMULON_BUILTIN_STYLE_NAMES_H_
