//
// Number-format codes Excel displays built-in format ids with in each UI
// locale. A file stores only the id, so these codes are for rendering and
// never for writing.

#ifndef FORMULON_NUM_FMT_LOCALE_H_
#define FORMULON_NUM_FMT_LOCALE_H_

#include <cstdint>
#include <optional>
#include <string_view>

#include "excel_profile.h"

namespace formulon {

/// Stored-dialect code Excel renders built-in `id` with under `profile`; empty
/// for an id that has no built-in code (rendered as General). The locale's separators, month names
/// and currency glyphs are applied by the renderer; this supplies the code's
/// shape.
std::string_view localized_builtin_num_fmt(std::uint16_t id, ExcelProfile profile) noexcept;

/// True when `id` is a built-in format with a code in at least one locale.
bool is_builtin_num_fmt_id(std::uint16_t id) noexcept;

/// Lowest built-in id (above General) whose en-US code is `code`. A cache
/// stores formats that way, so this maps one back to the id to render it by.
std::optional<std::uint16_t> builtin_num_fmt_id_for_en_us_code(std::string_view code) noexcept;

}  // namespace formulon

#endif  // FORMULON_NUM_FMT_LOCALE_H_
