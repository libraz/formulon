#ifndef FORMULON_UTILS_SBCS_CODEPAGE_H_
#define FORMULON_UTILS_SBCS_CODEPAGE_H_

#include <cstdint>

#include "excel_locale.h"

namespace formulon {

/// Returns the Unicode codepoint assigned to one byte in a single-byte code
/// page. ASCII bytes are shared by both tables; undefined Windows-1252 slots
/// decode to U+FFFD.
std::uint32_t sbcs_decode_byte(SbcsCodepage codepage, std::uint8_t byte) noexcept;

/// Returns the byte assigned to `codepoint`, or -1 when the codepoint is not
/// representable in the requested single-byte code page.
int sbcs_encode_codepoint(SbcsCodepage codepage, std::uint32_t codepoint) noexcept;

}  // namespace formulon

#endif  // FORMULON_UTILS_SBCS_CODEPAGE_H_
