//
// `RichStr` flag bits shared by the shared-string table reader and writer.

#ifndef FORMULON_IO_XLSB_RICH_STR_H_
#define FORMULON_IO_XLSB_RICH_STR_H_

#include <cstdint>

namespace formulon {
namespace io {
namespace xlsb {

/// `RichStr` flag bits ([MS-XLSB] §2.5.87): rich-text runs follow the
/// string when the first is set, a phonetic guide when the second is.
inline constexpr std::uint8_t kRichStrRichRuns = 0x01U;
inline constexpr std::uint8_t kRichStrPhonetic = 0x02U;

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_RICH_STR_H_
