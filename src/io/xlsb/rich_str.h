//
// `RichStr` ([MS-XLSB] §2.5.87) codec shared by the shared-string table and
// the comment text records: a flags byte, the string, then the rich-text runs
// and the phonetic guide the flags announce.

#ifndef FORMULON_IO_XLSB_RICH_STR_H_
#define FORMULON_IO_XLSB_RICH_STR_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "io/xlsb/record.h"
#include "phonetic.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// `RichStr` flag bits: rich-text runs follow the string when the first is
/// set, a phonetic guide when the second is.
inline constexpr std::uint8_t kRichStrRichRuns = 0x01U;
inline constexpr std::uint8_t kRichStrPhonetic = 0x02U;

/// One `StrRun`: the text from UTF-16 offset `ich` on uses font `ifnt`.
struct StrRun {
  std::uint16_t ich = 0;
  std::uint16_t ifnt = 0;
};

/// A decoded `RichStr`.
struct RichStr {
  std::string text;
  std::vector<StrRun> runs;
  std::vector<PhoneticRun> phonetic;
  PhoneticProperties phonetic_props;
};

/// Reads one `RichStr` from the front of `cursor`.
Expected<RichStr, Error> read_rich_str(ByteSpan& cursor);

/// Appends a `RichStr`; the runs flag is set when `runs` is non-empty and the
/// phonetic flag when `phonetic` is.
void emit_rich_str(std::vector<std::uint8_t>& out, std::string_view text, const std::vector<StrRun>& runs,
                   const std::vector<PhoneticRun>& phonetic, PhoneticProperties phonetic_props);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_RICH_STR_H_
