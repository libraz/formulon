//
// MS-XLSB shared-string table decoding: `xl/sharedStrings.bin` (`BrtSSTItem`
// `RichStr` entries) into workbook-lifetime text plus each entry's phonetic
// guide. The counterpart of `io/xlsb/sst_writer.h`.

#ifndef FORMULON_IO_XLSB_SST_READER_H_
#define FORMULON_IO_XLSB_SST_READER_H_

#include <cstdint>
#include <deque>
#include <string>
#include <string_view>
#include <vector>

#include "phonetic.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Decodes `xl/sharedStrings.bin` into an in-order list of string
/// payloads, appending each one into `text_storage` so cells can take
/// non-owning views. `out_phonetic` is filled in parallel with `entries`
/// -- one (possibly empty) run list per SST index, exactly as the OOXML
/// reader's `phonetic_for_entries` is, and `out_phonetic_props` beside it.
Expected<std::vector<std::string_view>, Error> DecodeSharedStringsBin(
    const std::vector<std::uint8_t>& body, std::deque<std::string>& text_storage,
    std::vector<std::vector<PhoneticRun>>& out_phonetic, std::vector<PhoneticProperties>& out_phonetic_props);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_SST_READER_H_
