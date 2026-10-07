// BrtHLink codec: one hyperlink record to and from `Hyperlink`.

#ifndef FORMULON_IO_XLSB_HYPERLINK_RECORDS_H_
#define FORMULON_IO_XLSB_HYPERLINK_RECORDS_H_

#include <cstdint>
#include <string_view>
#include <vector>

#include "io/xlsb/record.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Sheet;
struct Hyperlink;

namespace io {
namespace xlsb {

/// Decodes one BrtHLink record and appends the hyperlink to `sheet`.
Expected<void, Error> decode_hyperlink(const XlsbRecord& rec, Sheet& sheet);

/// Emits `hyperlink` as one BrtHLink record onto `dst`; `rid` is the
/// relationship id written in the record (empty for an internal link).
Expected<void, Error> emit_hyperlink(std::vector<std::uint8_t>& dst, const Hyperlink& hyperlink, std::string_view rid);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_HYPERLINK_RECORDS_H_
