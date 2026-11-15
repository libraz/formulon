// Cursor over a flat RawRecord list shared by the XLSB pivot block decoders.

#ifndef FORMULON_IO_XLSB_PIVOT_RECORD_CURSOR_H_
#define FORMULON_IO_XLSB_PIVOT_RECORD_CURSOR_H_

#include <cstddef>
#include <cstdint>
#include <vector>

#include "io/xlsb/pivot_records_common.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace pivot_records {

/// Names a decoded block in error text: `context` is the error context
/// ("xlsb_pivot_area"), `noun` the phrase in messages ("pivot area").
struct BlockNames {
  const char* context;
  const char* noun;
};

/// True when `records[pos]` exists and has record id `type`.
bool at(const std::vector<RawRecord>& records, std::size_t pos, std::uint16_t type);

/// Consumes the record at `*pos`, which must have id `type`.
Expected<const RawRecord*, Error> take(const BlockNames& block, const std::vector<RawRecord>& records, std::size_t* pos,
                                       std::uint16_t type);

/// Consumes an end record (id `type`), which must carry no payload.
Expected<void, Error> take_end(const BlockNames& block, const std::vector<RawRecord>& records, std::size_t* pos,
                               std::uint16_t type);

/// Appends a record with id `type` to `out`.
void push(std::vector<RawRecord>* out, std::uint16_t type, std::vector<std::uint8_t> payload = {});

}  // namespace pivot_records
}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PIVOT_RECORD_CURSOR_H_
