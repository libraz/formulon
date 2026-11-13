//
// dxf references inside framed BIFF12 record streams. One table maps each
// record type that carries an index into the styles part's BrtDXF list to the
// byte offset of every such slot, each offset the one the record's reader
// decodes; `scan_dxf_slots` reports them and `rewrite_dxf_slots` maps them
// through a `DxfSavePlan` in place, so framing and every other byte stay
// unchanged. A slot holding kNoDxf is not a reference.
//
// A dxf-bearing record no reader decodes cannot be rewritten; a stream
// holding one is reported so the caller can keep the dxf table as it is.

#ifndef FORMULON_IO_XLSB_DXF_RECORD_SLOTS_H_
#define FORMULON_IO_XLSB_DXF_RECORD_SLOTS_H_

#include <cstddef>
#include <cstdint>
#include <vector>

#include "io/dxf_save_plan.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon::io::xlsb {

/// One dxf slot of a record stream.
struct DxfRecordSlot {
  std::uint16_t record_type = 0;
  std::size_t record_index = 0;  ///< Ordinal of the record in the stream.
  std::size_t slot = 0;          ///< Ordinal of the slot within the record's slot list.
  std::size_t offset = 0;        ///< Byte offset of the u32 within the stream.
  std::uint32_t value = 0;
};

/// What a stream holds that bears on the dxf table.
struct DxfSlotScan {
  std::vector<DxfRecordSlot> slots;        ///< Every slot not holding kNoDxf, in stream order.
  std::size_t dxf_records = 0;             ///< BrtDXF records.
  std::vector<std::uint16_t> undecodable;  ///< Types of undecodable dxf-bearing records, in stream order.
};

/// True for a record type that carries a dxf index no reader decodes.
bool is_undecodable_dxf_record(std::uint16_t type);

/// The dxf slots of `stream`. Fails when the stream does not frame whole or a
/// slot-bearing record is too short or shaped so its slots cannot be located.
Expected<DxfSlotScan, Error> scan_dxf_slots(const std::vector<std::uint8_t>& stream);

/// `stream` with every slot mapped through `plan`; a slot `plan` maps to
/// nothing becomes kNoDxf. Fails as `scan_dxf_slots` does.
Expected<std::vector<std::uint8_t>, Error> rewrite_dxf_slots(const std::vector<std::uint8_t>& stream,
                                                             const DxfSavePlan& plan);

}  // namespace formulon::io::xlsb

#endif  // FORMULON_IO_XLSB_DXF_RECORD_SLOTS_H_
