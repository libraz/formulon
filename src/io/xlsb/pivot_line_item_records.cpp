#include "io/xlsb/pivot_line_item_records.h"

#include <utility>

#include "io/xlsb/record_writer.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr const char* kContext = "context=xlsb_pivot_line_items";
constexpr std::size_t kLineItemSize = 12;

Error Corrupt(const char* message) {
  return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, message, kContext);
}

Error Unexpected(const char* message) {
  return make_error(FormulonErrorCode::kIoXlsbUnknownRecord, message, kContext);
}

std::uint16_t BeginId(SxLineItemAxis axis) {
  return axis == SxLineItemAxis::kRows ? pivot_records::kBeginSXLIRws : pivot_records::kBeginSXLICols;
}

std::uint16_t EndId(SxLineItemAxis axis) {
  return axis == SxLineItemAxis::kRows ? pivot_records::kEndSXLIRws : pivot_records::kEndSXLICols;
}

/// Consumes `records[*pos]` when it is an empty record of `type`.
bool TakeEnd(const std::vector<XlsbRecord>& records, std::size_t* pos, std::uint16_t type) {
  if (*pos >= records.size() || records[*pos].type != type || records[*pos].payload.size != 0U) {
    return false;
  }
  ++*pos;
  return true;
}

}  // namespace

Expected<SxLineItems, Error> decode_line_items(const std::vector<XlsbRecord>& records, std::size_t* pos,
                                               SxLineItemAxis axis) {
  if (*pos >= records.size() || records[*pos].type != BeginId(axis)) {
    return Unexpected("xlsb pivot line-item block does not start with its begin record");
  }
  const ByteSpan begin = records[*pos].payload;
  if (begin.size != 4U) {
    return Corrupt("xlsb pivot line-item block count has an unexpected size");
  }
  const std::uint32_t declared = load_u32(begin.data);
  ++*pos;
  SxLineItems out;
  while (*pos < records.size() && records[*pos].type == pivot_records::kBeginSXLI) {
    const ByteSpan head = records[*pos].payload;
    if (head.size != kLineItemSize) {
      return Corrupt("xlsb pivot line item has an unexpected size");
    }
    SxLineItem item;
    item.r = load_u16(head.data);
    item.t = head.data[2];
    if (item.t > kSxLineItemTypeMax || head.data[3] != 0U) {
      return Corrupt("xlsb pivot line item carries a type or flag this codec has not characterised");
    }
    const std::uint32_t count = load_u32(head.data + 4);
    item.i = load_u32(head.data + 8);
    ++*pos;
    if (count != 0U) {
      if (*pos >= records.size() || records[*pos].type != pivot_records::kBeginISXVIs) {
        return Unexpected("xlsb pivot line item lacks its member index array");
      }
      const ByteSpan indices = records[*pos].payload;
      if (indices.size != static_cast<std::size_t>(count) * 4U) {
        return Corrupt("xlsb pivot line item index array does not match its count");
      }
      item.x.reserve(count);
      for (std::uint32_t k = 0; k < count; ++k) {
        item.x.push_back(load_u32(indices.data + static_cast<std::size_t>(k) * 4U));
      }
      ++*pos;
      if (!TakeEnd(records, pos, pivot_records::kEndISXVIs)) {
        return Unexpected("xlsb pivot line item index array is not closed");
      }
    }
    if (!TakeEnd(records, pos, pivot_records::kEndSXLI)) {
      return Unexpected("xlsb pivot line item is not closed");
    }
    out.items.push_back(std::move(item));
  }
  if (!TakeEnd(records, pos, EndId(axis))) {
    return Unexpected("xlsb pivot line-item block carries an uncharacterised record");
  }
  if (out.items.size() != declared) {
    return Corrupt("xlsb pivot line-item count does not match the items decoded");
  }
  return out;
}

void encode_line_items(std::vector<std::uint8_t>& dst, const SxLineItems& items, SxLineItemAxis axis) {
  std::vector<std::uint8_t> payload;
  emit_u32(payload, static_cast<std::uint32_t>(items.items.size()));
  emit_record(dst, BeginId(axis), payload);
  const std::vector<std::uint8_t> empty;
  for (const SxLineItem& item : items.items) {
    payload.clear();
    emit_u16(payload, item.r);
    emit_u8(payload, item.t);
    emit_u8(payload, 0);
    emit_u32(payload, static_cast<std::uint32_t>(item.x.size()));
    emit_u32(payload, item.i);
    emit_record(dst, pivot_records::kBeginSXLI, payload);
    if (!item.x.empty()) {
      payload.clear();
      for (const std::uint32_t v : item.x) {
        emit_u32(payload, v);
      }
      emit_record(dst, pivot_records::kBeginISXVIs, payload);
      emit_record(dst, pivot_records::kEndISXVIs, empty);
    }
    emit_record(dst, pivot_records::kEndSXLI, empty);
  }
  emit_record(dst, EndId(axis), empty);
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
