#include "io/xlsb/pivot_calc_item_records.h"

#include <utility>

#include "io/xlsb/pivot_record_cursor.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace pivot_calc {

namespace {

using pivot_records::RawRecord;

constexpr pivot_records::BlockNames kBlock{"xlsb_pivot_calc_item", "pivot calculated item"};

constexpr std::size_t kCountSize = 4U;
// BrtBeginPName: field u32, flags u16.
constexpr std::size_t kNameSize = 6U;
// BrtBeginPNPair: flags u8, field u32, item u32.
constexpr std::size_t kPairSize = 9U;

Error Corrupt(const char* message) {
  return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, message, "context=xlsb_pivot_calc_item");
}

Expected<std::uint32_t, Error> TakeCount(const std::vector<RawRecord>& records, std::size_t* pos, std::uint16_t type) {
  ASSIGN_OR_RETURN(const RawRecord* rec, pivot_records::take(kBlock, records, pos, type));
  if (rec->payload.size() != kCountSize) {
    return Corrupt("xlsb pivot calculated item count record has an unexpected size");
  }
  return load_u32(rec->payload.data());
}

void PushCount(std::vector<RawRecord>* out, std::uint16_t type, std::size_t count) {
  std::vector<std::uint8_t> payload;
  emit_u32(payload, static_cast<std::uint32_t>(count));
  pivot_records::push(out, type, std::move(payload));
}

// Reads `cb` bytes counted by a leading u32 from `cursor` into `*out`.
bool ReadCounted(ByteSpan& cursor, std::vector<std::uint8_t>* out) {
  if (cursor.size < kCountSize) {
    return false;
  }
  const std::uint32_t size = load_u32(cursor.data);
  if (cursor.size - kCountSize < size) {
    return false;
  }
  out->assign(cursor.data + kCountSize, cursor.data + kCountSize + size);
  cursor.data += kCountSize + size;
  cursor.size -= kCountSize + size;
  return true;
}

Expected<PName, Error> DecodeName(const std::vector<RawRecord>& records, std::size_t* pos) {
  ASSIGN_OR_RETURN(const RawRecord* head, pivot_records::take(kBlock, records, pos, pivot_records::kBeginPName));
  if (head->payload.size() != kNameSize) {
    return Corrupt("xlsb pivot calculated item name record has an unexpected size");
  }
  PName name;
  name.field = load_u32(head->payload.data());
  name.flags = load_u16(head->payload.data() + 4);
  if (!pivot_records::at(records, *pos, pivot_records::kBeginPNPairs)) {
    RETURN_IF_ERROR(pivot_records::take_end(kBlock, records, pos, pivot_records::kEndPName));
    return name;
  }
  ASSIGN_OR_RETURN(const std::uint32_t count, TakeCount(records, pos, pivot_records::kBeginPNPairs));
  name.pairs.emplace();
  while (pivot_records::at(records, *pos, pivot_records::kBeginPNPair)) {
    const RawRecord& rec = records[(*pos)++];
    if (rec.payload.size() != kPairSize) {
      return Corrupt("xlsb pivot calculated item name pair has an unexpected size");
    }
    PNPair pair;
    pair.flags = rec.payload[0];
    pair.field = load_u32(rec.payload.data() + 1);
    pair.item = load_u32(rec.payload.data() + 5);
    name.pairs->push_back(pair);
    RETURN_IF_ERROR(pivot_records::take_end(kBlock, records, pos, pivot_records::kEndPNPair));
  }
  RETURN_IF_ERROR(pivot_records::take_end(kBlock, records, pos, pivot_records::kEndPNPairs));
  if (name.pairs->size() != count) {
    return Corrupt("xlsb pivot calculated item name pair count does not match its pairs");
  }
  RETURN_IF_ERROR(pivot_records::take_end(kBlock, records, pos, pivot_records::kEndPName));
  return name;
}

Expected<std::vector<PName>, Error> DecodeNames(const std::vector<RawRecord>& records, std::size_t* pos) {
  ASSIGN_OR_RETURN(const std::uint32_t count, TakeCount(records, pos, pivot_records::kBeginPNames));
  std::vector<PName> names;
  while (pivot_records::at(records, *pos, pivot_records::kBeginPName)) {
    ASSIGN_OR_RETURN(PName name, DecodeName(records, pos));
    names.push_back(std::move(name));
  }
  RETURN_IF_ERROR(pivot_records::take_end(kBlock, records, pos, pivot_records::kEndPNames));
  if (names.size() != count) {
    return Corrupt("xlsb pivot calculated item name count does not match its names");
  }
  return names;
}

Expected<CalcItem, Error> DecodeItem(const std::vector<RawRecord>& records, std::size_t* pos) {
  ASSIGN_OR_RETURN(const RawRecord* head, pivot_records::take(kBlock, records, pos, pivot_records::kBeginPCDCalcItem));
  ByteSpan cursor{head->payload.data(), head->payload.size()};
  CalcItem item;
  ASSIGN_OR_RETURN(item.field, read_u32(cursor));
  if (!ReadCounted(cursor, &item.rgce) || !ReadCounted(cursor, &item.rgcb) || cursor.size != 0U) {
    return Corrupt("xlsb pivot calculated item formula does not fill its record");
  }
  if (pivot_records::at(records, *pos, pivot_records::kBeginPRule)) {
    ASSIGN_OR_RETURN(pivot_area::PivotArea area, pivot_area::decode_pivot_area(records, pos));
    item.area = std::move(area);
  }
  if (pivot_records::at(records, *pos, pivot_records::kBeginPNames)) {
    ASSIGN_OR_RETURN(std::vector<PName> names, DecodeNames(records, pos));
    item.names = std::move(names);
  }
  RETURN_IF_ERROR(pivot_records::take_end(kBlock, records, pos, pivot_records::kEndPCDCalcItem));
  return item;
}

}  // namespace

Expected<std::vector<CalcItem>, Error> decode_calc_items(const std::vector<RawRecord>& records, std::size_t* pos) {
  ASSIGN_OR_RETURN(const std::uint32_t count, TakeCount(records, pos, pivot_records::kBeginPCDCalcItems));
  std::vector<CalcItem> items;
  while (pivot_records::at(records, *pos, pivot_records::kBeginPCDCalcItem)) {
    auto item = DecodeItem(records, pos);
    if (!item) {
      return std::move(item.error());
    }
    items.push_back(std::move(item.value()));
  }
  RETURN_IF_ERROR(pivot_records::take_end(kBlock, records, pos, pivot_records::kEndPCDCalcItems));
  if (items.size() != count) {
    return Corrupt("xlsb pivot calculated item count does not match its items");
  }
  return items;
}

void encode_calc_items(const std::vector<CalcItem>& items, std::vector<RawRecord>* out) {
  PushCount(out, pivot_records::kBeginPCDCalcItems, items.size());
  for (const CalcItem& item : items) {
    std::vector<std::uint8_t> head;
    emit_u32(head, item.field);
    emit_cell_parsed_formula(head, item.rgce, item.rgcb);
    pivot_records::push(out, pivot_records::kBeginPCDCalcItem, std::move(head));
    if (item.area) {
      pivot_area::encode_pivot_area(*item.area, out);
    }
    if (item.names) {
      encode_names(*item.names, out);
    }
    pivot_records::push(out, pivot_records::kEndPCDCalcItem);
  }
  pivot_records::push(out, pivot_records::kEndPCDCalcItems);
}

Expected<std::vector<PName>, Error> decode_names(const std::vector<RawRecord>& records) {
  std::size_t pos = 0;
  ASSIGN_OR_RETURN(std::vector<PName> names, DecodeNames(records, &pos));
  if (pos != records.size()) {
    return Corrupt("xlsb pivot calculated field names are followed by other records");
  }
  return names;
}

void encode_names(const std::vector<PName>& names, std::vector<RawRecord>* out) {
  PushCount(out, pivot_records::kBeginPNames, names.size());
  for (const PName& name : names) {
    std::vector<std::uint8_t> name_head;
    emit_u32(name_head, name.field);
    emit_u16(name_head, name.flags);
    pivot_records::push(out, pivot_records::kBeginPName, std::move(name_head));
    if (name.pairs) {
      PushCount(out, pivot_records::kBeginPNPairs, name.pairs->size());
      for (const PNPair& pair : *name.pairs) {
        std::vector<std::uint8_t> payload;
        emit_u8(payload, pair.flags);
        emit_u32(payload, pair.field);
        emit_u32(payload, pair.item);
        pivot_records::push(out, pivot_records::kBeginPNPair, std::move(payload));
        pivot_records::push(out, pivot_records::kEndPNPair);
      }
      pivot_records::push(out, pivot_records::kEndPNPairs);
    }
    pivot_records::push(out, pivot_records::kEndPName);
  }
  pivot_records::push(out, pivot_records::kEndPNames);
}

}  // namespace pivot_calc
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
