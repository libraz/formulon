#include "io/xlsb/pivot_calc_item_records.h"

#include <string>
#include <utility>

#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace pivot_calc {

namespace {

using pivot_records::RawRecord;

constexpr std::size_t kCountSize = 4U;
// BrtBeginPName: field u32, flags u16.
constexpr std::size_t kNameSize = 6U;
// BrtBeginPNPair: field u32, flags u8, item u32.
constexpr std::size_t kPairSize = 9U;
// Payload-free operator ptgs: PtgAdd (0x03) through PtgParen (0x15).
constexpr std::uint8_t kFirstOperatorPtg = 0x03U;
constexpr std::uint8_t kLastOperatorPtg = 0x15U;
constexpr std::size_t kSxNameSize = 6U;

Error Corrupt(const char* message) {
  return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, message, "context=xlsb_pivot_calc_item");
}

Error Unexpected(std::uint16_t type) {
  return make_error(FormulonErrorCode::kIoXlsbUnknownRecord,
                    "xlsb pivot calculated item carries a record outside its grammar",
                    "context=xlsb_pivot_calc_item record_id=" + std::to_string(type));
}

bool At(const std::vector<RawRecord>& records, std::size_t pos, std::uint16_t type) {
  return pos < records.size() && records[pos].type == type;
}

Expected<const RawRecord*, Error> Take(const std::vector<RawRecord>& records, std::size_t* pos, std::uint16_t type) {
  if (*pos >= records.size()) {
    return Corrupt("xlsb pivot calculated item block ends before its end record");
  }
  const RawRecord& rec = records[*pos];
  if (rec.type != type) {
    return Unexpected(rec.type);
  }
  ++*pos;
  return &rec;
}

Expected<void, Error> TakeEnd(const std::vector<RawRecord>& records, std::size_t* pos, std::uint16_t type) {
  ASSIGN_OR_RETURN(const RawRecord* rec, Take(records, pos, type));
  if (!rec->payload.empty()) {
    return Corrupt("xlsb pivot calculated item end record carries a payload");
  }
  return Expected<void, Error>::Ok();
}

Expected<std::uint32_t, Error> TakeCount(const std::vector<RawRecord>& records, std::size_t* pos, std::uint16_t type) {
  ASSIGN_OR_RETURN(const RawRecord* rec, Take(records, pos, type));
  if (rec->payload.size() != kCountSize) {
    return Corrupt("xlsb pivot calculated item count record has an unexpected size");
  }
  return load_u32(rec->payload.data());
}

void Push(std::vector<RawRecord>* out, std::uint16_t type, std::vector<std::uint8_t> payload = {}) {
  out->push_back(RawRecord{type, std::move(payload)});
}

void PushCount(std::vector<RawRecord>* out, std::uint16_t type, std::size_t count) {
  std::vector<std::uint8_t> payload;
  emit_u32(payload, static_cast<std::uint32_t>(count));
  Push(out, type, std::move(payload));
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
  ASSIGN_OR_RETURN(const RawRecord* head, Take(records, pos, pivot_records::kBeginPName));
  if (head->payload.size() != kNameSize) {
    return Corrupt("xlsb pivot calculated item name record has an unexpected size");
  }
  PName name;
  name.field = load_u32(head->payload.data());
  name.flags = load_u16(head->payload.data() + 4);
  ASSIGN_OR_RETURN(const std::uint32_t count, TakeCount(records, pos, pivot_records::kBeginPNPairs));
  while (At(records, *pos, pivot_records::kBeginPNPair)) {
    const RawRecord& rec = records[(*pos)++];
    if (rec.payload.size() != kPairSize) {
      return Corrupt("xlsb pivot calculated item name pair has an unexpected size");
    }
    PNPair pair;
    pair.field = load_u32(rec.payload.data());
    pair.flags = rec.payload[4];
    pair.item = load_u32(rec.payload.data() + 5);
    name.pairs.push_back(pair);
    RETURN_IF_ERROR(TakeEnd(records, pos, pivot_records::kEndPNPair));
  }
  RETURN_IF_ERROR(TakeEnd(records, pos, pivot_records::kEndPNPairs));
  if (name.pairs.size() != count) {
    return Corrupt("xlsb pivot calculated item name pair count does not match its pairs");
  }
  RETURN_IF_ERROR(TakeEnd(records, pos, pivot_records::kEndPName));
  return name;
}

Expected<std::vector<PName>, Error> DecodeNames(const std::vector<RawRecord>& records, std::size_t* pos) {
  ASSIGN_OR_RETURN(const std::uint32_t count, TakeCount(records, pos, pivot_records::kBeginPNames));
  std::vector<PName> names;
  while (At(records, *pos, pivot_records::kBeginPName)) {
    ASSIGN_OR_RETURN(PName name, DecodeName(records, pos));
    names.push_back(std::move(name));
  }
  RETURN_IF_ERROR(TakeEnd(records, pos, pivot_records::kEndPNames));
  if (names.size() != count) {
    return Corrupt("xlsb pivot calculated item name count does not match its names");
  }
  return names;
}

Expected<CalcItem, Error> DecodeItem(const std::vector<RawRecord>& records, std::size_t* pos) {
  ASSIGN_OR_RETURN(const RawRecord* head, Take(records, pos, pivot_records::kBeginPCDCalcItem));
  ByteSpan cursor{head->payload.data(), head->payload.size()};
  CalcItem item;
  ASSIGN_OR_RETURN(item.field, read_u32(cursor));
  if (!ReadCounted(cursor, &item.rgce) || !ReadCounted(cursor, &item.rgcb) || cursor.size != 0U) {
    return Corrupt("xlsb pivot calculated item formula does not fill its record");
  }
  if (At(records, *pos, pivot_records::kBeginPRule)) {
    ASSIGN_OR_RETURN(pivot_area::PivotArea area, pivot_area::decode_pivot_area(records, pos));
    item.area = std::move(area);
  }
  if (At(records, *pos, pivot_records::kBeginPNames)) {
    ASSIGN_OR_RETURN(std::vector<PName> names, DecodeNames(records, pos));
    item.names = std::move(names);
  }
  RETURN_IF_ERROR(TakeEnd(records, pos, pivot_records::kEndPCDCalcItem));
  return item;
}

}  // namespace

Expected<std::vector<CalcItem>, Error> decode_calc_items(const std::vector<RawRecord>& records, std::size_t* pos) {
  ASSIGN_OR_RETURN(const std::uint32_t count, TakeCount(records, pos, pivot_records::kBeginPCDCalcItems));
  std::vector<CalcItem> items;
  while (At(records, *pos, pivot_records::kBeginPCDCalcItem)) {
    auto item = DecodeItem(records, pos);
    if (!item) {
      return std::move(item.error());
    }
    items.push_back(std::move(item.value()));
  }
  RETURN_IF_ERROR(TakeEnd(records, pos, pivot_records::kEndPCDCalcItems));
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
    Push(out, pivot_records::kBeginPCDCalcItem, std::move(head));
    if (item.area) {
      pivot_area::encode_pivot_area(*item.area, out);
    }
    if (item.names) {
      PushCount(out, pivot_records::kBeginPNames, item.names->size());
      for (const PName& name : *item.names) {
        std::vector<std::uint8_t> name_head;
        emit_u32(name_head, name.field);
        emit_u16(name_head, name.flags);
        Push(out, pivot_records::kBeginPName, std::move(name_head));
        PushCount(out, pivot_records::kBeginPNPairs, name.pairs.size());
        for (const PNPair& pair : name.pairs) {
          std::vector<std::uint8_t> payload;
          emit_u32(payload, pair.field);
          emit_u8(payload, pair.flags);
          emit_u32(payload, pair.item);
          Push(out, pivot_records::kBeginPNPair, std::move(payload));
          Push(out, pivot_records::kEndPNPair);
        }
        Push(out, pivot_records::kEndPNPairs);
        Push(out, pivot_records::kEndPName);
      }
      Push(out, pivot_records::kEndPNames);
    }
    Push(out, pivot_records::kEndPCDCalcItem);
  }
  Push(out, pivot_records::kEndPCDCalcItems);
}

Expected<std::vector<FormulaToken>, Error> walk_calc_item_formula(const std::vector<std::uint8_t>& rgce) {
  std::vector<FormulaToken> tokens;
  std::size_t i = 0;
  while (i < rgce.size()) {
    const std::uint8_t ptg = rgce[i];
    if (ptg == kPtgSxName) {
      if (rgce.size() - i < kSxNameSize || rgce[i + 1] != kEptgSxName) {
        return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                          "xlsb pivot calculated item formula carries a malformed PtgSxName",
                          "context=xlsb_pivot_calc_item");
      }
      tokens.push_back(FormulaToken{ptg, load_u32(rgce.data() + i + 2)});
      i += kSxNameSize;
    } else if (ptg >= kFirstOperatorPtg && ptg <= kLastOperatorPtg) {
      tokens.push_back(FormulaToken{ptg, 0U});
      ++i;
    } else {
      return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                        "xlsb pivot calculated item formula carries a token outside the walked set",
                        "context=xlsb_pivot_calc_item ptg=" + std::to_string(ptg));
    }
  }
  return tokens;
}

}  // namespace pivot_calc
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
