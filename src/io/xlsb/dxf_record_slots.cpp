#include "io/xlsb/dxf_record_slots.h"

#include <array>
#include <initializer_list>
#include <string>

#include "io/xlsb/dxf_props.h"
#include "io/xlsb/frt_formula.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"

namespace formulon::io::xlsb {
namespace {

/// Slot offsets within one record's payload.
struct SlotOffsets {
  std::array<std::size_t, 6> at{};
  std::size_t count = 0;
};

using Locator = bool (*)(ByteSpan payload, SlotOffsets* out);

bool Fixed(ByteSpan payload, std::initializer_list<std::size_t> offsets, SlotOffsets* out) {
  for (const std::size_t offset : offsets) {
    if (payload.size < offset + 4U) {
      return false;
    }
    out->at[out->count++] = offset;
  }
  return true;
}

/// RfX, list type, id, header and totals row counts, flags, then the six dxf ids.
bool BeginList(ByteSpan payload, SlotOffsets* out) {
  return Fixed(payload, {36, 40, 44, 48, 52, 56}, out);
}

/// Field id, totals function, then the header, data and totals dxf ids.
bool BeginListCol(ByteSpan payload, SlotOffsets* out) {
  return Fixed(payload, {8, 12, 16}, out);
}

/// iType, iTemplate, then dxfId.
bool BeginCfRule(ByteSpan payload, SlotOffsets* out) {
  return Fixed(payload, {8}, out);
}

/// FRTHeader (with an optional FRTFormula list), iType, iTemplate, then dxfId.
bool BeginCfRule14(ByteSpan payload, SlotOffsets* out) {
  ByteSpan p = payload;
  auto header = read_u32(p);
  if (!header || (header.value() != 0U && header.value() != kFrtHeaderFormulas)) {
    return false;
  }
  if (header.value() == kFrtHeaderFormulas) {
    auto count = read_u32(p);
    if (!count) {
      return false;
    }
    for (std::uint32_t i = 0; i < count.value(); ++i) {
      if (p.size < 12U) {
        return false;
      }
      const std::size_t cce = load_u32(p.data + 4);
      const std::size_t cb = load_u32(p.data + 8);
      if (p.size - 12U < cce || p.size - 12U - cce < cb) {
        return false;
      }
      const std::size_t size = 12U + cce + cb;
      p.data += size;
      p.size -= size;
    }
  }
  const std::size_t offset = payload.size - p.size + 8U;
  return Fixed(payload, {offset}, out);
}

/// dxfId first.
bool ColorFilter(ByteSpan payload, SlotOffsets* out) {
  return Fixed(payload, {0}, out);
}

/// tseType, size, then the dxf index.
bool TableStyleElement(ByteSpan payload, SlotOffsets* out) {
  return Fixed(payload, {8}, out);
}

/// Flags, RfX, then a dxf id only when sortOn is cell or font colour.
bool BeginSortCond(ByteSpan payload, SlotOffsets* out) {
  constexpr std::uint32_t kSortOnCellColor = 1U;
  constexpr std::uint32_t kSortOnFontColor = 2U;
  constexpr std::uint32_t kSortOnIcon = 3U;
  if (payload.size < 2U) {
    return false;
  }
  const std::uint32_t sort_on = (load_u16(payload.data) >> 1U) & 0x0FU;
  if (sort_on > kSortOnIcon) {
    return false;
  }
  if (sort_on != kSortOnCellColor && sort_on != kSortOnFontColor) {
    return true;
  }
  return Fixed(payload, {18}, out);
}

struct SlotRow {
  std::uint16_t type;
  Locator locate;
};

constexpr SlotRow kSlotRows[] = {
    {kBrtBeginList, BeginList},
    {kBrtBeginListCol, BeginListCol},
    {kBrtBeginCFRule, BeginCfRule},
    {kBeginCfRule14, BeginCfRule14},
    {kBrtColorFilter, ColorFilter},
    {kBrtBeginSortCond, BeginSortCond},
    {kBrtTableStyleElement, TableStyleElement},
};

// Dxf-bearing records without a decoder.
constexpr std::array<std::uint16_t, 0> kUndecodableRecords{};

Locator FindLocator(std::uint16_t type) {
  for (const SlotRow& row : kSlotRows) {
    if (row.type == type) {
      return row.locate;
    }
  }
  return nullptr;
}

Error Corrupt(const char* message, std::uint16_t type) {
  return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, message,
                    "context=xlsb_dxf_slots record_id=" + std::to_string(type));
}

}  // namespace

bool is_undecodable_dxf_record(std::uint16_t type) {
  for (const std::uint16_t undecodable : kUndecodableRecords) {
    if (undecodable == type) {
      return true;
    }
  }
  return false;
}

Expected<DxfSlotScan, Error> scan_dxf_slots(const std::vector<std::uint8_t>& stream) {
  const std::vector<FramedRecord> recs = split_records(stream);
  if (recs.empty() && !stream.empty()) {
    return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb stream does not frame whole",
                      "context=xlsb_dxf_slots");
  }
  DxfSlotScan scan;
  for (std::size_t i = 0; i < recs.size(); ++i) {
    const FramedRecord& rec = recs[i];
    if (rec.type == kBrtDxf) {
      ++scan.dxf_records;
      continue;
    }
    if (is_undecodable_dxf_record(rec.type)) {
      scan.undecodable.push_back(rec.type);
      continue;
    }
    const Locator locate = FindLocator(rec.type);
    if (locate == nullptr) {
      continue;
    }
    SlotOffsets offsets;
    if (!locate(rec.payload, &offsets)) {
      return Corrupt("xlsb record's dxf slot cannot be located", rec.type);
    }
    const std::size_t base = static_cast<std::size_t>(rec.payload.data - stream.data());
    for (std::size_t s = 0; s < offsets.count; ++s) {
      const std::uint32_t value = load_u32(rec.payload.data + offsets.at[s]);
      if (value != kNoDxf) {
        scan.slots.push_back(DxfRecordSlot{rec.type, i, s, base + offsets.at[s], value});
      }
    }
  }
  return scan;
}

Expected<std::vector<std::uint8_t>, Error> rewrite_dxf_slots(const std::vector<std::uint8_t>& stream,
                                                             const DxfSavePlan& plan) {
  auto scan = scan_dxf_slots(stream);
  if (!scan) {
    return std::move(scan.error());
  }
  std::vector<std::uint8_t> out = stream;
  for (const DxfRecordSlot& slot : scan.value().slots) {
    store_u32(out.data() + slot.offset, plan.map(slot.value).value_or(kNoDxf));
  }
  return out;
}

}  // namespace formulon::io::xlsb
