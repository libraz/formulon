#include "io/xlsb/pivot_area_records.h"

#include <utility>

#include "io/xlsb/pivot_record_cursor.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace pivot_area {

namespace {

using pivot_records::RawRecord;

constexpr pivot_records::BlockNames kBlock{"xlsb_pivot_area", "pivot area"};

// BrtBeginPRule: isxvd u32, type u8, 24-bit flags.
constexpr std::size_t kRuleSize = 8U;
// BrtBeginPRFilter: isxvd u32, item count u32, 24-bit flags.
constexpr std::size_t kFilterSize = 11U;
constexpr std::size_t kCountSize = 4U;

Error Corrupt(const char* message) {
  return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, message, "context=xlsb_pivot_area");
}

std::uint32_t Load24(const std::uint8_t* p) {
  return static_cast<std::uint32_t>(p[0]) | (static_cast<std::uint32_t>(p[1]) << 8) |
         (static_cast<std::uint32_t>(p[2]) << 16);
}

void Emit24(std::vector<std::uint8_t>& dst, std::uint32_t v) {
  emit_u8(dst, static_cast<std::uint8_t>(v));
  emit_u8(dst, static_cast<std::uint8_t>(v >> 8));
  emit_u8(dst, static_cast<std::uint8_t>(v >> 16));
}

Expected<PivotAreaFilter, Error> DecodeFilter(const std::vector<RawRecord>& records, std::size_t* pos) {
  ASSIGN_OR_RETURN(const RawRecord* begin, pivot_records::take(kBlock, records, pos, pivot_records::kBeginPRFilter));
  if (begin->payload.size() != kFilterSize) {
    return Corrupt("xlsb pivot area filter record has an unexpected size");
  }
  PivotAreaFilter filter;
  filter.field = load_u32(begin->payload.data());
  const std::uint32_t count = load_u32(begin->payload.data() + 4);
  filter.flags = Load24(begin->payload.data() + 8);
  while (pivot_records::at(records, *pos, pivot_records::kBeginPRFItem)) {
    const RawRecord& item = records[(*pos)++];
    if (item.payload.size() != kCountSize) {
      return Corrupt("xlsb pivot area filter item has an unexpected size");
    }
    filter.items.push_back(load_u32(item.payload.data()));
    RETURN_IF_ERROR(pivot_records::take_end(kBlock, records, pos, pivot_records::kEndPRFItem));
  }
  RETURN_IF_ERROR(pivot_records::take_end(kBlock, records, pos, pivot_records::kEndPRFilter));
  if (filter.items.size() != count) {
    return Corrupt("xlsb pivot area filter item count does not match its items");
  }
  return filter;
}

}  // namespace

Expected<PivotArea, Error> decode_pivot_area(const std::vector<RawRecord>& records, std::size_t* pos) {
  ASSIGN_OR_RETURN(const RawRecord* rule, pivot_records::take(kBlock, records, pos, pivot_records::kBeginPRule));
  if (rule->payload.size() != kRuleSize) {
    return Corrupt("xlsb pivot area rule record has an unexpected size");
  }
  PivotArea area;
  area.field = load_u32(rule->payload.data());
  area.type = rule->payload[4];
  area.flags = Load24(rule->payload.data() + 5);
  if (pivot_records::at(records, *pos, pivot_records::kBeginPRFilters)) {
    const RawRecord& list = records[(*pos)++];
    if (list.payload.size() != kCountSize) {
      return Corrupt("xlsb pivot area filter list has an unexpected size");
    }
    const std::uint32_t count = load_u32(list.payload.data());
    std::vector<PivotAreaFilter> filters;
    while (pivot_records::at(records, *pos, pivot_records::kBeginPRFilter)) {
      ASSIGN_OR_RETURN(PivotAreaFilter filter, DecodeFilter(records, pos));
      filters.push_back(std::move(filter));
    }
    RETURN_IF_ERROR(pivot_records::take_end(kBlock, records, pos, pivot_records::kEndPRFilters));
    if (filters.size() != count) {
      return Corrupt("xlsb pivot area filter count does not match its filters");
    }
    area.filters = std::move(filters);
  }
  RETURN_IF_ERROR(pivot_records::take_end(kBlock, records, pos, pivot_records::kEndPRule));
  return area;
}

void encode_pivot_area(const PivotArea& area, std::vector<RawRecord>* out) {
  std::vector<std::uint8_t> rule;
  emit_u32(rule, area.field);
  emit_u8(rule, area.type);
  Emit24(rule, area.flags);
  pivot_records::push(out, pivot_records::kBeginPRule, std::move(rule));
  if (area.filters) {
    std::vector<std::uint8_t> list;
    emit_u32(list, static_cast<std::uint32_t>(area.filters->size()));
    pivot_records::push(out, pivot_records::kBeginPRFilters, std::move(list));
    for (const PivotAreaFilter& filter : *area.filters) {
      std::vector<std::uint8_t> head;
      emit_u32(head, filter.field);
      emit_u32(head, static_cast<std::uint32_t>(filter.items.size()));
      Emit24(head, filter.flags);
      pivot_records::push(out, pivot_records::kBeginPRFilter, std::move(head));
      for (const std::uint32_t item : filter.items) {
        std::vector<std::uint8_t> index;
        emit_u32(index, item);
        pivot_records::push(out, pivot_records::kBeginPRFItem, std::move(index));
        pivot_records::push(out, pivot_records::kEndPRFItem);
      }
      pivot_records::push(out, pivot_records::kEndPRFilter);
    }
    pivot_records::push(out, pivot_records::kEndPRFilters);
  }
  pivot_records::push(out, pivot_records::kEndPRule);
}

}  // namespace pivot_area
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
