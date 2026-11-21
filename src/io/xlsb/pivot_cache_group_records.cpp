#include "io/xlsb/pivot_cache_group_records.h"

#include <utility>

#include "io/xlsb/pivot_record_cursor.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace pivot_group {

namespace {

using pivot_records::RawRecord;

constexpr pivot_records::BlockNames kBlock{"xlsb_pivot_group", "pivot field group"};

// BrtBeginPCDFGroup: par u32, base u32.
constexpr std::size_t kGroupSize = 8U;
// BrtBeginPCDFGRange: groupBy u8, flags u8, start/end/interval doubles.
constexpr std::size_t kRangeSize = 26U;
constexpr std::size_t kCountSize = 4U;
// BrtPCDIIndex inside BrtBeginPCDFGDiscrete: one u32 group item index.
constexpr std::size_t kIndexSize = 4U;
// BrtBeginPCDIRun item type of a string run.
constexpr std::uint16_t kRunStrings = 0x0002U;

Error Corrupt(const char* message) {
  return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, message, "context=xlsb_pivot_group");
}

// The BrtBeginPCDFGItems list: a count, then that many strings, as one string run from
// pivot_records::kMinItemRunLength items and as single BrtPCDIString records below it.
Expected<std::vector<std::string>, Error> DecodeItems(const std::vector<RawRecord>& records, std::size_t* pos) {
  ASSIGN_OR_RETURN(const RawRecord* list, pivot_records::take(kBlock, records, pos, pivot_records::kBeginPCDFGItems));
  if (list->payload.size() != kCountSize) {
    return Corrupt("xlsb pivot group item list has an unexpected size");
  }
  const std::uint32_t count = load_u32(list->payload.data());
  std::vector<std::string> items;
  if (count < pivot_records::kMinItemRunLength) {
    for (std::uint32_t i = 0; i < count; ++i) {
      ASSIGN_OR_RETURN(const RawRecord* single, pivot_records::take(kBlock, records, pos, pivot_records::kPCDIString));
      ByteSpan cursor{single->payload.data(), single->payload.size()};
      ASSIGN_OR_RETURN(std::string text, read_xlwidestring(cursor));
      if (cursor.size != 0U) {
        return Corrupt("xlsb pivot group item has trailing bytes");
      }
      items.push_back(std::move(text));
    }
    RETURN_IF_ERROR(pivot_records::take_end(kBlock, records, pos, pivot_records::kEndPCDFGItems));
    return items;
  }
  ASSIGN_OR_RETURN(const RawRecord* run, pivot_records::take(kBlock, records, pos, pivot_records::kBeginPCDIRun));
  ByteSpan cursor{run->payload.data(), run->payload.size()};
  ASSIGN_OR_RETURN(const std::uint16_t kind, read_u16(cursor));
  if (kind != kRunStrings) {
    return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt,
                      "xlsb pivot group items use a run type whose layout is unmeasured",
                      "context=xlsb_pivot_group run_type=" + std::to_string(kind));
  }
  ASSIGN_OR_RETURN(const std::uint32_t run_count, read_u32(cursor));
  if (run_count != count) {
    return Corrupt("xlsb pivot group item run length does not match its list count");
  }
  items.reserve(run_count);
  for (std::uint32_t i = 0; i < run_count; ++i) {
    ASSIGN_OR_RETURN(std::string text, read_xlwidestring(cursor));
    items.push_back(std::move(text));
  }
  if (cursor.size != 0U) {
    return Corrupt("xlsb pivot group item run has trailing bytes");
  }
  RETURN_IF_ERROR(pivot_records::take_end(kBlock, records, pos, pivot_records::kEndPCDIRun));
  RETURN_IF_ERROR(pivot_records::take_end(kBlock, records, pos, pivot_records::kEndPCDFGItems));
  return items;
}

// The BrtBeginPCDFGDiscrete map: a count, then that many single BrtPCDIIndex records at any count (no run form).
Expected<std::vector<std::uint32_t>, Error> DecodeDiscrete(const std::vector<RawRecord>& records, std::size_t* pos) {
  ASSIGN_OR_RETURN(const RawRecord* head,
                   pivot_records::take(kBlock, records, pos, pivot_records::kBeginPCDFGDiscrete));
  if (head->payload.size() != kCountSize) {
    return Corrupt("xlsb pivot group discrete map has an unexpected size");
  }
  const std::uint32_t count = load_u32(head->payload.data());
  std::vector<std::uint32_t> map;
  for (std::uint32_t i = 0; i < count; ++i) {
    ASSIGN_OR_RETURN(const RawRecord* index, pivot_records::take(kBlock, records, pos, pivot_records::kPCDIIndex));
    if (index->payload.size() != kIndexSize) {
      return Corrupt("xlsb pivot group discrete index has an unexpected size");
    }
    map.push_back(load_u32(index->payload.data()));
  }
  RETURN_IF_ERROR(pivot_records::take_end(kBlock, records, pos, pivot_records::kEndPCDFGDiscrete));
  return map;
}

}  // namespace

Expected<FieldGroup, Error> decode_field_group(const std::vector<RawRecord>& records, std::size_t* pos) {
  ASSIGN_OR_RETURN(const RawRecord* head, pivot_records::take(kBlock, records, pos, pivot_records::kBeginPCDFGroup));
  if (head->payload.size() != kGroupSize) {
    return Corrupt("xlsb pivot field group record has an unexpected size");
  }
  FieldGroup group;
  group.par = load_u32(head->payload.data());
  group.base = load_u32(head->payload.data() + 4);
  if (pivot_records::at(records, *pos, pivot_records::kBeginPCDFGRange)) {
    const RawRecord& rec = records[(*pos)++];
    if (rec.payload.size() != kRangeSize) {
      return Corrupt("xlsb pivot group range record has an unexpected size");
    }
    RangeGroup range;
    range.group_by = rec.payload[0];
    range.flags = rec.payload[1];
    range.start = load_f64(rec.payload.data() + 2);
    range.end = load_f64(rec.payload.data() + 10);
    range.interval = load_f64(rec.payload.data() + 18);
    RETURN_IF_ERROR(pivot_records::take_end(kBlock, records, pos, pivot_records::kEndPCDFGRange));
    group.range = range;
  } else if (pivot_records::at(records, *pos, pivot_records::kBeginPCDFGDiscrete)) {
    ASSIGN_OR_RETURN(std::vector<std::uint32_t> map, DecodeDiscrete(records, pos));
    group.discrete = std::move(map);
  }
  if (pivot_records::at(records, *pos, pivot_records::kBeginPCDFGItems)) {
    ASSIGN_OR_RETURN(std::vector<std::string> items, DecodeItems(records, pos));
    group.items = std::move(items);
  }
  RETURN_IF_ERROR(pivot_records::take_end(kBlock, records, pos, pivot_records::kEndPCDFGroup));
  return group;
}

void encode_field_group(const FieldGroup& group, std::vector<RawRecord>* out) {
  std::vector<std::uint8_t> head;
  emit_u32(head, group.par);
  emit_u32(head, group.base);
  pivot_records::push(out, pivot_records::kBeginPCDFGroup, std::move(head));
  if (group.range) {
    std::vector<std::uint8_t> range;
    emit_u8(range, group.range->group_by);
    emit_u8(range, group.range->flags);
    emit_double(range, group.range->start);
    emit_double(range, group.range->end);
    emit_double(range, group.range->interval);
    pivot_records::push(out, pivot_records::kBeginPCDFGRange, std::move(range));
    pivot_records::push(out, pivot_records::kEndPCDFGRange);
  }
  if (group.discrete) {
    std::vector<std::uint8_t> count;
    emit_u32(count, static_cast<std::uint32_t>(group.discrete->size()));
    pivot_records::push(out, pivot_records::kBeginPCDFGDiscrete, std::move(count));
    for (const std::uint32_t index : *group.discrete) {
      std::vector<std::uint8_t> entry;
      emit_u32(entry, index);
      pivot_records::push(out, pivot_records::kPCDIIndex, std::move(entry));
    }
    pivot_records::push(out, pivot_records::kEndPCDFGDiscrete);
  }
  if (group.items) {
    const auto count = static_cast<std::uint32_t>(group.items->size());
    std::vector<std::uint8_t> list;
    emit_u32(list, count);
    pivot_records::push(out, pivot_records::kBeginPCDFGItems, std::move(list));
    if (count < pivot_records::kMinItemRunLength) {
      for (const std::string& text : *group.items) {
        std::vector<std::uint8_t> single;
        emit_xlwidestring(single, text);
        pivot_records::push(out, pivot_records::kPCDIString, std::move(single));
      }
    } else {
      std::vector<std::uint8_t> run;
      emit_u16(run, kRunStrings);
      emit_u32(run, count);
      for (const std::string& text : *group.items) {
        emit_xlwidestring(run, text);
      }
      pivot_records::push(out, pivot_records::kBeginPCDIRun, std::move(run));
      pivot_records::push(out, pivot_records::kEndPCDIRun);
    }
    pivot_records::push(out, pivot_records::kEndPCDFGItems);
  }
  pivot_records::push(out, pivot_records::kEndPCDFGroup);
}

}  // namespace pivot_group
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
