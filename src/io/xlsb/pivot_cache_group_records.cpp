#include "io/xlsb/pivot_cache_group_records.h"

#include <utility>

#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace pivot_group {

namespace {

using pivot_records::RawRecord;

// BrtBeginPCDFGroup: par u32, base u32.
constexpr std::size_t kGroupSize = 8U;
// BrtBeginPCDFGRange: groupBy u8, flags u8, start/end/interval doubles.
constexpr std::size_t kRangeSize = 26U;
constexpr std::size_t kCountSize = 4U;
// BrtBeginPCDIRun item type of a string run.
constexpr std::uint16_t kRunStrings = 0x0002U;

Error Corrupt(const char* message) {
  return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, message, "context=xlsb_pivot_group");
}

Error Unexpected(std::uint16_t type) {
  return make_error(FormulonErrorCode::kIoXlsbUnknownRecord,
                    "xlsb pivot field group carries a record outside its grammar",
                    "context=xlsb_pivot_group record_id=" + std::to_string(type));
}

bool At(const std::vector<RawRecord>& records, std::size_t pos, std::uint16_t type) {
  return pos < records.size() && records[pos].type == type;
}

Expected<const RawRecord*, Error> Take(const std::vector<RawRecord>& records, std::size_t* pos, std::uint16_t type) {
  if (*pos >= records.size()) {
    return Corrupt("xlsb pivot field group block ends before its end record");
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
    return Corrupt("xlsb pivot field group end record carries a payload");
  }
  return Expected<void, Error>::Ok();
}

void Push(std::vector<RawRecord>* out, std::uint16_t type, std::vector<std::uint8_t> payload = {}) {
  out->push_back(RawRecord{type, std::move(payload)});
}

// The BrtBeginPCDFGItems list: a count, then one string run of that length.
Expected<std::vector<std::string>, Error> DecodeItems(const std::vector<RawRecord>& records, std::size_t* pos) {
  ASSIGN_OR_RETURN(const RawRecord* list, Take(records, pos, pivot_records::kBeginPCDFGItems));
  if (list->payload.size() != kCountSize) {
    return Corrupt("xlsb pivot group item list has an unexpected size");
  }
  const std::uint32_t count = load_u32(list->payload.data());
  ASSIGN_OR_RETURN(const RawRecord* run, Take(records, pos, pivot_records::kBeginPCDIRun));
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
  std::vector<std::string> items;
  items.reserve(run_count);
  for (std::uint32_t i = 0; i < run_count; ++i) {
    ASSIGN_OR_RETURN(std::string text, read_xlwidestring(cursor));
    items.push_back(std::move(text));
  }
  if (cursor.size != 0U) {
    return Corrupt("xlsb pivot group item run has trailing bytes");
  }
  RETURN_IF_ERROR(TakeEnd(records, pos, pivot_records::kEndPCDIRun));
  RETURN_IF_ERROR(TakeEnd(records, pos, pivot_records::kEndPCDFGItems));
  return items;
}

}  // namespace

Expected<FieldGroup, Error> decode_field_group(const std::vector<RawRecord>& records, std::size_t* pos) {
  ASSIGN_OR_RETURN(const RawRecord* head, Take(records, pos, pivot_records::kBeginPCDFGroup));
  if (head->payload.size() != kGroupSize) {
    return Corrupt("xlsb pivot field group record has an unexpected size");
  }
  FieldGroup group;
  group.par = load_u32(head->payload.data());
  group.base = load_u32(head->payload.data() + 4);
  if (At(records, *pos, pivot_records::kBeginPCDFGRange)) {
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
    RETURN_IF_ERROR(TakeEnd(records, pos, pivot_records::kEndPCDFGRange));
    group.range = range;
  }
  if (At(records, *pos, pivot_records::kBeginPCDFGItems)) {
    ASSIGN_OR_RETURN(std::vector<std::string> items, DecodeItems(records, pos));
    group.items = std::move(items);
  }
  RETURN_IF_ERROR(TakeEnd(records, pos, pivot_records::kEndPCDFGroup));
  return group;
}

void encode_field_group(const FieldGroup& group, std::vector<RawRecord>* out) {
  std::vector<std::uint8_t> head;
  emit_u32(head, group.par);
  emit_u32(head, group.base);
  Push(out, pivot_records::kBeginPCDFGroup, std::move(head));
  if (group.range) {
    std::vector<std::uint8_t> range;
    emit_u8(range, group.range->group_by);
    emit_u8(range, group.range->flags);
    emit_double(range, group.range->start);
    emit_double(range, group.range->end);
    emit_double(range, group.range->interval);
    Push(out, pivot_records::kBeginPCDFGRange, std::move(range));
    Push(out, pivot_records::kEndPCDFGRange);
  }
  if (group.items) {
    const auto count = static_cast<std::uint32_t>(group.items->size());
    std::vector<std::uint8_t> list;
    emit_u32(list, count);
    Push(out, pivot_records::kBeginPCDFGItems, std::move(list));
    std::vector<std::uint8_t> run;
    emit_u16(run, kRunStrings);
    emit_u32(run, count);
    for (const std::string& text : *group.items) {
      emit_xlwidestring(run, text);
    }
    Push(out, pivot_records::kBeginPCDIRun, std::move(run));
    Push(out, pivot_records::kEndPCDIRun);
    Push(out, pivot_records::kEndPCDFGItems);
  }
  Push(out, pivot_records::kEndPCDFGroup);
}

}  // namespace pivot_group
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
