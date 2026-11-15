#include "io/xlsb/pivot_record_cursor.h"

#include <string>
#include <utility>

#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace pivot_records {

bool at(const std::vector<RawRecord>& records, std::size_t pos, std::uint16_t type) {
  return pos < records.size() && records[pos].type == type;
}

Expected<const RawRecord*, Error> take(const BlockNames& block, const std::vector<RawRecord>& records, std::size_t* pos,
                                       std::uint16_t type) {
  const std::string context = std::string("context=") + block.context;
  if (*pos >= records.size()) {
    return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt,
                      std::string("xlsb ") + block.noun + " block ends before its end record", context);
  }
  const RawRecord& rec = records[*pos];
  if (rec.type != type) {
    return make_error(FormulonErrorCode::kIoXlsbUnknownRecord,
                      std::string("xlsb ") + block.noun + " carries a record outside its grammar",
                      context + " record_id=" + std::to_string(rec.type));
  }
  ++*pos;
  return &rec;
}

Expected<void, Error> take_end(const BlockNames& block, const std::vector<RawRecord>& records, std::size_t* pos,
                               std::uint16_t type) {
  ASSIGN_OR_RETURN(const RawRecord* rec, take(block, records, pos, type));
  if (!rec->payload.empty()) {
    return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt,
                      std::string("xlsb ") + block.noun + " end record carries a payload",
                      std::string("context=") + block.context);
  }
  return Expected<void, Error>::Ok();
}

void push(std::vector<RawRecord>* out, std::uint16_t type, std::vector<std::uint8_t> payload) {
  out->push_back(RawRecord{type, std::move(payload)});
}

}  // namespace pivot_records
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
