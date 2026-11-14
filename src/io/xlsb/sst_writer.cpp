//
// Implementation of the XLSB shared-string-table builder + emitter.

#include "io/xlsb/sst_writer.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/ooxml/shared_strings_writer.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/rich_str.h"
#include "phonetic.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

std::uint32_t SstBuilder::intern(std::string_view text, const std::vector<PhoneticRun>& phonetic,
                                 PhoneticProperties phonetic_props) {
  ++total_count_;
  std::string key = SharedStringKey(text, phonetic, phonetic_props);
  if (auto it = index_.find(key); it != index_.end()) {
    return it->second;
  }
  const std::uint32_t idx = static_cast<std::uint32_t>(entries_.size());
  entries_.push_back(SstEntry{std::string(text), phonetic, phonetic_props});
  index_.emplace(std::move(key), idx);
  return idx;
}

Expected<std::vector<std::uint8_t>, Error> emit_sst(const SstBuilder& sst) {
  std::vector<std::uint8_t> body;
  // Reserve a generous lower-bound: 16 bytes for the framing
  // (begin + end records) plus 8 for the begin payload, plus a
  // per-item average of ~8 bytes of framing on top of the string
  // body. Resizing is amortised; this just prevents pathological
  // small reallocations on workbooks with many strings.
  body.reserve(32U + sst.size() * 16U);

  // BrtBeginSst payload: (u32 cstTotal, u32 cstUnique), cell references then distinct strings.
  {
    std::vector<std::uint8_t> p;
    emit_u32(p, sst.total_count());
    emit_u32(p, sst.size());
    emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginSst), p);
  }

  // One BrtSSTItem per entry: (u8 flags, XLWideString[, phonetic tail]).
  std::vector<std::uint8_t> p;
  for (const SstEntry& entry : sst.entries()) {
    p.clear();
    emit_rich_str(p, entry.text, {}, entry.phonetic, entry.phonetic_props);
    emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtSSTItem), p);
  }

  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtEndSst), ByteSpan{});
  return body;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
