// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the XLSB shared-string-table builder + emitter.

#include "io/xlsb/sst_writer.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

std::uint32_t SstBuilder::intern(std::string_view text) {
  // Hashing keys directly off the `string_view` requires either a
  // custom hash type or temporarily materialising a `std::string`. We
  // pick the latter for simplicity — interning is O(text-cells) at
  // workbook write time, which is dwarfed by the rest of the writer.
  std::string key(text);
  if (auto it = index_.find(key); it != index_.end()) {
    return it->second;
  }
  const std::uint32_t idx = static_cast<std::uint32_t>(entries_.size());
  entries_.push_back(key);
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

  // BrtBeginSst payload: (u32 cstTotal, u32 cstUnique). We don't
  // separately track multiplicity (the interner has already deduped
  // every text cell to a single index), so cstTotal == cstUnique.
  {
    std::vector<std::uint8_t> p;
    emit_u32(p, sst.size());
    emit_u32(p, sst.size());
    emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginSst), p);
  }

  // One BrtSSTItem per entry: (u8 flags=0, XLWideString).
  for (const std::string& s : sst.entries()) {
    std::vector<std::uint8_t> p;
    emit_u8(p, 0);  // flags: not rich, no phonetic guide.
    emit_xlwidestring(p, s);
    emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtSSTItem), p);
  }

  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtEndSst), ByteSpan{});
  return body;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
