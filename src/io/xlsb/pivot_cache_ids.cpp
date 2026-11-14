#include "io/xlsb/pivot_cache_ids.h"

#include <algorithm>
#include <cstddef>

#include "io/xlsb/frt_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::uint16_t kBeginPivotCacheIds = 384;
constexpr std::uint16_t kEndPivotCacheIds = 385;
constexpr std::uint16_t kBeginPivotCacheId = 386;
constexpr std::uint16_t kEndPivotCacheId = 387;
constexpr std::uint16_t kBrtFileRecover = 155;

}  // namespace

std::vector<std::uint8_t> read_book_pivot_cache_ids(ByteSpan workbook_bin) {
  const std::vector<std::uint8_t> body(workbook_bin.data, workbook_bin.data + workbook_bin.size);
  const std::vector<FramedRecord> recs = split_records(body);
  const auto begin =
      std::find_if(recs.begin(), recs.end(), [](const FramedRecord& r) { return r.type == kBeginPivotCacheIds; });
  const auto end = std::find_if(begin, recs.end(), [](const FramedRecord& r) { return r.type == kEndPivotCacheIds; });
  if (end == recs.end()) {
    return {};
  }
  return std::vector<std::uint8_t>(body.begin() + static_cast<std::ptrdiff_t>(begin->begin),
                                   body.begin() + static_cast<std::ptrdiff_t>(end->end));
}

void splice_book_pivot_cache_ids(std::vector<std::uint8_t>& workbook_bin, const std::vector<std::uint8_t>& retained,
                                 const std::vector<std::pair<std::string, std::string>>& id_map) {
  std::vector<std::uint8_t> block;
  std::size_t ids = 0;
  for (const FramedRecord& rec : split_records(retained)) {
    if (rec.type != kBeginPivotCacheId) {
      if (rec.type != kEndPivotCacheId) {
        append_record(block, retained, rec);
      }
      continue;
    }
    if (rec.payload.size < 4U) {
      continue;
    }
    ByteSpan cursor{rec.payload.data + 4, rec.payload.size - 4U};
    const auto id = read_xlwidestring(cursor);
    const auto hit =
        id ? std::find_if(id_map.begin(), id_map.end(), [&id](const auto& e) { return e.first == id.value(); })
           : id_map.end();
    if (hit == id_map.end()) {
      continue;
    }
    std::vector<std::uint8_t> payload(rec.payload.data, rec.payload.data + 4);
    emit_xlwidestring(payload, hit->second);
    payload.insert(payload.end(), cursor.data, cursor.data + cursor.size);
    emit_record(block, kBeginPivotCacheId, payload);
    emit_record(block, kEndPivotCacheId, ByteSpan{});
    ++ids;
  }
  if (ids == 0U) {
    return;
  }
  const std::vector<FramedRecord> dst = split_records(workbook_bin);
  const auto at = std::find_if(dst.begin(), dst.end(), [](const FramedRecord& r) {
    return r.type == kBrtFileRecover || r.type == kFrtBegin || r.type == kBrtEndBook;
  });
  const std::size_t offset = at == dst.end() ? workbook_bin.size() : at->begin;
  workbook_bin.insert(workbook_bin.begin() + static_cast<std::ptrdiff_t>(offset), block.begin(), block.end());
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
