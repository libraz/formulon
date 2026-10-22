// @size-budget: 4 KB

#include "io/xlsb/slicer_records.h"

#include <algorithm>
#include <array>
#include <cstddef>
#include <optional>

#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::uint16_t kFrtBegin = 35;
constexpr std::uint16_t kFrtEnd = 36;
constexpr std::uint16_t kEndBook = 132;
constexpr std::uint16_t kBeginSlicerCacheIds = 1070;
constexpr std::uint16_t kBeginSlicerCacheId = 1072;
constexpr std::uint16_t kEndSlicerCacheId = 1073;
constexpr std::uint16_t kTableSlicerCacheIds = 2075;
constexpr std::uint16_t kTableSlicerCacheId = 2076;
constexpr std::uint32_t kFrtRelId = 0x08U;

// FRTWORKBOOK blocks in grammar order, by the record that opens each.
constexpr std::array<std::uint16_t, 14> kBookFrtOrder = {
    1036, 1113, kBeginSlicerCacheIds, 1117, 2048, 2051, 2080, 2083, kTableSlicerCacheIds, 2091, 2121, 5095, 5099, 5117};

std::size_t Rank(std::uint16_t first) {
  return static_cast<std::size_t>(std::find(kBookFrtOrder.begin(), kBookFrtOrder.end(), first) - kBookFrtOrder.begin());
}

std::uint32_t LoadU32(const std::uint8_t* p) {
  return static_cast<std::uint32_t>(p[0]) | (static_cast<std::uint32_t>(p[1]) << 8) |
         (static_cast<std::uint32_t>(p[2]) << 16) | (static_cast<std::uint32_t>(p[3]) << 24);
}

/// The relationship id of an FRTHeader carrying only fRelID, and the offset
/// past it.
std::optional<std::string> RelId(ByteSpan p, std::size_t& rest) {
  if (p.size < 6U || LoadU32(p.data) != kFrtRelId) {
    return std::nullopt;
  }
  const std::size_t cch = static_cast<std::size_t>(p.data[4] | (p.data[5] << 8U));
  if (p.size - 6U < 2U * cch) {
    return std::nullopt;
  }
  std::string id;
  for (std::size_t i = 0; i < cch; ++i) {
    const auto unit = static_cast<std::uint16_t>(p.data[6 + 2 * i] | (p.data[7 + 2 * i] << 8U));
    if (unit == 0U || unit > 0x7FU) {
      return std::nullopt;
    }
    id.push_back(static_cast<char>(unit));
  }
  rest = 6U + 2U * cch;
  return id;
}

std::vector<std::uint8_t> WithRelId(ByteSpan p, std::size_t rest, const std::string& id) {
  std::vector<std::uint8_t> out;
  emit_u32(out, kFrtRelId);
  emit_u16(out, static_cast<std::uint16_t>(id.size()));
  for (const char c : id) {
    emit_u16(out, static_cast<std::uint16_t>(static_cast<unsigned char>(c)));
  }
  out.insert(out.end(), p.data + rest, p.data + p.size);
  return out;
}

/// Index of the BrtFRTEnd closing the block opened at `i`, or `recs.size()`.
std::size_t BlockEnd(const std::vector<FramedRecord>& recs, std::size_t i) {
  for (std::size_t k = i + 1U; k < recs.size(); ++k) {
    if (recs[k].type == kFrtBegin) {
      return recs.size();
    }
    if (recs[k].type == kFrtEnd) {
      return k;
    }
  }
  return recs.size();
}

}  // namespace

std::vector<std::uint8_t> read_book_slicer_cache_ids(ByteSpan workbook_bin) {
  const std::vector<std::uint8_t> body(workbook_bin.data, workbook_bin.data + workbook_bin.size);
  const std::vector<FramedRecord> recs = split_records(body);
  std::vector<std::uint8_t> out;
  for (std::size_t i = 0; i + 1U < recs.size(); ++i) {
    if (recs[i].type != kFrtBegin ||
        (recs[i + 1U].type != kBeginSlicerCacheIds && recs[i + 1U].type != kTableSlicerCacheIds)) {
      continue;
    }
    const std::size_t end = BlockEnd(recs, i);
    if (end == recs.size()) {
      continue;
    }
    out.insert(out.end(), body.begin() + static_cast<std::ptrdiff_t>(recs[i].begin),
               body.begin() + static_cast<std::ptrdiff_t>(recs[end].end));
    i = end;
  }
  return out;
}

void splice_book_slicer_cache_ids(std::vector<std::uint8_t>& workbook_bin, const std::vector<std::uint8_t>& retained,
                                  const std::vector<std::pair<std::string, std::string>>& id_map) {
  const std::vector<FramedRecord> src = split_records(retained);
  for (std::size_t i = 0; i + 1U < src.size(); ++i) {
    if (src[i].type != kFrtBegin) {
      continue;
    }
    const std::size_t end = BlockEnd(src, i);
    if (end == src.size()) {
      continue;
    }
    const std::uint16_t first = src[i + 1U].type;
    std::vector<std::uint8_t> block;
    std::size_t ids = 0;
    for (std::size_t k = i; k <= end; ++k) {
      const FramedRecord& rec = src[k];
      if (rec.type != kBeginSlicerCacheId && rec.type != kTableSlicerCacheId) {
        append_record(block, retained, rec);
        continue;
      }
      std::size_t rest = 0;
      const std::optional<std::string> id = RelId(rec.payload, rest);
      const auto hit = id ? std::find_if(id_map.begin(), id_map.end(), [&id](const auto& e) { return e.first == *id; })
                          : id_map.end();
      if (hit == id_map.end()) {
        if (rec.type == kBeginSlicerCacheId && k + 1U <= end && src[k + 1U].type == kEndSlicerCacheId) {
          ++k;
        }
        continue;
      }
      emit_record(block, rec.type, WithRelId(rec.payload, rest, hit->second));
      ++ids;
    }
    i = end;
    if (ids == 0U) {
      continue;
    }
    const std::vector<FramedRecord> dst = split_records(workbook_bin);
    std::size_t at = dst.empty() || dst.back().type != kEndBook ? workbook_bin.size() : dst.back().begin;
    for (std::size_t k = 0; k + 1U < dst.size(); ++k) {
      const std::size_t rank = Rank(dst[k + 1U].type);
      if (dst[k].type == kFrtBegin && rank < kBookFrtOrder.size() && rank > Rank(first)) {
        at = dst[k].begin;
        break;
      }
    }
    workbook_bin.insert(workbook_bin.begin() + static_cast<std::ptrdiff_t>(at), block.begin(), block.end());
  }
}

bool is_xml_slicer_part(std::string_view path) {
  const auto under = [path](std::string_view dir) {
    return path.size() > dir.size() + 4U && path.compare(0, dir.size(), dir) == 0 &&
           path.compare(path.size() - 4U, 4U, ".xml") == 0;
  };
  return under("xl/slicers/") || under("xl/slicerCaches/");
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
