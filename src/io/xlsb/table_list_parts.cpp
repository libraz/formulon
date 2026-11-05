//
// Worksheet-tail LISTPARTS handling for XLSB tables: finding the block,
// splicing in BrtListPart records for newly written tables and releasing the
// ones a model table replaces.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/sheet_reader.h"
#include "io/xlsb/table_records.h"
#include "io/xlsb/table_records_internal.h"
#include "sheet.h"
#include "unknown_relationship.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using detail::FindListParts;
using detail::ListPartsBlock;

enum : std::uint16_t {
  kBrtBeginListParts = 660,
  kBrtListPart = 661,
  kBrtEndListParts = 662,
};

/// Records the worksheet grammar places after the hyperlinks and before
/// LISTPARTS.
bool PrecedesListParts(std::uint16_t type) {
  return is_after_hyperlinks(type) && type != kBrtBeginListParts && type != kBrtListPart && type != kBrtEndListParts;
}

std::vector<std::uint8_t> ListPartRecord(const std::string& rel_id) {
  std::vector<std::uint8_t> payload;
  emit_xlnullablewidestring(payload, std::string_view(rel_id));
  std::vector<std::uint8_t> out;
  emit_record(out, kBrtListPart, payload);
  return out;
}

}  // namespace

namespace detail {

std::optional<ListPartsBlock> FindListParts(const std::vector<FramedRecord>& recs) {
  for (std::size_t i = 0; i < recs.size(); ++i) {
    if (recs[i].type != kBrtBeginListParts) {
      continue;
    }
    ListPartsBlock block;
    block.begin = i;
    for (std::size_t j = i + 1U; j < recs.size(); ++j) {
      if (recs[j].type == kBrtEndListParts) {
        block.end = j;
        return block;
      }
      if (recs[j].type != kBrtListPart) {
        return std::nullopt;
      }
      ByteSpan payload = recs[j].payload;
      auto id = read_xlnullablewidestring(payload);
      if (!id || id.value().empty()) {
        return std::nullopt;
      }
      block.rel_ids.push_back(std::move(id.value()));
    }
    return std::nullopt;
  }
  return std::nullopt;
}

/// Removes the BrtListPart records naming `rel_ids` from the sheet's LISTPARTS
/// block (the whole block once none is left) and the relationships behind
/// them from the sheet.
void ReleaseSheetTables(Sheet& sheet, const std::vector<std::string>& rel_ids) {
  auto released = [&rel_ids](const std::string& id) {
    return std::find(rel_ids.begin(), rel_ids.end(), id) != rel_ids.end();
  };
  XlsbSheetTail tail = sheet.xlsb_tail();
  const std::vector<FramedRecord> recs = split_records(tail.after_hyperlinks);
  if (const std::optional<ListPartsBlock> block = FindListParts(recs)) {
    std::vector<std::string> kept_ids;
    for (const std::string& id : block->rel_ids) {
      if (!released(id)) {
        kept_ids.push_back(id);
      }
    }
    std::vector<std::uint8_t> kept;
    kept.reserve(tail.after_hyperlinks.size());
    for (std::size_t i = 0; i < recs.size(); ++i) {
      if (i < block->begin || i > block->end) {
        append_record(kept, tail.after_hyperlinks, recs[i]);
      } else if (i == block->begin && !kept_ids.empty()) {
        std::vector<std::uint8_t> payload;
        emit_u32(payload, static_cast<std::uint32_t>(kept_ids.size()));
        emit_record(kept, kBrtBeginListParts, payload);
        for (const std::string& id : kept_ids) {
          const std::vector<std::uint8_t> rec = ListPartRecord(id);
          kept.insert(kept.end(), rec.begin(), rec.end());
        }
        emit_record(kept, kBrtEndListParts, ByteSpan{});
      }
    }
    tail.after_hyperlinks = std::move(kept);
    sheet.set_xlsb_tail(std::move(tail));
  }
  std::vector<UnknownRelationship> rels = sheet.unknown_relationships();
  rels.erase(std::remove_if(rels.begin(), rels.end(),
                            [&released](const UnknownRelationship& rel) { return released(rel.id); }),
             rels.end());
  sheet.set_unknown_relationships(std::move(rels));
}

}  // namespace detail

void splice_list_parts(std::vector<std::uint8_t>& after_hyperlinks, const std::vector<std::string>& rel_ids) {
  if (rel_ids.empty()) {
    return;
  }
  std::vector<std::uint8_t> parts;
  for (const std::string& rel_id : rel_ids) {
    const std::vector<std::uint8_t> rec = ListPartRecord(rel_id);
    parts.insert(parts.end(), rec.begin(), rec.end());
  }
  const std::vector<FramedRecord> recs = split_records(after_hyperlinks);
  std::vector<std::uint8_t> out;
  out.reserve(after_hyperlinks.size() + parts.size() + 16U);
  if (const std::optional<ListPartsBlock> block = FindListParts(recs)) {
    // A retained block keeps its own parts; ours join it before its end.
    for (std::size_t i = 0; i < recs.size(); ++i) {
      if (i == block->begin) {
        std::vector<std::uint8_t> payload;
        emit_u32(payload, static_cast<std::uint32_t>(block->rel_ids.size() + rel_ids.size()));
        emit_record(out, kBrtBeginListParts, payload);
        continue;
      }
      if (i == block->end) {
        out.insert(out.end(), parts.begin(), parts.end());
      }
      append_record(out, after_hyperlinks, recs[i]);
    }
    after_hyperlinks = std::move(out);
    return;
  }
  std::size_t insert_at = 0;
  for (std::size_t i = 0; i < recs.size(); ++i) {
    if (PrecedesListParts(recs[i].type)) {
      insert_at = i + 1U;
    }
  }
  std::vector<std::uint8_t> block;
  std::vector<std::uint8_t> payload;
  emit_u32(payload, static_cast<std::uint32_t>(rel_ids.size()));
  emit_record(block, kBrtBeginListParts, payload);
  block.insert(block.end(), parts.begin(), parts.end());
  emit_record(block, kBrtEndListParts, ByteSpan{});
  if (recs.empty()) {
    // Empty, or not whole records: append rather than split it.
    after_hyperlinks.insert(after_hyperlinks.end(), block.begin(), block.end());
    return;
  }
  for (std::size_t i = 0; i < recs.size(); ++i) {
    if (i == insert_at) {
      out.insert(out.end(), block.begin(), block.end());
    }
    append_record(out, after_hyperlinks, recs[i]);
  }
  if (insert_at == recs.size()) {
    out.insert(out.end(), block.begin(), block.end());
  }
  after_hyperlinks = std::move(out);
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
