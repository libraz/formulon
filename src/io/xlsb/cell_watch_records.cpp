#include "io/xlsb/cell_watch_records.h"

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "utils/a1_ref.h"
#include "utils/structured_log.h"

namespace formulon::io::xlsb {
namespace {

constexpr std::string_view kCellWatches = "cellWatches";
constexpr std::size_t kCellWatchSize = 8U;

bool IsCellWatches(const WorksheetRawChild& child) {
  return child.xml.compare(0, kCellWatches.size() + 1U, "<cellWatches") == 0;
}

/// Records the worksheet grammar puts ahead of the cell-watch block in the post-hyperlink tail.
bool PrecedesCellWatches(std::uint16_t type) {
  switch (type) {
    case kBrtPrintOptions:
    case kBrtMargins:
    case kBrtPageSetup:
    case kBrtBeginHeaderFooter:
    case kBrtEndHeaderFooter:
    case kBrtBeginRwBrk:
    case kBrtBrk:
    case kBrtEndRwBrk:
    case kBrtBeginColBrk:
    case kBrtEndColBrk:
    case kBrtBigName:
      return true;
    default:
      return false;
  }
}

}  // namespace

std::optional<std::string> decode_cell_watch(ByteSpan payload) {
  if (payload.size != kCellWatchSize) {
    return std::nullopt;
  }
  const std::uint32_t row = load_u32(payload.data);
  const std::uint32_t col = load_u32(payload.data + 4);
  if (row >= a1::kMaxRows || col >= a1::kMaxCols) {
    return std::nullopt;
  }
  return a1::encode_a1(row, col);
}

void write_cell_watches(WorksheetRawExtensions& children, const std::vector<std::string>& refs) {
  for (std::size_t i = children.size(); i-- > 0;) {
    if (IsCellWatches(children[i])) {
      children.erase(children.begin() + static_cast<std::ptrdiff_t>(i));
    }
  }
  if (refs.empty()) {
    return;
  }
  std::string xml = "<cellWatches>";
  for (const std::string& ref : refs) {
    xml.append("<cellWatch");
    append_xml_attr(xml, "r", ref);
    xml.append("/>");
  }
  xml.append("</cellWatches>");
  insert_raw_child(
      children, WorksheetRawChild{static_cast<std::uint32_t>(worksheet_child::slot_of(kCellWatches)), std::move(xml)});
}

void splice_cell_watches(std::vector<std::uint8_t>& after_hyperlinks, const WorksheetRawExtensions& children) {
  std::vector<std::uint8_t> block;
  for (const WorksheetRawChild& child : children) {
    pugi::xml_document doc;
    if (!IsCellWatches(child) || !doc.load_buffer(child.xml.data(), child.xml.size())) {
      continue;
    }
    for (const pugi::xml_node watch : doc.child(kCellWatches.data()).children("cellWatch")) {
      std::uint32_t row = 0;
      std::uint32_t col = 0;
      if (!a1::parse_a1_ref(attr_str(watch, "r"), &row, &col)) {
        StructuredLog("xlsb.cell_watch.skipped").field("ref", attr_str(watch, "r")).warn();
        continue;
      }
      std::vector<std::uint8_t> payload;
      emit_u32(payload, row);
      emit_u32(payload, col);
      emit_record(block, kBrtCellWatch, payload);
    }
  }
  if (block.empty()) {
    return;
  }
  const std::vector<FramedRecord> recs = split_records(after_hyperlinks);
  if (!after_hyperlinks.empty() && recs.empty()) {
    return;
  }
  std::size_t insert_at = 0;
  for (const FramedRecord& rec : recs) {
    if (PrecedesCellWatches(rec.type)) {
      insert_at = rec.end;
    }
  }
  std::vector<std::uint8_t> framed;
  emit_record(framed, kBrtBeginCellWatches, ByteSpan{});
  framed.insert(framed.end(), block.begin(), block.end());
  emit_record(framed, kBrtEndCellWatches, ByteSpan{});
  after_hyperlinks.insert(after_hyperlinks.begin() + static_cast<std::ptrdiff_t>(insert_at), framed.begin(),
                          framed.end());
}

}  // namespace formulon::io::xlsb
