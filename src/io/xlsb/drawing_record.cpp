//
// Implementation of the worksheet-tail BrtDrawing helpers. See
// `io/xlsb/drawing_record.h` for the contract.

#include "io/xlsb/drawing_record.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "chart/chart_ns.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/sheet_reader.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// Records the worksheet grammar places after the hyperlinks and before
/// BrtDrawing, or before BrtLegacyDrawing when `placed` is that.
bool PrecedesDrawing(std::uint16_t type, std::uint16_t placed) {
  if (!is_after_hyperlinks(type)) {
    return false;
  }
  switch (type) {
    case kBrtDrawing:
      return placed == kBrtLegacyDrawing;
    case kBrtLegacyDrawing:
    case kBrtLegacyDrawingHF:
    case kBrtBkhim:
    case kBrtBeginOleObjects:
    case kBrtOleObject:
    case kBrtEndOleObjects:
    case kBrtBeginActiveXControls:
    case kBrtActiveX:
    case kBrtEndActiveXControls:
    case kBrtBeginWebPubItems:
    case kBrtBeginWebPubItem:
    case kBrtEndWebPubItem:
    case kBrtEndWebPubItems:
    case kBrtBeginTableParts:
    case kBrtTablePart:
    case kBrtEndTableParts:
      return false;
    default:
      return true;
  }
}

/// True when `buf` is empty or splits whole into `recs`.
bool IsSplit(const std::vector<std::uint8_t>& buf, const std::vector<FramedRecord>& recs) {
  return buf.empty() || !recs.empty();
}

}  // namespace

DrawingRecordScan scan_drawing_records(const std::vector<std::uint8_t>& after_hyperlinks, std::uint16_t type) {
  DrawingRecordScan scan;
  const std::vector<FramedRecord> recs = split_records(after_hyperlinks);
  if (!IsSplit(after_hyperlinks, recs)) {
    scan.parsed = false;
    return scan;
  }
  ByteSpan payload{};
  for (const FramedRecord& rec : recs) {
    if (rec.type == type) {
      ++scan.count;
      payload = rec.payload;
    }
  }
  if (scan.count != 1U) {
    return scan;
  }
  // A null or overlong length fails the read; trailing bytes are an unknown layout.
  auto rel_id = read_xlwidestring(payload);
  if (rel_id && payload.size == 0U) {
    scan.rel_id = std::move(rel_id.value());
  }
  return scan;
}

bool strip_drawing_records(std::vector<std::uint8_t>& after_hyperlinks, std::uint16_t type) {
  const std::vector<FramedRecord> recs = split_records(after_hyperlinks);
  if (!IsSplit(after_hyperlinks, recs)) {
    return false;
  }
  std::vector<std::uint8_t> kept;
  kept.reserve(after_hyperlinks.size());
  for (const FramedRecord& rec : recs) {
    if (rec.type != type) {
      append_record(kept, after_hyperlinks, rec);
    }
  }
  after_hyperlinks = std::move(kept);
  return true;
}

namespace {

bool SpliceDrawingRecord(std::vector<std::uint8_t>& after_hyperlinks, std::string_view rel_id, std::uint16_t type) {
  const std::vector<FramedRecord> recs = split_records(after_hyperlinks);
  if (!IsSplit(after_hyperlinks, recs)) {
    return false;
  }
  std::vector<std::uint8_t> payload;
  emit_xlwidestring(payload, rel_id);
  std::vector<std::uint8_t> record;
  emit_record(record, type, payload);
  std::size_t insert_at = 0;
  for (std::size_t i = 0; i < recs.size(); ++i) {
    if (PrecedesDrawing(recs[i].type, type)) {
      insert_at = i + 1U;
    }
  }
  std::vector<std::uint8_t> out;
  out.reserve(after_hyperlinks.size() + record.size());
  for (std::size_t i = 0; i < recs.size(); ++i) {
    if (i == insert_at) {
      out.insert(out.end(), record.begin(), record.end());
    }
    if (recs[i].type != type) {
      append_record(out, after_hyperlinks, recs[i]);
    }
  }
  if (insert_at == recs.size()) {
    out.insert(out.end(), record.begin(), record.end());
  }
  after_hyperlinks = std::move(out);
  return true;
}

}  // namespace

bool splice_drawing(std::vector<std::uint8_t>& after_hyperlinks, std::string_view rel_id) {
  return SpliceDrawingRecord(after_hyperlinks, rel_id, kBrtDrawing);
}

bool splice_legacy_drawing(std::vector<std::uint8_t>& after_hyperlinks, std::string_view rel_id) {
  return SpliceDrawingRecord(after_hyperlinks, rel_id, kBrtLegacyDrawing);
}

bool drawing_frames_are_writable(std::string_view drawing_xml) {
  pugi::xml_document doc;
  if (!doc.load_buffer(drawing_xml.data(), drawing_xml.size(), pugi::parse_default, pugi::encoding_utf8)) {
    return false;
  }
  const pugi::xml_node other = doc.find_node([](const pugi::xml_node& node) {
    return node.type() == pugi::node_element && local_name(node) == "graphicData" &&
           std::string_view(node.attribute("uri").value()) != chart_ns::kNsC;
  });
  return !other;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
