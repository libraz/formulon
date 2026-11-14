#include "io/xlsb/user_book_views.h"

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstring>
#include <set>
#include <string>
#include <string_view>

#include "io/ooxml/workbook_part_reader.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "sheet.h"
#include "sheet_passthrough.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::uint16_t kBrtUserBookView = 397;
constexpr std::uint16_t kBrtBeginUserShView = 423;
constexpr std::uint16_t kBrtFileRecover = 155;
constexpr std::size_t kBookViewGuidOffset = 24;
constexpr std::size_t kGuidSize = 16;

using Guid = std::array<std::uint8_t, kGuidSize>;

Guid GuidAt(const std::uint8_t* data) {
  Guid guid{};
  std::memcpy(guid.data(), data, kGuidSize);
  return guid;
}

/// GUIDs of the custom views the sheets' retained tails carry.
std::set<Guid> SheetViewGuids(const Workbook& workbook) {
  std::set<Guid> guids;
  for (std::size_t i = 0; i < workbook.sheet_count(); ++i) {
    const XlsbSheetTail& tail = workbook.sheet(i).xlsb_tail();
    for (const std::vector<std::uint8_t>* slot :
         {&tail.before_merges, &tail.after_merges_before_hyperlinks, &tail.after_hyperlinks}) {
      for (const FramedRecord& rec : split_records(*slot)) {
        if (rec.type == kBrtBeginUserShView && rec.payload.size >= kGuidSize) {
          guids.insert(GuidAt(rec.payload.data));
        }
      }
    }
  }
  return guids;
}

/// Number of `<tag ...>` elements in `xml`, not counting a longer name that
/// starts with `tag` (the `<tag>s` container).
std::size_t CountElements(std::string_view xml, std::string_view tag) {
  const std::string open = "<" + std::string(tag);
  std::size_t count = 0;
  for (std::size_t at = xml.find(open); at != std::string_view::npos; at = xml.find(open, at + open.size())) {
    const std::size_t next = at + open.size();
    if (next < xml.size() && (xml[next] == ' ' || xml[next] == '/' || xml[next] == '>' || xml[next] == '\t' ||
                              xml[next] == '\n' || xml[next] == '\r')) {
      ++count;
    }
  }
  return count;
}

std::size_t CountRawChildElements(const WorksheetRawExtensions& children, std::size_t slot, std::string_view tag) {
  std::size_t count = 0;
  for (const WorksheetRawChild& child : children) {
    if (child.slot == slot) {
      count += CountElements(child.xml, tag);
    }
  }
  return count;
}

}  // namespace

std::vector<std::uint8_t> read_book_user_views(ByteSpan workbook_bin) {
  const std::vector<std::uint8_t> body(workbook_bin.data, workbook_bin.data + workbook_bin.size);
  std::vector<std::uint8_t> out;
  for (const FramedRecord& rec : split_records(body)) {
    if (rec.type == kBrtUserBookView) {
      append_record(out, body, rec);
    }
  }
  return out;
}

std::size_t splice_book_user_views(std::vector<std::uint8_t>& workbook_bin, const Workbook& workbook) {
  const std::vector<std::uint8_t>& retained = workbook.xlsb_user_book_views();
  if (retained.empty()) {
    return 0U;
  }
  const std::set<Guid> sheet_views = SheetViewGuids(workbook);
  std::vector<std::uint8_t> block;
  std::size_t dropped = 0;
  for (const FramedRecord& rec : split_records(retained)) {
    if (rec.payload.size >= kBookViewGuidOffset + kGuidSize &&
        sheet_views.count(GuidAt(rec.payload.data + kBookViewGuidOffset)) != 0U) {
      append_record(block, retained, rec);
    } else {
      ++dropped;
    }
  }
  if (block.empty()) {
    return dropped;
  }
  const std::vector<FramedRecord> dst = split_records(workbook_bin);
  const auto at = std::find_if(dst.begin(), dst.end(), [](const FramedRecord& r) {
    return r.type == kBrtFileRecover || r.type == kFrtBegin || r.type == kBrtEndBook;
  });
  const std::size_t offset = at == dst.end() ? workbook_bin.size() : at->begin;
  workbook_bin.insert(workbook_bin.begin() + static_cast<std::ptrdiff_t>(offset), block.begin(), block.end());
  return dropped;
}

std::size_t unwritten_xml_book_views(const Workbook& workbook) {
  return CountRawChildElements(workbook.raw_children(), ooxml::workbook_child::slot_of("customWorkbookViews"),
                               "customWorkbookView");
}

std::size_t unwritten_xml_sheet_views(const Sheet& sheet) {
  return CountRawChildElements(sheet.raw_extensions(), worksheet_child::slot_of("customSheetViews"), "customSheetView");
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
