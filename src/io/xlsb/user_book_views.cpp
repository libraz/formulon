#include "io/xlsb/user_book_views.h"

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstring>
#include <optional>
#include <set>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/ooxml/workbook_part_reader.h"
#include "io/xlsb/custom_view_records.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/guid_text.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "sheet_passthrough.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

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

/// Visits each `element` under the raw children of `slot`; false when a child does not parse.
template <typename Children, typename Fn>
bool ForEachRawElement(const Children& children, std::size_t slot, const char* element, Fn&& fn) {
  for (const auto& child : children) {
    if (child.slot != slot) {
      continue;
    }
    pugi::xml_document doc;
    if (!doc.load_buffer(child.xml.data(), child.xml.size())) {
      return false;
    }
    for (const pugi::xml_node& node : doc.document_element().children(element)) {
      fn(node);
    }
  }
  return true;
}

std::size_t GuidCount(const std::vector<Guid>& guids, const Guid& guid) {
  return static_cast<std::size_t>(std::count(guids.begin(), guids.end(), guid));
}

struct EncodedView {
  Guid guid{};
  std::vector<std::uint8_t> records;
  bool ok = false;
};

EncodedView EncodeView(const pugi::xml_node& node, bool (*encode)(const pugi::xml_node&, std::vector<std::uint8_t>&)) {
  EncodedView view;
  view.ok = ParseGuid(node.attribute("guid").value(), view.guid) && encode(node, view.records);
  return view;
}

}  // namespace

XmlCustomViewPlan plan_xml_custom_views(const Workbook& workbook) {
  const std::size_t sheet_count = workbook.sheet_count();
  XmlCustomViewPlan plan;
  plan.sheet_blocks.resize(sheet_count);
  plan.orphan_sheet_views.assign(sheet_count, 0U);
  if (!workbook.xlsb_user_book_views().empty()) {
    return plan;
  }
  std::vector<EncodedView> books;
  // A view set that does not parse names no GUID, so nothing can be paired with it.
  bool readable = ForEachRawElement(
      workbook.raw_children(), ooxml::workbook_child::slot_of("customWorkbookViews"), "customWorkbookView",
      [&](const pugi::xml_node& node) { books.push_back(EncodeView(node, encode_book_view)); });
  std::vector<std::vector<EncodedView>> sheets(sheet_count);
  for (std::size_t i = 0; i < sheet_count; ++i) {
    readable = ForEachRawElement(
                   workbook.sheet(i).raw_extensions(), worksheet_child::slot_of("customSheetViews"), "customSheetView",
                   [&](const pugi::xml_node& node) { sheets[i].push_back(EncodeView(node, encode_sheet_view)); }) &&
               readable;
  }
  std::vector<Guid> book_guids;
  for (const EncodedView& book : books) {
    book_guids.push_back(book.guid);
  }
  std::set<Guid> writable;
  for (const EncodedView& book : books) {
    bool carried = false;
    bool encodes = readable && book.ok && GuidCount(book_guids, book.guid) == 1U;
    for (const std::vector<EncodedView>& views : sheets) {
      for (const EncodedView& view : views) {
        if (view.guid == book.guid) {
          carried = true;
          encodes = encodes && view.ok;
        }
      }
    }
    if (carried && encodes) {
      writable.insert(book.guid);
      plan.book_records.insert(plan.book_records.end(), book.records.begin(), book.records.end());
    } else {
      ++plan.refused_book_views;
    }
  }
  for (std::size_t i = 0; i < sheet_count; ++i) {
    std::vector<std::uint8_t> block;
    for (const EncodedView& view : sheets[i]) {
      if (writable.count(view.guid) != 0U) {
        block.insert(block.end(), view.records.begin(), view.records.end());
      } else if (GuidCount(book_guids, view.guid) == 0U) {
        ++plan.orphan_sheet_views[i];
      }
    }
    if (!block.empty()) {
      emit_record(plan.sheet_blocks[i], kBrtBeginUserShViews, ByteSpan{});
      plan.sheet_blocks[i].insert(plan.sheet_blocks[i].end(), block.begin(), block.end());
      emit_record(plan.sheet_blocks[i], kBrtEndUserShViews, ByteSpan{});
    }
  }
  return plan;
}

namespace {

/// Each view of the sheet's retained BrtBeginUserShViews blocks as its records, BrtBeginUserShView through
/// BrtEndUserShView; false when a block holds anything outside a view.
bool SheetRecordViews(const XlsbSheetTail& tail,
                      std::vector<std::pair<std::vector<FramedRecord>, const std::vector<std::uint8_t>*>>* views) {
  for (const std::vector<std::uint8_t>* slot :
       {&tail.before_merges, &tail.after_merges_before_hyperlinks, &tail.after_hyperlinks}) {
    bool in_block = false;
    std::vector<FramedRecord> view;
    for (const FramedRecord& rec : split_records(*slot)) {
      if (!in_block) {
        in_block = rec.type == kBrtBeginUserShViews;
        continue;
      }
      if (rec.type == kBrtEndUserShViews) {
        if (!view.empty()) {
          return false;
        }
        in_block = false;
      } else if (view.empty() && rec.type != kBrtBeginUserShView) {
        return false;
      } else {
        view.push_back(rec);
        if (rec.type == kBrtEndUserShView) {
          views->emplace_back(std::move(view), slot);
          view.clear();
        }
      }
    }
    if (in_block) {
      return false;
    }
  }
  return true;
}

struct DecodedView {
  Guid guid{};
  std::optional<std::string> xml;
};

}  // namespace

RecordCustomViewPlan plan_record_custom_views(const Workbook& workbook) {
  const std::size_t sheet_count = workbook.sheet_count();
  RecordCustomViewPlan plan;
  plan.sheet_xml.resize(sheet_count);
  const std::vector<std::uint8_t>& book = workbook.xlsb_user_book_views();
  if (book.empty() || xml_book_view_count(workbook) != 0U) {
    return plan;
  }
  std::vector<DecodedView> books;
  for (const FramedRecord& rec : split_records(book)) {
    DecodedView view;
    if (rec.type == kBrtUserBookView && rec.payload.size >= kBookViewGuidOffset + kGuidSize) {
      view.guid = GuidAt(rec.payload.data + kBookViewGuidOffset);
      view.xml = decode_book_view(rec.payload);
    }
    books.push_back(std::move(view));
  }
  bool readable = true;
  std::vector<std::vector<DecodedView>> sheets(sheet_count);
  for (std::size_t i = 0; i < sheet_count; ++i) {
    std::vector<std::pair<std::vector<FramedRecord>, const std::vector<std::uint8_t>*>> views;
    readable = SheetRecordViews(workbook.sheet(i).xlsb_tail(), &views) && readable;
    for (const auto& [records, slot] : views) {
      DecodedView view;
      if (records.front().payload.size >= kGuidSize) {
        view.guid = GuidAt(records.front().payload.data);
        view.xml = decode_sheet_view(records, *slot);
      }
      sheets[i].push_back(std::move(view));
    }
  }
  std::vector<Guid> book_guids;
  for (const DecodedView& view : books) {
    book_guids.push_back(view.guid);
  }
  std::set<Guid> writable;
  std::string book_views;
  for (const DecodedView& view : books) {
    bool carried = false;
    bool decodes = readable && view.xml.has_value() && GuidCount(book_guids, view.guid) == 1U;
    for (const std::vector<DecodedView>& sheet_views : sheets) {
      for (const DecodedView& sheet_view : sheet_views) {
        if (sheet_view.guid == view.guid) {
          carried = true;
          decodes = decodes && sheet_view.xml.has_value();
        }
      }
    }
    if (carried && decodes) {
      writable.insert(view.guid);
      book_views.append(*view.xml);
    }
  }
  if (book_views.empty()) {
    return plan;
  }
  plan.book_xml = "<customWorkbookViews>" + book_views + "</customWorkbookViews>";
  for (std::size_t i = 0; i < sheet_count; ++i) {
    std::string views;
    for (const DecodedView& view : sheets[i]) {
      if (writable.count(view.guid) != 0U) {
        views.append(*view.xml);
      }
    }
    if (!views.empty()) {
      plan.sheet_xml[i] = "<customSheetViews>" + views + "</customSheetViews>";
    }
  }
  return plan;
}

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

std::size_t splice_book_user_views(std::vector<std::uint8_t>& workbook_bin, const Workbook& workbook,
                                   const std::vector<std::uint8_t>& generated) {
  const std::vector<std::uint8_t>& retained = workbook.xlsb_user_book_views();
  std::vector<std::uint8_t> block = retained.empty() ? generated : std::vector<std::uint8_t>{};
  std::size_t dropped = 0;
  const std::set<Guid> sheet_views = retained.empty() ? std::set<Guid>{} : SheetViewGuids(workbook);
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

std::size_t xml_book_view_count(const Workbook& workbook) {
  return CountRawChildElements(workbook.raw_children(), ooxml::workbook_child::slot_of("customWorkbookViews"),
                               "customWorkbookView");
}

std::size_t xml_sheet_view_count(const Sheet& sheet) {
  return CountRawChildElements(sheet.raw_extensions(), worksheet_child::slot_of("customSheetViews"), "customSheetView");
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
