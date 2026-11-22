// Loading of the external link parts an XLSB workbook's supporting-book list
// names, and the projection of those links onto the Ptg decoder's name tables.

#include <cstdint>
#include <limits>
#include <map>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "external_link.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/part_graph.h"
#include "io/ooxml/relationship_xml.h"
#include "io/ooxml/rels_walker.h"
#include "io/ooxml_defs.h"
#include "io/xlsb/external_link_reader.h"
#include "io/xlsb/package_loader.h"
#include "io/xlsb/part_paths.h"
#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/reader_internal.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/workbook_bin_reader.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/index_sort.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace reader_detail {
namespace {

using detail::FindRelationship;

/// Reads the remote workbook URLs recorded in an external link part's own
/// rels file into `record`: the relationship `record.body_rel_id` names is
/// the target, and the ones `record.absolute_rel_id` / `record.relative_rel_id`
/// name (from the part's alternate-URL record) are the absolute and relative
/// URLs Excel records beside it. Without a body rel id
/// the first external-path relationship is the target and no absolute URL
/// is read. Leaves both empty when the part has no rels, which is what an
/// unresolvable link already looks like.
void ReadExternalLinkTargets(const ZipReader& zip, ExternalLinkRecord& record) {
  const std::string rels_path = ooxml::rels_path_for_part(record.part_path);
  if (!zip.has_entry(rels_path)) {
    return;
  }
  const bool body_known = !record.body_rel_id.empty();
  auto status = ooxml::visit_relationship_nodes(
      zip, rels_path, "external link rels", "xlsb_reader", [&](const pugi::xml_node& rel) -> Expected<void, Error> {
        const std::string_view type = rel.attribute("Type").value();
        const bool missing = type == kRelExternalLinkPathMissing;
        if (type != kRelExternalLinkPath && !missing) {
          return Expected<void, Error>::Ok();
        }
        const std::string_view id = rel.attribute("Id").value();
        const bool is_body = body_known ? id == record.body_rel_id : record.target.empty();
        if (is_body) {
          record.target = rel.attribute("Target").value();
          record.target_path_missing = missing;
        } else if (!missing && body_known && !record.absolute_rel_id.empty() && id == record.absolute_rel_id) {
          record.absolute_target = rel.attribute("Target").value();
        } else if (!missing && body_known && !record.relative_rel_id.empty() && id == record.relative_rel_id) {
          record.relative_target = rel.attribute("Target").value();
        }
        return Expected<void, Error>::Ok();
      });
  (void)status;
}

/// `BrtBeginExternalBook` `sbt` values other than a workbook (0).
constexpr std::uint16_t kSbtDde = 1;
constexpr std::uint16_t kSbtOle = 2;

/// The `sbt` of the part's opening `BrtBeginExternalBook`; 0 when the part
/// does not open with one.
std::uint16_t ExternalLinkSbt(ByteSpan part) {
  auto rec_or = read_record(part);
  if (!rec_or || rec_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtBeginExternalBook)) {
    return 0;
  }
  ByteSpan payload = rec_or.value().payload;
  auto sbt_or = read_u16(payload);
  return sbt_or ? sbt_or.value() : std::uint16_t{0};
}

}  // namespace

/// Renumbers the external books by first appearance in the `BrtExternSheet` list; books no entry
/// names follow in stored order. `books` and `ranges` are rewritten to the same new numbers.
void NumberExternalBooksByFirstUse(std::vector<XlsbSupBook>& books, std::vector<XlsbSheetRange>& ranges) {
  std::map<std::uint32_t, std::uint32_t> renumber;
  std::uint32_t next = 1;
  const auto assign = [&](std::uint32_t old_number) {
    if (old_number != 0 && renumber.emplace(old_number, next).second) {
      ++next;
    }
  };
  for (const XlsbSheetRange& range : ranges) {
    if (range.external_book != std::numeric_limits<std::uint32_t>::max()) {
      assign(range.external_book);
    }
  }
  for (const XlsbSupBook& sup : books) {
    assign(sup.external_book);
  }
  for (XlsbSupBook& sup : books) {
    if (sup.external_book != 0) {
      sup.external_book = renumber[sup.external_book];
    }
  }
  for (XlsbSheetRange& range : ranges) {
    const auto it = renumber.find(range.external_book);
    if (it != renumber.end()) {
      range.external_book = it->second;
    }
  }
}

/// Loads every external link part the supporting-book list names, in
/// `[N]` order, so `external_links()[N - 1]` is the book a formula's
/// `[N]` prefix selects.
///
/// A part that is missing, unreadable or uses an encoding the decoder
/// does not support contributes an entry with an empty cache rather than
/// none: the position has to be held or every later `[N]` would shift by
/// one and resolve against the wrong workbook. A reference into an empty
/// cache reads `#REF!`, which is the behaviour before this existed.
///
/// A link whose `BrtBeginExternalBook` carries `sbt` 1 (DDE) or 2 (OLE) is
/// recorded as that kind and its body is not read as a supporting book. The
/// XLSB writer regenerates the parts of external-book links only, so a DDE
/// or OLE part keeps round-tripping verbatim through passthrough.
void LoadExternalLinkParts(const ZipReader& zip, Workbook& wb, const std::vector<XlsbSupBook>& books) {
  std::vector<const XlsbSupBook*> ordered;
  for (const XlsbSupBook& sup : books) {
    if (sup.external_book != 0) {  // 0 is this workbook.
      ordered.push_back(&sup);
    }
  }
  sort_by_index(ordered,
                [](const XlsbSupBook* l, const XlsbSupBook* r) { return l->external_book < r->external_book; });
  std::vector<ExternalLinkRecord> links;
  for (const XlsbSupBook* sup_ptr : ordered) {
    const XlsbSupBook& sup = *sup_ptr;
    ExternalLinkRecord record;
    record.index = sup.external_book;
    record.rel_id = sup.rel_id;
    auto path_or = FindRelationship(zip, kWorkbookBinPath, "Id", sup.rel_id, "external link rels");
    if (path_or && !path_or.value().empty() && zip.has_entry(path_or.value())) {
      record.part_path = path_or.value();
      auto bytes_or = zip.read_entry(record.part_path);
      if (bytes_or) {
        const std::vector<std::uint8_t>& bytes = bytes_or.value();
        const std::uint16_t sbt = ExternalLinkSbt(ByteSpan{bytes.data(), bytes.size()});
        if (sbt == kSbtDde || sbt == kSbtOle) {
          record.kind = sbt == kSbtDde ? ExternalLinkRecord::Kind::kDdeLink : ExternalLinkRecord::Kind::kOleLink;
        } else if (auto book_or = read_external_link_bin(ByteSpan{bytes.data(), bytes.size()}, &record.body_rel_id,
                                                         &record.absolute_rel_id, &record.relative_rel_id)) {
          record.kind = ExternalLinkRecord::Kind::kExternalBook;
          record.book = std::move(book_or.value());
        } else {
          record.body_rel_id.clear();
          record.absolute_rel_id.clear();
          record.relative_rel_id.clear();
        }
      }
      ReadExternalLinkTargets(zip, record);
    }
    links.push_back(std::move(record));
  }
  if (!links.empty()) {
    wb.set_external_links(std::move(links));
  }
}

/// Projects the loaded external links onto the two name tables
/// `decode_ptgs` consults. The projection exists so the Ptg decoder does
/// not have to know about `Workbook` or the link records at all.
XlsbExternalBooks CollectExternalBookTables(const Workbook& wb) {
  XlsbExternalBooks books;
  books.reserve(wb.external_links().size());
  for (const ExternalLinkRecord& link : wb.external_links()) {
    XlsbExternalBook entry;
    entry.sheet_names = link.book.sheet_names;
    entry.names.reserve(link.book.names.size());
    for (const ExternalBookName& name : link.book.names) {
      entry.names.push_back(name.name);
    }
    books.push_back(std::move(entry));
  }
  return books;
}

}  // namespace reader_detail
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
