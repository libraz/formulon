// Structural checks on a saved .xlsb package: cross-part and per-record rules
// Excel enforces at open (it refuses the file, crashes, or silently reads a
// different state) plus the workbook record order Excel writes. Each rule has
// its own function over the unzipped parts that appends "<rule>: <detail>"
// entries; ValidateXlsbPackage unzips once and runs them all. The rule name
// before ": " is a stable key.
#ifndef FORMULON_TESTS_SUPPORT_XLSB_PACKAGE_VALIDATOR_H_
#define FORMULON_TESTS_SUPPORT_XLSB_PACKAGE_VALIDATOR_H_

#include <cstddef>
#include <cstdint>
#include <map>
#include <set>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/xlsb/guid_text.h"
#include "io/xlsb/pivot_records_common.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/sheet_record_flags.h"
#include "io/xlsb/workbook_props_records.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "sheet_layout.h"

namespace formulon::test {

/// The unzipped parts of a package, by part name.
struct XlsbParts {
  std::map<std::string, std::vector<std::uint8_t>> parts;

  /// The bytes of `name`, or null when the package has no such part.
  const std::vector<std::uint8_t>* find(const std::string& name) const {
    const auto it = parts.find(name);
    return it == parts.end() ? nullptr : &it->second;
  }
};

namespace xlsb_validator_detail {

using io::ByteSpan;
using io::xlsb::FramedRecord;
using io::xlsb::load_u32;
using io::xlsb::XlsbRecordType;

inline constexpr const char* kWorkbookPart = "xl/workbook.bin";
inline constexpr const char* kWorkbookRels = "xl/_rels/workbook.bin.rels";
inline constexpr const char* kStylesPart = "xl/styles.bin";

inline constexpr std::uint16_t kBrtFileRecover = 155;
inline constexpr std::uint16_t kBrtBeginPivotCacheIds = 384;
inline constexpr std::uint16_t kBrtEndPivotCacheIds = 385;
inline constexpr std::uint16_t kBrtBeginPivotCacheId = 386;
inline constexpr std::uint16_t kBrtUserBookView = 397;
inline constexpr std::uint16_t kBrtBeginUserShView = 423;

/// BrtUserBookView: iTabid (the active sheet's sheetId) and the view GUID.
inline constexpr std::size_t kUserBookViewActiveSheetOffset = 16;
inline constexpr std::size_t kUserBookViewGuidOffset = 24;
/// BrtBookView: itabCur, the seventh u32.
inline constexpr std::size_t kBookViewItabCurOffset = 24;
/// BrtBeginSXView: cacheId.
inline constexpr std::size_t kSxViewCacheIdOffset = 28;
/// BrtRowHdr: ixfe, and the flag byte holding fGhostDirty.
inline constexpr std::size_t kRowHdrXfOffset = 4;
inline constexpr std::size_t kRowHdrFlags2Offset = 11;
/// BrtWsFmtInfo: dxGCol's "no default width" sentinel and units, and the flag word.
inline constexpr std::uint32_t kAbsentDefaultColumnWidth = 0xFFFFFFFFU;
inline constexpr double kColumnWidthUnits = 256.0;
inline constexpr std::size_t kWsFmtFlagsOffset = 8;
/// Breaks Excel accepts in one row or column break block.
inline constexpr std::size_t kMaxBreaksPerBlock = 1024;

inline bool StartsWith(const std::string& s, std::string_view prefix) {
  return s.size() >= prefix.size() && s.compare(0, prefix.size(), prefix) == 0;
}

inline bool EndsWith(const std::string& s, std::string_view suffix) {
  return s.size() >= suffix.size() && s.compare(s.size() - suffix.size(), suffix.size(), suffix) == 0;
}

/// True when `name` sits directly in `dir` (which ends in '/') and ends in `.bin`.
inline bool IsBinIn(const std::string& name, std::string_view dir) {
  return StartsWith(name, dir) && EndsWith(name, ".bin") && name.find('/', dir.size()) == std::string::npos;
}

/// Records of `part`; empty when the part is absent or does not frame whole.
inline std::vector<FramedRecord> Records(const XlsbParts& pkg, const std::string& part) {
  const std::vector<std::uint8_t>* bytes = pkg.find(part);
  return bytes == nullptr ? std::vector<FramedRecord>{} : io::xlsb::split_records(*bytes);
}

/// Resolves a relationship target of a part in `xl/`, collapsing "." and "..".
inline std::string ResolveFromXl(const std::string& target) {
  const std::string full = (!target.empty() && target[0] == '/') ? target.substr(1) : "xl/" + target;
  std::vector<std::string> segs;
  std::size_t start = 0;
  while (start <= full.size()) {
    std::size_t end = full.find('/', start);
    if (end == std::string::npos) {
      end = full.size();
    }
    const std::string seg = full.substr(start, end - start);
    if (seg == "..") {
      if (!segs.empty()) {
        segs.pop_back();
      }
    } else if (!seg.empty() && seg != ".") {
      segs.push_back(seg);
    }
    start = end + 1U;
  }
  std::string joined;
  for (std::size_t i = 0; i < segs.size(); ++i) {
    joined += (i == 0 ? "" : "/") + segs[i];
  }
  return joined;
}

/// One workbook relationship: the last segment of its type URI and its resolved part name.
struct WorkbookRel {
  std::string type;
  std::string part;
};

/// Workbook relationships by id; empty when the rels part is absent or malformed.
inline std::map<std::string, WorkbookRel> WorkbookRels(const XlsbParts& pkg) {
  std::map<std::string, WorkbookRel> out;
  const std::vector<std::uint8_t>* bytes = pkg.find(kWorkbookRels);
  pugi::xml_document doc;
  if (bytes == nullptr || !doc.load_buffer(bytes->data(), bytes->size())) {
    return out;
  }
  for (const pugi::xml_node rel : doc.child("Relationships").children("Relationship")) {
    const std::string type = rel.attribute("Type").as_string();
    const std::size_t slash = type.rfind('/');
    out[rel.attribute("Id").as_string()] = WorkbookRel{slash == std::string::npos ? type : type.substr(slash + 1U),
                                                       ResolveFromXl(rel.attribute("Target").as_string())};
  }
  return out;
}

/// Worksheet parts the workbook relationships name, in name order.
inline std::vector<std::string> WorksheetParts(const XlsbParts& pkg) {
  std::set<std::string> parts;
  for (const auto& entry : WorkbookRels(pkg)) {
    if (entry.second.type == "worksheet" && pkg.find(entry.second.part) != nullptr) {
      parts.insert(entry.second.part);
    }
  }
  return std::vector<std::string>(parts.begin(), parts.end());
}

/// sheetId (iTabID) of every BrtBundleSh, in order.
inline std::vector<std::uint32_t> SheetIds(const std::vector<FramedRecord>& book) {
  std::vector<std::uint32_t> ids;
  for (const FramedRecord& rec : book) {
    if (rec.type == static_cast<std::uint16_t>(XlsbRecordType::BrtBundleSh) && rec.payload.size >= 8U) {
      ids.push_back(load_u32(rec.payload.data + 4));
    }
  }
  return ids;
}

inline std::string JoinIds(const std::vector<std::uint32_t>& ids) {
  std::string out;
  for (std::size_t i = 0; i < ids.size(); ++i) {
    out += (i == 0 ? "" : ",") + std::to_string(ids[i]);
  }
  return out;
}

/// Unzips `bytes` into `out`; on failure appends a `package_unreadable` entry
/// to `findings` and returns false.
inline bool UnzipXlsbParts(ByteSpan bytes, XlsbParts& out, std::vector<std::string>& findings) {
  io::ZipReader zip;
  if (!zip.open(bytes)) {
    findings.push_back("package_unreadable: not a zip archive");
    return false;
  }
  for (const std::string& name : zip.list_entries()) {
    auto part = zip.read_entry(name);
    if (!part) {
      findings.push_back("package_unreadable: " + name + " does not extract");
      return false;
    }
    out.parts[name] = std::move(part.value());
  }
  if (out.find(kWorkbookPart) == nullptr) {
    findings.push_back("package_unreadable: no xl/workbook.bin");
    return false;
  }
  return true;
}

/// The BrtUserBookView (397) GUIDs of `xl/workbook.bin` equal the union of the
/// BrtBeginUserShView (423) GUIDs over the worksheets.
/// Evidence: accept retained_children_edit/children_0_base.xlsb; refuse it
/// without 397 (refused) or without the sheet's 422..425 (crash).
inline void CheckUserBookViewPairing(const XlsbParts& pkg, std::vector<std::string>& out) {
  std::set<std::string> book;
  for (const FramedRecord& rec : Records(pkg, kWorkbookPart)) {
    if (rec.type == kBrtUserBookView && rec.payload.size >= kUserBookViewGuidOffset + io::xlsb::kGuidBytes) {
      book.insert(io::xlsb::FormatGuid(rec.payload.data + kUserBookViewGuidOffset));
    }
  }
  std::set<std::string> sheets;
  for (const std::string& part : WorksheetParts(pkg)) {
    for (const FramedRecord& rec : Records(pkg, part)) {
      if (rec.type == kBrtBeginUserShView && rec.payload.size >= io::xlsb::kGuidBytes) {
        sheets.insert(io::xlsb::FormatGuid(rec.payload.data));
      }
    }
  }
  for (const std::string& guid : book) {
    if (sheets.count(guid) == 0U) {
      out.push_back("user_book_view_pairing: " + guid + " in workbook.bin has no sheet view");
    }
  }
  for (const std::string& guid : sheets) {
    if (book.count(guid) == 0U) {
      out.push_back("user_book_view_pairing: " + guid + " in a sheet has no workbook view");
    }
  }
}

/// Every BrtPivotCacheID (386) resolves to an existing cache definition part,
/// every cache definition part is named by one, idSx values are unique, and
/// every pivot table's BrtBeginSXView cacheId equals some idSx.
/// Evidence: accept xlsb_parity/pivot/*.xlsb; refuse a dangling extra 386, an
/// idSx no pivot table names, and a cache with no 386 (refused).
inline void CheckPivotCacheIdResolves(const XlsbParts& pkg, std::vector<std::string>& out) {
  const std::map<std::string, WorkbookRel> rels = WorkbookRels(pkg);
  std::set<std::uint32_t> ids;
  std::set<std::string> named;
  bool in_block = false;
  for (const FramedRecord& rec : Records(pkg, kWorkbookPart)) {
    if (rec.type == kBrtBeginPivotCacheIds || rec.type == kBrtEndPivotCacheIds) {
      in_block = rec.type == kBrtBeginPivotCacheIds;
      continue;
    }
    if (!in_block || rec.type != kBrtBeginPivotCacheId || rec.payload.size < 4U) {
      continue;
    }
    const std::uint32_t id_sx = load_u32(rec.payload.data);
    const std::string label = "idSx " + std::to_string(id_sx);
    if (!ids.insert(id_sx).second) {
      out.push_back("pivot_cache_id_resolves: " + label + " repeats");
    }
    ByteSpan cursor{rec.payload.data + 4, rec.payload.size - 4U};
    auto rid = io::xlsb::read_xlwidestring(cursor);
    if (!rid) {
      out.push_back("pivot_cache_id_resolves: " + label + " has no readable r:id");
      continue;
    }
    const auto rel = rels.find(rid.value());
    if (rel == rels.end() || rel->second.type != "pivotCacheDefinition") {
      out.push_back("pivot_cache_id_resolves: " + label + " r:id " + rid.value() +
                    " names no pivotCacheDefinition relationship");
    } else if (pkg.find(rel->second.part) == nullptr) {
      out.push_back("pivot_cache_id_resolves: " + label + " r:id " + rid.value() + " targets missing " +
                    rel->second.part);
    } else {
      named.insert(rel->second.part);
    }
  }
  for (const auto& entry : pkg.parts) {
    const std::string& name = entry.first;
    if (IsBinIn(name, "xl/pivotCache/") && StartsWith(name, "xl/pivotCache/pivotCacheDefinition") &&
        named.count(name) == 0U) {
      out.push_back("pivot_cache_id_resolves: " + name + " is named by no BrtPivotCacheID");
    }
  }
  for (const auto& entry : pkg.parts) {
    if (!IsBinIn(entry.first, "xl/pivotTables/")) {
      continue;
    }
    for (const FramedRecord& rec : io::xlsb::split_records(entry.second)) {
      if (rec.type == io::xlsb::pivot_records::kBeginSXView && rec.payload.size >= kSxViewCacheIdOffset + 4U) {
        const std::uint32_t cache_id = load_u32(rec.payload.data + kSxViewCacheIdOffset);
        if (ids.count(cache_id) == 0U) {
          out.push_back("pivot_cache_id_resolves: " + entry.first + " cacheId " + std::to_string(cache_id) +
                        " matches no idSx");
        }
      }
    }
  }
}

/// Each BrtUserBookView (397) iTabid is a BrtBundleSh sheetId, and each
/// BrtBookView (158) itabCur is below the sheet count.
/// Evidence: accept retained_children_edit/children_0_base.xlsb and
/// xlsb_parity/pivot/*.xlsb; refuse an unknown sheet, which opens as another
/// active sheet (silent-wrong).
inline void CheckActiveSheetIsId(const XlsbParts& pkg, std::vector<std::string>& out) {
  const std::vector<FramedRecord> book = Records(pkg, kWorkbookPart);
  const std::vector<std::uint32_t> ids = SheetIds(book);
  for (const FramedRecord& rec : book) {
    if (rec.type == kBrtUserBookView && rec.payload.size >= kUserBookViewActiveSheetOffset + 4U) {
      const std::uint32_t active = load_u32(rec.payload.data + kUserBookViewActiveSheetOffset);
      bool found = false;
      for (const std::uint32_t id : ids) {
        found = found || id == active;
      }
      if (!found) {
        out.push_back("active_sheet_is_id: 397 names sheetId " + std::to_string(active) + ", book has " + JoinIds(ids));
      }
    }
    if (rec.type == io::xlsb::kBrtBookView && rec.payload.size >= kBookViewItabCurOffset + 4U) {
      const std::uint32_t itab_cur = load_u32(rec.payload.data + kBookViewItabCurOffset);
      if (itab_cur >= ids.size()) {
        out.push_back("active_sheet_is_id: 158 itabCur " + std::to_string(itab_cur) + " with " +
                      std::to_string(ids.size()) + " sheets");
      }
    }
  }
}

/// Each BrtWsFmtInfo (485) row and column outline summary level is at most 7.
/// Evidence: accept xlsb_parity/format/outline_row_7.xlsb; refuse a summary of
/// 8 (refused).
inline void CheckOutlineLevel(const XlsbParts& pkg, std::vector<std::string>& out) {
  for (const std::string& part : WorksheetParts(pkg)) {
    for (const FramedRecord& rec : Records(pkg, part)) {
      if (rec.type != static_cast<std::uint16_t>(XlsbRecordType::BrtWsFmtInfo) ||
          rec.payload.size < kWsFmtFlagsOffset + 4U) {
        continue;
      }
      const std::uint32_t flags = load_u32(rec.payload.data + kWsFmtFlagsOffset);
      const unsigned row_level = (flags >> 16U) & 0xFFU;
      const unsigned col_level = (flags >> 24U) & 0xFFU;
      if (row_level > kMaxOutlineLevel || col_level > kMaxOutlineLevel) {
        out.push_back("outline_level: " + part + " row " + std::to_string(row_level) + " col " +
                      std::to_string(col_level));
      }
    }
  }
}

/// Each BrtWsFmtInfo (485) dxGCol is the "no default" sentinel or at most 255
/// characters.
/// Evidence: accept xlsb_parity/format/width_255.xlsb, base_255.xlsb and
/// height_409_5.xlsb; refuse width_record_max.xlsb, dxGCol 256 characters
/// (refused).
inline void CheckWsFmtInfoRange(const XlsbParts& pkg, std::vector<std::string>& out) {
  for (const std::string& part : WorksheetParts(pkg)) {
    for (const FramedRecord& rec : Records(pkg, part)) {
      if (rec.type != static_cast<std::uint16_t>(XlsbRecordType::BrtWsFmtInfo) || rec.payload.size < 4U) {
        continue;
      }
      const std::uint32_t dx_g_col = load_u32(rec.payload.data);
      if (dx_g_col != kAbsentDefaultColumnWidth &&
          static_cast<double>(dx_g_col) / kColumnWidthUnits > kMaxColumnWidthChars) {
        out.push_back("ws_fmt_info_range: " + part + " dxGCol " + std::to_string(dx_g_col));
      }
    }
  }
}

/// A row (392..393) or column (394..395) break block holds at most 1024
/// BrtBrk records.
/// Evidence: accept xlsb_parity/print/breaks_row_1024.xlsb and
/// breaks_col_1024.xlsb; refuse 1025 breaks (refused).
inline void CheckPageBreakCount(const XlsbParts& pkg, std::vector<std::string>& out) {
  for (const std::string& part : WorksheetParts(pkg)) {
    std::uint16_t open = 0;
    std::size_t count = 0;
    for (const FramedRecord& rec : Records(pkg, part)) {
      if (rec.type == io::xlsb::kBrtBeginRwBrk || rec.type == io::xlsb::kBrtBeginColBrk) {
        open = rec.type;
        count = 0;
      } else if (rec.type == io::xlsb::kBrtBrk && open != 0U) {
        ++count;
      } else if ((rec.type == io::xlsb::kBrtEndRwBrk || rec.type == io::xlsb::kBrtEndColBrk) && open != 0U) {
        if (count > kMaxBreaksPerBlock) {
          out.push_back("page_break_count: " + part + (open == io::xlsb::kBrtBeginRwBrk ? " row" : " column") +
                        " block has " + std::to_string(count) + " breaks");
        }
        open = 0;
      }
    }
  }
}

/// Every cell record's iStyleRef and every styled BrtRowHdr's ixfe (fGhostDirty
/// set) in the sheet data is below the number of cell XFs in `xl/styles.bin`.
/// Evidence: accept every xlsb_parity fixture; refuse an index past the cell XF
/// table (refused).
inline void CheckSheetXfInRange(const XlsbParts& pkg, std::vector<std::string>& out) {
  if (pkg.find(kStylesPart) == nullptr) {
    return;
  }
  std::size_t xf_count = 0;
  bool in_cell_xfs = false;
  for (const FramedRecord& rec : Records(pkg, kStylesPart)) {
    const auto type = static_cast<XlsbRecordType>(rec.type);
    if (type == XlsbRecordType::BrtBeginCellXFs || type == XlsbRecordType::BrtEndCellXFs) {
      in_cell_xfs = type == XlsbRecordType::BrtBeginCellXFs;
    } else if (in_cell_xfs && type == XlsbRecordType::BrtXF) {
      ++xf_count;
    }
  }
  for (const std::string& part : WorksheetParts(pkg)) {
    bool in_data = false;
    std::uint32_t row = 0;
    for (const FramedRecord& rec : Records(pkg, part)) {
      const auto type = static_cast<XlsbRecordType>(rec.type);
      if (type == XlsbRecordType::BrtBeginSheetData || type == XlsbRecordType::BrtEndSheetData) {
        in_data = type == XlsbRecordType::BrtBeginSheetData;
        continue;
      }
      if (!in_data) {
        continue;
      }
      if (type == XlsbRecordType::BrtRowHdr) {
        if (rec.payload.size >= 4U) {
          row = load_u32(rec.payload.data);
        }
        if (rec.payload.size > kRowHdrFlags2Offset &&
            (rec.payload.data[kRowHdrFlags2Offset] & io::xlsb::kRowGhostDirty) != 0U) {
          const std::uint32_t xf = load_u32(rec.payload.data + kRowHdrXfOffset);
          if (xf >= xf_count) {
            out.push_back("sheet_xf_in_range: " + part + " row " + std::to_string(row) + " xf " + std::to_string(xf) +
                          " of " + std::to_string(xf_count));
          }
        }
        continue;
      }
      const bool is_cell = rec.type >= static_cast<std::uint16_t>(XlsbRecordType::BrtCellBlank) &&
                           rec.type <= static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaError);
      if (!is_cell || rec.payload.size < 8U) {
        continue;
      }
      // Cell header: column u32, then iStyleRef as a u24.
      const std::uint32_t xf = static_cast<std::uint32_t>(rec.payload.data[4]) |
                               (static_cast<std::uint32_t>(rec.payload.data[5]) << 8U) |
                               (static_cast<std::uint32_t>(rec.payload.data[6]) << 16U);
      if (xf >= xf_count) {
        out.push_back("sheet_xf_in_range: " + part + " row " + std::to_string(row) + " col " +
                      std::to_string(load_u32(rec.payload.data)) + " xf " + std::to_string(xf) + " of " +
                      std::to_string(xf_count));
      }
    }
  }
}

/// In `xl/workbook.bin` the records present among BrtCalcProp (157),
/// BrtUserBookView (397), BrtBeginPivotCacheIDs (384), BrtFileRecover (155)
/// and BrtEndBook (132) come in that order.
/// Evidence: accept retained_children_edit/children_0_base.xlsb and
/// xlsb_parity/pivot/*.xlsb; Excel opens another order (writer-shape).
inline void CheckWorkbookRecordOrder(const XlsbParts& pkg, std::vector<std::string>& out) {
  constexpr std::uint16_t kOrder[] = {io::xlsb::kBrtCalcProp, kBrtUserBookView, kBrtBeginPivotCacheIds, kBrtFileRecover,
                                      io::xlsb::kBrtEndBook};
  constexpr std::size_t kOrderSize = sizeof(kOrder) / sizeof(kOrder[0]);
  std::size_t last_rank = 0;
  for (const FramedRecord& rec : Records(pkg, kWorkbookPart)) {
    for (std::size_t rank = 0; rank < kOrderSize; ++rank) {
      if (rec.type != kOrder[rank]) {
        continue;
      }
      if (rank < last_rank) {
        out.push_back("workbook_record_order: " + std::to_string(rec.type) + " after " +
                      std::to_string(kOrder[last_rank]));
      } else {
        last_rank = rank;
      }
    }
  }
}

}  // namespace xlsb_validator_detail

using xlsb_validator_detail::CheckActiveSheetIsId;
using xlsb_validator_detail::CheckOutlineLevel;
using xlsb_validator_detail::CheckPageBreakCount;
using xlsb_validator_detail::CheckPivotCacheIdResolves;
using xlsb_validator_detail::CheckSheetXfInRange;
using xlsb_validator_detail::CheckUserBookViewPairing;
using xlsb_validator_detail::CheckWorkbookRecordOrder;
using xlsb_validator_detail::CheckWsFmtInfoRange;
using xlsb_validator_detail::UnzipXlsbParts;

/// Runs every rule over `xlsb_bytes`; one entry per violation.
inline std::vector<std::string> ValidateXlsbPackage(io::ByteSpan xlsb_bytes) {
  std::vector<std::string> out;
  XlsbParts pkg;
  if (!UnzipXlsbParts(xlsb_bytes, pkg, out)) {
    return out;
  }
  CheckUserBookViewPairing(pkg, out);
  CheckPivotCacheIdResolves(pkg, out);
  CheckActiveSheetIsId(pkg, out);
  CheckOutlineLevel(pkg, out);
  CheckWsFmtInfoRange(pkg, out);
  CheckPageBreakCount(pkg, out);
  CheckSheetXfInRange(pkg, out);
  CheckWorkbookRecordOrder(pkg, out);
  return out;
}

}  // namespace formulon::test

#endif  // FORMULON_TESTS_SUPPORT_XLSB_PACKAGE_VALIDATOR_H_
