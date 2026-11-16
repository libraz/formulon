// Structural checks on a saved .xlsx package: rules Excel enforces at open
// (it repairs or refuses a package that breaks one). Each rule has its own
// function returning "<rule>: <detail>" entries; ValidateXlsxPackage runs all.
#ifndef FORMULON_TESTS_SUPPORT_XLSX_PACKAGE_VALIDATOR_H_
#define FORMULON_TESTS_SUPPORT_XLSX_PACKAGE_VALIDATOR_H_

#include <cstddef>
#include <set>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/zip_reader.h"
#include "pugixml.hpp"

namespace formulon::test {

using XlsxFindings = std::vector<std::string>;

namespace xlsx_validator_detail {

struct Package {
  io::ZipReader zip;
  std::vector<std::string> names;
};

inline bool Open(Package& pkg, io::ByteSpan bytes) {
  if (!pkg.zip.open(bytes)) {
    return false;
  }
  pkg.names = pkg.zip.list_entries();
  return true;
}

inline bool EndsWith(const std::string& s, std::string_view suffix) {
  return s.size() >= suffix.size() && s.compare(s.size() - suffix.size(), suffix.size(), suffix) == 0;
}

inline bool Parse(const Package& pkg, const std::string& name, pugi::xml_document& doc) {
  auto part = pkg.zip.read_entry(name);
  if (!part) {
    return false;
  }
  const auto& data = part.value();
  return static_cast<bool>(doc.load_buffer(data.data(), data.size()));
}

inline bool IsWorksheetPart(const std::string& name) {
  return name.rfind("xl/worksheets/", 0) == 0 && EndsWith(name, ".xml") && name.find('/', 14) == std::string::npos;
}

// UTF-16 code units of a UTF-8 string.
inline std::size_t Utf16Length(const std::string& s) {
  std::size_t units = 0;
  for (const char ch : s) {
    const auto c = static_cast<unsigned char>(ch);
    if ((c & 0xC0U) != 0x80U) {
      units += (c >= 0xF0U) ? 2U : 1U;
    }
  }
  return units;
}

inline std::string AsciiLower(std::string s) {
  for (char& c : s) {
    if (c >= 'A' && c <= 'Z') {
      c = static_cast<char>(c - 'A' + 'a');
    }
  }
  return s;
}

// Directory part of `part` including the trailing '/', or empty.
inline std::string DirOf(const std::string& part) {
  const std::size_t slash = part.rfind('/');
  return slash == std::string::npos ? std::string() : part.substr(0, slash + 1);
}

// Resolves `target` against `base_dir`, collapsing "." and "..".
inline std::string Resolve(const std::string& base_dir, const std::string& target) {
  const std::string full = (!target.empty() && target[0] == '/') ? target.substr(1) : base_dir + target;
  std::vector<std::string> out;
  std::size_t start = 0;
  while (start <= full.size()) {
    std::size_t end = full.find('/', start);
    if (end == std::string::npos) {
      end = full.size();
    }
    const std::string seg = full.substr(start, end - start);
    if (seg == "..") {
      if (!out.empty()) {
        out.pop_back();
      }
    } else if (!seg.empty() && seg != ".") {
      out.push_back(seg);
    }
    start = end + 1U;
  }
  std::string joined;
  for (std::size_t i = 0; i < out.size(); ++i) {
    joined += (i == 0 ? "" : "/") + out[i];
  }
  return joined;
}

/// Every worksheet's `<sheetFormatPr>` carries `defaultRowHeight`.
inline XlsxFindings CheckDefaultRowHeightPresent(io::ByteSpan bytes) {
  XlsxFindings out;
  Package pkg;
  if (!Open(pkg, bytes)) {
    return out;
  }
  for (const std::string& name : pkg.names) {
    pugi::xml_document doc;
    if (!IsWorksheetPart(name) || !Parse(pkg, name, doc)) {
      continue;
    }
    const pugi::xml_node fmt = doc.child("worksheet").child("sheetFormatPr");
    if (fmt && !fmt.attribute("defaultRowHeight")) {
      out.push_back("default_row_height_present: " + name + " has sheetFormatPr without defaultRowHeight");
    }
  }
  return out;
}

/// `<workbookView activeTab>` (default 0) names an existing sheet.
inline XlsxFindings CheckActiveTabInRange(io::ByteSpan bytes) {
  XlsxFindings out;
  Package pkg;
  pugi::xml_document doc;
  if (!Open(pkg, bytes) || !Parse(pkg, "xl/workbook.xml", doc)) {
    return out;
  }
  const pugi::xml_node wb = doc.child("workbook");
  unsigned sheets = 0;
  for (const pugi::xml_node s : wb.child("sheets").children("sheet")) {
    (void)s;
    ++sheets;
  }
  for (const pugi::xml_node view : wb.child("bookViews").children("workbookView")) {
    const unsigned active = view.attribute("activeTab").as_uint(0);
    if (active >= sheets) {
      out.push_back("active_tab_in_range: activeTab " + std::to_string(active) + " with " + std::to_string(sheets) +
                    " sheets");
    }
  }
  return out;
}

/// Every `<sheet sheetId>` is present, positive and unique.
inline XlsxFindings CheckSheetIdUnique(io::ByteSpan bytes) {
  XlsxFindings out;
  Package pkg;
  pugi::xml_document doc;
  if (!Open(pkg, bytes) || !Parse(pkg, "xl/workbook.xml", doc)) {
    return out;
  }
  std::set<unsigned> seen;
  for (const pugi::xml_node s : doc.child("workbook").child("sheets").children("sheet")) {
    const std::string name = s.attribute("name").as_string();
    const unsigned id = s.attribute("sheetId").as_uint(0);
    if (id == 0) {
      out.push_back("sheet_id_unique: sheet '" + name + "' has a missing or zero sheetId");
    } else if (!seen.insert(id).second) {
      out.push_back("sheet_id_unique: sheetId " + std::to_string(id) + " repeats");
    }
  }
  return out;
}

/// Sheet names are non-empty, at most 31 UTF-16 units, unique ignoring ASCII case.
inline XlsxFindings CheckSheetNameUnique(io::ByteSpan bytes) {
  XlsxFindings out;
  Package pkg;
  pugi::xml_document doc;
  if (!Open(pkg, bytes) || !Parse(pkg, "xl/workbook.xml", doc)) {
    return out;
  }
  std::set<std::string> seen;
  for (const pugi::xml_node s : doc.child("workbook").child("sheets").children("sheet")) {
    const std::string name = s.attribute("name").as_string();
    if (name.empty()) {
      out.push_back("sheet_name_unique: empty sheet name");
    } else if (Utf16Length(name) > 31U) {
      out.push_back("sheet_name_unique: '" + name + "' is longer than 31 characters");
    }
    if (!seen.insert(AsciiLower(name)).second) {
      out.push_back("sheet_name_unique: '" + name + "' repeats");
    }
  }
  return out;
}

/// Every internal relationship target resolves to a part in the package.
inline XlsxFindings CheckRelTargetExists(io::ByteSpan bytes) {
  XlsxFindings out;
  Package pkg;
  if (!Open(pkg, bytes)) {
    return out;
  }
  const std::set<std::string> parts(pkg.names.begin(), pkg.names.end());
  for (const std::string& name : pkg.names) {
    pugi::xml_document doc;
    if (!EndsWith(name, ".rels") || !Parse(pkg, name, doc)) {
      continue;
    }
    // The source part's directory is the parent of the "_rels" folder.
    const std::string rels_dir = DirOf(name);
    const std::string base =
        rels_dir.size() >= 6U && EndsWith(rels_dir, "_rels/") ? rels_dir.substr(0, rels_dir.size() - 6U) : rels_dir;
    for (const pugi::xml_node rel : doc.child("Relationships").children("Relationship")) {
      if (std::string_view(rel.attribute("TargetMode").as_string()) == "External") {
        continue;
      }
      const std::string target = rel.attribute("Target").as_string();
      if (parts.count(Resolve(base, target)) == 0) {
        out.push_back("rel_target_exists: " + name + " -> " + target);
      }
    }
  }
  return out;
}

/// Every part but `[Content_Types].xml` has an Override or a Default extension.
inline XlsxFindings CheckContentTypeCovered(io::ByteSpan bytes) {
  XlsxFindings out;
  Package pkg;
  pugi::xml_document doc;
  if (!Open(pkg, bytes) || !Parse(pkg, "[Content_Types].xml", doc)) {
    return out;
  }
  std::set<std::string> overrides;
  std::set<std::string> defaults;
  for (const pugi::xml_node n : doc.child("Types").children("Override")) {
    overrides.insert(n.attribute("PartName").as_string());
  }
  for (const pugi::xml_node n : doc.child("Types").children("Default")) {
    defaults.insert(AsciiLower(n.attribute("Extension").as_string()));
  }
  for (const std::string& name : pkg.names) {
    if (name == "[Content_Types].xml" || (!name.empty() && name.back() == '/')) {
      continue;
    }
    const std::size_t dot = name.rfind('.');
    const std::string ext = dot == std::string::npos ? std::string() : AsciiLower(name.substr(dot + 1U));
    if (overrides.count("/" + name) == 0 && defaults.count(ext) == 0) {
      out.push_back("content_type_covered: " + name);
    }
  }
  return out;
}

/// `<rowBreaks>` / `<colBreaks>` hold at most 1024 `<brk>` each.
inline XlsxFindings CheckPageBreaksMax1024(io::ByteSpan bytes) {
  XlsxFindings out;
  Package pkg;
  if (!Open(pkg, bytes)) {
    return out;
  }
  for (const std::string& name : pkg.names) {
    pugi::xml_document doc;
    if (!IsWorksheetPart(name) || !Parse(pkg, name, doc)) {
      continue;
    }
    for (const char* tag : {"rowBreaks", "colBreaks"}) {
      unsigned count = 0;
      for (const pugi::xml_node b : doc.child("worksheet").child(tag).children("brk")) {
        (void)b;
        ++count;
      }
      if (count > 1024U) {
        out.push_back("page_breaks_max_1024: " + name + " " + tag + " has " + std::to_string(count) + " breaks");
      }
    }
  }
  return out;
}

/// Every `<protectedRange name>` is non-empty and at most 255 UTF-16 units.
inline XlsxFindings CheckProtectedRangeTitle(io::ByteSpan bytes) {
  XlsxFindings out;
  Package pkg;
  if (!Open(pkg, bytes)) {
    return out;
  }
  for (const std::string& name : pkg.names) {
    pugi::xml_document doc;
    if (!IsWorksheetPart(name) || !Parse(pkg, name, doc)) {
      continue;
    }
    for (const pugi::xml_node r : doc.child("worksheet").child("protectedRanges").children("protectedRange")) {
      const std::string title = r.attribute("name").as_string();
      if (title.empty() || Utf16Length(title) > 255U) {
        out.push_back("protected_range_title: " + name + " title of " + std::to_string(Utf16Length(title)) + " units");
      }
    }
  }
  return out;
}

/// `baseColWidth` and every `<col width>` are at most 255, `defaultColWidth` at most 255.99609375.
inline XlsxFindings CheckColumnWidthMax255(io::ByteSpan bytes) {
  XlsxFindings out;
  Package pkg;
  if (!Open(pkg, bytes)) {
    return out;
  }
  for (const std::string& name : pkg.names) {
    pugi::xml_document doc;
    if (!IsWorksheetPart(name) || !Parse(pkg, name, doc)) {
      continue;
    }
    const pugi::xml_node ws = doc.child("worksheet");
    const pugi::xml_node fmt = ws.child("sheetFormatPr");
    // Excel opens a default width of 255 + 255/256 and refuses 256.
    for (const auto& [attr, max] : {std::pair<const char*, double>{"baseColWidth", 255.0},
                                    std::pair<const char*, double>{"defaultColWidth", 255.99609375}}) {
      if (fmt.attribute(attr).as_double() > max) {
        out.push_back(std::string("column_width_max_255: ") + name + " " + attr + " " +
                      fmt.attribute(attr).as_string());
      }
    }
    for (const pugi::xml_node col : ws.child("cols").children("col")) {
      if (col.attribute("width").as_double() > 255.0) {
        out.push_back("column_width_max_255: " + name + " col width " + col.attribute("width").as_string());
      }
    }
  }
  return out;
}

/// The sheet's outline summaries (`sheetFormatPr` outlineLevelRow / outlineLevelCol) are at most 7. Excel refuses
/// a summary of 8 but opens a row or column at level 8 under a summary of 7 or less.
inline XlsxFindings CheckOutlineLevelMax7(io::ByteSpan bytes) {
  XlsxFindings out;
  Package pkg;
  if (!Open(pkg, bytes)) {
    return out;
  }
  for (const std::string& name : pkg.names) {
    pugi::xml_document doc;
    if (!IsWorksheetPart(name) || !Parse(pkg, name, doc)) {
      continue;
    }
    const pugi::xml_node fmt = doc.child("worksheet").child("sheetFormatPr");
    for (const char* attr : {"outlineLevelRow", "outlineLevelCol"}) {
      if (fmt.attribute(attr).as_uint(0) > 7U) {
        out.push_back(std::string("outline_level_max_7: ") + name + " " + attr + " " + fmt.attribute(attr).as_string());
      }
    }
  }
  return out;
}

/// `<workbookProtection>` comes after every `mc:AlternateContent` and `xr:revisionPtr` sibling.
inline XlsxFindings CheckWorkbookProtectionOrder(io::ByteSpan bytes) {
  XlsxFindings out;
  Package pkg;
  pugi::xml_document doc;
  if (!Open(pkg, bytes) || !Parse(pkg, "xl/workbook.xml", doc)) {
    return out;
  }
  bool protection_seen = false;
  for (const pugi::xml_node child : doc.child("workbook").children()) {
    const std::string_view tag = child.name();
    if (tag == "workbookProtection") {
      protection_seen = true;
    } else if (protection_seen && (tag == "mc:AlternateContent" || tag == "xr:revisionPtr")) {
      out.push_back("workbook_protection_order: workbookProtection precedes " + std::string(tag));
    }
  }
  return out;
}

}  // namespace xlsx_validator_detail

using xlsx_validator_detail::CheckActiveTabInRange;
using xlsx_validator_detail::CheckColumnWidthMax255;
using xlsx_validator_detail::CheckContentTypeCovered;
using xlsx_validator_detail::CheckDefaultRowHeightPresent;
using xlsx_validator_detail::CheckOutlineLevelMax7;
using xlsx_validator_detail::CheckPageBreaksMax1024;
using xlsx_validator_detail::CheckProtectedRangeTitle;
using xlsx_validator_detail::CheckRelTargetExists;
using xlsx_validator_detail::CheckSheetIdUnique;
using xlsx_validator_detail::CheckSheetNameUnique;
using xlsx_validator_detail::CheckWorkbookProtectionOrder;

/// Runs every rule; empty means the package is valid.
inline XlsxFindings ValidateXlsxPackage(io::ByteSpan bytes) {
  XlsxFindings all;
  const auto add = [&all](const XlsxFindings& found) { all.insert(all.end(), found.begin(), found.end()); };
  add(CheckDefaultRowHeightPresent(bytes));
  add(CheckActiveTabInRange(bytes));
  add(CheckSheetIdUnique(bytes));
  add(CheckSheetNameUnique(bytes));
  add(CheckWorkbookProtectionOrder(bytes));
  add(CheckRelTargetExists(bytes));
  add(CheckContentTypeCovered(bytes));
  add(CheckPageBreaksMax1024(bytes));
  add(CheckProtectedRangeTitle(bytes));
  add(CheckColumnWidthMax255(bytes));
  add(CheckOutlineLevelMax7(bytes));
  return all;
}

}  // namespace formulon::test

#endif  // FORMULON_TESTS_SUPPORT_XLSX_PACKAGE_VALIDATOR_H_
