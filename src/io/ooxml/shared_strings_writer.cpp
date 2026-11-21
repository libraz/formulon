//
// Implementation of the OOXML shared-string table builder and writer.

#include "io/ooxml/shared_strings_writer.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml_defs.h"
#include "io/phonetic_pr.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "io/xsd_int.h"
#include "passthrough_part.h"
#include "phonetic.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/a1_ref.h"
#include "utils/index_sort.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {

std::string SharedStringKey(std::string_view text, const std::vector<PhoneticRun>& phonetic,
                            PhoneticProperties phonetic_props) {
  std::string key;
  key.reserve(text.size() + 32U + phonetic.size() * 16U);
  key.append(std::to_string(text.size()));
  key.push_back(':');
  key.append(text);
  // Length-prefix each run's kana for the same reason the surface text is
  // prefixed: without it, adjacent fields could be re-cut into the same
  // byte sequence and collide two distinct annotations onto one <si>.
  for (const PhoneticRun& run : phonetic) {
    key.push_back(';');
    key.append(std::to_string(run.sb));
    key.push_back(',');
    key.append(std::to_string(run.eb));
    key.push_back(',');
    key.append(std::to_string(run.text.size()));
    key.push_back(':');
    key.append(run.text);
  }
  if (!phonetic.empty()) {
    key.push_back('|');
    key.append(std::to_string(phonetic_props.font_id));
    key.push_back(',');
    key.append(std::to_string(phonetic_props.type));
    key.push_back(',');
    key.append(std::to_string(phonetic_props.alignment));
  }
  return key;
}

std::uint32_t SharedStrings::intern(std::string_view text, const std::vector<PhoneticRun>& phonetic,
                                    PhoneticProperties phonetic_props) {
  ++total_count_;
  std::string key = SharedStringKey(text, phonetic, phonetic_props);
  const auto found = index_.find(key);
  if (found != index_.end()) {
    return found->second;
  }
  const std::uint32_t index = static_cast<std::uint32_t>(entries_.size());
  entries_.push_back(SharedStringEntry{std::string(text), phonetic, phonetic_props});
  index_.emplace(std::move(key), index);
  return index;
}

std::uint32_t SharedStrings::index_of(std::string_view text, const std::vector<PhoneticRun>& phonetic,
                                      PhoneticProperties phonetic_props) const {
  const auto found = index_.find(SharedStringKey(text, phonetic, phonetic_props));
  // `BuildSharedStrings` visits the same literal-cell set before any
  // worksheet is written, so this fallback is unreachable in normal use.
  // Keep the writer exception-free even if a future caller violates that
  // internal ordering contract.
  return found == index_.end() ? 0U : found->second;
}

void SharedStrings::pin(std::uint32_t index, std::string_view text, const std::vector<PhoneticRun>& phonetic,
                        PhoneticProperties phonetic_props) {
  ++total_count_;
  if (index < entries_.size() && !entries_[index].text.empty()) {
    return;
  }
  if (index >= entries_.size()) {
    entries_.resize(static_cast<std::size_t>(index) + 1U);
  }
  entries_[index] = SharedStringEntry{std::string(text), phonetic, phonetic_props};
  index_.emplace(SharedStringKey(text, phonetic, phonetic_props), index);
}

namespace {

/// Pins the strings an opaque sheet's verbatim part names through `t="s"` cells at the indices it stores, taking
/// each cell's text from the model, so the part still resolves against the rebuilt table.
void PinOpaqueSheetStrings(const Workbook& workbook, const Sheet& sheet, SharedStrings& table) {
  const PassthroughPart* part = ooxml::find_passthrough_part(workbook, sheet.opaque_ooxml_part_path());
  pugi::xml_document doc;
  if (part == nullptr || !strings::ends_with(part->path, ".xml") ||
      !doc.load_buffer(part->bytes.data(), part->bytes.size())) {
    return;
  }
  for (const pugi::xml_node row : doc.document_element().child("sheetData").children("row")) {
    for (const pugi::xml_node c : row.children("c")) {
      std::uint32_t index = 0;
      std::uint32_t r = 0;
      std::uint32_t col = 0;
      if (std::string_view(c.attribute("t").value()) != "s" || !parse_xsd_nonneg_int(c.child_value("v"), &index) ||
          !a1::parse_a1_ref(c.attribute("r").value(), &r, &col)) {
        continue;
      }
      const Cell* cell = sheet.cell_at(r, col);
      if (cell != nullptr && cell->cached_value.is_text()) {
        table.pin(index, cell->cached_value.as_text(), cell->phonetic_runs, cell->phonetic_props);
      } else {
        table.pin(index, std::string_view{}, {}, PhoneticProperties{});
      }
    }
  }
}

}  // namespace

SharedStrings BuildSharedStrings(const Workbook& workbook) {
  SharedStrings strings;
  for (std::size_t i = 0; i < workbook.sheet_count(); ++i) {
    if (workbook.sheet(i).is_opaque_ooxml_sheet()) {
      PinOpaqueSheetStrings(workbook, workbook.sheet(i), strings);
    }
  }
  for (std::size_t i = 0; i < workbook.sheet_count(); ++i) {
    const Sheet& sheet = workbook.sheet(i);
    if (sheet.is_opaque_ooxml_sheet()) {
      continue;
    }
    std::vector<std::uint32_t> rows;
    rows.reserve(sheet.rows().size());
    for (const auto& [row, cells] : sheet.rows()) {
      (void)cells;
      rows.push_back(row);
    }
    sort_ascending(rows);
    for (const std::uint32_t row : rows) {
      const auto row_it = sheet.rows().find(row);
      if (row_it == sheet.rows().end()) {
        continue;
      }
      const RowCells& cells = row_it->second;
      for (const auto& [col, cell] : cells) {
        (void)col;
        if (cell.formula_text.empty() && cell.cached_value.is_text()) {
          strings.intern(cell.cached_value.as_text(), cell.phonetic_runs, cell.phonetic_props);
        }
      }
    }
  }
  return strings;
}

std::string WriteSharedStrings(const SharedStrings& strings) {
  std::string out;
  out.reserve(128U + strings.entries().size() * 32U);
  out.append(kXmlDecl);
  out.append("<sst");
  append_xml_ns_decl(out, "", kNsSpreadsheetMl);
  out.append(" count=\"");
  out.append(std::to_string(strings.total_count()));
  out.append("\" uniqueCount=\"");
  out.append(std::to_string(strings.entries().size()));
  out.append("\">");
  for (const SharedStringEntry& entry : strings.entries()) {
    out.append("<si><t xml:space=\"preserve\">");
    AppendXmlEscaped(out, entry.text);
    out.append("</t>");
    append_phonetic_runs(out, entry.phonetic, entry.phonetic_props);
    out.append("</si>");
  }
  out.append("</sst>\n");
  return out;
}

}  // namespace io
}  // namespace formulon
