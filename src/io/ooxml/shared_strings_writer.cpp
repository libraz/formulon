//
// Implementation of the OOXML shared-string table builder and writer.

#include "io/ooxml/shared_strings_writer.h"

#include <algorithm>
#include <cstdint>
#include <string>
#include <string_view>

#include "cell.h"
#include "eval/utf8_length.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {

std::string SharedStrings::key_for(std::string_view text, std::string_view phonetic) {
  std::string key;
  key.reserve(text.size() + phonetic.size() + 32U);
  key.append(std::to_string(text.size()));
  key.push_back(':');
  key.append(text);
  key.append(phonetic);
  return key;
}

std::uint32_t SharedStrings::intern(std::string_view text, std::string_view phonetic) {
  ++total_count_;
  const std::string key = key_for(text, phonetic);
  const auto found = index_.find(key);
  if (found != index_.end()) {
    return found->second;
  }
  const std::uint32_t index = static_cast<std::uint32_t>(entries_.size());
  entries_.push_back(SharedStringEntry{std::string(text), std::string(phonetic)});
  index_.emplace(std::move(key), index);
  return index;
}

std::uint32_t SharedStrings::index_of(std::string_view text, std::string_view phonetic) const {
  const auto found = index_.find(key_for(text, phonetic));
  // `BuildSharedStrings` visits the same literal-cell set before any
  // worksheet is written, so this fallback is unreachable in normal use.
  // Keep the writer exception-free even if a future caller violates that
  // internal ordering contract.
  return found == index_.end() ? 0U : found->second;
}

SharedStrings BuildSharedStrings(const Workbook& workbook) {
  SharedStrings strings;
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
    std::sort(rows.begin(), rows.end());
    for (const std::uint32_t row : rows) {
      const auto row_it = sheet.rows().find(row);
      if (row_it == sheet.rows().end()) {
        continue;
      }
      const std::vector<Cell>& cells = row_it->second;
      for (const Cell& cell : cells) {
        if (cell.formula_text.empty() && cell.cached_value.is_text()) {
          strings.intern(cell.cached_value.as_text(), cell.phonetic_text);
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
  out.append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"");
  out.append(std::to_string(strings.total_count()));
  out.append("\" uniqueCount=\"");
  out.append(std::to_string(strings.entries().size()));
  out.append("\">");
  for (const SharedStringEntry& entry : strings.entries()) {
    out.append("<si><t xml:space=\"preserve\">");
    AppendXmlEscaped(out, entry.text);
    out.append("</t>");
    if (!entry.phonetic.empty()) {
      out.append("<rPh sb=\"0\" eb=\"");
      out.append(std::to_string(eval::utf16_units_in(entry.text)));
      out.append("\"><t xml:space=\"preserve\">");
      AppendXmlEscaped(out, entry.phonetic);
      out.append("</t></rPh>");
    }
    out.append("</si>");
  }
  out.append("</sst>\n");
  return out;
}

}  // namespace io
}  // namespace formulon
