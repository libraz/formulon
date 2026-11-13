#include "io/sheet_phonetic.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>

#include "io/phonetic_pr.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "sheet_passthrough.h"
#include "workbook.h"

namespace formulon {
namespace io {

std::optional<PhoneticProperties> sheet_phonetic_properties(const Sheet& sheet) {
  const auto slot = static_cast<std::uint32_t>(worksheet_child::slot_of("phoneticPr"));
  for (const WorksheetRawChild& child : sheet.raw_extensions()) {
    pugi::xml_document doc;
    if (child.slot == slot && doc.load_buffer(child.xml.data(), child.xml.size()) && doc.child("phoneticPr")) {
      return read_phonetic_properties(doc);
    }
  }
  return std::nullopt;
}

std::string sheet_phonetic_xml(const PhoneticProperties& props) {
  std::string xml = "<phoneticPr fontId=\"" + std::to_string(props.font_id) + "\"";
  if (props.type != 1U) {
    append_xml_attr(xml, "type", phonetic_type_name(props.type));
  }
  if (props.alignment != 1U) {
    append_xml_attr(xml, "alignment", phonetic_alignment_name(props.alignment));
  }
  xml.append("/>");
  return xml;
}

void set_sheet_phonetic_properties(Sheet& sheet, const PhoneticProperties& props) {
  std::string xml = sheet_phonetic_xml(props);
  WorksheetRawExtensions& children = sheet.mutable_raw_extensions();
  const auto slot = static_cast<std::uint32_t>(worksheet_child::slot_of("phoneticPr"));
  children.erase(std::remove_if(children.begin(), children.end(),
                                [slot](const WorksheetRawChild& child) { return child.slot == slot; }),
                 children.end());
  insert_raw_child(children, WorksheetRawChild{slot, std::move(xml)});
}

void clamp_sheet_phonetic_fonts(Workbook& wb) {
  const std::size_t font_count = wb.styles().fonts.size();
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    std::optional<PhoneticProperties> props = sheet_phonetic_properties(wb.sheet(i));
    if (props && props->font_id != 0U && props->font_id >= font_count) {
      props->font_id = 0U;
      set_sheet_phonetic_properties(wb.sheet(i), *props);
    }
  }
}

}  // namespace io
}  // namespace formulon
