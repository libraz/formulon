#include "io/dynamic_array_formula.h"

#include <string_view>

#include "io/xml_utils.h"
#include "pugixml.hpp"

namespace formulon {
namespace io {

bool is_dynamic_array_formula(const Cell& cell) {
  return cell.dynamic_array && !cell.formula_text.empty();
}

bool has_dynamic_array_formula(const Workbook& wb) {
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    for (const auto& [row, cells] : wb.sheet(i).rows()) {
      (void)row;
      for (std::uint32_t col = 0; col < cells.size(); ++col) {
        if (is_dynamic_array_formula(cells[col])) {
          return true;
        }
      }
    }
  }
  return false;
}

std::uint32_t xldapr_cell_metadata_index(const std::vector<std::uint8_t>& metadata_xml) {
  pugi::xml_document doc;
  if (!load_xml_buffer(doc, metadata_xml, "dynamic_array_formula", "xl/metadata.xml")) {
    return 0U;
  }
  const pugi::xml_node root = doc.child("metadata");
  std::uint32_t xldapr_ordinal = 0U;
  std::uint32_t type_ordinal = 0U;
  for (pugi::xml_node type = root.child("metadataTypes").child("metadataType"); type;
       type = type.next_sibling("metadataType")) {
    ++type_ordinal;
    if (xldapr_ordinal == 0U && std::string_view(type.attribute("name").value()) == "XLDAPR") {
      xldapr_ordinal = type_ordinal;
    }
  }
  if (xldapr_ordinal == 0U) {
    return 0U;
  }
  std::uint32_t bk_index = 0U;
  for (pugi::xml_node bk = root.child("cellMetadata").child("bk"); bk; bk = bk.next_sibling("bk")) {
    ++bk_index;
    for (pugi::xml_node rc = bk.child("rc"); rc; rc = rc.next_sibling("rc")) {
      if (rc.attribute("t").as_uint(0U) == xldapr_ordinal) {
        return bk_index;
      }
    }
  }
  return 0U;
}

void apply_loaded_dynamic_array_marks(Sheet& sheet, const std::unordered_set<std::uint64_t>& marked) {
  for (const CellAddress address : sheet.formula_cells_in(0U, 0U, Sheet::kMaxRows - 1U, Sheet::kMaxCols - 1U)) {
    sheet.set_cell_dynamic_array(address.row, address.col,
                                 marked.count(dynamic_array_cell_key(address.row, address.col)) != 0U);
  }
}

}  // namespace io
}  // namespace formulon
