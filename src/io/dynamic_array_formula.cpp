#include "io/dynamic_array_formula.h"

#include <string_view>
#include <utility>

#include "io/xml_utils.h"
#include "parser/parser.h"
#include "pugixml.hpp"
#include "utils/arena.h"
#include "utils/strings.h"

namespace formulon {
namespace io {
namespace {

// A defined name's formula standing for one value: a one-cell reference or a constant.
bool NameFormulaIsScalar(std::string_view formula) {
  if (!formula.empty() && formula.front() == '=') {
    formula.remove_prefix(1);
  }
  Arena arena;
  const parser::AstNode* body = parser::parse_strict(formula, arena);
  if (body != nullptr && body->kind() == parser::NodeKind::UnaryOp) {
    body = &body->as_unary_operand();
  }
  if (body == nullptr) {
    return false;
  }
  switch (body->kind()) {
    case parser::NodeKind::Literal:
    case parser::NodeKind::ErrorLiteral:
      return true;
    case parser::NodeKind::Ref:
      return !body->as_ref().is_full_col && !body->as_ref().is_full_row;
    case parser::NodeKind::Ref3D:
      return !body->as_ref3d_is_range() && body->as_ref3d_sheet_begin() == body->as_ref3d_sheet_end();
    default:
      return false;
  }
}

}  // namespace

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

xlsb::NameIsScalar legacy_name_shapes(const Workbook& wb, std::size_t sheet_index) {
  const std::vector<DefinedName>& names = wb.defined_names();
  std::vector<bool> scalar(names.size());
  for (std::size_t i = 0; i < names.size(); ++i) {
    scalar[i] = NameFormulaIsScalar(names[i].formula);
  }
  return [&wb, sheet_index, scalar = std::move(scalar)](const parser::AstNode& node) {
    const std::vector<DefinedName>& defined = wb.defined_names();
    const bool name_ref = node.kind() == parser::NodeKind::NameRef;
    const std::string_view name = name_ref ? node.as_name() : node.as_external_ref_name();
    const std::string_view qualifier = name_ref ? node.as_name_sheet() : std::string_view();
    const auto scope = static_cast<std::int32_t>(qualifier.empty() ? sheet_index : wb.sheet_index_by_name(qualifier));
    std::size_t found = defined.size();
    for (std::size_t i = 0; i < defined.size() && i < scalar.size(); ++i) {
      if (!strings::case_insensitive_eq(defined[i].name, name)) {
        continue;
      }
      if (defined[i].local_sheet_id == scope) {
        found = i;
        break;
      }
      if (defined[i].local_sheet_id < 0) {
        found = i;
      }
    }
    return found < scalar.size() && scalar[found];
  };
}

}  // namespace io
}  // namespace formulon
