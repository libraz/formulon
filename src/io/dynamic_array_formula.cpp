#include "io/dynamic_array_formula.h"

#include <algorithm>
#include <functional>
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

// What a defined name's formula shows of it without evaluation.
xlsb::NameShape NameFormulaShape(std::string_view formula) {
  if (!formula.empty() && formula.front() == '=') {
    formula.remove_prefix(1);
  }
  Arena arena;
  const parser::AstNode* body = parser::parse_strict(formula, arena);
  xlsb::NameShape shape;
  if (body == nullptr) {
    return shape;
  }
  if (body->kind() == parser::NodeKind::Literal && body->as_literal().is_number()) {
    const double d = body->as_literal().as_number();
    if (d >= 1.0 && d <= 65535.0 && d == static_cast<double>(static_cast<std::uint32_t>(d))) {
      shape.positive_int = static_cast<std::uint32_t>(d);
    }
  }
  if (body->kind() == parser::NodeKind::UnaryOp) {
    body = &body->as_unary_operand();
  }
  auto extent = [&shape](const parser::Reference& a, const parser::Reference& b) {
    shape.rows = a.is_full_col ? Sheet::kMaxRows : (a.row > b.row ? a.row - b.row : b.row - a.row) + 1U;
    shape.cols = a.is_full_row ? Sheet::kMaxCols : (a.col > b.col ? a.col - b.col : b.col - a.col) + 1U;
  };
  switch (body->kind()) {
    case parser::NodeKind::Literal:
    case parser::NodeKind::ErrorLiteral:
      shape.scalar = true;
      break;
    case parser::NodeKind::Ref:
      extent(body->as_ref(), body->as_ref());
      shape.scalar = shape.rows == 1U && shape.cols == 1U;
      break;
    case parser::NodeKind::Ref3D:
      shape.scalar = !body->as_ref3d_is_range() && body->as_ref3d_sheet_begin() == body->as_ref3d_sheet_end();
      break;
    case parser::NodeKind::RangeOp:
      if (body->as_range_lhs().kind() == parser::NodeKind::Ref &&
          body->as_range_rhs().kind() == parser::NodeKind::Ref) {
        extent(body->as_range_lhs().as_ref(), body->as_range_rhs().as_ref());
      }
      break;
    default:
      break;
  }
  return shape;
}

// The defined name a `NameRef`, `[0]!Name` or name-call node refers to.
std::string_view NameOf(const parser::AstNode& node) {
  switch (node.kind()) {
    case parser::NodeKind::NameRef:
      return node.as_name();
    case parser::NodeKind::Call:
      return node.as_call_name();
    default:
      return node.as_external_ref_name();
  }
}

// Index in `names` of `name` as a formula in `scope` (a sheet index, or -1
// for workbook scope) resolves it: sheet-local before workbook-wide;
// `names.size()` when undefined.
std::size_t FindDefinedName(const std::vector<DefinedName>& names, std::string_view name, std::int32_t scope) {
  std::size_t found = names.size();
  for (std::size_t i = 0; i < names.size(); ++i) {
    if (!strings::case_insensitive_eq(names[i].name, name)) {
      continue;
    }
    if (names[i].local_sheet_id == scope) {
      return i;
    }
    if (names[i].local_sheet_id < 0) {
      found = i;
    }
  }
  return found;
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

std::vector<xlsb::NameShape> defined_name_shapes(const Workbook& wb) {
  const std::vector<DefinedName>& names = wb.defined_names();
  std::vector<xlsb::NameShape> shapes(names.size());
  for (std::size_t i = 0; i < names.size(); ++i) {
    shapes[i] = NameFormulaShape(names[i].formula);
    shapes[i].defined = true;
  }
  // fCalcExp and constant recalculation follow the names a formula refers
  // to, so they resolve in dependency order; a name reached again while it
  // is being resolved counts as clear.
  enum class State : std::uint8_t { kPending, kResolving, kDone };
  std::vector<State> state(names.size(), State::kPending);
  std::function<void(std::size_t)> resolve = [&](std::size_t i) {
    if (state[i] != State::kPending) {
      return;
    }
    state[i] = State::kResolving;
    std::string_view formula = names[i].formula;
    if (!formula.empty() && formula.front() == '=') {
      formula.remove_prefix(1);
    }
    const xlsb::NameShapes referenced = [&](const parser::AstNode& node) {
      const std::size_t found = FindDefinedName(names, NameOf(node), names[i].local_sheet_id);
      if (found >= names.size()) {
        return xlsb::NameShape{};
      }
      resolve(found);
      xlsb::NameShape shape = shapes[found];
      if (state[found] != State::kDone) {
        shape.calc_exp = shape.always_calculates = false;
      }
      return shape;
    };
    Arena arena;
    if (const parser::AstNode* body = parser::parse_strict(formula, arena); body != nullptr) {
      shapes[i].calc_exp = xlsb::name_sets_calc_exp(*body, referenced);
      shapes[i].always_calculates = xlsb::formula_always_calculates(*body, referenced);
    }
    state[i] = State::kDone;
  };
  for (std::size_t i = 0; i < names.size(); ++i) {
    resolve(i);
  }
  return shapes;
}

xlsb::NameShapes name_shapes(const Workbook& wb, std::size_t sheet_index) {
  return [&wb, sheet_index, shapes = defined_name_shapes(wb)](const parser::AstNode& node) {
    const std::string_view qualifier =
        node.kind() == parser::NodeKind::NameRef ? node.as_name_sheet() : std::string_view();
    const auto scope = static_cast<std::int32_t>(qualifier.empty() ? sheet_index : wb.sheet_index_by_name(qualifier));
    const std::size_t found = FindDefinedName(wb.defined_names(), NameOf(node), scope);
    return found < shapes.size() ? shapes[found] : xlsb::NameShape{};
  };
}

bool formula_cell_always_calculates(const Sheet& sheet, std::uint32_t row, std::uint32_t col, const Cell& cell,
                                    const xlsb::NameShapes& names) {
  if (cell.formula_text.empty()) {
    return false;
  }
  if (is_dynamic_array_formula(cell)) {
    const std::vector<CellAddress> blocked = sheet.blocked_spill_anchors_intersecting(row, col, 1U, 1U);
    if (std::any_of(blocked.begin(), blocked.end(),
                    [&](const CellAddress& a) { return a.row == row && a.col == col; })) {
      return true;
    }
  }
  std::string_view formula = cell.formula_text;
  if (formula.front() == '=') {
    formula.remove_prefix(1);
  }
  Arena arena;
  parser::Parser parser(formula, arena);
  const parser::AstNode* root = parser.parse();
  return root != nullptr && parser.errors().empty() && xlsb::formula_always_calculates(*root, names);
}

bool entered_as_dynamic_array(const Workbook& wb, std::size_t sheet_index, const parser::AstNode& root) {
  return xlsb::formula_is_dynamic_array(root, name_shapes(wb, sheet_index));
}

}  // namespace io
}  // namespace formulon
