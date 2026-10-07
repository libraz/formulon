//
// Implementation of the slot-class and result-shape model shared by the Ptg
// encoder and the formula-flag analyses. See `io/xlsb/ptg_operand_class.h`.

#include "io/xlsb/ptg_operand_class.h"

#include <cstdint>
#include <optional>
#include <string_view>
#include <utility>

#include "io/future_functions.h"
#include "io/xlsb/func_id_table.h"
#include "io/xlsb/ptg_targets.h"
#include "parser/ast.h"
#include "utils/strings.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace detail {
namespace {

/// Built-ins that return one of their `S` arguments, or an error in its place.
bool PassesArgumentThrough(std::string_view name) {
  for (const std::string_view fn : {"IF", "CHOOSE", "IFERROR", "IFNA"}) {
    if (strings::case_insensitive_eq(name, fn)) {
      return true;
    }
  }
  return false;
}

}  // namespace

std::uint8_t SlotClass(Slot slot, bool area, PtgEvaluation evaluation) {
  if (slot.operand && (evaluation == PtgEvaluation::kLegacy || evaluation == PtgEvaluation::kLegacyArray)) {
    // A legacy formula intersects an operand unless the parameter forces arrays.
    return slot.letter == 'F' ? kPtgArrayClass : kPtgValueClass;
  }
  if (slot.operand && evaluation == PtgEvaluation::kConditionalFormat && slot.letter != 0 && slot.letter != 'V') {
    return kPtgArrayClass;
  }
  if (slot.operand) {
    switch (slot.letter) {
      case 'S':
      case 'I':
      case 'Y':
        return area ? kPtgArrayClass : kPtgValueClass;
      case 'F':
      case 'X':
        return kPtgArrayClass;
      default:
        return kPtgValueClass;
    }
  }
  switch (slot.letter) {
    case 'V':
      return kPtgValueClass;
    case 'F':
      return kPtgArrayClass;
    case 'I':
      return area ? kPtgArrayClass : kPtgValueClass;
    case 'Y':
      return area ? kPtgValueClass : kPtgReferenceClass;
    default:
      return kPtgReferenceClass;
  }
}

bool IsBuiltin(std::string_view name) {
  return lookup_func_by_name(name) != nullptr || UsesHiddenNameRoute(name);
}

Shape Shapes::of(const parser::AstNode& node) {
  using parser::NodeKind;
  switch (node.kind()) {
    case NodeKind::Ref:
      return node.as_ref().is_full_col || node.as_ref().is_full_row ? Shape::kReference : Shape::kScalar;
    case NodeKind::RangeOp:
    case NodeKind::SpillRef:
      return Shape::kReference;
    case NodeKind::Ref3D:
      return node.as_ref3d_is_range() || node.as_ref3d_sheet_begin() != node.as_ref3d_sheet_end() ? Shape::kReference
                                                                                                  : Shape::kScalar;
    case NodeKind::IntersectOp:
      // One cell on either side bounds the intersection to it.
      return of(node.as_intersect_lhs()) != Shape::kScalar && of(node.as_intersect_rhs()) != Shape::kScalar
                 ? Shape::kReference
                 : Shape::kScalar;
    case NodeKind::NameRef:
      return name_shape(node);
    case NodeKind::ExternalRef:
      if (parser::is_self_book_name_ref(node)) {
        return name_shape(node);
      }
      return node.as_external_ref_is_range() || !node.as_external_ref_name().empty() ||
                     !node.as_external_ref_sheet_end().empty() || node.as_external_ref_cell().is_full_col ||
                     node.as_external_ref_cell().is_full_row
                 ? Shape::kReference
                 : Shape::kScalar;
    case NodeKind::StructuredRef:
      return node.as_structured_ref_modifier() == parser::StructuredRefModifier::At ? Shape::kScalar
                                                                                    : Shape::kReference;
    case NodeKind::UnaryOp:
      return !lifting_ || of(node.as_unary_operand()) == Shape::kScalar ? Shape::kScalar : Shape::kArray;
    case NodeKind::BinaryOp:
      return !lifting_ || (of(node.as_binary_lhs()) == Shape::kScalar && of(node.as_binary_rhs()) == Shape::kScalar)
                 ? Shape::kScalar
                 : Shape::kArray;
    case NodeKind::ArrayLiteral:
      return lifting_ ? Shape::kArray : Shape::kScalar;
    case NodeKind::LambdaCall:
      return Shape::kArray;
    case NodeKind::LetBinding: {
      const std::uint32_t n = node.as_let_binding_count();
      for (std::uint32_t i = 0; i < n; ++i) {
        bind(node.as_let_binding_name(i), of(node.as_let_binding_expr(i)));
      }
      const Shape body = of(node.as_let_body());
      unbind(n);
      return body;
    }
    case NodeKind::Call:
      return call_shape(node);
    case NodeKind::Literal:
    case NodeKind::ErrorLiteral:
    case NodeKind::ErrorPlaceholder:
    case NodeKind::UnionOp:
    case NodeKind::ImplicitIntersection:
    case NodeKind::Lambda:
      return Shape::kScalar;
  }
  return Shape::kReference;
}

const std::pair<std::string_view, Shape>* Shapes::binding(const parser::AstNode& node) const {
  if (node.kind() != parser::NodeKind::NameRef || !node.as_name_sheet().empty()) {
    return nullptr;
  }
  for (auto it = scope_.rbegin(); it != scope_.rend(); ++it) {
    if (strings::case_insensitive_eq(it->first, node.as_name())) {
      return &*it;
    }
  }
  return nullptr;
}

Shape Shapes::name_shape(const parser::AstNode& node) const {
  if (const auto* b = binding(node)) {
    return b->second;
  }
  return defined(node).scalar ? Shape::kScalar : Shape::kReference;
}

bool Shapes::positive_int(const parser::AstNode& node, std::uint32_t want) const {
  std::uint32_t n = 0;
  if (node.kind() == parser::NodeKind::Literal && node.as_literal().is_number()) {
    const double d = node.as_literal().as_number();
    n = d >= 1.0 && d <= 65535.0 && d == static_cast<double>(static_cast<std::uint32_t>(d))
            ? static_cast<std::uint32_t>(d)
            : 0U;
  } else if (node.kind() == parser::NodeKind::NameRef && binding(node) == nullptr) {
    n = defined(node).positive_int;
  }
  return n != 0U && (want == 0U || n == want);
}

std::optional<std::pair<std::uint32_t, std::uint32_t>> Shapes::extent(const parser::AstNode& node) const {
  constexpr std::uint32_t kRows = 1048576U;
  constexpr std::uint32_t kCols = 16384U;
  auto one = [&](const parser::Reference& r) {
    return std::make_pair(r.is_full_col ? kRows : 1U, r.is_full_row ? kCols : 1U);
  };
  switch (node.kind()) {
    case parser::NodeKind::Ref:
      return one(node.as_ref());
    case parser::NodeKind::RangeOp: {
      const parser::AstNode& lhs = node.as_range_lhs();
      const parser::AstNode& rhs = node.as_range_rhs();
      if (lhs.kind() != parser::NodeKind::Ref || rhs.kind() != parser::NodeKind::Ref) {
        return std::nullopt;
      }
      const parser::Reference& a = lhs.as_ref();
      const parser::Reference& b = rhs.as_ref();
      const std::uint32_t rows = a.is_full_col ? kRows : (a.row > b.row ? a.row - b.row : b.row - a.row) + 1U;
      const std::uint32_t cols = a.is_full_row ? kCols : (a.col > b.col ? a.col - b.col : b.col - a.col) + 1U;
      return std::make_pair(rows, cols);
    }
    case parser::NodeKind::NameRef: {
      if (binding(node) != nullptr) {
        return std::nullopt;
      }
      const NameShape s = defined(node);
      if (s.rows == 0U) {
        return std::nullopt;
      }
      return std::make_pair(s.rows, s.cols);
    }
    default:
      return std::nullopt;
  }
}

Shape Shapes::call_shape(const parser::AstNode& node) {
  const std::string_view name = canonical_function_name(node.as_call_name());
  if (bound(node.as_call_name()) || !IsBuiltin(name)) {
    return Shape::kArray;  // a LAMBDA through a name, or a function Excel does not know
  }
  const std::uint32_t arity = node.as_call_arity();
  auto is = [&](std::string_view fn) { return strings::case_insensitive_eq(name, fn); };
  Shape passed = Shape::kScalar;
  for (std::uint32_t i = 0; i < arity; ++i) {
    const char letter = xlsb_parameter_class(name, i);
    const Shape arg = of(node.as_call_arg(i));
    if (lifting_ && lifts_argument(name, i, letter, arg)) {
      return Shape::kArray;
    }
    // Without lifting, a multi-valued `I` argument still reaches the result
    // of a call that passes its arguments through (`IF(A1:A2,1,0)`).
    if ((letter == 'S' || (!lifting_ && letter == 'I')) && arg != Shape::kScalar && passed != Shape::kArray) {
      passed = arg;
    }
  }
  if (PassesArgumentThrough(name)) {
    return passed;
  }
  if (is("INDEX")) {
    // A reference indexed by positive integer constants is one cell.
    const Shape source = arity == 0U ? Shape::kReference : of(node.as_call_arg(0));
    if (source != Shape::kReference) {
      return source;
    }
    for (std::uint32_t i = 1; i < arity; ++i) {
      if (!positive_int(node.as_call_arg(i))) {
        return Shape::kReference;
      }
    }
    return arity < 2U ? Shape::kReference : Shape::kScalar;
  }
  if (is("OFFSET")) {
    // One cell from a one-cell base whose height and width are 1 or absent.
    if (arity < 3U || of(node.as_call_arg(0)) != Shape::kScalar ||
        (arity >= 4U && !positive_int(node.as_call_arg(3), 1U)) ||
        (arity >= 5U && !positive_int(node.as_call_arg(4), 1U))) {
      return Shape::kReference;
    }
    return Shape::kScalar;
  }
  if (is("INDIRECT")) {
    return Shape::kReference;
  }
  if (is("ROW") || is("COLUMN")) {
    return arity != 0U && of(node.as_call_arg(0)) != Shape::kScalar ? Shape::kArray : Shape::kScalar;
  }
  if (is("XLOOKUP")) {
    // One cell when the return array runs along the lookup array: one
    // column for a column (or a single cell), one row for a row.
    if (arity < 3U || passed != Shape::kScalar) {
      return Shape::kReference;
    }
    const auto lookup = extent(node.as_call_arg(1));
    const auto result = extent(node.as_call_arg(2));
    if (!lookup || !result) {
      return Shape::kReference;
    }
    if (lookup->second == 1U) {
      return result->second == 1U ? Shape::kScalar : Shape::kReference;
    }
    if (lookup->first == 1U) {
      return result->first == 1U ? Shape::kScalar : Shape::kReference;
    }
    return Shape::kReference;
  }
  return xlsb_returns_array(name) ? Shape::kArray : Shape::kScalar;
}

}  // namespace detail
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
