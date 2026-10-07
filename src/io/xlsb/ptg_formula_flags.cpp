//
// The static judgements behind a formula record's flags: the dynamic-array
// mark, `fCalcExp`, always-calculate, a volatile call, and where a legacy
// formula shows `@`.
// Declared in `io/xlsb/ptg_writer.h`.

#include <algorithm>
#include <cstdint>
#include <string_view>
#include <unordered_map>
#include <vector>

#include "io/future_functions.h"
#include "io/xlsb/func_id_table.h"
#include "io/xlsb/ptg_operand_class.h"
#include "io/xlsb/ptg_writer.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "utils/strings.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using detail::ArrayScope;
using detail::IsBuiltin;
using detail::kPtgValueClass;
using detail::Shape;
using detail::Shapes;
using detail::Slot;
using detail::SlotClass;

/// Walker behind `formula_is_dynamic_array`: true when `node`, at `slot`
/// (`root` for the formula's own node), is multi-valued where Excel would
/// otherwise take one value: the root, an operator's operand, or an argument
/// its parameter evaluates element by element (`Shapes::lifts`).
bool MarksDynamic(Shapes& shapes, const parser::AstNode& node, Slot slot, bool root) {
  using parser::NodeKind;
  const Shape shape = shapes.of(node);
  if (shape != Shape::kScalar && (root || slot.operand || Shapes::lifts(slot.letter, shape))) {
    return true;
  }
  switch (node.kind()) {
    case NodeKind::UnaryOp:
      return MarksDynamic(shapes, node.as_unary_operand(), Slot{slot.letter, true}, false);
    case NodeKind::BinaryOp:
      return MarksDynamic(shapes, node.as_binary_lhs(), Slot{slot.letter, true}, false) ||
             MarksDynamic(shapes, node.as_binary_rhs(), Slot{slot.letter, true}, false);
    case NodeKind::RangeOp:
      return MarksDynamic(shapes, node.as_range_lhs(), Slot{'R', false}, false) ||
             MarksDynamic(shapes, node.as_range_rhs(), Slot{'R', false}, false);
    case NodeKind::IntersectOp:
      return MarksDynamic(shapes, node.as_intersect_lhs(), Slot{'R', false}, false) ||
             MarksDynamic(shapes, node.as_intersect_rhs(), Slot{'R', false}, false);
    case NodeKind::UnionOp:
      for (std::uint32_t i = 0; i < node.as_union_arity(); ++i) {
        if (MarksDynamic(shapes, node.as_union_child(i), Slot{'R', false}, false)) {
          return true;
        }
      }
      return false;
    case NodeKind::Call: {
      const std::string_view name = canonical_function_name(node.as_call_name());
      const bool builtin = !shapes.bound(node.as_call_name()) && IsBuiltin(name);
      for (std::uint32_t i = 0; i < node.as_call_arity(); ++i) {
        // A LAMBDA's arguments are bound whole.
        const Slot arg = builtin ? Slot{xlsb_parameter_class(name, i), false} : Slot{};
        if ((builtin && Shapes::lifts_argument(name, i, arg.letter, shapes.of(node.as_call_arg(i)))) ||
            MarksDynamic(shapes, node.as_call_arg(i), arg, false)) {
          return true;
        }
      }
      return false;
    }
    case NodeKind::LambdaCall:
      for (std::uint32_t i = 0; i < node.as_lambda_call_arity(); ++i) {
        if (MarksDynamic(shapes, node.as_lambda_call_arg(i), Slot{}, false)) {
          return true;
        }
      }
      return false;
    case NodeKind::LetBinding: {
      const std::uint32_t n = node.as_let_binding_count();
      bool marks = false;
      for (std::uint32_t i = 0; i < n; ++i) {
        marks = marks || MarksDynamic(shapes, node.as_let_binding_expr(i), Slot{'S', false}, false);
        shapes.bind(node.as_let_binding_name(i), shapes.of(node.as_let_binding_expr(i)));
      }
      marks = marks || MarksDynamic(shapes, node.as_let_body(), Slot{'S', false}, root);
      shapes.unbind(n);
      return marks;
    }
    case NodeKind::ImplicitIntersection:
      // The operand is intersected; what it is computed from is not.
      return MarksDynamic(shapes, node.as_implicit_intersection_operand(), Slot{'Y', false}, false);
    default:
      // Leaves, `#`, and a LAMBDA, whose body is evaluated where it is called.
      return false;
  }
}

/// Walker behind `legacy_intersections`: where Excel 365 shows an `@` in a
/// formula saved without the dynamic-array mark (measured cell by cell
/// against its formula2 text). A multi-valued node takes one wherever an
/// area would be value class.
class LegacyIntersections {
 public:
  LegacyIntersections(const NameShapes& names, std::vector<const parser::AstNode*>& out)
      : shapes_(names, /*lifting=*/false), out_(out) {}

  void walk(const parser::AstNode& node, Slot slot, bool root) {
    // Everything under a forced-array parameter is evaluated as an array,
    // nested calls included (`SUMPRODUCT(ABS(A1:A2))` shows no `@`).
    const bool value =
        !in_forced_array_ && (root || SlotClass(slot, /*area=*/true, PtgEvaluation::kLegacy) == kPtgValueClass);
    if (node.kind() == parser::NodeKind::ImplicitIntersection) {
      // A written `@` is listed when it is the one Excel would show; its
      // operand never takes a second one.
      const parser::AstNode& operand = node.as_implicit_intersection_operand();
      if (value && shapes_.of(operand) != Shape::kScalar) {
        out_.push_back(&node);
      }
      walk_children(operand, slot, root);
      return;
    }
    if (value && shapes_.of(node) != Shape::kScalar) {
      out_.push_back(&node);
    }
    walk_children(node, slot, root);
  }

 private:
  void walk_children(const parser::AstNode& node, Slot slot, bool root) {
    using parser::NodeKind;
    switch (node.kind()) {
      case NodeKind::UnaryOp:
        walk(node.as_unary_operand(), Slot{slot.letter, true}, false);
        return;
      case NodeKind::BinaryOp:
        walk(node.as_binary_lhs(), Slot{slot.letter, true}, false);
        walk(node.as_binary_rhs(), Slot{slot.letter, true}, false);
        return;
      case NodeKind::RangeOp:
        walk(node.as_range_lhs(), Slot{'R', false}, false);
        walk(node.as_range_rhs(), Slot{'R', false}, false);
        return;
      case NodeKind::IntersectOp:
        walk(node.as_intersect_lhs(), Slot{'R', false}, false);
        walk(node.as_intersect_rhs(), Slot{'R', false}, false);
        return;
      case NodeKind::UnionOp:
        for (std::uint32_t i = 0; i < node.as_union_arity(); ++i) {
          walk(node.as_union_child(i), Slot{'R', false}, false);
        }
        return;
      case NodeKind::Call: {
        const std::string_view name = canonical_function_name(node.as_call_name());
        const bool builtin = !shapes_.bound(node.as_call_name()) && IsBuiltin(name);
        for (std::uint32_t i = 0; i < node.as_call_arity(); ++i) {
          const Slot arg = builtin ? Slot{xlsb_parameter_class(name, i), false} : Slot{};
          const ArrayScope forced(in_forced_array_, arg.letter == 'F');
          walk(node.as_call_arg(i), arg, false);
        }
        return;
      }
      case NodeKind::LambdaCall:
        for (std::uint32_t i = 0; i < node.as_lambda_call_arity(); ++i) {
          walk(node.as_lambda_call_arg(i), Slot{}, false);
        }
        return;
      case NodeKind::LetBinding: {
        const std::uint32_t n = node.as_let_binding_count();
        for (std::uint32_t i = 0; i < n; ++i) {
          walk(node.as_let_binding_expr(i), Slot{}, false);
          shapes_.bind(node.as_let_binding_name(i), shapes_.of(node.as_let_binding_expr(i)));
        }
        walk(node.as_let_body(), slot, root);
        shapes_.unbind(n);
        return;
      }
      default:
        // Leaves; a LAMBDA body is evaluated where it is called, not here.
        return;
    }
  }

  Shapes shapes_;
  std::vector<const parser::AstNode*>& out_;
  bool in_forced_array_ = false;
};

}  // namespace

bool formula_is_dynamic_array(const parser::AstNode& root, const NameShapes& names) {
  Shapes shapes(names);
  return MarksDynamic(shapes, root, Slot{}, /*root=*/true);
}

namespace {

/// `node` is the value `name_sets_calc_exp` looks for, or reaches one through operators.
bool ValueSetsCalcExp(const parser::AstNode& node,
                      const std::unordered_map<const parser::AstNode*, std::uint8_t>& parens) {
  using parser::NodeKind;
  if (parens.count(&node) != 0U) {
    return false;
  }
  switch (node.kind()) {
    case NodeKind::Call: {
      // A call through a name (a LAMBDA, or one Excel does not know) counts as a LAMBDA call.
      const std::string_view name = canonical_function_name(node.as_call_name());
      return !IsBuiltin(name) || xlsb_sets_calc_exp(name);
    }
    case NodeKind::LetBinding:
    case NodeKind::Lambda:
    case NodeKind::LambdaCall:
    case NodeKind::SpillRef:
      return true;
    case NodeKind::UnaryOp:
      return ValueSetsCalcExp(node.as_unary_operand(), parens);
    case NodeKind::BinaryOp:
      return ValueSetsCalcExp(node.as_binary_lhs(), parens) || ValueSetsCalcExp(node.as_binary_rhs(), parens);
    case NodeKind::RangeOp:
      return ValueSetsCalcExp(node.as_range_lhs(), parens) || ValueSetsCalcExp(node.as_range_rhs(), parens);
    default:
      return false;
  }
}

/// `node` refers, anywhere within it, to a defined name `names` flags.
bool RefersToCalcExpName(const parser::AstNode& node, const NameShapes& names) {
  const bool name_ref =
      node.kind() == parser::NodeKind::NameRef ||
      (node.kind() == parser::NodeKind::Call && !IsBuiltin(canonical_function_name(node.as_call_name())));
  if (name_ref && names && names(node).calc_exp) {
    return true;
  }
  for (const parser::AstNode* child : parser::child_nodes(node)) {
    if (RefersToCalcExpName(*child, names)) {
      return true;
    }
  }
  return false;
}

/// Walker behind `formula_always_calculates`; `bound` holds the LET / LAMBDA
/// names in scope, which a call may go through without being unknown.
bool AlwaysCalculates(const parser::AstNode& node, const NameShapes& names, std::vector<std::string_view>& bound) {
  using parser::NodeKind;
  auto is_bound = [&](std::string_view name) {
    return std::any_of(bound.begin(), bound.end(),
                       [&](std::string_view b) { return strings::case_insensitive_eq(b, name); });
  };
  switch (node.kind()) {
    case NodeKind::Call: {
      const std::string_view name = node.as_call_name();
      if (parser::is_volatile_function_name(name)) {
        return true;
      }
      if (!IsBuiltin(canonical_function_name(name)) && !is_bound(name)) {
        const NameShape shape = names ? names(node) : NameShape{};
        if (!shape.defined || shape.always_calculates) {
          return true;
        }
      }
      break;
    }
    case NodeKind::NameRef:
    case NodeKind::ExternalRef:
      if ((node.kind() == NodeKind::ExternalRef && !parser::is_self_book_name_ref(node)) ||
          (node.kind() == NodeKind::NameRef && is_bound(node.as_name()))) {
        break;
      }
      if (names && names(node).always_calculates) {
        return true;
      }
      break;
    case NodeKind::LambdaCall: {
      // A call through a reference (`A1(1)`) calls nothing Excel knows.
      const NodeKind callee = node.as_lambda_call_callee().kind();
      if (callee != NodeKind::Lambda && callee != NodeKind::LambdaCall) {
        return true;
      }
      break;
    }
    case NodeKind::LetBinding: {
      const std::size_t depth = bound.size();
      bool found = false;
      for (std::uint32_t i = 0; i < node.as_let_binding_count() && !found; ++i) {
        found = AlwaysCalculates(node.as_let_binding_expr(i), names, bound);
        bound.push_back(node.as_let_binding_name(i));
      }
      found = found || AlwaysCalculates(node.as_let_body(), names, bound);
      bound.resize(depth);
      return found;
    }
    case NodeKind::Lambda: {
      const std::size_t depth = bound.size();
      for (std::uint32_t i = 0; i < node.as_lambda_param_count(); ++i) {
        bound.push_back(node.as_lambda_param(i));
      }
      const bool found = AlwaysCalculates(node.as_lambda_body(), names, bound);
      bound.resize(depth);
      return found;
    }
    default:
      break;
  }
  for (const parser::AstNode* child : parser::child_nodes(node)) {
    if (AlwaysCalculates(*child, names, bound)) {
      return true;
    }
  }
  return false;
}

/// True when `node` calls one of Excel's volatile functions anywhere.
bool ContainsVolatileCall(const parser::AstNode& node) {
  if (node.kind() == parser::NodeKind::Call && parser::is_volatile_function_name(node.as_call_name())) {
    return true;
  }
  for (const parser::AstNode* child : parser::child_nodes(node)) {
    if (ContainsVolatileCall(*child)) {
      return true;
    }
  }
  return false;
}

}  // namespace

bool name_sets_calc_exp(const parser::AstNode& root, const NameShapes& names) {
  std::unordered_map<const parser::AstNode*, std::uint8_t> parens;
  parser::collect_parenthesized_nodes(root, parens);
  return ValueSetsCalcExp(root, parens) || RefersToCalcExpName(root, names);
}

bool formula_calls_volatile(const parser::AstNode& root) {
  return ContainsVolatileCall(root);
}

bool formula_always_calculates(const parser::AstNode& root, const NameShapes& names) {
  std::vector<std::string_view> bound;
  return AlwaysCalculates(root, names, bound);
}

std::vector<const parser::AstNode*> legacy_intersections(const parser::AstNode& root, const NameShapes& names) {
  std::vector<const parser::AstNode*> out;
  LegacyIntersections(names, out).walk(root, Slot{}, /*root=*/true);
  return out;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
