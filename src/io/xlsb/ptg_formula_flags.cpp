//
// The static judgements behind a formula record's flags: the dynamic-array
// mark, `fCalcExp`, always-calculate, a volatile call, and where a legacy
// formula shows `@`.
// Declared in `io/xlsb/ptg_writer.h`.

#include <algorithm>
#include <cstdint>
#include <initializer_list>
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

using detail::FlagScope;
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
      return parser::any_child_node(node, [&](const parser::AstNode& operand) {
        return MarksDynamic(shapes, operand, Slot{slot.letter, true}, false);
      });
    case NodeKind::RangeOp:
      return parser::any_child_node(
          node, [&](const parser::AstNode& operand) { return MarksDynamic(shapes, operand, Slot{'R', false}, false); });
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

bool IsOneOf(std::string_view name, std::initializer_list<std::string_view> names) {
  return std::any_of(names.begin(), names.end(),
                     [&](std::string_view fn) { return strings::case_insensitive_eq(name, fn); });
}

/// Where a reference a node produces is used: `iterated` by a parameter that
/// takes an array of references in an array context, or `kept` as a
/// reference by a parameter that is neither value nor array class.
struct Use {
  bool iterated = false;
  bool kept = false;
  /// A value parameter of a formula saved without the mark that takes an array in Excel 365 (`TRANSPOSE`).
  bool narrowed = false;
  /// Evaluated as an array only because an array parameter (`ROWS`) holds it directly.
  bool forced_by_array_parameter = false;
};

/// Walker behind `legacy_intersections`: where Excel 365 shows an `@` in a
/// formula saved without the dynamic-array mark (cell by cell
/// as in its formula2 text). A multi-valued node takes one wherever an
/// area would be value class.
class LegacyIntersections {
 public:
  LegacyIntersections(const NameShapes& names, std::vector<const parser::AstNode*>& out)
      : shapes_(names, /*lifting=*/false), out_(out) {}

  /// `covered` when the value of `node` is already the value of an
  /// intersected node, the body of an intersected LET or LAMBDA call: it
  /// takes no second `@` (`@LET(a,A1:A3,a)`).
  void walk(const parser::AstNode& node, Slot slot, bool root, bool covered = false, Use use = {}) {
    // Everything under a forced-array or array parameter is evaluated as an
    // array, nested calls included (`SUMPRODUCT(ABS(A1:A2))`,
    // `XMATCH(TRUE,ISBLANK(B1:B3))` show no `@`).
    // A computed array takes the class of a value where an area would be array class (`TYPE(@INDEX({1;2;3},2))`).
    const bool intersects_spread = this->intersects_spread(node, use);
    const Shape shape = intersects_spread ? shapes_.of_unspread(node) : shapes_.of(node);
    bool value_slot = root || SlotClass(slot, shape == Shape::kReference, PtgEvaluation::kLegacy) == kPtgValueClass;
    if (!root && !slot.operand) {
      if (node.kind() == parser::NodeKind::ArrayLiteral && (slot.letter == 'I' || (slot.letter == 'V' && array_ok_))) {
        // An array constant is an array in an `I` parameter (`CHOOSE({1;2},1,2)`), and in a value parameter
        // where an array is accepted (`SUM(ABS({-1,-2}))`).
        value_slot = false;
      } else if (shape == Shape::kArray && node.kind() == parser::NodeKind::Call) {
        const std::string_view name = canonical_function_name(node.as_call_name());
        // ROW and COLUMN over several cells give one value to an `S` parameter (`SUM(@ROW(A1:A3))`), and an array
        // a built-in always returns gives one to a reference parameter outside an array context (`N(@SEQUENCE(3))`).
        value_slot = value_slot || (slot.letter == 'S' && IsOneOf(name, {"ROW", "COLUMN"})) ||
                     (slot.letter == 'R' && !array_ok_ && xlsb_returns_array(name));
      }
    }
    // An array of references can never be an array where Excel 365 alone takes one
    // (`TRANSPOSE(@OFFSET(A1:A3,{0;1},0))`).
    const bool narrowed = use.narrowed && spreads_references(node) && !intersects_spread;
    const bool value = !covered && (narrowed || (!in_forced_array_ && value_slot));
    if (node.kind() == parser::NodeKind::ImplicitIntersection) {
      // A written `@` is listed when it is the one Excel would show; its
      // operand never takes a second one.
      const parser::AstNode& operand = node.as_implicit_intersection_operand();
      if (value && shapes_.of(operand) != Shape::kScalar) {
        out_.push_back(&node);
      }
      walk_children(operand, slot, root, /*covered=*/true, use);
      return;
    }
    const bool intersected = value && shape != Shape::kScalar;
    if (intersected) {
      out_.push_back(&node);
    }
    walk_children(node, slot, root, covered || intersected, use);
  }

 private:
  void walk_children(const parser::AstNode& node, Slot slot, bool root, bool covered, Use use) {
    using parser::NodeKind;
    switch (node.kind()) {
      case NodeKind::UnaryOp:
        walk(node.as_unary_operand(), Slot{slot.letter, true}, false);
        return;
      case NodeKind::BinaryOp:
        for (std::uint32_t i = 0; i < node.as_binary_arity(); ++i) {
          walk(node.as_binary_operand(i), Slot{slot.letter, true}, false);
        }
        return;
      case NodeKind::RangeOp:
        for (std::uint32_t i = 0; i < node.as_range_arity(); ++i) {
          walk(node.as_range_operand(i), Slot{'R', false}, false, false, Use{false, true});
        }
        return;
      case NodeKind::IntersectOp:
        walk(node.as_intersect_lhs(), Slot{'R', false}, false, false, Use{false, true});
        walk(node.as_intersect_rhs(), Slot{'R', false}, false, false, Use{false, true});
        return;
      case NodeKind::UnionOp:
        for (std::uint32_t i = 0; i < node.as_union_arity(); ++i) {
          walk(node.as_union_child(i), Slot{'R', false}, false, false, Use{false, true});
        }
        return;
      case NodeKind::Call:
        walk_call(node, slot, use);
        return;
      case NodeKind::LambdaCall: {
        for (std::uint32_t i = 0; i < node.as_lambda_call_arity(); ++i) {
          walk(node.as_lambda_call_arg(i), Slot{}, false, false, Use{false, true});
        }
        // A LAMBDA called where it is written runs its body here, and a
        // parameter may be bound to a range (`@LAMBDA(x,@x+1)(2)`).
        const parser::AstNode& callee = node.as_lambda_call_callee();
        if (callee.kind() == NodeKind::Lambda) {
          const std::uint32_t n = callee.as_lambda_param_count();
          for (std::uint32_t i = 0; i < n; ++i) {
            shapes_.bind(callee.as_lambda_param(i), Shape::kReference);
          }
          walk(callee.as_lambda_body(), slot, root, covered);
          shapes_.unbind(n);
        }
        return;
      }
      case NodeKind::LetBinding: {
        const std::uint32_t n = node.as_let_binding_count();
        for (std::uint32_t i = 0; i < n; ++i) {
          walk(node.as_let_binding_expr(i), Slot{}, false);
          shapes_.bind(node.as_let_binding_name(i), shapes_.of(node.as_let_binding_expr(i)));
        }
        walk(node.as_let_body(), slot, root, covered);
        shapes_.unbind(n);
        return;
      }
      default:
        // Leaves; a LAMBDA body is evaluated where it is called, not here.
        return;
    }
  }

  void walk_call(const parser::AstNode& node, Slot slot, Use use) {
    const std::string_view name = canonical_function_name(node.as_call_name());
    const bool builtin = !shapes_.bound(node.as_call_name()) && IsBuiltin(name);
    // A call that is an array-class operand (`AGGREGATE(15,6,ROW(A1:A3)/ISNUMBER(MATCH(A1:A3,B1:B3,0)),1)`)
    // evaluates its arguments as arrays.
    const bool array_operand = slot.operand && SlotClass(slot, /*area=*/true, PtgEvaluation::kLegacy) != kPtgValueClass;
    const bool builds = builtin && IsOneOf(name, {"OFFSET", "INDIRECT"});
    const bool intersect_spread = intersects_spread(node, use);
    // INDEX and XLOOKUP give one reference: an array constant they would spread over is intersected where the result
    // is used as a reference (`SUBTOTAL(9,INDEX(A1:B3,0,@{1,2}))`).
    const bool picks = builtin && (use.kept || use.iterated) && IsOneOf(name, {"INDEX", "XLOOKUP"});
    const bool per_element = builtin && IsOneOf(name, {"MAP", "REDUCE", "SCAN", "BYROW", "BYCOL"});
    const bool call_array_ok = array_ok_;
    for (std::uint32_t i = 0; i < node.as_call_arity(); ++i) {
      const parser::AstNode& arg = node.as_call_arg(i);
      const Slot param = builtin ? Slot{shapes_.parameter_class(name, i), false} : Slot{};
      if (per_element && arg.kind() == parser::NodeKind::Lambda) {
        walk_per_element_body(arg);
        continue;
      }
      const bool forced_outside = in_forced_array_;
      const FlagScope forced(in_forced_array_,
                             in_forced_array_ || array_operand || param.letter == 'F' || param.letter == 'X');
      const FlagScope array_ok(array_ok_, array_ok_ || in_forced_array_ || param.letter == 'S');
      if ((builds || picks) && param.letter == 'V' && arg.kind() == parser::NodeKind::ArrayLiteral) {
        if (intersect_spread || picks) {
          out_.push_back(&arg);
        }
        continue;
      }
      if (builds && param.letter == 'V' && shapes_.of_array_operand(arg) == Shape::kArray) {
        // A computed array is intersected at its own operands (`OFFSET(A1,@ROW(A1:A3)-1,0)`) unless a forced-array
        // parameter around the call evaluates it as an array; an array parameter holding the call does not.
        const FlagScope lifted(in_forced_array_,
                               in_forced_array_ && !intersects_arguments(use) && !use.forced_by_array_parameter);
        walk(arg, param, false);
        continue;
      }
      Use arg_use{false, param.letter != 'V' && param.letter != 'F'};
      arg_use.iterated = builtin && call_array_ok && xlsb_iterates_references(name, i);
      arg_use.narrowed = builtin && param.letter == 'V' && xlsb_parameter_class(name, i) != 'V';
      arg_use.forced_by_array_parameter = param.letter == 'X' && !forced_outside && !array_operand;
      walk(arg, param, false, false, arg_use);
    }
  }

  /// OFFSET or INDIRECT spread over an array argument builds an array of references. Where that array is neither
  /// iterated nor turned into values, each array argument is intersected instead (`ROWS(OFFSET(A1,@{0;1},0))`).
  bool intersects_spread(const parser::AstNode& node, Use use) {
    return intersects_arguments(use) && spreads_references(node);
  }

  bool intersects_arguments(Use use) const { return !use.iterated && !use.narrowed && (use.kept || !array_ok_); }

  bool spreads_references(const parser::AstNode& node) {
    if (node.kind() != parser::NodeKind::Call || shapes_.bound(node.as_call_name()) ||
        !IsOneOf(canonical_function_name(node.as_call_name()), {"OFFSET", "INDIRECT"})) {
      return false;
    }
    // OFFSET's base is a reference however many cells it spans. An operator is a value here, as in `Shapes`, until
    // a forced array evaluates it.
    const std::uint32_t first =
        strings::case_insensitive_eq(canonical_function_name(node.as_call_name()), "OFFSET") ? 1U : 0U;
    for (std::uint32_t i = first; i < node.as_call_arity(); ++i) {
      if (shapes_.of(node.as_call_arg(i)) == Shape::kArray) {
        return true;
      }
    }
    return false;
  }

  /// The body of a LAMBDA a built-in calls once per element (`MAP(A1:A2,LAMBDA(x,@ROW(x)))`), whose parameters may be
  /// bound to ranges; a body that is a parameter returns that element as it is.
  void walk_per_element_body(const parser::AstNode& lambda) {
    const std::uint32_t n = lambda.as_lambda_param_count();
    for (std::uint32_t i = 0; i < n; ++i) {
      shapes_.bind(lambda.as_lambda_param(i), Shape::kReference);
    }
    const parser::AstNode& body = lambda.as_lambda_body();
    const bool parameter =
        body.kind() == parser::NodeKind::NameRef && body.as_name_sheet().empty() && shapes_.bound(body.as_name());
    {
      const FlagScope array_ok(array_ok_, false);
      const FlagScope forced(in_forced_array_, false);
      walk(body, Slot{}, /*root=*/true, parameter);
    }
    shapes_.unbind(n);
  }

  Shapes shapes_;
  std::vector<const parser::AstNode*>& out_;
  bool in_forced_array_ = false;
  /// An array is accepted here: under an argument evaluated as an array, or an `S` parameter's.
  bool array_ok_ = false;
};

}  // namespace

bool formula_is_dynamic_array(const parser::AstNode& root, const NameShapes& names) {
  Shapes shapes(names);
  return MarksDynamic(shapes, root, Slot{}, /*root=*/true);
}

namespace {

/// `node` is the value `name_sets_calc_exp` looks for, or reaches one through operators.
bool ValueSetsCalcExp(const parser::AstNode& node, const parser::ParenCounts& parens) {
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
      return parser::any_child_node(node,
                                    [&](const parser::AstNode& operand) { return ValueSetsCalcExp(operand, parens); });
    case NodeKind::RangeOp:
      return parser::any_child_node(node,
                                    [&](const parser::AstNode& operand) { return ValueSetsCalcExp(operand, parens); });
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
      // A call through a reference (`A1(1)`) calls nothing Excel knows; a call
      // of a computed function value (`CHOOSE(1,SUM,ABS)(5)`) does.
      const NodeKind callee = node.as_lambda_call_callee().kind();
      if (callee != NodeKind::Lambda && callee != NodeKind::LambdaCall && callee != NodeKind::Call) {
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

/// Walker behind `function_value_refs`; `bound` holds the LET / LAMBDA names in scope.
void CollectFunctionValues(const parser::AstNode& node, const NameShapes& names, std::vector<std::string_view>& bound,
                           std::vector<const parser::AstNode*>& out) {
  using parser::NodeKind;
  const std::size_t depth = bound.size();
  switch (node.kind()) {
    case NodeKind::NameRef: {
      const std::string_view name = node.as_name();
      const bool is_bound = std::any_of(bound.begin(), bound.end(),
                                        [&](std::string_view b) { return strings::case_insensitive_eq(b, name); });
      if (node.as_name_sheet().empty() && !is_bound && !(names && names(node).defined) &&
          IsBuiltin(canonical_function_name(name))) {
        out.push_back(&node);
      }
      return;
    }
    case NodeKind::LetBinding:
      for (std::uint32_t i = 0; i < node.as_let_binding_count(); ++i) {
        CollectFunctionValues(node.as_let_binding_expr(i), names, bound, out);
        bound.push_back(node.as_let_binding_name(i));
      }
      CollectFunctionValues(node.as_let_body(), names, bound, out);
      bound.resize(depth);
      return;
    case NodeKind::Lambda:
      for (std::uint32_t i = 0; i < node.as_lambda_param_count(); ++i) {
        bound.push_back(node.as_lambda_param(i));
      }
      CollectFunctionValues(node.as_lambda_body(), names, bound, out);
      bound.resize(depth);
      return;
    default:
      for (const parser::AstNode* child : parser::child_nodes(node)) {
        CollectFunctionValues(*child, names, bound, out);
      }
      return;
  }
}

}  // namespace

std::vector<const parser::AstNode*> function_value_refs(const parser::AstNode& root, const NameShapes& names) {
  std::vector<std::string_view> bound;
  std::vector<const parser::AstNode*> out;
  CollectFunctionValues(root, names, bound, out);
  return out;
}

bool name_sets_calc_exp(const parser::AstNode& root, const NameShapes& names) {
  parser::ParenCounts parens;
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
