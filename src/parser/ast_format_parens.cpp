//
// Parenthesisation pass: decides once how many parenthesis pairs each AST node
// prints with (written pairs plus those a slot's binding power demands). The
// canonical formatter and the XLSB `PtgParen` encoder both read the result.

#include <algorithm>
#include <cstdint>
#include <string_view>

#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/ast_format_internal.h"
#include "parser/parser_detail.h"
#include "parser/reference.h"

namespace formulon {
namespace parser {
namespace detail {
namespace {

// The identifier a `:` endpoint opens with, when the endpoint is written as a
// bare identifier rather than as a reference. A call contributes its name
// because the parser sees the name before the argument list.
std::string_view LeadingIdentifier(const AstNode& node) noexcept {
  switch (node.kind()) {
    case NodeKind::NameRef:
      // `Sheet1!Name` opens with its sheet; a quoted qualifier is no identifier.
      if (!node.as_name_sheet().empty()) {
        return node.as_name_sheet_quoted() || local_sheet_needs_quoting_a1(node.as_name_sheet()) ? std::string_view{}
                                                                                                 : node.as_name_sheet();
      }
      return node.as_name();
    case NodeKind::Call:
      return node.as_call_name();
    default:
      return {};
  }
}

// `A:C` is written as two bare column identifiers, so the parser folds
// `<Ident>:<Ident>` into a single whole-column range whenever both sides name
// a column. A defined name or a LAMBDA parameter spelled like a column takes
// part in that fold as well: `RangeOp(NameRef RO, NameRef r)` emitted as
// `RO:r` reads back as the whole-column range RO:R, and `RangeOp(NameRef LE,
// Call NA())` emitted as `LE:NA()` reads back as a range with a stray `()`.
// Parenthesising the left endpoint keeps the two operands apart.
bool ColonEndpointsWouldFold(const AstNode& lhs, const AstNode& rhs) noexcept {
  return lhs.kind() == NodeKind::NameRef && is_a1_column_shaped(lhs.as_name()) &&
         is_a1_column_shaped(LeadingIdentifier(rhs));
}

// Excel holds a column-shaped name apart from a name on its left by
// parenthesising the right endpoint instead: `A:(B)`, `_xlnm.A1:(A)`.
bool RightEndpointTakesParens(const AstNode& lhs, const AstNode& rhs) noexcept {
  return lhs.kind() == NodeKind::NameRef && rhs.kind() == NodeKind::NameRef && rhs.as_name_sheet().empty() &&
         is_a1_column_shaped(rhs.as_name());
}

}  // namespace

// True when operands `i` and `i + 1` of `range` print as one whole-axis
// token (`A:C`, `1:3`): a spliced pair of whole-column or whole-row Refs.
//
// A pair that is one whole-axis Ref (`whole_axis_pair_is_single`) never
// prints spliced: `A:A` reads back as the single Ref rather than as a pair,
// so such a pair prints in full (`A:A:A:A`), which reads back as itself.
bool PrintsSpliced(const AstNode& range, std::uint32_t i) noexcept {
  if (!range.as_range_spliced(i)) {
    return false;
  }
  const AstNode& lhs = range.as_range_operand(i);
  const AstNode& rhs = range.as_range_operand(i + 1U);
  if (lhs.kind() != NodeKind::Ref || rhs.kind() != NodeKind::Ref) {
    return false;
  }
  const Reference& lr = lhs.as_ref();
  const Reference& rr = rhs.as_ref();
  const bool same_axis = (lr.is_full_col && rr.is_full_col) || (lr.is_full_row && rr.is_full_row);
  return same_axis && !whole_axis_pair_is_single(lr, rr);
}

namespace {

// A callee that re-parses as the same callee without parentheses. An
// unqualified cell spelled like a function (`LOG10`) would re-parse as that
// function's call, so it keeps them.
bool CalleePrintsBare(const AstNode& callee) {
  switch (callee.kind()) {
    case NodeKind::Lambda:
    case NodeKind::NameRef:
    case NodeKind::LambdaCall:
    case NodeKind::UnionOp:  // prints its own parentheses
      return true;
    case NodeKind::Call:
      // `CHOOSE(1,SUM,ABS)(5)`; a trim reference prints as a range.
      return trim_ref_call_mode(callee) == TrimRefMode::None;
    case NodeKind::Ref:
      return !callee.as_ref().sheet.empty() || !is_cellref_shaped_function_name(format_a1(callee.as_ref()));
    default:
      return is_self_book_name_ref(callee);
  }
}

// The one place that decides how many parenthesis pairs each node prints
// with: the pairs written around it, or one where its slot's binding power
// demands it. The formatters and the XLSB `PtgParen`s both read the result.
void Parenthesize(const AstNode& node, bool needed, ParenCounts& out) {
  const std::uint16_t count = std::max<std::uint16_t>(node.paren_depth(), needed ? 1U : 0U);
  if (count != 0U) {
    std::uint16_t& slot = out[&node];
    slot = std::max(slot, count);
  }
}

}  // namespace

void CollectParens(const AstNode& node, int min_bp, ParenCounts& out) {
  auto wrap_if = [&](int bp) { Parenthesize(node, bp < min_bp, out); };
  switch (node.kind()) {
    case NodeKind::Literal:
      Parenthesize(node, node.as_literal().is_number() && node.as_literal().as_number() < 0.0, out);
      return;
    case NodeKind::SpillRef:
      wrap_if(kBpPostfixHash);
      if (const AstNode* anchor = node.as_spill_ref_anchor_expr(); anchor != nullptr) {
        CollectParens(*anchor, kBpPostfixHash, out);
      }
      return;
    case NodeKind::UnaryOp: {
      const int bp = node.as_unary_op() == UnaryOp::Percent ? kBpPostfixPercent : kBpUnaryPrefix;
      wrap_if(bp);
      // A prefix `@` under a prefix sign can only parse as its operand: `-@{1,2}`.
      const AstNode& operand = node.as_unary_operand();
      const bool prefix_at = node.as_unary_op() != UnaryOp::Percent && operand.kind() == NodeKind::ImplicitIntersection;
      CollectParens(operand, prefix_at ? kBpAtPrefix : bp, out);
      return;
    }
    case NodeKind::BinaryOp: {
      const int bp = binop_binding_power(node.as_binary_op(0));
      wrap_if(bp);
      // Left-associative: only the leading operand may be a bare same-power chain.
      for (std::uint32_t i = 0; i < node.as_binary_arity(); ++i) {
        CollectParens(node.as_binary_operand(i), i == 0U ? bp : bp + 1, out);
      }
      return;
    }
    case NodeKind::RangeOp: {
      wrap_if(kBpRange);
      const std::uint32_t n = node.as_range_arity();
      for (std::uint32_t i = 0; i < n; ++i) {
        const AstNode& operand = node.as_range_operand(i);
        if (i + 1U < n && !PrintsSpliced(node, i) &&
            !RightEndpointTakesParens(operand, node.as_range_operand(i + 1U))) {
          Parenthesize(operand, ColonEndpointsWouldFold(operand, node.as_range_operand(i + 1U)), out);
        }
        if (i > 0U && RightEndpointTakesParens(node.as_range_operand(i - 1U), operand)) {
          Parenthesize(operand, true, out);
        }
        CollectParens(operand, i == 0U ? kBpRange : kBpRange + 1, out);
      }
      return;
    }
    case NodeKind::UnionOp:
      // A union only parses inside parentheses.
      Parenthesize(node, true, out);
      for (std::uint32_t i = 0; i < node.as_union_arity(); ++i) {
        CollectParens(node.as_union_child(i), 0, out);
      }
      return;
    case NodeKind::IntersectOp:
      wrap_if(kBpIntersect);
      CollectParens(node.as_intersect_lhs(), kBpIntersect, out);
      CollectParens(node.as_intersect_rhs(), kBpIntersect + 1, out);
      return;
    case NodeKind::ImplicitIntersection:
      wrap_if(kBpAtPrefix);
      CollectParens(node.as_implicit_intersection_operand(), kBpAtPrefix, out);
      return;
    case NodeKind::Call:
      // A trim reference prints as a range operator.
      Parenthesize(node, trim_ref_call_mode(node) != TrimRefMode::None && kBpRange < min_bp, out);
      for (std::uint32_t i = 0; i < node.as_call_arity(); ++i) {
        CollectParens(node.as_call_arg(i), 0, out);
      }
      return;
    case NodeKind::Lambda:
      Parenthesize(node, false, out);
      CollectParens(node.as_lambda_body(), 0, out);
      return;
    case NodeKind::LetBinding:
      Parenthesize(node, false, out);
      for (std::uint32_t i = 0; i < node.as_let_binding_count(); ++i) {
        CollectParens(node.as_let_binding_expr(i), 0, out);
      }
      CollectParens(node.as_let_body(), 0, out);
      return;
    case NodeKind::LambdaCall: {
      Parenthesize(node, false, out);
      const AstNode& callee = node.as_lambda_call_callee();
      Parenthesize(callee, !CalleePrintsBare(callee), out);
      CollectParens(callee, 0, out);
      for (std::uint32_t i = 0; i < node.as_lambda_call_arity(); ++i) {
        CollectParens(node.as_lambda_call_arg(i), 0, out);
      }
      return;
    }
    // Leaves, and array constants, whose elements never take parentheses.
    case NodeKind::Ref:
    case NodeKind::Ref3D:
    case NodeKind::ExternalRef:
    case NodeKind::StructuredRef:
    case NodeKind::NameRef:
    case NodeKind::ArrayLiteral:
    case NodeKind::ErrorLiteral:
    case NodeKind::ErrorPlaceholder:
      Parenthesize(node, false, out);
      return;
  }
}

}  // namespace detail
}  // namespace parser
}  // namespace formulon
