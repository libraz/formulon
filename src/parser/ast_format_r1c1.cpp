//
// R1C1 formatter. Reference-bearing nodes and the operators around them are
// rendered here; leaves that carry no cell reference (literals, names,
// structured references, array constants) are delegated to `format_formula`.
// Parenthesis counts come from `collect_parenthesized_nodes`, so grouping
// matches the A1 formatter exactly.

#include "parser/ast_format_r1c1.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_map>

#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/reference.h"

namespace formulon {
namespace parser {
namespace {

using ParenCounts = std::unordered_map<const AstNode*, std::uint8_t>;

struct Host {
  std::int64_t row;
  std::int64_t col;
};

// One axis of a cell reference: `R3` (absolute), `R[-1]` (relative), `R` (same).
void AppendAxis(char axis, std::int64_t index, std::int64_t host, bool absolute, std::string& out) {
  out.push_back(axis);
  if (absolute) {
    out.append(std::to_string(index + 1));
    return;
  }
  const std::int64_t delta = index - host;
  if (delta != 0) {
    out.push_back('[');
    out.append(std::to_string(delta));
    out.push_back(']');
  }
}

// The cell part of `r` (no sheet): a whole column / row prints its axis only.
void AppendCell(const Reference& r, const Host& host, std::string& out) {
  if (!r.is_full_col) {
    AppendAxis('R', r.row, host.row, r.row_abs, out);
  }
  if (!r.is_full_row) {
    AppendAxis('C', r.col, host.col, r.col_abs, out);
  }
}

void AppendSheetQualifier(const Reference& r, std::string& out) {
  if (r.sheet.empty()) {
    return;
  }
  append_sheet_name(r.sheet, r.sheet_quoted, out);
  out.push_back('!');
}

void AppendRef(const Reference& r, const Host& host, std::string& out) {
  AppendSheetQualifier(r, out);
  AppendCell(r, host, out);
}

void AppendExternalRef(const AstNode& node, const Host& host, std::string& out) {
  const std::string_view sheet = node.as_external_ref_sheet();
  const std::string_view name = node.as_external_ref_name();
  append_external_qualifier(node.as_external_ref_book(), sheet, out);
  if (!name.empty()) {
    out.append(name);
    return;
  }
  AppendCell(node.as_external_ref_cell(), host, out);
  if (node.as_external_ref_is_range()) {
    out.push_back(':');
    AppendCell(node.as_external_ref_cell_end(), host, out);
  }
}

// Excel's R1C1 reading quotes each sheet of a 3-D span on its own
// (`Data:'My Sheet'!`), unlike the single quoted unit its A1 storage uses.
void AppendRef3D(const AstNode& node, const Host& host, std::string& out) {
  append_sheet_name(node.as_ref3d_sheet_begin(), false, out);
  out.push_back(':');
  append_sheet_name(node.as_ref3d_sheet_end(), false, out);
  out.push_back('!');
  AppendCell(node.as_ref3d_cell(), host, out);
  if (node.as_ref3d_is_range()) {
    out.push_back(':');
    AppendCell(node.as_ref3d_cell_end(), host, out);
  }
}

// True for a node whose `@` Excel does not echo back in `FormulaR1C1`.
bool IsReferenceShaped(const AstNode& node) noexcept {
  switch (node.kind()) {
    case NodeKind::Ref:
    case NodeKind::Ref3D:
    case NodeKind::ExternalRef:
    case NodeKind::RangeOp:
      return true;
    default:
      return false;
  }
}

struct Emitter {
  Host host;
  ParenCounts parens;

  void emit(const AstNode& node, std::string& out) const {
    switch (node.kind()) {
      case NodeKind::Literal:
      case NodeKind::StructuredRef:
      case NodeKind::NameRef:
      case NodeKind::ArrayLiteral:
      case NodeKind::ErrorLiteral:
      case NodeKind::ErrorPlaceholder:
        // No cell reference inside; the A1 formatter's text (own parentheses
        // included) is already correct.
        out.append(format_formula(node));
        return;
      default:
        break;
    }
    const auto it = parens.find(&node);
    const std::uint8_t n = it == parens.end() ? 0U : it->second;
    out.append(n, '(');
    emit_bare(node, out);
    out.append(n, ')');
  }

  void emit_args(std::uint32_t count, const AstNode& (AstNode::*arg)(std::uint32_t) const, const AstNode& node,
                 std::string& out) const {
    for (std::uint32_t i = 0; i < count; ++i) {
      if (i > 0) {
        out.push_back(',');
      }
      emit((node.*arg)(i), out);
    }
  }

  void emit_bare(const AstNode& node, std::string& out) const {
    switch (node.kind()) {
      case NodeKind::Ref:
        AppendRef(node.as_ref(), host, out);
        return;
      case NodeKind::SpillRef:
        if (const AstNode* anchor = node.as_spill_ref_anchor_expr(); anchor != nullptr) {
          emit(*anchor, out);
        } else {
          AppendRef(node.as_spill_ref(), host, out);
        }
        out.push_back('#');
        return;
      case NodeKind::Ref3D:
        AppendRef3D(node, host, out);
        return;
      case NodeKind::ExternalRef:
        AppendExternalRef(node, host, out);
        return;
      case NodeKind::UnaryOp:
        if (node.as_unary_op() == UnaryOp::Percent) {
          emit(node.as_unary_operand(), out);
          out.push_back('%');
        } else {
          out.push_back(node.as_unary_op() == UnaryOp::Plus ? '+' : '-');
          emit(node.as_unary_operand(), out);
        }
        return;
      case NodeKind::BinaryOp:
        emit(node.as_binary_lhs(), out);
        out.append(binop_token(node.as_binary_op()));
        emit(node.as_binary_rhs(), out);
        return;
      case NodeKind::RangeOp:
        emit(node.as_range_lhs(), out);
        out.push_back(':');
        emit(node.as_range_rhs(), out);
        return;
      case NodeKind::UnionOp:
        emit_args(node.as_union_arity(), &AstNode::as_union_child, node, out);
        return;
      case NodeKind::IntersectOp:
        emit(node.as_intersect_lhs(), out);
        out.push_back(' ');
        emit(node.as_intersect_rhs(), out);
        return;
      case NodeKind::ImplicitIntersection:
        if (!IsReferenceShaped(node.as_implicit_intersection_operand())) {
          out.push_back('@');
        }
        emit(node.as_implicit_intersection_operand(), out);
        return;
      case NodeKind::Call:
        out.append(node.as_call_name());
        out.push_back('(');
        emit_args(node.as_call_arity(), &AstNode::as_call_arg, node, out);
        out.push_back(')');
        return;
      case NodeKind::Lambda: {
        out.append("LAMBDA(");
        const std::uint32_t n = node.as_lambda_param_count();
        const std::uint32_t first_optional = n - node.as_lambda_optional_count();
        for (std::uint32_t i = 0; i < n; ++i) {
          if (i > 0) {
            out.push_back(',');
          }
          const bool optional = i >= first_optional;
          if (optional) {
            out.push_back('[');
          }
          out.append(node.as_lambda_param(i));
          if (optional) {
            out.push_back(']');
          }
        }
        if (n > 0) {
          out.push_back(',');
        }
        emit(node.as_lambda_body(), out);
        out.push_back(')');
        return;
      }
      case NodeKind::LetBinding:
        out.append("LET(");
        for (std::uint32_t i = 0; i < node.as_let_binding_count(); ++i) {
          out.append(node.as_let_binding_name(i));
          out.push_back(',');
          emit(node.as_let_binding_expr(i), out);
          out.push_back(',');
        }
        emit(node.as_let_body(), out);
        out.push_back(')');
        return;
      case NodeKind::LambdaCall:
        emit(node.as_lambda_call_callee(), out);
        out.push_back('(');
        emit_args(node.as_lambda_call_arity(), &AstNode::as_lambda_call_arg, node, out);
        out.push_back(')');
        return;
      case NodeKind::Literal:
      case NodeKind::StructuredRef:
      case NodeKind::NameRef:
      case NodeKind::ArrayLiteral:
      case NodeKind::ErrorLiteral:
      case NodeKind::ErrorPlaceholder:
        // Delegated in `emit`.
        return;
    }
  }
};

}  // namespace

std::string format_formula_r1c1(const AstNode& node, std::uint32_t anchor_row, std::uint32_t anchor_col) {
  if (!ast_depth_within_limit(node, kMaxFormulaAstDepth)) {
    return "#REF!";
  }
  Emitter emitter{{static_cast<std::int64_t>(anchor_row), static_cast<std::int64_t>(anchor_col)}, {}};
  collect_parenthesized_nodes(node, emitter.parens);
  std::string out;
  out.reserve(64);
  emitter.emit(node, out);
  return out;
}

}  // namespace parser
}  // namespace formulon
