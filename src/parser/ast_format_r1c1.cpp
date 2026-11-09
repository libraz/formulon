//
// R1C1 formatter. Reference-bearing nodes, names and the operators around
// them are rendered here; leaves that carry no cell reference (literals,
// stored-form structured references, array constants) are delegated to
// `format_formula`. A cell-shaped name prints bare: in R1C1 it is no cell.
// Parenthesis counts come from `collect_parenthesized_nodes`, so grouping
// matches the A1 formatter exactly.

#include "parser/ast_format_r1c1.h"

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>

#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/ast_format_internal.h"
#include "parser/formula_prefix.h"
#include "parser/reference.h"
#include "parser/structured_ref_payload.h"
#include "sheet_name.h"

namespace formulon {
namespace parser {
namespace {

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

// A local sheet name under the R1C1 quoting rule. The source's quoting
// hint is not consulted: `'S2'!A1` reads back as `S2!RC`.
void AppendSheetName(std::string_view sheet, std::string& out) {
  append_sheet_name_quoted_if(sheet, local_sheet_needs_quoting_r1c1(sheet), out);
}

void AppendSheetQualifier(std::string_view sheet, std::string& out) {
  if (sheet.empty()) {
    return;
  }
  AppendSheetName(sheet, out);
  out.push_back('!');
}

void AppendRef(const Reference& r, const Host& host, std::string& out) {
  AppendSheetQualifier(r.sheet, out);
  AppendCell(r, host, out);
}

void AppendExternalRef(const AstNode& node, const Host& host, std::string& out) {
  const std::string_view name = node.as_external_ref_name();
  append_external_qualifier(node, /*r1c1=*/true, out);
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
  AppendSheetName(node.as_ref3d_sheet_begin(), out);
  out.push_back(':');
  AppendSheetName(node.as_ref3d_sheet_end(), out);
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
  const StructuredRefDisplay* display;

  void emit(const AstNode& node, std::string& out) const {
    switch (node.kind()) {
      case NodeKind::StructuredRef:
        if (display != nullptr) {
          break;
        }
        out.append(format_formula(node));
        return;
      case NodeKind::Literal:
      case NodeKind::ArrayLiteral:
      case NodeKind::ErrorPlaceholder:
        // No cell reference inside; the A1 formatter's text (own parentheses
        // included) is already correct.
        out.append(format_formula(node));
        return;
      default:
        break;
    }
    const std::uint16_t n = detail::ParenCount(parens, node);
    out.append(n, '(');
    emit_bare(node, out);
    out.append(n, ')');
  }

  // `Sheet3!#REF!` under the R1C1 quoting rule; a dead sheet's address
  // relative to the formula's cell (`#REF!R[-1]C[-4]`, `#REF!C[-4]`).
  void emit_error_literal(const AstNode& node, std::string& out) const {
    const RefErrorSite* site = node.as_ref_error_site();
    if (site != nullptr && site->form == RefErrorSite::Form::kQualified) {
      AppendSheetQualifier(site->sheet, out);
    }
    out.append(display_name(node.as_error_literal()));
    if (site == nullptr || site->form != RefErrorSite::Form::kDeadSheet) {
      return;
    }
    if (!site->has_address) {
      out.append(display_name(ErrorCode::Ref));
      return;
    }
    AppendCell(site->first, host, out);
    if (site->is_range) {
      out.push_back(':');
      AppendCell(site->last, host, out);
    }
  }

  // Only reached with `display` set; a payload that does not parse prints as stored.
  void emit_structured_ref(const AstNode& node, std::string& out) const {
    const std::string_view table = node.as_structured_ref_table();
    const std::optional<StructuredRefPayload> payload =
        parse_structured_ref_payload_text(node.as_structured_ref_column());
    if (!payload) {
      out.append(table).append("[").append(node.as_structured_ref_column()).push_back(']');
      return;
    }
    const bool inside = !display->own_table.empty() && collection_names::equal(table, display->own_table);
    out.append(display_structured_ref(table, *payload, inside));
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
        emit(node.as_binary_operand(0), out);
        for (std::uint32_t i = 1; i < node.as_binary_arity(); ++i) {
          out.append(binop_token(node.as_binary_op(i - 1U)));
          emit(node.as_binary_operand(i), out);
        }
        return;
      case NodeKind::RangeOp:
        // A whole column or row is one token in R1C1 (`C8`), so a whole-axis
        // token and two whole-axis operands both print as `C8:C5`.
        for (std::uint32_t i = 0; i < node.as_range_arity(); ++i) {
          if (i > 0U) {
            out.push_back(':');
          }
          emit(node.as_range_operand(i), out);
        }
        return;
      case NodeKind::UnionOp:
        detail::EmitCommaList(node.as_union_arity(), out, [&](std::uint32_t i) { emit(node.as_union_child(i), out); });
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
        if (const TrimRefMode mode = trim_ref_call_mode(node); mode != TrimRefMode::None &&
                                                               node.as_call_arg(0).kind() == NodeKind::RangeOp &&
                                                               parens.count(&node.as_call_arg(0)) == 0U) {
          // The chain's last `:` takes the operator.
          const AstNode& range = node.as_call_arg(0);
          for (std::uint32_t i = 0; i < range.as_range_arity(); ++i) {
            if (i > 0U) {
              out.append(i + 1U == range.as_range_arity() ? trim_ref_operator(mode) : ":");
            }
            emit(range.as_range_operand(i), out);
          }
          return;
        }
        out.append(cell_shaped_name_bare(node.as_call_name()));
        out.push_back('(');
        detail::EmitCommaList(node.as_call_arity(), out, [&](std::uint32_t i) { emit(node.as_call_arg(i), out); });
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
          out.append(cell_shaped_name_bare(node.as_lambda_param(i)));
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
          out.append(cell_shaped_name_bare(node.as_let_binding_name(i)));
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
        detail::EmitCommaList(node.as_lambda_call_arity(), out,
                              [&](std::uint32_t i) { emit(node.as_lambda_call_arg(i), out); });
        out.push_back(')');
        return;
      case NodeKind::NameRef:
        // A qualifier follows the R1C1 quoting rule.
        AppendSheetQualifier(node.as_name_sheet(), out);
        out.append(cell_shaped_name_bare(node.as_name()));
        return;
      case NodeKind::ErrorLiteral:
        emit_error_literal(node, out);
        return;
      case NodeKind::StructuredRef:
        emit_structured_ref(node, out);
        return;
      case NodeKind::Literal:
      case NodeKind::ArrayLiteral:
      case NodeKind::ErrorPlaceholder:
        // Delegated in `emit`.
        return;
    }
  }
};

}  // namespace

std::string format_formula_r1c1(const AstNode& node, std::uint32_t anchor_row, std::uint32_t anchor_col,
                                const StructuredRefDisplay* display) {
  if (!ast_depth_within_limit(node, kMaxFormulaAstDepth)) {
    return std::string(display_name(ErrorCode::Ref));
  }
  Emitter emitter{{static_cast<std::int64_t>(anchor_row), static_cast<std::int64_t>(anchor_col)}, {}, display};
  collect_parenthesized_nodes(node, emitter.parens);
  std::string out;
  out.reserve(64);
  emitter.emit(node, out);
  return out;
}

}  // namespace parser
}  // namespace formulon
