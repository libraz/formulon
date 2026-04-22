// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the S-expression dumper.

#include "parser/ast_dump.h"

#include <cmath>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>

#include "parser/ast.h"
#include "parser/reference.h"
#include "utils/expected.h"  // FM_CHECK
#include "value.h"

namespace formulon {
namespace parser {
namespace {

// Forward declaration: dumper recurses through every variable-arity arm.
void DumpInto(const AstNode& node, std::string& out);

// Formats a double with locale-independent rules:
//   * Exact integer in [-1e16, 1e16] -> integer form, no decimal point.
//   * NaN / +inf / -inf -> "nan" / "inf" / "-inf".
//   * Negative zero -> "0".
//   * Otherwise -> std::to_string with trailing zeros and a stranded
//     trailing dot trimmed.
//
// std::to_string is locale-dependent in principle but produces the C-locale
// "1.234" form on every libc Formulon supports; we accept the dependency
// for now and will swap in double-conversion when we need exact roundtripping.
void AppendNumber(std::string& out, double v) {
  if (std::isnan(v)) {
    out.append("nan");
    return;
  }
  if (std::isinf(v)) {
    out.append(v < 0.0 ? "-inf" : "inf");
    return;
  }
  // Negative zero collapses to plain "0" for stable goldens.
  if (v == 0.0) {
    out.push_back('0');
    return;
  }
  // Integer fast path.
  if (std::abs(v) < 1e16) {
    const double truncated = std::trunc(v);
    if (truncated == v) {
      const std::int64_t as_int = static_cast<std::int64_t>(truncated);
      out.append(std::to_string(as_int));
      return;
    }
  }
  std::string s = std::to_string(v);
  // Strip trailing zeros after a decimal point, then a trailing dot.
  const auto dot = s.find('.');
  if (dot != std::string::npos) {
    std::size_t last = s.size();
    while (last > dot + 1 && s[last - 1] == '0') {
      --last;
    }
    if (last > 0 && s[last - 1] == '.') {
      --last;
    }
    s.resize(last);
  }
  out.append(s);
}

void AppendValueLiteral(std::string& out, const Value& v) {
  switch (v.kind()) {
    case ValueKind::Blank:
      out.append("(blank)");
      return;
    case ValueKind::Number:
      out.append("(num ");
      AppendNumber(out, v.as_number());
      out.push_back(')');
      return;
    case ValueKind::Bool:
      out.append(v.as_boolean() ? "(bool true)" : "(bool false)");
      return;
    case ValueKind::Error:
      out.append("(err ");
      out.append(display_name(v.as_error()));
      out.push_back(')');
      return;
    case ValueKind::Text:
    case ValueKind::Array:
    case ValueKind::Ref:
    case ValueKind::Lambda:
      // The parser currently only emits scalar Value literals. The remaining
      // kinds are reserved for follow-up work; treat them as opaque so the
      // dumper still terminates if a future caller stuffs one into a Literal.
      out.append("(value ?)");
      return;
  }
  out.append("(value ?)");
}

const char* BinOpToken(BinOp op) {
  switch (op) {
    case BinOp::Add:
      return "+";
    case BinOp::Sub:
      return "-";
    case BinOp::Mul:
      return "*";
    case BinOp::Div:
      return "/";
    case BinOp::Pow:
      return "^";
    case BinOp::Concat:
      return "&";
    case BinOp::Eq:
      return "=";
    case BinOp::NotEq:
      return "<>";
    case BinOp::Lt:
      return "<";
    case BinOp::LtEq:
      return "<=";
    case BinOp::Gt:
      return ">";
    case BinOp::GtEq:
      return ">=";
  }
  return "?";
}

const char* UnaryOpToken(UnaryOp op) {
  switch (op) {
    case UnaryOp::Plus:
      return "+";
    case UnaryOp::Minus:
      return "-";
    case UnaryOp::Percent:
      return "%";
  }
  return "?";
}

const char* StructuredModifierToken(StructuredRefModifier m) {
  switch (m) {
    case StructuredRefModifier::None:
      return "";
    case StructuredRefModifier::At:
      return "@";
    case StructuredRefModifier::Headers:
      return "#headers";
    case StructuredRefModifier::Data:
      return "#data";
    case StructuredRefModifier::Totals:
      return "#totals";
    case StructuredRefModifier::All:
      return "#all";
  }
  return "";
}

void DumpInto(const AstNode& node, std::string& out) {
  switch (node.kind()) {
    case NodeKind::Literal:
      AppendValueLiteral(out, node.as_literal());
      return;

    case NodeKind::Ref:
      out.append("(ref ");
      out.append(format_a1(node.as_ref()));
      out.push_back(')');
      return;

    case NodeKind::ExternalRef: {
      out.append("(ext-ref [");
      out.append(std::to_string(node.as_external_ref_book_id()));
      out.append("] ");
      out.append(node.as_external_ref_sheet());
      out.push_back(' ');
      // Inner cell ref is formatted without its sheet: external ref already
      // exposes sheet as a separate field, so we strip it for the cell part.
      Reference cell_no_sheet = node.as_external_ref_cell();
      cell_no_sheet.sheet = {};
      cell_no_sheet.sheet_quoted = false;
      out.append(format_a1(cell_no_sheet));
      out.push_back(')');
      return;
    }

    case NodeKind::StructuredRef: {
      out.append("(struct-ref ");
      out.append(node.as_structured_ref_table());
      const std::string_view col = node.as_structured_ref_column();
      const StructuredRefModifier mod = node.as_structured_ref_modifier();
      if (!col.empty()) {
        out.push_back(' ');
        out.append(col);
      }
      if (mod != StructuredRefModifier::None) {
        out.push_back(' ');
        out.append(StructuredModifierToken(mod));
      }
      out.push_back(')');
      return;
    }

    case NodeKind::NameRef:
      out.append("(name ");
      out.append(node.as_name());
      out.push_back(')');
      return;

    case NodeKind::UnaryOp:
      out.append("(unary ");
      out.append(UnaryOpToken(node.as_unary_op()));
      out.push_back(' ');
      DumpInto(node.as_unary_operand(), out);
      out.push_back(')');
      return;

    case NodeKind::BinaryOp:
      out.append("(binary ");
      out.append(BinOpToken(node.as_binary_op()));
      out.push_back(' ');
      DumpInto(node.as_binary_lhs(), out);
      out.push_back(' ');
      DumpInto(node.as_binary_rhs(), out);
      out.push_back(')');
      return;

    case NodeKind::RangeOp:
      out.append("(range ");
      DumpInto(node.as_range_lhs(), out);
      out.push_back(' ');
      DumpInto(node.as_range_rhs(), out);
      out.push_back(')');
      return;

    case NodeKind::UnionOp: {
      out.append("(union");
      const std::uint32_t n = node.as_union_arity();
      for (std::uint32_t i = 0; i < n; ++i) {
        out.push_back(' ');
        DumpInto(node.as_union_child(i), out);
      }
      out.push_back(')');
      return;
    }

    case NodeKind::IntersectOp:
      out.append("(intersect ");
      DumpInto(node.as_intersect_lhs(), out);
      out.push_back(' ');
      DumpInto(node.as_intersect_rhs(), out);
      out.push_back(')');
      return;

    case NodeKind::ImplicitIntersection:
      out.append("(at ");
      DumpInto(node.as_implicit_intersection_operand(), out);
      out.push_back(')');
      return;

    case NodeKind::Call: {
      out.append("(call ");
      out.append(node.as_call_name());
      const std::uint32_t n = node.as_call_arity();
      for (std::uint32_t i = 0; i < n; ++i) {
        out.push_back(' ');
        DumpInto(node.as_call_arg(i), out);
      }
      out.push_back(')');
      return;
    }

    case NodeKind::ArrayLiteral: {
      out.append("(array ");
      out.append(std::to_string(node.as_array_rows()));
      out.push_back(' ');
      out.append(std::to_string(node.as_array_cols()));
      const std::uint32_t rows = node.as_array_rows();
      const std::uint32_t cols = node.as_array_cols();
      for (std::uint32_t r = 0; r < rows; ++r) {
        for (std::uint32_t c = 0; c < cols; ++c) {
          out.push_back(' ');
          DumpInto(node.as_array_element(r, c), out);
        }
      }
      out.push_back(')');
      return;
    }

    case NodeKind::Lambda: {
      out.append("(lambda (");
      const std::uint32_t n = node.as_lambda_param_count();
      for (std::uint32_t i = 0; i < n; ++i) {
        if (i > 0) {
          out.push_back(' ');
        }
        out.append(node.as_lambda_param(i));
      }
      out.append(") ");
      DumpInto(node.as_lambda_body(), out);
      out.push_back(')');
      return;
    }

    case NodeKind::LetBinding: {
      out.append("(let (");
      const std::uint32_t n = node.as_let_binding_count();
      for (std::uint32_t i = 0; i < n; ++i) {
        if (i > 0) {
          out.push_back(' ');
        }
        out.push_back('(');
        out.append(node.as_let_binding_name(i));
        out.push_back(' ');
        DumpInto(node.as_let_binding_expr(i), out);
        out.push_back(')');
      }
      out.append(") ");
      DumpInto(node.as_let_body(), out);
      out.push_back(')');
      return;
    }

    case NodeKind::LambdaCall: {
      out.append("(lambda-call ");
      DumpInto(node.as_lambda_call_callee(), out);
      const std::uint32_t n = node.as_lambda_call_arity();
      for (std::uint32_t i = 0; i < n; ++i) {
        out.push_back(' ');
        DumpInto(node.as_lambda_call_arg(i), out);
      }
      out.push_back(')');
      return;
    }

    case NodeKind::ErrorLiteral:
      out.append("(err-lit ");
      out.append(display_name(node.as_error_literal()));
      out.push_back(')');
      return;

    case NodeKind::ErrorPlaceholder:
      out.append("(error)");
      return;
  }
  // Defensive: unreachable while every NodeKind is covered above.
  out.append("(unknown)");
}

}  // namespace

std::string dump_sexpr(const AstNode& node) {
  std::string out;
  out.reserve(64);
  DumpInto(node, out);
  return out;
}

}  // namespace parser
}  // namespace formulon
