//
// Implementation of the S-expression dumper.

#include "parser/ast_dump.h"

#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>

#include "parser/ast.h"
#include "parser/reference.h"
#include "utils/double_format.h"
#include "utils/expected.h"  // FM_CHECK
#include "value.h"

namespace formulon {
namespace parser {
namespace {

// Forward declaration: dumper recurses through every variable-arity arm.
void DumpInto(const AstNode& node, std::string& out);

void AppendValueLiteral(std::string& out, const Value& v) {
  switch (v.kind()) {
    case ValueKind::Blank:
      out.append("(blank)");
      return;
    case ValueKind::Number:
      out.append("(num ");
      format_double(out, v.as_number());
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
    case ValueKind::Text: {
      // Quote the payload and escape `"` as `\"` and `\` as `\\` so the
      // dump is unambiguous when the text contains either character. This
      // is C-string-style escaping for stable goldens, not Excel's `""`
      // doubling (which is reserved for source-formula display).
      out.append("(text \"");
      const std::string_view s = v.as_text();
      for (char c : s) {
        if (c == '\\' || c == '"') {
          out.push_back('\\');
        }
        out.push_back(c);
      }
      out.append("\")");
      return;
    }
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

    case NodeKind::SpillRef:
      // Spilled-range reference (`A1#`): dumped with a trailing `#` after
      // the anchor's A1 spelling so the corpus stays unambiguous against
      // ordinary refs.
      out.append("(spill-ref ");
      out.append(format_a1(node.as_spill_ref()));
      out.append("#)");
      return;

    case NodeKind::Ref3D: {
      out.append("(ref3d ");
      out.append(node.as_ref3d_sheet_begin());
      out.push_back(':');
      out.append(node.as_ref3d_sheet_end());
      out.push_back(' ');
      out.append(format_a1(node.as_ref3d_cell()));
      if (node.as_ref3d_is_range()) {
        const Reference& begin = node.as_ref3d_cell();
        const Reference& end = node.as_ref3d_cell_end();
        if ((begin.is_full_col && end.is_full_col) || (begin.is_full_row && end.is_full_row)) {
          const std::string end_text = format_a1(end);
          const std::size_t begin_axis_end = out.rfind(':');
          out.erase(begin_axis_end);
          out.push_back(':');
          out.append(end_text, 0, end_text.find(':'));
        } else {
          out.push_back(':');
          out.append(format_a1(end));
        }
      }
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
      const std::uint32_t opt = node.as_lambda_optional_count();
      const std::uint32_t first_optional = n - opt;
      for (std::uint32_t i = 0; i < n; ++i) {
        if (i > 0) {
          out.push_back(' ');
        }
        if (i >= first_optional) {
          out.push_back('[');
          out.append(node.as_lambda_param(i));
          out.push_back(']');
        } else {
          out.append(node.as_lambda_param(i));
        }
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
