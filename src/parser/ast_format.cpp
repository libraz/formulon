// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Excel-compatible AST → formula text formatter. The formatter mirrors the
// Pratt parser's recursive structure: each kind emits its surface form and
// recurses into children, attaching parentheses whenever the child's
// effective binding power is below the parent slot's minimum. The output
// re-parses to a structurally equivalent AST (verified by the round-trip
// tests in `tests/unit/parser/ast_format_test.cpp`).

#include "parser/ast_format.h"

#include <cstdint>
#include <string>
#include <string_view>

#include "parser/ast.h"
#include "parser/parser_detail.h"
#include "parser/reference.h"
#include "utils/double_format.h"
#include "value.h"

namespace formulon {
namespace parser {
namespace {

using detail::kBpAddSub;
using detail::kBpAtPrefix;
using detail::kBpComparison;
using detail::kBpConcat;
using detail::kBpIntersect;
using detail::kBpMulDiv;
using detail::kBpPostfixHash;
using detail::kBpPostfixPercent;
using detail::kBpPow;
using detail::kBpRange;
using detail::kBpUnaryPrefix;

const char* BinOpToken(BinOp op) noexcept {
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
  return "+";
}

// Returns the binding-power slot for the given binary operator. Right-
// associativity is encoded in the per-side `min_bp` calls in `FormatBinary`,
// not here.
int BinOpBp(BinOp op) noexcept {
  switch (op) {
    case BinOp::Pow:
      return kBpPow;
    case BinOp::Mul:
    case BinOp::Div:
      return kBpMulDiv;
    case BinOp::Add:
    case BinOp::Sub:
      return kBpAddSub;
    case BinOp::Concat:
      return kBpConcat;
    case BinOp::Eq:
    case BinOp::NotEq:
    case BinOp::Lt:
    case BinOp::LtEq:
    case BinOp::Gt:
    case BinOp::GtEq:
      return kBpComparison;
  }
  return kBpComparison;
}

// Forward declaration: the full recursion target. `min_bp` is the minimum
// binding power the parent requires; if `EffectiveBp(node) < min_bp` we wrap
// the rendered text in `(...)`.
void FormatNode(const AstNode& node, std::string& out, int min_bp);

void FormatNumberLiteral(const Value& v, std::string& out) {
  // Negative numerics are written as a unary minus by the parser, so they
  // never appear inside a Literal payload coming from real source. We still
  // handle the sign here defensively for hand-built ASTs (the test suite
  // does this for completeness).
  const double n = v.as_number();
  if (n < 0.0) {
    out.push_back('(');
    format_double(out, n);
    out.push_back(')');
    return;
  }
  format_double(out, n);
}

void FormatTextLiteral(std::string_view s, std::string& out) {
  // Excel string literals double embedded `"`; that is the only escape.
  out.push_back('"');
  for (char c : s) {
    if (c == '"') {
      out.push_back('"');
    }
    out.push_back(c);
  }
  out.push_back('"');
}

void FormatLiteral(const Value& v, std::string& out) {
  switch (v.kind()) {
    case ValueKind::Blank:
      // Bare blank literals are not directly representable as Excel source;
      // treat as the empty string so any round-trip stays parseable.
      out.append("\"\"");
      return;
    case ValueKind::Number:
      FormatNumberLiteral(v, out);
      return;
    case ValueKind::Bool:
      out.append(v.as_boolean() ? "TRUE" : "FALSE");
      return;
    case ValueKind::Text:
      FormatTextLiteral(v.as_text(), out);
      return;
    case ValueKind::Error:
      out.append(display_name(v.as_error()));
      return;
    case ValueKind::Array:
    case ValueKind::Ref:
    case ValueKind::Lambda:
      // The parser does not emit non-scalar Value literals; defensively
      // render an Excel-side `#VALUE!` so an exotic input still produces
      // parseable text rather than truncating the formula.
      out.append(display_name(ErrorCode::Value));
      return;
  }
}

void FormatRef(const Reference& r, std::string& out) {
  // format_a1 already handles sheet-quoted vs bare and full-col / full-row.
  out.append(format_a1(r));
}

void FormatExternalRef(const AstNode& node, std::string& out) {
  out.push_back('[');
  out.append(std::to_string(node.as_external_ref_book_id()));
  out.push_back(']');
  // The external ref's sheet field carries the sheet name verbatim; use the
  // same quoting heuristic that `format_a1` applies to the inline sheet
  // field. We bypass format_a1's sheet handling here because the cell sub-
  // ref's own `sheet` is normally empty (the parser strips it).
  const std::string_view sheet = node.as_external_ref_sheet();
  // Quote when the sheet contains anything outside [A-Za-z0-9_.] or is empty.
  bool needs_quote = sheet.empty();
  for (char c : sheet) {
    const bool ok = (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '_' || c == '.';
    if (!ok) {
      needs_quote = true;
      break;
    }
  }
  if (needs_quote) {
    out.push_back('\'');
    for (char c : sheet) {
      if (c == '\'') {
        out.push_back('\'');
      }
      out.push_back(c);
    }
    out.push_back('\'');
  } else {
    out.append(sheet);
  }
  out.push_back('!');
  Reference cell_no_sheet = node.as_external_ref_cell();
  cell_no_sheet.sheet = {};
  cell_no_sheet.sheet_quoted = false;
  out.append(format_a1(cell_no_sheet));
}

void FormatStructuredRef(const AstNode& node, std::string& out) {
  out.append(node.as_structured_ref_table());
  // Two AST shapes need handling:
  //
  // 1. modifier == None: the parser packs the entire bracket payload into
  //    the `column` slot verbatim. We re-emit it inside a single bracket
  //    pair so the output round-trips through the parser ("Tbl[Region]",
  //    "Tbl[[#Headers],[Region]]", "Tbl[@Region]", ...).
  // 2. modifier != None: the factory was used directly with structured
  //    fields. Excel's surface form for "specifier + column" wraps each
  //    specifier in its own bracket pair and groups them inside an outer
  //    `[...]`: e.g. `Tbl[[#Data],[Region]]`. The bare-specifier shape is
  //    `Tbl[#Data]`. The `@` modifier is the lone exception: it sits
  //    directly before the column without an outer wrap (`Tbl[@Region]`).
  const std::string_view col = node.as_structured_ref_column();
  const StructuredRefModifier mod = node.as_structured_ref_modifier();
  if (mod == StructuredRefModifier::None) {
    out.push_back('[');
    out.append(col);
    out.push_back(']');
    return;
  }
  if (mod == StructuredRefModifier::At) {
    out.push_back('[');
    out.push_back('@');
    if (!col.empty()) {
      out.append(col);
    }
    out.push_back(']');
    return;
  }
  const char* spec = nullptr;
  switch (mod) {
    case StructuredRefModifier::Headers:
      spec = "#Headers";
      break;
    case StructuredRefModifier::Data:
      spec = "#Data";
      break;
    case StructuredRefModifier::Totals:
      spec = "#Totals";
      break;
    case StructuredRefModifier::All:
      spec = "#All";
      break;
    case StructuredRefModifier::None:
    case StructuredRefModifier::At:
      // Handled above.
      return;
  }
  if (col.empty()) {
    out.push_back('[');
    out.append(spec);
    out.push_back(']');
    return;
  }
  out.append("[[");
  out.append(spec);
  out.append("],[");
  out.append(col);
  out.append("]]");
}

void FormatUnary(const AstNode& node, std::string& out, int min_bp) {
  const UnaryOp op = node.as_unary_op();
  if (op == UnaryOp::Percent) {
    // Postfix: render operand at percent-level binding power then append `%`.
    const bool wrap = kBpPostfixPercent < min_bp;
    if (wrap) {
      out.push_back('(');
    }
    FormatNode(node.as_unary_operand(), out, kBpPostfixPercent);
    out.push_back('%');
    if (wrap) {
      out.push_back(')');
    }
    return;
  }
  // Prefix `+` / `-`.
  const bool wrap = kBpUnaryPrefix < min_bp;
  if (wrap) {
    out.push_back('(');
  }
  out.push_back(op == UnaryOp::Plus ? '+' : '-');
  FormatNode(node.as_unary_operand(), out, kBpUnaryPrefix);
  if (wrap) {
    out.push_back(')');
  }
}

void FormatBinary(const AstNode& node, std::string& out, int min_bp) {
  const BinOp op = node.as_binary_op();
  const int bp = BinOpBp(op);
  const bool wrap = bp < min_bp;
  if (wrap) {
    out.push_back('(');
  }
  // Power is right-associative: LHS demands `bp + 1` (so a same-precedence
  // child on the left would be wrapped), RHS demands `bp` (so the same
  // child does not need wrapping). Every other binary is left-associative.
  const int lhs_min = (op == BinOp::Pow) ? bp + 1 : bp;
  const int rhs_min = (op == BinOp::Pow) ? bp : bp + 1;
  FormatNode(node.as_binary_lhs(), out, lhs_min);
  out.append(BinOpToken(op));
  FormatNode(node.as_binary_rhs(), out, rhs_min);
  if (wrap) {
    out.push_back(')');
  }
}

void FormatRangeOp(const AstNode& node, std::string& out, int min_bp) {
  const bool wrap = kBpRange < min_bp;
  if (wrap) {
    out.push_back('(');
  }
  FormatNode(node.as_range_lhs(), out, kBpRange);
  out.push_back(':');
  FormatNode(node.as_range_rhs(), out, kBpRange + 1);
  if (wrap) {
    out.push_back(')');
  }
}

void FormatIntersect(const AstNode& node, std::string& out, int min_bp) {
  const bool wrap = kBpIntersect < min_bp;
  if (wrap) {
    out.push_back('(');
  }
  FormatNode(node.as_intersect_lhs(), out, kBpIntersect);
  out.push_back(' ');
  FormatNode(node.as_intersect_rhs(), out, kBpIntersect + 1);
  if (wrap) {
    out.push_back(')');
  }
}

void FormatUnion(const AstNode& node, std::string& out) {
  // Union is only well-formed inside a parenthesised expression. Always
  // wrap so no caller can accidentally absorb the comma into a higher-
  // precedence outer slot.
  out.push_back('(');
  const std::uint32_t n = node.as_union_arity();
  for (std::uint32_t i = 0; i < n; ++i) {
    if (i > 0) {
      out.push_back(',');
    }
    // Inside parens the slot is permissive; let the child decide whether
    // it needs internal parens via its own min_bp.
    FormatNode(node.as_union_child(i), out, 0);
  }
  out.push_back(')');
}

void FormatImplicitIntersection(const AstNode& node, std::string& out, int min_bp) {
  const bool wrap = kBpAtPrefix < min_bp;
  if (wrap) {
    out.push_back('(');
  }
  out.push_back('@');
  FormatNode(node.as_implicit_intersection_operand(), out, kBpAtPrefix);
  if (wrap) {
    out.push_back(')');
  }
}

void FormatCall(const AstNode& node, std::string& out) {
  out.append(node.as_call_name());
  out.push_back('(');
  const std::uint32_t n = node.as_call_arity();
  for (std::uint32_t i = 0; i < n; ++i) {
    if (i > 0) {
      out.push_back(',');
    }
    FormatNode(node.as_call_arg(i), out, 0);
  }
  out.push_back(')');
}

void FormatArrayLiteral(const AstNode& node, std::string& out) {
  out.push_back('{');
  const std::uint32_t rows = node.as_array_rows();
  const std::uint32_t cols = node.as_array_cols();
  for (std::uint32_t r = 0; r < rows; ++r) {
    if (r > 0) {
      out.push_back(';');
    }
    for (std::uint32_t c = 0; c < cols; ++c) {
      if (c > 0) {
        out.push_back(',');
      }
      FormatNode(node.as_array_element(r, c), out, 0);
    }
  }
  out.push_back('}');
}

void FormatLambda(const AstNode& node, std::string& out) {
  out.append("LAMBDA(");
  const std::uint32_t n = node.as_lambda_param_count();
  const std::uint32_t opt = node.as_lambda_optional_count();
  const std::uint32_t first_optional = n - opt;
  for (std::uint32_t i = 0; i < n; ++i) {
    if (i > 0) {
      out.push_back(',');
    }
    if (i >= first_optional) {
      out.push_back('[');
      out.append(node.as_lambda_param(i));
      out.push_back(']');
    } else {
      out.append(node.as_lambda_param(i));
    }
  }
  if (n > 0) {
    out.push_back(',');
  }
  FormatNode(node.as_lambda_body(), out, 0);
  out.push_back(')');
}

void FormatLet(const AstNode& node, std::string& out) {
  out.append("LET(");
  const std::uint32_t n = node.as_let_binding_count();
  for (std::uint32_t i = 0; i < n; ++i) {
    out.append(node.as_let_binding_name(i));
    out.push_back(',');
    FormatNode(node.as_let_binding_expr(i), out, 0);
    out.push_back(',');
  }
  FormatNode(node.as_let_body(), out, 0);
  out.push_back(')');
}

void FormatLambdaCall(const AstNode& node, std::string& out) {
  // Excel's immediately-invoked lambda renders as `<callee>(args...)`. The
  // callee itself may need parens when it is not a bare ident / lambda.
  const AstNode& callee = node.as_lambda_call_callee();
  // A bare Lambda atom is parseable directly without parens; everything
  // else (paren'd expression, prior LambdaCall, NameRef) is also fine
  // because the parser treats the postfix `(` as a high-precedence
  // operator. Still, wrap operator children defensively.
  if (callee.kind() == NodeKind::Lambda || callee.kind() == NodeKind::NameRef ||
      callee.kind() == NodeKind::LambdaCall) {
    FormatNode(callee, out, 0);
  } else {
    out.push_back('(');
    FormatNode(callee, out, 0);
    out.push_back(')');
  }
  out.push_back('(');
  const std::uint32_t n = node.as_lambda_call_arity();
  for (std::uint32_t i = 0; i < n; ++i) {
    if (i > 0) {
      out.push_back(',');
    }
    FormatNode(node.as_lambda_call_arg(i), out, 0);
  }
  out.push_back(')');
}

void FormatNode(const AstNode& node, std::string& out, int min_bp) {
  switch (node.kind()) {
    case NodeKind::Literal:
      FormatLiteral(node.as_literal(), out);
      return;
    case NodeKind::Ref:
      FormatRef(node.as_ref(), out);
      return;
    case NodeKind::SpillRef: {
      // Postfix `#`. Wrap when the parent slot demands tighter binding than
      // the postfix-hash level (very rare since `#` sits near the top).
      const bool wrap = kBpPostfixHash < min_bp;
      if (wrap) {
        out.push_back('(');
      }
      FormatRef(node.as_spill_ref(), out);
      out.push_back('#');
      if (wrap) {
        out.push_back(')');
      }
      return;
    }
    case NodeKind::ExternalRef:
      FormatExternalRef(node, out);
      return;
    case NodeKind::StructuredRef:
      FormatStructuredRef(node, out);
      return;
    case NodeKind::NameRef:
      out.append(node.as_name());
      return;
    case NodeKind::UnaryOp:
      FormatUnary(node, out, min_bp);
      return;
    case NodeKind::BinaryOp:
      FormatBinary(node, out, min_bp);
      return;
    case NodeKind::RangeOp:
      FormatRangeOp(node, out, min_bp);
      return;
    case NodeKind::UnionOp:
      FormatUnion(node, out);
      return;
    case NodeKind::IntersectOp:
      FormatIntersect(node, out, min_bp);
      return;
    case NodeKind::ImplicitIntersection:
      FormatImplicitIntersection(node, out, min_bp);
      return;
    case NodeKind::Call:
      FormatCall(node, out);
      return;
    case NodeKind::ArrayLiteral:
      FormatArrayLiteral(node, out);
      return;
    case NodeKind::Lambda:
      FormatLambda(node, out);
      return;
    case NodeKind::LetBinding:
      FormatLet(node, out);
      return;
    case NodeKind::LambdaCall:
      FormatLambdaCall(node, out);
      return;
    case NodeKind::ErrorLiteral:
      out.append(display_name(node.as_error_literal()));
      return;
    case NodeKind::ErrorPlaceholder:
      // The placeholder represents a parse failure; rendering it as #REF!
      // keeps the round-trip from producing unparseable text. Real callers
      // should not be re-formatting trees that contain placeholders.
      out.append(display_name(ErrorCode::Ref));
      return;
  }
}

}  // namespace

std::string format_formula(const AstNode& node) {
  std::string out;
  out.reserve(64);
  FormatNode(node, out, 0);
  return out;
}

}  // namespace parser
}  // namespace formulon
