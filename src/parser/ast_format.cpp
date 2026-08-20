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
#include <vector>

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

// Returns the binding-power slot for the given binary operator. Every
// binary operator is left-associative (including `^`, matching Excel 365's
// left-to-right evaluation of a chained power); the per-side `min_bp` calls
// in `FormatBinary` / `StorageEmitter::emit_binary` encode that uniformly.
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

// Renders a Number payload. Outside an array constant a negative value is
// parenthesised so its leading `-` cannot glue onto an adjacent operator;
// `parenthesise_negative` turns that off for slots where Excel's grammar
// forbids parentheses.
//
// The parser writes a negative numeric as a unary minus over a positive
// literal everywhere except inside an array constant, where the sign is
// folded into the element value.
void FormatNumberLiteral(const Value& v, std::string& out, bool parenthesise_negative) {
  const double n = v.as_number();
  if (n < 0.0 && parenthesise_negative) {
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

void FormatLiteral(const Value& v, std::string& out, bool parenthesise_negative = true) {
  switch (v.kind()) {
    case ValueKind::Blank:
      // A Blank literal represents an omitted argument. Its surrounding
      // call/array formatter emits the separators, which preserves the empty
      // slot (for example, FN(a,,b)). Rendering it as "" changes semantics:
      // many Excel functions distinguish an omitted argument from empty text.
      return;
    case ValueKind::Number:
      FormatNumberLiteral(v, out, parenthesise_negative);
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

// Appends `s`, doubling any embedded single quotes per Excel's escaping
// convention.
void AppendQuoteEscaped(std::string_view s, std::string& out) {
  for (char c : s) {
    if (c == '\'') {
      out.push_back('\'');
    }
    out.push_back(c);
  }
}

void FormatRef3D(const AstNode& node, std::string& out) {
  // A 3-D range's sheet span (`SheetFrom:SheetTo`) quotes as a SINGLE
  // unit when either endpoint needs quoting -- `'Data:S2'!B1`, not
  // `Data:'S2'!B1` -- rather than quoting each sheet name independently.
  // Verified against a real Excel-365-produced package: a genuine 3-D
  // range from sheet `Data` to sheet `S2` (the latter ambiguous with a
  // cell reference) serialises as `SUM('Data:S2'!B1)`.
  const std::string_view begin = node.as_ref3d_sheet_begin();
  const std::string_view end = node.as_ref3d_sheet_end();
  if (sheet_name_needs_quoting(begin) || sheet_name_needs_quoting(end)) {
    out.push_back('\'');
    AppendQuoteEscaped(begin, out);
    out.push_back(':');
    AppendQuoteEscaped(end, out);
    out.push_back('\'');
  } else {
    out.append(begin);
    out.push_back(':');
    out.append(end);
  }
  out.push_back('!');
  Reference cell_no_sheet = node.as_ref3d_cell();
  cell_no_sheet.sheet = {};
  cell_no_sheet.sheet_quoted = false;
  out.append(format_a1(cell_no_sheet));
  // Range tail (`'Data:S2'!A1:B2`): append the bottom-right corner.
  if (node.as_ref3d_is_range()) {
    Reference end_no_sheet = node.as_ref3d_cell_end();
    end_no_sheet.sheet = {};
    end_no_sheet.sheet_quoted = false;
    if ((cell_no_sheet.is_full_col && end_no_sheet.is_full_col) ||
        (cell_no_sheet.is_full_row && end_no_sheet.is_full_row)) {
      const std::string end_ref_text = format_a1(end_no_sheet);
      const std::size_t begin_axis_end = out.find(':', out.size() - format_a1(cell_no_sheet).size());
      out.erase(begin_axis_end);
      out.push_back(':');
      out.append(end_ref_text, 0, end_ref_text.find(':'));
    } else {
      out.push_back(':');
      out.append(format_a1(end_no_sheet));
    }
  }
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
  // Left-associative: LHS demands `bp` (a same-precedence child on the left
  // does not need wrapping), RHS demands `bp + 1` (a same-precedence child
  // on the right must be wrapped so it does not silently re-associate on
  // reparse). Uniform across every binary operator, including `^`.
  const int lhs_min = bp;
  const int rhs_min = bp + 1;
  FormatNode(node.as_binary_lhs(), out, lhs_min);
  out.append(BinOpToken(op));
  FormatNode(node.as_binary_rhs(), out, rhs_min);
  if (wrap) {
    out.push_back(')');
  }
}

// True when `name` on its own is a bare column token -- one to three ASCII
// letters naming a column inside Excel's grid.
bool IsBareColumnToken(std::string_view name) noexcept {
  if (name.empty() || name.size() > 3) {
    return false;
  }
  std::uint32_t column = 0;
  for (char c : name) {
    if (!detail::IsAsciiLetter(c)) {
      return false;
    }
    const char upper = (c >= 'a' && c <= 'z') ? static_cast<char>(c - ('a' - 'A')) : c;
    column = column * 26U + static_cast<std::uint32_t>(upper - 'A') + 1U;
  }
  return column <= detail::kMaxColumn;
}

// The identifier a `:` endpoint opens with, when the endpoint is written as a
// bare identifier rather than as a reference. A call contributes its name
// because the parser sees the name before the argument list.
std::string_view LeadingIdentifier(const AstNode& node) noexcept {
  switch (node.kind()) {
    case NodeKind::NameRef:
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
  return lhs.kind() == NodeKind::NameRef && IsBareColumnToken(lhs.as_name()) &&
         IsBareColumnToken(LeadingIdentifier(rhs));
}

void FormatRangeOp(const AstNode& node, std::string& out, int min_bp) {
  const bool wrap = kBpRange < min_bp;
  if (wrap) {
    out.push_back('(');
  }
  const AstNode& lhs = node.as_range_lhs();
  const AstNode& rhs = node.as_range_rhs();
  // Multi-column (`A:C`) / multi-row (`1:3`) whole references are stored as a
  // RangeOp over two whole-column / whole-row Refs. Each endpoint's
  // `format_a1` output duplicates its own axis (`A:A`, `C:C`), so a naive
  // `<lhs>:<rhs>` join would emit `A:A:C:C`. Splice the two endpoints at
  // their leading axis token to reproduce the compact `A:C` / `1:3` form.
  //
  // Two endpoints sitting on the same axis are the one case the splice must
  // not touch. `RangeOp(A:A, A:A)` would splice to `A:A`, and `A:A` reads
  // back as the single whole-column Ref rather than as the pair it came
  // from. Anchoring does not rescue it -- `$A:A` is likewise one token, not
  // two -- so the test is on the axis index, not on the rendered text. The
  // parser only builds this shape from text that already spelled the pair
  // out (`=A:A:A:A`, `=$A:$A:A:A`), so compacting it discards what was
  // written; left unspliced the pair joins to text that reads back as
  // itself.
  if (lhs.kind() == NodeKind::Ref && rhs.kind() == NodeKind::Ref) {
    const Reference& lr = lhs.as_ref();
    const Reference& rr = rhs.as_ref();
    const bool both_full_col = lr.is_full_col && rr.is_full_col;
    const bool both_full_row = lr.is_full_row && rr.is_full_row;
    const bool same_axis = (both_full_col && lr.col == rr.col) || (both_full_row && lr.row == rr.row);
    if ((both_full_col || both_full_row) && !same_axis) {
      const std::string lhs_str = format_a1(lr);
      const std::string rhs_str = format_a1(rr);
      out.append(lhs_str, 0, lhs_str.find(':'));
      out.push_back(':');
      out.append(rhs_str, 0, rhs_str.find(':'));
      if (wrap) {
        out.push_back(')');
      }
      return;
    }
  }
  const bool split_endpoints = ColonEndpointsWouldFold(lhs, rhs);
  if (split_endpoints) {
    out.push_back('(');
  }
  FormatNode(lhs, out, kBpRange);
  if (split_endpoints) {
    out.push_back(')');
  }
  out.push_back(':');
  FormatNode(rhs, out, kBpRange + 1);
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

// An Excel array constant admits only literal constants: no parentheses and
// no expressions. A negative element therefore has to be written with a bare
// sign, or the emitted text is text Excel -- and this parser -- rejects.
void FormatArrayElement(const AstNode& node, std::string& out) {
  if (node.kind() == NodeKind::Literal) {
    FormatLiteral(node.as_literal(), out, /*parenthesise_negative=*/false);
    return;
  }
  FormatNode(node, out, 0);
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
      FormatArrayElement(node.as_array_element(r, c), out);
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
      if (const AstNode* anchor = node.as_spill_ref_anchor_expr(); anchor != nullptr) {
        // A computed anchor prints at the postfix level so a lower-binding
        // sub-expression parenthesises itself back into place.
        FormatNode(*anchor, out, kBpPostfixHash);
      } else {
        FormatRef(node.as_spill_ref(), out);
      }
      out.push_back('#');
      if (wrap) {
        out.push_back(')');
      }
      return;
    }
    case NodeKind::Ref3D:
      FormatRef3D(node, out);
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

// Storage-form emitter: mirrors `FormatNode`'s dispatch but spells each
// function name the way the file stores it (via the injected speller,
// which owns both the `_xlfn.` / `_xlfn._xlws.` prefixes and any
// name Excel accepts without storing) and applies the `_xlpm.` LET /
// LAMBDA parameter prefix (tracked through a lexical scope stack). Leaf /
// operator rendering reuses the same free helpers and precedence rules as
// the canonical formatter so the two stay in lockstep.
struct StorageEmitter {
  StorageFunctionNameSpeller spell;
  std::vector<std::string_view> scope;  // in-scope LET binding / LAMBDA param names

  bool in_scope(std::string_view name) const {
    for (const std::string_view s : scope) {
      if (s == name) {
        return true;
      }
    }
    return false;
  }

  void append_function_name(std::string& out, std::string_view canonical) const { out.append(spell(canonical)); }

  void emit(const AstNode& node, std::string& out, int min_bp) {
    switch (node.kind()) {
      case NodeKind::Literal:
        FormatLiteral(node.as_literal(), out);
        return;
      case NodeKind::Ref:
        FormatRef(node.as_ref(), out);
        return;
      case NodeKind::SpillRef: {
        const bool wrap = kBpPostfixHash < min_bp;
        if (wrap) {
          out.push_back('(');
        }
        if (const AstNode* anchor = node.as_spill_ref_anchor_expr(); anchor != nullptr) {
          // Recurse through `emit`, not `FormatNode`: a computed anchor may
          // itself call a function whose stored spelling carries a prefix.
          emit(*anchor, out, kBpPostfixHash);
        } else {
          FormatRef(node.as_spill_ref(), out);
        }
        out.push_back('#');
        if (wrap) {
          out.push_back(')');
        }
        return;
      }
      case NodeKind::Ref3D:
        FormatRef3D(node, out);
        return;
      case NodeKind::StructuredRef:
        FormatStructuredRef(node, out);
        return;
      case NodeKind::NameRef: {
        const std::string_view name = node.as_name();
        if (in_scope(name)) {
          out.append("_xlpm.");
        }
        out.append(name);
        return;
      }
      case NodeKind::UnaryOp:
        emit_unary(node, out, min_bp);
        return;
      case NodeKind::BinaryOp:
        emit_binary(node, out, min_bp);
        return;
      case NodeKind::RangeOp:
        emit_binary_ref(node, out, min_bp, kBpRange, ':');
        return;
      case NodeKind::UnionOp:
        emit_union(node, out);
        return;
      case NodeKind::IntersectOp:
        emit_binary_ref(node, out, min_bp, kBpIntersect, ' ');
        return;
      case NodeKind::ImplicitIntersection: {
        const bool wrap = kBpAtPrefix < min_bp;
        if (wrap) {
          out.push_back('(');
        }
        out.push_back('@');
        emit(node.as_implicit_intersection_operand(), out, kBpAtPrefix);
        if (wrap) {
          out.push_back(')');
        }
        return;
      }
      case NodeKind::Call:
        emit_call(node, out);
        return;
      case NodeKind::ArrayLiteral:
        emit_array(node, out);
        return;
      case NodeKind::Lambda:
        emit_lambda(node, out);
        return;
      case NodeKind::LetBinding:
        emit_let(node, out);
        return;
      case NodeKind::LambdaCall:
        emit_lambda_call(node, out);
        return;
      case NodeKind::ErrorLiteral:
        out.append(display_name(node.as_error_literal()));
        return;
      case NodeKind::ErrorPlaceholder:
        out.append(display_name(ErrorCode::Ref));
        return;
    }
  }

  void emit_unary(const AstNode& node, std::string& out, int min_bp) {
    const UnaryOp op = node.as_unary_op();
    if (op == UnaryOp::Percent) {
      const bool wrap = kBpPostfixPercent < min_bp;
      if (wrap) {
        out.push_back('(');
      }
      emit(node.as_unary_operand(), out, kBpPostfixPercent);
      out.push_back('%');
      if (wrap) {
        out.push_back(')');
      }
      return;
    }
    const bool wrap = kBpUnaryPrefix < min_bp;
    if (wrap) {
      out.push_back('(');
    }
    out.push_back(op == UnaryOp::Plus ? '+' : '-');
    emit(node.as_unary_operand(), out, kBpUnaryPrefix);
    if (wrap) {
      out.push_back(')');
    }
  }

  void emit_binary(const AstNode& node, std::string& out, int min_bp) {
    const BinOp op = node.as_binary_op();
    const int bp = BinOpBp(op);
    const bool wrap = bp < min_bp;
    if (wrap) {
      out.push_back('(');
    }
    // Left-associative, uniformly across every binary operator including
    // `^`; see `FormatBinary`'s comment for the rationale.
    const int lhs_min = bp;
    const int rhs_min = bp + 1;
    emit(node.as_binary_lhs(), out, lhs_min);
    out.append(BinOpToken(op));
    emit(node.as_binary_rhs(), out, rhs_min);
    if (wrap) {
      out.push_back(')');
    }
  }

  // Shared shape for the two reference binary operators (`:` range and the
  // space intersect): wrap at `bp`, emit lhs at `bp`, rhs at `bp + 1`.
  void emit_binary_ref(const AstNode& node, std::string& out, int min_bp, int bp, char sep) {
    const bool wrap = bp < min_bp;
    if (wrap) {
      out.push_back('(');
    }
    if (node.kind() == NodeKind::RangeOp) {
      const bool split_endpoints = ColonEndpointsWouldFold(node.as_range_lhs(), node.as_range_rhs());
      if (split_endpoints) {
        out.push_back('(');
      }
      emit(node.as_range_lhs(), out, bp);
      if (split_endpoints) {
        out.push_back(')');
      }
      out.push_back(sep);
      emit(node.as_range_rhs(), out, bp + 1);
    } else {
      emit(node.as_intersect_lhs(), out, bp);
      out.push_back(sep);
      emit(node.as_intersect_rhs(), out, bp + 1);
    }
    if (wrap) {
      out.push_back(')');
    }
  }

  void emit_union(const AstNode& node, std::string& out) {
    out.push_back('(');
    const std::uint32_t n = node.as_union_arity();
    for (std::uint32_t i = 0; i < n; ++i) {
      if (i > 0) {
        out.push_back(',');
      }
      emit(node.as_union_child(i), out, 0);
    }
    out.push_back(')');
  }

  void emit_call(const AstNode& node, std::string& out) {
    append_function_name(out, node.as_call_name());
    out.push_back('(');
    const std::uint32_t n = node.as_call_arity();
    for (std::uint32_t i = 0; i < n; ++i) {
      if (i > 0) {
        out.push_back(',');
      }
      emit(node.as_call_arg(i), out, 0);
    }
    out.push_back(')');
  }

  void emit_array(const AstNode& node, std::string& out) {
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
        if (node.as_array_element(r, c).kind() == NodeKind::Literal) {
          FormatLiteral(node.as_array_element(r, c).as_literal(), out, /*parenthesise_negative=*/false);
        } else {
          emit(node.as_array_element(r, c), out, 0);
        }
      }
    }
    out.push_back('}');
  }

  void emit_lambda(const AstNode& node, std::string& out) {
    append_function_name(out, "LAMBDA");
    out.push_back('(');
    const std::uint32_t n = node.as_lambda_param_count();
    const std::uint32_t opt = node.as_lambda_optional_count();
    const std::uint32_t first_optional = n - opt;
    const std::size_t base = scope.size();
    for (std::uint32_t i = 0; i < n; ++i) {
      if (i > 0) {
        out.push_back(',');
      }
      const std::string_view param = node.as_lambda_param(i);
      if (i >= first_optional) {
        out.push_back('[');
        out.append("_xlpm.");
        out.append(param);
        out.push_back(']');
      } else {
        out.append("_xlpm.");
        out.append(param);
      }
      scope.push_back(param);
    }
    if (n > 0) {
      out.push_back(',');
    }
    emit(node.as_lambda_body(), out, 0);
    out.push_back(')');
    scope.resize(base);
  }

  void emit_let(const AstNode& node, std::string& out) {
    append_function_name(out, "LET");
    out.push_back('(');
    const std::uint32_t n = node.as_let_binding_count();
    const std::size_t base = scope.size();
    for (std::uint32_t i = 0; i < n; ++i) {
      const std::string_view name = node.as_let_binding_name(i);
      out.append("_xlpm.");
      out.append(name);
      out.push_back(',');
      // A binding value may reference earlier bindings but not itself, so it
      // is emitted with the scope accumulated so far (before pushing `name`).
      emit(node.as_let_binding_expr(i), out, 0);
      out.push_back(',');
      scope.push_back(name);
    }
    emit(node.as_let_body(), out, 0);
    out.push_back(')');
    scope.resize(base);
  }

  void emit_lambda_call(const AstNode& node, std::string& out) {
    const AstNode& callee = node.as_lambda_call_callee();
    if (callee.kind() == NodeKind::Lambda || callee.kind() == NodeKind::NameRef ||
        callee.kind() == NodeKind::LambdaCall) {
      emit(callee, out, 0);
    } else {
      out.push_back('(');
      emit(callee, out, 0);
      out.push_back(')');
    }
    out.push_back('(');
    const std::uint32_t n = node.as_lambda_call_arity();
    for (std::uint32_t i = 0; i < n; ++i) {
      if (i > 0) {
        out.push_back(',');
      }
      emit(node.as_lambda_call_arg(i), out, 0);
    }
    out.push_back(')');
  }
};

}  // namespace

std::string format_formula(const AstNode& node) {
  if (!ast_depth_within_limit(node, kMaxFormulaAstDepth)) {
    return "#REF!";
  }
  std::string out;
  out.reserve(64);
  FormatNode(node, out, 0);
  return out;
}

std::string format_formula_storage(const AstNode& node, StorageFunctionNameSpeller spell) {
  if (!ast_depth_within_limit(node, kMaxFormulaAstDepth)) {
    return "#REF!";
  }
  StorageEmitter emitter{spell, {}};
  std::string out;
  out.reserve(64);
  emitter.emit(node, out, 0);
  return out;
}

}  // namespace parser
}  // namespace formulon
