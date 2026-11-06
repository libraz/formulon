//
// Excel-compatible AST → formula text formatter. The formatter mirrors the
// Pratt parser's recursive structure: each kind emits its surface form and
// recurses into children, attaching parentheses whenever the child's
// effective binding power is below the parent slot's minimum. The output
// re-parses to a structurally equivalent AST (verified by the round-trip
// tests in `tests/unit/parser/ast_format_test.cpp`).

#include "parser/ast_format.h"

#include <algorithm>
#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include "parser/ast.h"
#include "parser/ast_format_internal.h"
#include "parser/formula_prefix.h"
#include "parser/parser_detail.h"
#include "parser/reference.h"
#include "utils/arena.h"
#include "utils/double_format.h"
#include "utils/expected.h"  // FM_CHECK
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace parser {
namespace {

using detail::AppendBindingName;
using detail::AppendCellShapedCallName;
using detail::AppendNameRef;
using detail::AppendParamName;
using detail::CollectParens;
using detail::HasParamPrefix;
using detail::InScope;
using detail::kBpAtPrefix;
using detail::kBpIntersect;
using detail::kBpPostfixHash;
using detail::kBpPostfixPercent;
using detail::kBpRange;
using detail::kBpUnaryPrefix;
using detail::NameScope;
using detail::PrintsSpliced;

// `ParenCounts` is decided once by `CollectParens` for the formatters and the
// XLSB encoder alike.
std::uint16_t ParenCount(const ParenCounts& parens, const AstNode& node) {
  const auto it = parens.find(&node);
  return it == parens.end() ? 0U : it->second;
}

// What the canonical formatter prints around each node besides its text.
struct Layout {
  ParenCounts parens;
  const FormulaWhitespace* whitespace = nullptr;
  NameScope* scope = nullptr;
};

// Appends the whitespace runs kept at `slot` / `index` of `node`, if any.
void AppendWhitespace(const Layout& layout, const AstNode& node, WhitespaceSlot slot, std::uint16_t index,
                      std::string& out) {
  if (layout.whitespace == nullptr) {
    return;
  }
  for (const WhitespaceRun& run : *layout.whitespace) {
    if (run.node == &node && run.slot == slot && run.index == index) {
      out.append(run.count, run.ch);
    }
  }
}

// Forward declaration: the full recursion target, which prints `node`
// inside the parentheses `layout` gives it.
void FormatNode(const AstNode& node, std::string& out, const Layout& layout);
// `FormatNode` without the node's own parentheses.
void FormatBare(const AstNode& node, std::string& out, const Layout& layout);

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
    format_formula_number(out, n);
    out.push_back(')');
    return;
  }
  format_formula_number(out, n);
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

// Appends the sheet-less cell part of a reference tail in A1 notation: a
// cell, a rectangle, or a whole column / row span (`A:C`, never `A:A:C:C`).
void AppendSheetlessArea(const Reference& first, const Reference& last, bool is_range, std::string& out) {
  Reference cell_no_sheet = first;
  cell_no_sheet.sheet = {};
  cell_no_sheet.sheet_quoted = false;
  const std::string first_text = format_a1(cell_no_sheet);
  if (!is_range) {
    out.append(first_text);
    return;
  }
  Reference end_no_sheet = last;
  end_no_sheet.sheet = {};
  end_no_sheet.sheet_quoted = false;
  const std::string end_text = format_a1(end_no_sheet);
  if ((cell_no_sheet.is_full_col && end_no_sheet.is_full_col) ||
      (cell_no_sheet.is_full_row && end_no_sheet.is_full_row)) {
    out.append(first_text, 0, first_text.find(':'));
    out.push_back(':');
    out.append(end_text, 0, end_text.find(':'));
    return;
  }
  out.append(first_text);
  out.push_back(':');
  out.append(end_text);
}

// A `#REF!` with its collapse site: `Sheet3!#REF!` in both forms; a dead
// sheet's `#REF!A1` in the formula bar, bare `#REF!` in storage.
void FormatErrorLiteral(const AstNode& node, bool storage, std::string& out) {
  const RefErrorSite* site = node.as_ref_error_site();
  if (site != nullptr && site->form == RefErrorSite::Form::kQualified) {
    append_sheet_name(site->sheet, site->sheet_quoted, out);
    out.push_back('!');
  }
  out.append(display_name(node.as_error_literal()));
  if (site == nullptr || site->form != RefErrorSite::Form::kDeadSheet || storage) {
    return;
  }
  if (!site->has_address) {
    out.append(display_name(ErrorCode::Ref));
    return;
  }
  AppendSheetlessArea(site->first, site->last, site->is_range, out);
}

bool IsDecimal(std::string_view text) noexcept {
  return !text.empty() && detail::IsAllAsciiDigits(text);
}

// The qualifier of a cross-workbook reference through its `!`. With an
// `indexer` the book is the link index and the directory is dropped, the
// shape a stored formula carries.
void AppendExternalQualifier(const AstNode& node, bool r1c1, const ExternalBookIndexer* indexer, std::string& out) {
  if (is_self_book_name_ref(node)) {
    out.append("[0]!");
    return;
  }
  std::string_view path = node.as_external_ref_path();
  std::string_view book = node.as_external_ref_book();
  const std::string_view sheet = node.as_external_ref_sheet();
  const std::string_view sheet_end = node.as_external_ref_sheet_end();
  std::string index_text;
  if (indexer != nullptr) {
    const std::uint32_t index = indexer->index(indexer->ctx, path, book);
    FM_CHECK(index != 0U, "cross-workbook reference names a book with no external link");
    index_text = std::to_string(index);
    book = index_text;
    path = {};
  }
  if (sheet.empty()) {
    // A book-scope name: `Book.xlsx!Name`, but `[1]!Name` for a book that
    // is a link index, which only the storage indexer and the XLSB decoder
    // produce.
    if (path.empty() && IsDecimal(book)) {
      out.push_back('[');
      out.append(book);
      out.append("]!");
      return;
    }
    if (path.empty() && !external_book_needs_quoting(book)) {
      out.append(book);
    } else {
      out.push_back('\'');
      append_quote_escaped(path, out);
      append_quote_escaped(book, out);
      out.push_back('\'');
    }
    out.push_back('!');
    return;
  }
  const bool spans_sheets = !sheet_end.empty();
  const bool quoted = !path.empty() || external_book_needs_quoting(book) || external_sheet_needs_quoting(sheet) ||
                      external_sheet_needs_quoting(sheet_end) || (spans_sheets && !r1c1);
  if (quoted) {
    out.push_back('\'');
  }
  append_quote_escaped(path, out);
  out.push_back('[');
  append_quote_escaped(book, out);
  out.push_back(']');
  append_quote_escaped(sheet, out);
  if (spans_sheets) {
    out.push_back(':');
    append_quote_escaped(sheet_end, out);
  }
  if (quoted) {
    out.push_back('\'');
  }
  out.push_back('!');
}

void FormatExternalRef(const AstNode& node, std::string& out, const ExternalBookIndexer* indexer = nullptr) {
  AppendExternalQualifier(node, /*r1c1=*/false, indexer, out);
  if (const std::string_view name = node.as_external_ref_name(); !name.empty()) {
    out.append(name);
    return;
  }
  AppendSheetlessArea(node.as_external_ref_cell(), node.as_external_ref_cell_end(), node.as_external_ref_is_range(),
                      out);
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
  if (local_sheet_needs_quoting_a1(begin) || local_sheet_needs_quoting_a1(end)) {
    out.push_back('\'');
    append_quote_escaped(begin, out);
    out.push_back(':');
    append_quote_escaped(end, out);
    out.push_back('\'');
  } else {
    out.append(begin);
    out.push_back(':');
    out.append(end);
  }
  out.push_back('!');
  AppendSheetlessArea(node.as_ref3d_cell(), node.as_ref3d_cell_end(), node.as_ref3d_is_range(), out);
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

void FormatUnary(const AstNode& node, std::string& out, const Layout& layout) {
  const UnaryOp op = node.as_unary_op();
  if (op == UnaryOp::Percent) {
    FormatNode(node.as_unary_operand(), out, layout);
    AppendWhitespace(layout, node, WhitespaceSlot::kBeforeOperator, 0, out);
    out.push_back('%');
    return;
  }
  out.push_back(op == UnaryOp::Plus ? '+' : '-');
  FormatNode(node.as_unary_operand(), out, layout);
}

void FormatBinary(const AstNode& node, std::string& out, const Layout& layout) {
  FormatNode(node.as_binary_operand(0), out, layout);
  for (std::uint32_t i = 1; i < node.as_binary_arity(); ++i) {
    AppendWhitespace(layout, node, WhitespaceSlot::kBeforeOperator, static_cast<std::uint16_t>(i - 1U), out);
    out.append(binop_token(node.as_binary_op(i - 1U)));
    FormatNode(node.as_binary_operand(i), out, layout);
  }
}

// Prints a `:` chain with `emit` printing each operand. The two halves of a
// whole-axis token print as one (`A:C`, each half cut at its own `:`);
// every other operand prints in full, so `H:H:E:E` keeps both of its
// tokens. Both A1 emitters share this so the spelling cannot drift
// between them. Returns the offset of the last `:` written.
template <typename Emit>
std::size_t EmitRangeChain(const AstNode& node, std::string& out, Emit&& emit) {
  std::size_t last_colon = std::string::npos;
  const std::uint32_t n = node.as_range_arity();
  for (std::uint32_t i = 0; i < n; ++i) {
    if (i > 0U) {
      last_colon = out.size();
      out.push_back(':');
    }
    const AstNode& operand = node.as_range_operand(i);
    if ((i > 0U && PrintsSpliced(node, i - 1U)) || (i + 1U < n && PrintsSpliced(node, i))) {
      std::string text;
      emit(operand, text);
      out.append(text, 0, text.find(':'));
    } else {
      emit(operand, out);
    }
  }
  return last_colon;
}

void FormatRangeOp(const AstNode& node, std::string& out, const Layout& layout) {
  EmitRangeChain(node, out, [&](const AstNode& operand, std::string& dst) { FormatNode(operand, dst, layout); });
}

void FormatIntersect(const AstNode& node, std::string& out, const Layout& layout) {
  FormatNode(node.as_intersect_lhs(), out, layout);
  AppendWhitespace(layout, node, WhitespaceSlot::kBeforeOperator, 0, out);
  out.push_back(' ');
  FormatNode(node.as_intersect_rhs(), out, layout);
}

// Emits `count` items separated by commas; `emit_at(i)` writes item `i`.
template <typename EmitAt>
void EmitCommaList(std::uint32_t count, std::string& out, EmitAt&& emit_at) {
  for (std::uint32_t i = 0; i < count; ++i) {
    if (i > 0) {
      out.push_back(',');
    }
    emit_at(i);
  }
}

// Emits an array constant; `emit_element` writes one element.
template <typename EmitElement>
void EmitArrayConstant(const AstNode& node, std::string& out, EmitElement&& emit_element) {
  out.push_back('{');
  const std::uint32_t rows = node.as_array_rows();
  const std::uint32_t cols = node.as_array_cols();
  for (std::uint32_t r = 0; r < rows; ++r) {
    if (r > 0) {
      out.push_back(';');
    }
    EmitCommaList(cols, out, [&](std::uint32_t c) { emit_element(node.as_array_element(r, c)); });
  }
  out.push_back('}');
}

void FormatUnion(const AstNode& node, std::string& out, const Layout& layout) {
  EmitCommaList(node.as_union_arity(), out, [&](std::uint32_t i) { FormatNode(node.as_union_child(i), out, layout); });
}

// Prints a `_TRO_*` call over a range as its operator (`A1:.A10`, `A.:A`) by
// respelling the range's own `:`; false, printing nothing, for any other
// argument.
bool FormatTrimRef(const AstNode& node, TrimRefMode mode, std::string& out, const Layout& layout) {
  const AstNode& arg = node.as_call_arg(0);
  std::string text;
  std::size_t colon = std::string::npos;
  if (arg.kind() == NodeKind::RangeOp && ParenCount(layout.parens, arg) == 0U) {
    // The operator respelled is the chain's last `:`.
    colon =
        EmitRangeChain(arg, text, [&](const AstNode& operand, std::string& dst) { FormatNode(operand, dst, layout); });
  } else if (arg.kind() == NodeKind::Ref && ParenCount(layout.parens, arg) == 0U) {
    // A whole column or row prints as one token; sheet names hold no `:`.
    FormatBare(arg, text, layout);
    colon = text.rfind(':');
  }
  if (colon == std::string::npos) {
    return false;
  }
  out.append(text, 0, colon);
  out.append(trim_ref_operator(mode));
  out.append(text, colon + 1, std::string::npos);
  return true;
}

void FormatCall(const AstNode& node, std::string& out, const Layout& layout) {
  if (const TrimRefMode mode = trim_ref_call_mode(node);
      mode != TrimRefMode::None && FormatTrimRef(node, mode, out, layout)) {
    return;
  }
  if (!AppendCellShapedCallName(node.as_call_name(), *layout.scope, out)) {
    out.append(node.as_call_name());
  }
  out.push_back('(');
  EmitCommaList(node.as_call_arity(), out, [&](std::uint32_t i) { FormatNode(node.as_call_arg(i), out, layout); });
  AppendWhitespace(layout, node, WhitespaceSlot::kBeforeClose, 0, out);
  out.push_back(')');
}

// An Excel array constant admits only literal constants: no parentheses and
// no expressions. A negative element therefore has to be written with a bare
// sign, or the emitted text is text Excel -- and this parser -- rejects.
void FormatArrayElement(const AstNode& node, std::string& out, const Layout& layout) {
  if (node.kind() == NodeKind::Literal) {
    FormatLiteral(node.as_literal(), out, /*parenthesise_negative=*/false);
    return;
  }
  FormatNode(node, out, layout);
}

void FormatArrayLiteral(const AstNode& node, std::string& out, const Layout& layout) {
  EmitArrayConstant(node, out, [&](const AstNode& element) { FormatArrayElement(element, out, layout); });
}

void FormatLambda(const AstNode& node, std::string& out, const Layout& layout) {
  out.append("LAMBDA(");
  const std::uint32_t n = node.as_lambda_param_count();
  const std::uint32_t opt = node.as_lambda_optional_count();
  const std::uint32_t first_optional = n - opt;
  const std::size_t base = layout.scope->size();
  for (std::uint32_t i = 0; i < n; ++i) {
    if (i > 0) {
      out.push_back(',');
    }
    if (i >= first_optional) {
      out.push_back('[');
      AppendBindingName(node.as_lambda_param(i), /*storage=*/false, out);
      out.push_back(']');
    } else {
      AppendBindingName(node.as_lambda_param(i), /*storage=*/false, out);
    }
    layout.scope->push_back(cell_shaped_name_bare(node.as_lambda_param(i)));
  }
  if (n > 0) {
    out.push_back(',');
  }
  FormatNode(node.as_lambda_body(), out, layout);
  out.push_back(')');
  layout.scope->resize(base);
}

void FormatLet(const AstNode& node, std::string& out, const Layout& layout) {
  out.append("LET(");
  const std::uint32_t n = node.as_let_binding_count();
  const std::size_t base = layout.scope->size();
  for (std::uint32_t i = 0; i < n; ++i) {
    AppendBindingName(node.as_let_binding_name(i), /*storage=*/false, out);
    out.push_back(',');
    // A binding value sees the earlier bindings only.
    FormatNode(node.as_let_binding_expr(i), out, layout);
    out.push_back(',');
    layout.scope->push_back(cell_shaped_name_bare(node.as_let_binding_name(i)));
  }
  FormatNode(node.as_let_body(), out, layout);
  out.push_back(')');
  layout.scope->resize(base);
}

void FormatLambdaCall(const AstNode& node, std::string& out, const Layout& layout) {
  FormatNode(node.as_lambda_call_callee(), out, layout);
  out.push_back('(');
  EmitCommaList(node.as_lambda_call_arity(), out,
                [&](std::uint32_t i) { FormatNode(node.as_lambda_call_arg(i), out, layout); });
  out.push_back(')');
}

void FormatBare(const AstNode& node, std::string& out, const Layout& layout) {
  switch (node.kind()) {
    case NodeKind::Literal:
      FormatLiteral(node.as_literal(), out, /*parenthesise_negative=*/false);
      return;
    case NodeKind::Ref:
      FormatRef(node.as_ref(), out);
      return;
    case NodeKind::SpillRef:
      if (const AstNode* anchor = node.as_spill_ref_anchor_expr(); anchor != nullptr) {
        FormatNode(*anchor, out, layout);
      } else {
        FormatRef(node.as_spill_ref(), out);
      }
      out.push_back('#');
      return;
    case NodeKind::Ref3D:
      FormatRef3D(node, out);
      return;
    case NodeKind::ExternalRef:
      FormatExternalRef(node, out);
      return;
    case NodeKind::StructuredRef:
      FormatStructuredRef(node, out);
      return;
    case NodeKind::NameRef:
      AppendNameRef(node, *layout.scope, out);
      return;
    case NodeKind::UnaryOp:
      FormatUnary(node, out, layout);
      return;
    case NodeKind::BinaryOp:
      FormatBinary(node, out, layout);
      return;
    case NodeKind::RangeOp:
      FormatRangeOp(node, out, layout);
      return;
    case NodeKind::UnionOp:
      FormatUnion(node, out, layout);
      return;
    case NodeKind::IntersectOp:
      FormatIntersect(node, out, layout);
      return;
    case NodeKind::ImplicitIntersection:
      out.push_back('@');
      FormatNode(node.as_implicit_intersection_operand(), out, layout);
      return;
    case NodeKind::Call:
      FormatCall(node, out, layout);
      return;
    case NodeKind::ArrayLiteral:
      FormatArrayLiteral(node, out, layout);
      return;
    case NodeKind::Lambda:
      FormatLambda(node, out, layout);
      return;
    case NodeKind::LetBinding:
      FormatLet(node, out, layout);
      return;
    case NodeKind::LambdaCall:
      FormatLambdaCall(node, out, layout);
      return;
    case NodeKind::ErrorLiteral:
      FormatErrorLiteral(node, /*storage=*/false, out);
      return;
    case NodeKind::ErrorPlaceholder:
      // The placeholder represents a parse failure; rendering it as #REF!
      // keeps the round-trip from producing unparseable text. Real callers
      // should not be re-formatting trees that contain placeholders.
      out.append(display_name(ErrorCode::Ref));
      return;
  }
}

void FormatNode(const AstNode& node, std::string& out, const Layout& layout) {
  const std::uint16_t n = ParenCount(layout.parens, node);
  for (unsigned i = n; i >= 1U; --i) {
    AppendWhitespace(layout, node, WhitespaceSlot::kBeforeOpen, static_cast<std::uint16_t>(i), out);
    out.push_back('(');
  }
  AppendWhitespace(layout, node, WhitespaceSlot::kBeforeText, 0, out);
  FormatBare(node, out, layout);
  for (unsigned i = 1; i <= n; ++i) {
    AppendWhitespace(layout, node, WhitespaceSlot::kBeforeClose, static_cast<std::uint16_t>(i), out);
    out.push_back(')');
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
  const std::vector<const AstNode*>* omitted_at = nullptr;
  const ExternalBookIndexer* indexer = nullptr;
  const std::vector<const AstNode*>* function_values = nullptr;
  ParenCounts parens;

  // A NameRef spelled in a different case than its binding is still that
  // parameter and must carry the `_xlpm.` prefix.
  bool in_scope(std::string_view name) const { return InScope(scope, name); }

  void append_function_name(std::string& out, std::string_view canonical) const { out.append(spell(canonical)); }

  void emit(const AstNode& node, std::string& out) {
    const std::uint16_t n = ParenCount(parens, node);
    out.append(n, '(');
    emit_bare(node, out);
    out.append(n, ')');
  }

  void emit_bare(const AstNode& node, std::string& out) {
    switch (node.kind()) {
      case NodeKind::Literal:
        FormatLiteral(node.as_literal(), out, /*parenthesise_negative=*/false);
        return;
      case NodeKind::Ref:
        FormatRef(node.as_ref(), out);
        return;
      case NodeKind::SpillRef: {
        // Excel's on-disk spelling of the postfix `#` operator is a call to
        // the internal `ANCHORARRAY` function (see
        // eval/dynamic_array/anchor.h), not the bare `<ref>#` the canonical
        // formatter uses. `append_function_name` supplies the `_xlfn.`
        // prefix the same way it does for any other future function.
        append_function_name(out, "ANCHORARRAY");
        out.push_back('(');
        if (const AstNode* anchor = node.as_spill_ref_anchor_expr(); anchor != nullptr) {
          emit(*anchor, out);
        } else {
          FormatRef(node.as_spill_ref(), out);
        }
        out.push_back(')');
        return;
      }
      case NodeKind::Ref3D:
        FormatRef3D(node, out);
        return;
      case NodeKind::ExternalRef:
        FormatExternalRef(node, out, indexer);
        return;
      case NodeKind::StructuredRef:
        FormatStructuredRef(node, out);
        return;
      case NodeKind::NameRef: {
        const std::string_view name = node.as_name();
        const std::string_view sheet = node.as_name_sheet();
        // An extensionless book's book-scope name, written as a sheet.
        if (!sheet.empty() && indexer != nullptr && indexer->qualifier_index != nullptr) {
          const std::uint32_t index = indexer->qualifier_index(indexer->ctx, sheet);
          if (index != 0U) {
            out.push_back('[');
            out.append(std::to_string(index));
            out.append("]!");
            out.append(name);
            return;
          }
        }
        // A sheet-qualified name is a workbook name, never a LET / LAMBDA
        // parameter.
        const std::string_view bare = cell_shaped_name_bare(name);
        const bool param = sheet.empty() && in_scope(bare);
        if (param && !is_a1_cell_shaped(bare) && !HasParamPrefix(name)) {
          out.append(kParamPrefix);
        } else if (!param && function_values != nullptr &&
                   std::find(function_values->begin(), function_values->end(), &node) != function_values->end()) {
          // Excel upper-cases a built-in it stores as a value.
          out.append("_xleta.");
          out.append(strings::to_ascii_upper(name));
          return;
        }
        AppendNameRef(node, scope, out);
        return;
      }
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
        EmitRangeChain(node, out, [this](const AstNode& operand, std::string& dst) { emit(operand, dst); });
        return;
      case NodeKind::UnionOp:
        EmitCommaList(node.as_union_arity(), out, [&](std::uint32_t i) { emit(node.as_union_child(i), out); });
        return;
      case NodeKind::IntersectOp:
        emit(node.as_intersect_lhs(), out);
        out.push_back(' ');
        emit(node.as_intersect_rhs(), out);
        return;
      case NodeKind::ImplicitIntersection:
        if (omitted_at != nullptr && std::find(omitted_at->begin(), omitted_at->end(), &node) != omitted_at->end()) {
          emit(node.as_implicit_intersection_operand(), out);
          return;
        }
        // Excel stores a written `@` as a call to `_xlfn.SINGLE`.
        append_function_name(out, "SINGLE");
        out.push_back('(');
        emit(node.as_implicit_intersection_operand(), out);
        out.push_back(')');
        return;
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
        FormatErrorLiteral(node, /*storage=*/true, out);
        return;
      case NodeKind::ErrorPlaceholder:
        out.append(display_name(ErrorCode::Ref));
        return;
    }
  }

  void emit_call(const AstNode& node, std::string& out) {
    // `f(-2)` calls a LET / LAMBDA parameter.
    const std::string_view name = node.as_call_name();
    if (AppendCellShapedCallName(name, scope, out)) {
      // Spelled with its prefix already.
    } else if (in_scope(name)) {
      AppendParamName(name, out);
    } else {
      append_function_name(out, name);
    }
    out.push_back('(');
    EmitCommaList(node.as_call_arity(), out, [&](std::uint32_t i) { emit(node.as_call_arg(i), out); });
    out.push_back(')');
  }

  void emit_array(const AstNode& node, std::string& out) {
    EmitArrayConstant(node, out, [&](const AstNode& element) {
      if (element.kind() == NodeKind::Literal) {
        FormatLiteral(element.as_literal(), out, /*parenthesise_negative=*/false);
      } else {
        emit(element, out);
      }
    });
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
        AppendBindingName(param, /*storage=*/true, out);
        out.push_back(']');
      } else {
        AppendBindingName(param, /*storage=*/true, out);
      }
      scope.push_back(cell_shaped_name_bare(param));
    }
    if (n > 0) {
      out.push_back(',');
    }
    emit(node.as_lambda_body(), out);
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
      AppendBindingName(name, /*storage=*/true, out);
      out.push_back(',');
      // A binding value may reference earlier bindings but not itself, so it
      // is emitted with the scope accumulated so far (before pushing `name`).
      emit(node.as_let_binding_expr(i), out);
      out.push_back(',');
      scope.push_back(cell_shaped_name_bare(name));
    }
    emit(node.as_let_body(), out);
    out.push_back(')');
    scope.resize(base);
  }

  void emit_lambda_call(const AstNode& node, std::string& out) {
    emit(node.as_lambda_call_callee(), out);
    out.push_back('(');
    EmitCommaList(node.as_lambda_call_arity(), out, [&](std::uint32_t i) { emit(node.as_lambda_call_arg(i), out); });
    out.push_back(')');
  }
};

}  // namespace

const char* binop_token(BinOp op) noexcept {
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

void append_external_qualifier(const AstNode& node, bool r1c1, std::string& out) {
  AppendExternalQualifier(node, r1c1, nullptr, out);
}

void collect_parenthesized_nodes(const AstNode& root, ParenCounts& out) {
  if (ast_depth_within_limit(root, kMaxFormulaAstDepth)) {
    CollectParens(root, 0, out);
  }
}

const WhitespaceRun* find_whitespace(const FormulaWhitespace& whitespace, const AstNode& node, WhitespaceSlot slot,
                                     std::uint16_t index) {
  for (const WhitespaceRun& run : whitespace) {
    if (run.node == &node && run.slot == slot && run.index == index) {
      return &run;
    }
  }
  return nullptr;
}

std::string format_formula(const AstNode& node, const FormulaWhitespace* whitespace) {
  return format_formula_in_scope(node, {}, whitespace);
}

std::string format_formula_in_scope(const AstNode& node, const std::vector<std::string_view>& scope_names,
                                    const FormulaWhitespace* whitespace) {
  if (!ast_depth_within_limit(node, kMaxFormulaAstDepth)) {
    return "#REF!";
  }
  NameScope scope;
  scope.reserve(scope_names.size());
  for (const std::string_view name : scope_names) {
    scope.push_back(cell_shaped_name_bare(name));
  }
  Layout layout;
  layout.whitespace = whitespace != nullptr && !whitespace->empty() ? whitespace : nullptr;
  layout.scope = &scope;
  CollectParens(node, 0, layout.parens);
  std::string out;
  out.reserve(64);
  FormatNode(node, out, layout);
  return out;
}

std::string format_formula_storage(const AstNode& node, StorageFunctionNameSpeller spell,
                                   const std::vector<const AstNode*>* omitted_at, const ExternalBookIndexer* indexer,
                                   const std::vector<const AstNode*>* function_values) {
  if (!ast_depth_within_limit(node, kMaxFormulaAstDepth)) {
    return "#REF!";
  }
  StorageEmitter emitter{spell, {}, omitted_at, indexer, function_values, {}};
  CollectParens(node, 0, emitter.parens);
  std::string out;
  out.reserve(64);
  emitter.emit(node, out);
  return out;
}

bool formula_needs_storage_requote(const AstNode& root) {
  const auto bare_but_needs_quotes = [](std::string_view sheet, bool quoted) {
    return !sheet.empty() && !quoted && local_sheet_needs_quoting_a1(sheet);
  };
  std::vector<const AstNode*> pending{&root};
  while (!pending.empty()) {
    const AstNode& node = *pending.back();
    pending.pop_back();
    switch (node.kind()) {
      case NodeKind::Ref:
        if (bare_but_needs_quotes(node.as_ref().sheet, node.as_ref().sheet_quoted)) {
          return true;
        }
        break;
      case NodeKind::SpillRef:
        if (node.as_spill_ref_anchor_expr() == nullptr &&
            bare_but_needs_quotes(node.as_spill_ref().sheet, node.as_spill_ref().sheet_quoted)) {
          return true;
        }
        break;
      case NodeKind::NameRef:
        if (bare_but_needs_quotes(node.as_name_sheet(), node.as_name_sheet_quoted())) {
          return true;
        }
        break;
      case NodeKind::ErrorLiteral:
        if (const RefErrorSite* site = node.as_ref_error_site();
            site != nullptr && site->form == RefErrorSite::Form::kQualified &&
            bare_but_needs_quotes(site->sheet, site->sheet_quoted)) {
          return true;
        }
        break;
      case NodeKind::Ref3D:
        // The span keeps no quoting hint; requoting text that was already
        // quoted rewrites it to the same spelling.
        if (local_sheet_needs_quoting_a1(node.as_ref3d_sheet_begin()) ||
            local_sheet_needs_quoting_a1(node.as_ref3d_sheet_end())) {
          return true;
        }
        break;
      default:
        break;
    }
    if (node.kind() != NodeKind::ArrayLiteral) {
      any_child_node(node, [&](const AstNode& child) {
        pending.push_back(&child);
        return false;
      });
    }
  }
  return false;
}

}  // namespace parser
}  // namespace formulon
