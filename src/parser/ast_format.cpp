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

// Parenthesis pairs to print around each node, decided once by
// `CollectParens` for the formatters and the XLSB encoder alike.
using ParenCounts = std::unordered_map<const AstNode*, std::uint8_t>;

std::uint8_t ParenCount(const ParenCounts& parens, const AstNode& node) {
  const auto it = parens.find(&node);
  return it == parens.end() ? 0U : it->second;
}

// Forward declaration: the full recursion target, which prints `node`
// inside the parentheses `parens` gives it.
void FormatNode(const AstNode& node, std::string& out, const ParenCounts& parens);
// `FormatNode` without the node's own parentheses.
void FormatBare(const AstNode& node, std::string& out, const ParenCounts& parens);

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

// Appends the `Sheet1!` / `'My Sheet'!` qualifier of a sheet-scoped NameRef;
// nothing for an unqualified one.
void AppendNameSheetQualifier(const AstNode& node, std::string& out) {
  const std::string_view sheet = node.as_name_sheet();
  if (sheet.empty()) {
    return;
  }
  append_sheet_name(sheet, node.as_name_sheet_quoted(), out);
  out.push_back('!');
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

bool IsDecimal(std::string_view text) noexcept {
  if (text.empty()) {
    return false;
  }
  for (char c : text) {
    if (!detail::IsAsciiDigit(c)) {
      return false;
    }
  }
  return true;
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
      AppendQuoteEscaped(path, out);
      AppendQuoteEscaped(book, out);
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
  AppendQuoteEscaped(path, out);
  out.push_back('[');
  AppendQuoteEscaped(book, out);
  out.push_back(']');
  AppendQuoteEscaped(sheet, out);
  if (spans_sheets) {
    out.push_back(':');
    AppendQuoteEscaped(sheet_end, out);
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

void FormatUnary(const AstNode& node, std::string& out, const ParenCounts& parens) {
  const UnaryOp op = node.as_unary_op();
  if (op == UnaryOp::Percent) {
    FormatNode(node.as_unary_operand(), out, parens);
    out.push_back('%');
    return;
  }
  out.push_back(op == UnaryOp::Plus ? '+' : '-');
  FormatNode(node.as_unary_operand(), out, parens);
}

void FormatBinary(const AstNode& node, std::string& out, const ParenCounts& parens) {
  FormatNode(node.as_binary_lhs(), out, parens);
  out.append(binop_token(node.as_binary_op()));
  FormatNode(node.as_binary_rhs(), out, parens);
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
  return lhs.kind() == NodeKind::NameRef && IsBareColumnToken(lhs.as_name()) &&
         IsBareColumnToken(LeadingIdentifier(rhs));
}

// Renders a `RangeOp` over two whole-axis `Ref`s as the compact `A:C` /
// `1:3` form and reports whether it did.
//
// Multi-column (`A:C`) / multi-row (`1:3`) whole references are stored as a
// RangeOp over two whole-column / whole-row Refs. Each endpoint's
// `format_a1` output duplicates its own axis (`A:A`, `C:C`), so a naive
// `<lhs>:<rhs>` join emits `A:A:C:C`. That names the same rectangle, so no
// evaluated value changes, but it is not the spelling Excel stores and it
// is not a fixpoint of the formatter.
//
// Both emitters call this. Living in only one of them is what let the
// storage form drift from the canonical form while every value-level test
// stayed green.
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
bool TrySpliceWholeAxisPair(const AstNode& lhs, const AstNode& rhs, std::string& out) {
  if (lhs.kind() != NodeKind::Ref || rhs.kind() != NodeKind::Ref) {
    return false;
  }
  const Reference& lr = lhs.as_ref();
  const Reference& rr = rhs.as_ref();
  const bool both_full_col = lr.is_full_col && rr.is_full_col;
  const bool both_full_row = lr.is_full_row && rr.is_full_row;
  if (!both_full_col && !both_full_row) {
    return false;
  }
  if ((both_full_col && lr.col == rr.col) || (both_full_row && lr.row == rr.row)) {
    return false;
  }
  const std::string lhs_str = format_a1(lr);
  const std::string rhs_str = format_a1(rr);
  out.append(lhs_str, 0, lhs_str.find(':'));
  out.push_back(':');
  out.append(rhs_str, 0, rhs_str.find(':'));
  return true;
}

void FormatRangeOp(const AstNode& node, std::string& out, const ParenCounts& parens) {
  const AstNode& lhs = node.as_range_lhs();
  const AstNode& rhs = node.as_range_rhs();
  if (TrySpliceWholeAxisPair(lhs, rhs, out)) {
    return;
  }
  FormatNode(lhs, out, parens);
  out.push_back(':');
  FormatNode(rhs, out, parens);
}

void FormatIntersect(const AstNode& node, std::string& out, const ParenCounts& parens) {
  FormatNode(node.as_intersect_lhs(), out, parens);
  out.push_back(' ');
  FormatNode(node.as_intersect_rhs(), out, parens);
}

void FormatUnion(const AstNode& node, std::string& out, const ParenCounts& parens) {
  const std::uint32_t n = node.as_union_arity();
  for (std::uint32_t i = 0; i < n; ++i) {
    if (i > 0) {
      out.push_back(',');
    }
    FormatNode(node.as_union_child(i), out, parens);
  }
}

// Spells the trim-reference operator for `mode`.
const char* TrimRefOperator(TrimRefMode mode) noexcept {
  switch (mode) {
    case TrimRefMode::Leading:
      return ".:";
    case TrimRefMode::Trailing:
      return ":.";
    case TrimRefMode::Both:
      return ".:.";
    case TrimRefMode::None:
      break;
  }
  return ":";
}

// Prints a `_TRO_*` call over a range as its operator (`A1:.A10`, `A.:A`) by
// respelling the range's own `:`; false, printing nothing, for any other
// argument.
bool FormatTrimRef(const AstNode& node, TrimRefMode mode, std::string& out, const ParenCounts& parens) {
  const AstNode& arg = node.as_call_arg(0);
  std::string text;
  std::size_t colon = std::string::npos;
  if (arg.kind() == NodeKind::RangeOp && ParenCount(parens, arg) == 0U &&
      !TrySpliceWholeAxisPair(arg.as_range_lhs(), arg.as_range_rhs(), text)) {
    FormatNode(arg.as_range_lhs(), text, parens);
    colon = text.size();
    text.push_back(':');
    FormatNode(arg.as_range_rhs(), text, parens);
  } else if ((arg.kind() == NodeKind::Ref || arg.kind() == NodeKind::RangeOp) && ParenCount(parens, arg) == 0U) {
    // A whole column or row prints as one token; sheet names hold no `:`.
    if (text.empty()) {
      FormatBare(arg, text, parens);
    }
    colon = text.rfind(':');
  }
  if (colon == std::string::npos) {
    return false;
  }
  out.append(text, 0, colon);
  out.append(TrimRefOperator(mode));
  out.append(text, colon + 1, std::string::npos);
  return true;
}

void FormatCall(const AstNode& node, std::string& out, const ParenCounts& parens) {
  if (const TrimRefMode mode = trim_ref_call_mode(node);
      mode != TrimRefMode::None && FormatTrimRef(node, mode, out, parens)) {
    return;
  }
  out.append(node.as_call_name());
  out.push_back('(');
  const std::uint32_t n = node.as_call_arity();
  for (std::uint32_t i = 0; i < n; ++i) {
    if (i > 0) {
      out.push_back(',');
    }
    FormatNode(node.as_call_arg(i), out, parens);
  }
  out.push_back(')');
}

// An Excel array constant admits only literal constants: no parentheses and
// no expressions. A negative element therefore has to be written with a bare
// sign, or the emitted text is text Excel -- and this parser -- rejects.
void FormatArrayElement(const AstNode& node, std::string& out, const ParenCounts& parens) {
  if (node.kind() == NodeKind::Literal) {
    FormatLiteral(node.as_literal(), out, /*parenthesise_negative=*/false);
    return;
  }
  FormatNode(node, out, parens);
}

void FormatArrayLiteral(const AstNode& node, std::string& out, const ParenCounts& parens) {
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
      FormatArrayElement(node.as_array_element(r, c), out, parens);
    }
  }
  out.push_back('}');
}

void FormatLambda(const AstNode& node, std::string& out, const ParenCounts& parens) {
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
  FormatNode(node.as_lambda_body(), out, parens);
  out.push_back(')');
}

void FormatLet(const AstNode& node, std::string& out, const ParenCounts& parens) {
  out.append("LET(");
  const std::uint32_t n = node.as_let_binding_count();
  for (std::uint32_t i = 0; i < n; ++i) {
    out.append(node.as_let_binding_name(i));
    out.push_back(',');
    FormatNode(node.as_let_binding_expr(i), out, parens);
    out.push_back(',');
  }
  FormatNode(node.as_let_body(), out, parens);
  out.push_back(')');
}

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

void FormatLambdaCall(const AstNode& node, std::string& out, const ParenCounts& parens) {
  FormatNode(node.as_lambda_call_callee(), out, parens);
  out.push_back('(');
  const std::uint32_t n = node.as_lambda_call_arity();
  for (std::uint32_t i = 0; i < n; ++i) {
    if (i > 0) {
      out.push_back(',');
    }
    FormatNode(node.as_lambda_call_arg(i), out, parens);
  }
  out.push_back(')');
}

void FormatBare(const AstNode& node, std::string& out, const ParenCounts& parens) {
  switch (node.kind()) {
    case NodeKind::Literal:
      FormatLiteral(node.as_literal(), out, /*parenthesise_negative=*/false);
      return;
    case NodeKind::Ref:
      FormatRef(node.as_ref(), out);
      return;
    case NodeKind::SpillRef:
      if (const AstNode* anchor = node.as_spill_ref_anchor_expr(); anchor != nullptr) {
        FormatNode(*anchor, out, parens);
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
      AppendNameSheetQualifier(node, out);
      out.append(node.as_name());
      return;
    case NodeKind::UnaryOp:
      FormatUnary(node, out, parens);
      return;
    case NodeKind::BinaryOp:
      FormatBinary(node, out, parens);
      return;
    case NodeKind::RangeOp:
      FormatRangeOp(node, out, parens);
      return;
    case NodeKind::UnionOp:
      FormatUnion(node, out, parens);
      return;
    case NodeKind::IntersectOp:
      FormatIntersect(node, out, parens);
      return;
    case NodeKind::ImplicitIntersection:
      out.push_back('@');
      FormatNode(node.as_implicit_intersection_operand(), out, parens);
      return;
    case NodeKind::Call:
      FormatCall(node, out, parens);
      return;
    case NodeKind::ArrayLiteral:
      FormatArrayLiteral(node, out, parens);
      return;
    case NodeKind::Lambda:
      FormatLambda(node, out, parens);
      return;
    case NodeKind::LetBinding:
      FormatLet(node, out, parens);
      return;
    case NodeKind::LambdaCall:
      FormatLambdaCall(node, out, parens);
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

void FormatNode(const AstNode& node, std::string& out, const ParenCounts& parens) {
  const std::uint8_t n = ParenCount(parens, node);
  out.append(n, '(');
  FormatBare(node, out, parens);
  out.append(n, ')');
}

// The one place that decides how many parenthesis pairs each node prints
// with: the pairs written around it, or one where its slot's binding power
// demands it. The formatters and the XLSB `PtgParen`s both read the result.
void Parenthesize(const AstNode& node, bool needed, ParenCounts& out) {
  const std::uint8_t count = std::max<std::uint8_t>(node.paren_depth(), needed ? 1U : 0U);
  if (count != 0U) {
    std::uint8_t& slot = out[&node];
    slot = std::max(slot, count);
  }
}

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
      CollectParens(node.as_unary_operand(), bp, out);
      return;
    }
    case NodeKind::BinaryOp: {
      const int bp = BinOpBp(node.as_binary_op());
      wrap_if(bp);
      CollectParens(node.as_binary_lhs(), bp, out);
      CollectParens(node.as_binary_rhs(), bp + 1, out);
      return;
    }
    case NodeKind::RangeOp: {
      wrap_if(kBpRange);
      const AstNode& lhs = node.as_range_lhs();
      const AstNode& rhs = node.as_range_rhs();
      std::string scratch;
      if (TrySpliceWholeAxisPair(lhs, rhs, scratch)) {
        return;
      }
      Parenthesize(lhs, ColonEndpointsWouldFold(lhs, rhs), out);
      CollectParens(lhs, kBpRange, out);
      CollectParens(rhs, kBpRange + 1, out);
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

  bool in_scope(std::string_view name) const {
    // Case-insensitive: LET/LAMBDA parameter names resolve case-insensitively
    // (Excel folds ASCII case on name resolution -- see
    // `lookup_lexical_binding` in eval/compiler.cpp), so a NameRef spelled
    // in a different case than its binding is still the same parameter and
    // must carry the `_xlpm.` prefix, not be written as a free name.
    for (const std::string_view s : scope) {
      if (strings::case_insensitive_eq(s, name)) {
        return true;
      }
    }
    return false;
  }

  void append_function_name(std::string& out, std::string_view canonical) const { out.append(spell(canonical)); }

  void emit(const AstNode& node, std::string& out) {
    const std::uint8_t n = ParenCount(parens, node);
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
        if (sheet.empty() && in_scope(name)) {
          out.append("_xlpm.");
        } else if (function_values != nullptr &&
                   std::find(function_values->begin(), function_values->end(), &node) != function_values->end()) {
          // Excel upper-cases a built-in it stores as a value.
          out.append("_xleta.");
          out.append(strings::to_ascii_upper(name));
          return;
        }
        AppendNameSheetQualifier(node, out);
        out.append(name);
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
        emit(node.as_binary_lhs(), out);
        out.append(binop_token(node.as_binary_op()));
        emit(node.as_binary_rhs(), out);
        return;
      case NodeKind::RangeOp:
        // A whole-axis pair has no function name inside it, so the shared
        // splice renders both forms.
        if (!TrySpliceWholeAxisPair(node.as_range_lhs(), node.as_range_rhs(), out)) {
          emit(node.as_range_lhs(), out);
          out.push_back(':');
          emit(node.as_range_rhs(), out);
        }
        return;
      case NodeKind::UnionOp:
        for (std::uint32_t i = 0; i < node.as_union_arity(); ++i) {
          if (i > 0) {
            out.push_back(',');
          }
          emit(node.as_union_child(i), out);
        }
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
        out.append(display_name(node.as_error_literal()));
        return;
      case NodeKind::ErrorPlaceholder:
        out.append(display_name(ErrorCode::Ref));
        return;
    }
  }

  void emit_call(const AstNode& node, std::string& out) {
    // `f(-2)` calls a LET / LAMBDA parameter.
    if (in_scope(node.as_call_name())) {
      out.append("_xlpm.");
      out.append(node.as_call_name());
    } else {
      append_function_name(out, node.as_call_name());
    }
    out.push_back('(');
    const std::uint32_t n = node.as_call_arity();
    for (std::uint32_t i = 0; i < n; ++i) {
      if (i > 0) {
        out.push_back(',');
      }
      emit(node.as_call_arg(i), out);
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
          emit(node.as_array_element(r, c), out);
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
      out.append("_xlpm.");
      out.append(name);
      out.push_back(',');
      // A binding value may reference earlier bindings but not itself, so it
      // is emitted with the scope accumulated so far (before pushing `name`).
      emit(node.as_let_binding_expr(i), out);
      out.push_back(',');
      scope.push_back(name);
    }
    emit(node.as_let_body(), out);
    out.push_back(')');
    scope.resize(base);
  }

  void emit_lambda_call(const AstNode& node, std::string& out) {
    emit(node.as_lambda_call_callee(), out);
    out.push_back('(');
    const std::uint32_t n = node.as_lambda_call_arity();
    for (std::uint32_t i = 0; i < n; ++i) {
      if (i > 0) {
        out.push_back(',');
      }
      emit(node.as_lambda_call_arg(i), out);
    }
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

void append_sheet_name(std::string_view sheet, bool force_quote, std::string& out) {
  if (force_quote || local_sheet_needs_quoting_a1(sheet)) {
    out.push_back('\'');
    AppendQuoteEscaped(sheet, out);
    out.push_back('\'');
  } else {
    out.append(sheet);
  }
}

void append_external_qualifier(const AstNode& node, bool r1c1, std::string& out) {
  AppendExternalQualifier(node, r1c1, nullptr, out);
}

void collect_parenthesized_nodes(const AstNode& root, std::unordered_map<const AstNode*, std::uint8_t>& out) {
  if (ast_depth_within_limit(root, kMaxFormulaAstDepth)) {
    CollectParens(root, 0, out);
  }
}

std::string format_formula(const AstNode& node) {
  if (!ast_depth_within_limit(node, kMaxFormulaAstDepth)) {
    return "#REF!";
  }
  ParenCounts parens;
  CollectParens(node, 0, parens);
  std::string out;
  out.reserve(64);
  FormatNode(node, out, parens);
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
    for (const AstNode* child : child_nodes(node)) {
      pending.push_back(child);
    }
  }
  return false;
}

}  // namespace parser
}  // namespace formulon
