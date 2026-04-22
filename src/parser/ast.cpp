// Copyright 2026 libraz. Licensed under the MIT License.
//
// AST factories and per-kind accessor implementations. Storage policy: every
// pointer array and every string view stored in the AST lives in the arena
// passed to the factory, so the caller's source buffers and argv arrays may
// be released or mutated after the call returns.

#include "parser/ast.h"

#include <cstring>
#include <string>
#include <string_view>

#include "parser/reference.h"
#include "parser/token.h"
#include "utils/arena.h"
#include "utils/expected.h"  // FM_CHECK
#include "value.h"

namespace formulon {
namespace parser {
namespace {

// Copies an array of N const AstNode pointers into arena-owned storage.
// Returns nullptr on allocation failure or when n == 0 (callers handle the
// empty case explicitly).
const AstNode* const* CopyChildArray(Arena& arena, const AstNode* const* src, std::uint32_t n) {
  if (n == 0) {
    return nullptr;
  }
  auto* dst = arena.create_array<const AstNode*>(n);
  if (dst == nullptr) {
    return nullptr;
  }
  std::memcpy(dst, src, sizeof(const AstNode*) * n);
  return dst;
}

// Copies an array of N string_view payloads into arena-owned storage. Each
// string is itself re-interned so the AST owns both the pointer array and
// the byte storage. Returns nullptr on allocation failure or when n == 0.
const std::string_view* CopyNameArray(Arena& arena, const std::string_view* src, std::uint32_t n) {
  if (n == 0) {
    return nullptr;
  }
  auto* dst = arena.create_array<std::string_view>(n);
  if (dst == nullptr) {
    return nullptr;
  }
  for (std::uint32_t i = 0; i < n; ++i) {
    dst[i] = arena.intern(src[i]);
  }
  return dst;
}

}  // namespace

// ---------------------------------------------------------------------------
// Factories
// ---------------------------------------------------------------------------

AstNode* make_literal(Arena& arena, Value v) {
  AstNode* n = arena.create<AstNode>();
  if (n == nullptr) {
    return nullptr;
  }
  n->kind_ = NodeKind::Literal;
  n->data_.literal = v;
  return n;
}

AstNode* make_ref(Arena& arena, const Reference& r) {
  AstNode* n = arena.create<AstNode>();
  if (n == nullptr) {
    return nullptr;
  }
  n->kind_ = NodeKind::Ref;
  n->data_.ref = r;
  n->data_.ref.sheet = arena.intern(r.sheet);
  return n;
}

AstNode* make_external_ref(Arena& arena, std::uint32_t book_id, std::string_view sheet, const Reference& cell) {
  // Heap-allocate the payload so the AstNode union stays small (see the size
  // budget asserted in ast.h).
  auto* payload = arena.create<AstNode::ExternalRefPayload>();
  if (payload == nullptr) {
    return nullptr;
  }
  payload->book_id = book_id;
  payload->sheet = arena.intern(sheet);
  payload->cell = cell;
  payload->cell.sheet = arena.intern(cell.sheet);
  AstNode* n = arena.create<AstNode>();
  if (n == nullptr) {
    return nullptr;
  }
  n->kind_ = NodeKind::ExternalRef;
  n->data_.external_ref = payload;
  return n;
}

AstNode* make_structured_ref(Arena& arena, std::string_view table, std::string_view column,
                             StructuredRefModifier modifier) {
  AstNode* n = arena.create<AstNode>();
  if (n == nullptr) {
    return nullptr;
  }
  n->kind_ = NodeKind::StructuredRef;
  n->data_.structured_ref.table = arena.intern(table);
  n->data_.structured_ref.column = arena.intern(column);
  n->data_.structured_ref.modifier = modifier;
  return n;
}

AstNode* make_name_ref(Arena& arena, std::string_view name) {
  AstNode* n = arena.create<AstNode>();
  if (n == nullptr) {
    return nullptr;
  }
  n->kind_ = NodeKind::NameRef;
  n->data_.name = arena.intern(name);
  return n;
}

AstNode* make_unary_op(Arena& arena, UnaryOp op, AstNode* operand) {
  FM_CHECK(operand != nullptr, "make_unary_op: operand must be non-null");
  AstNode* n = arena.create<AstNode>();
  if (n == nullptr) {
    return nullptr;
  }
  n->kind_ = NodeKind::UnaryOp;
  n->data_.unary.op = op;
  n->data_.unary.operand = operand;
  return n;
}

AstNode* make_binary_op(Arena& arena, BinOp op, AstNode* lhs, AstNode* rhs) {
  FM_CHECK(lhs != nullptr && rhs != nullptr, "make_binary_op: lhs/rhs must be non-null");
  AstNode* n = arena.create<AstNode>();
  if (n == nullptr) {
    return nullptr;
  }
  n->kind_ = NodeKind::BinaryOp;
  n->data_.binary.op = op;
  n->data_.binary.lhs = lhs;
  n->data_.binary.rhs = rhs;
  return n;
}

AstNode* make_range_op(Arena& arena, AstNode* lhs, AstNode* rhs) {
  FM_CHECK(lhs != nullptr && rhs != nullptr, "make_range_op: lhs/rhs must be non-null");
  AstNode* n = arena.create<AstNode>();
  if (n == nullptr) {
    return nullptr;
  }
  n->kind_ = NodeKind::RangeOp;
  n->data_.range.lhs = lhs;
  n->data_.range.rhs = rhs;
  return n;
}

AstNode* make_union_op(Arena& arena, const AstNode* const* children, std::uint32_t count) {
  FM_CHECK(count >= 2, "make_union_op: union requires at least 2 children");
  FM_CHECK(children != nullptr, "make_union_op: children must be non-null");
  const AstNode* const* copied = CopyChildArray(arena, children, count);
  if (copied == nullptr) {
    return nullptr;
  }
  AstNode* n = arena.create<AstNode>();
  if (n == nullptr) {
    return nullptr;
  }
  n->kind_ = NodeKind::UnionOp;
  n->data_.variadic.children = copied;
  n->data_.variadic.count = count;
  return n;
}

AstNode* make_intersect_op(Arena& arena, AstNode* lhs, AstNode* rhs) {
  FM_CHECK(lhs != nullptr && rhs != nullptr, "make_intersect_op: lhs/rhs must be non-null");
  AstNode* n = arena.create<AstNode>();
  if (n == nullptr) {
    return nullptr;
  }
  n->kind_ = NodeKind::IntersectOp;
  // Reuse the binary RangePayload slot (same shape: two children).
  n->data_.range.lhs = lhs;
  n->data_.range.rhs = rhs;
  return n;
}

AstNode* make_implicit_intersection(Arena& arena, AstNode* operand) {
  FM_CHECK(operand != nullptr, "make_implicit_intersection: operand must be non-null");
  AstNode* n = arena.create<AstNode>();
  if (n == nullptr) {
    return nullptr;
  }
  n->kind_ = NodeKind::ImplicitIntersection;
  // Reuse the unary slot; UnaryOp::Plus is just a placeholder discriminator
  // never read by the implicit-intersection accessor.
  n->data_.unary.op = UnaryOp::Plus;
  n->data_.unary.operand = operand;
  return n;
}

AstNode* make_call(Arena& arena, std::string_view name, const AstNode* const* args, std::uint32_t arity) {
  const AstNode* const* copied = nullptr;
  if (arity > 0) {
    FM_CHECK(args != nullptr, "make_call: args must be non-null when arity > 0");
    copied = CopyChildArray(arena, args, arity);
    if (copied == nullptr) {
      return nullptr;
    }
  }
  AstNode* n = arena.create<AstNode>();
  if (n == nullptr) {
    return nullptr;
  }
  n->kind_ = NodeKind::Call;
  n->data_.call.name = arena.intern(name);
  n->data_.call.args = copied;
  n->data_.call.arity = arity;
  return n;
}

AstNode* make_array_literal(Arena& arena, std::uint32_t rows, std::uint32_t cols,
                            const AstNode* const* elems_row_major) {
  FM_CHECK(rows >= 1 && cols >= 1, "make_array_literal: rows and cols must be >= 1");
  FM_CHECK(elems_row_major != nullptr, "make_array_literal: elements must be non-null");
  const std::uint32_t total = rows * cols;
  // Overflow guard: if rows * cols overflows uint32 we cannot represent it.
  FM_CHECK(rows == 0 || total / rows == cols, "make_array_literal: rows*cols overflow");
  const AstNode* const* copied = CopyChildArray(arena, elems_row_major, total);
  if (copied == nullptr) {
    return nullptr;
  }
  AstNode* n = arena.create<AstNode>();
  if (n == nullptr) {
    return nullptr;
  }
  n->kind_ = NodeKind::ArrayLiteral;
  n->data_.array.elements = copied;
  n->data_.array.rows = rows;
  n->data_.array.cols = cols;
  return n;
}

AstNode* make_lambda(Arena& arena, const std::string_view* params, std::uint32_t param_count, AstNode* body) {
  FM_CHECK(body != nullptr, "make_lambda: body must be non-null");
  const std::string_view* copied_params = nullptr;
  if (param_count > 0) {
    FM_CHECK(params != nullptr, "make_lambda: params must be non-null when param_count > 0");
    copied_params = CopyNameArray(arena, params, param_count);
    if (copied_params == nullptr) {
      return nullptr;
    }
  }
  AstNode* n = arena.create<AstNode>();
  if (n == nullptr) {
    return nullptr;
  }
  n->kind_ = NodeKind::Lambda;
  n->data_.lambda.params = copied_params;
  n->data_.lambda.param_count = param_count;
  n->data_.lambda.body = body;
  return n;
}

AstNode* make_let_binding(Arena& arena, const std::string_view* names, const AstNode* const* exprs,
                          std::uint32_t binding_count, AstNode* body) {
  FM_CHECK(binding_count >= 1, "make_let_binding: binding_count must be >= 1");
  FM_CHECK(names != nullptr && exprs != nullptr, "make_let_binding: names/exprs must be non-null");
  FM_CHECK(body != nullptr, "make_let_binding: body must be non-null");
  const std::string_view* copied_names = CopyNameArray(arena, names, binding_count);
  if (copied_names == nullptr) {
    return nullptr;
  }
  const AstNode* const* copied_exprs = CopyChildArray(arena, exprs, binding_count);
  if (copied_exprs == nullptr) {
    return nullptr;
  }
  AstNode* n = arena.create<AstNode>();
  if (n == nullptr) {
    return nullptr;
  }
  n->kind_ = NodeKind::LetBinding;
  n->data_.let.names = copied_names;
  n->data_.let.exprs = copied_exprs;
  n->data_.let.binding_count = binding_count;
  n->data_.let.body = body;
  return n;
}

AstNode* make_lambda_call(Arena& arena, AstNode* callee, const AstNode* const* args, std::uint32_t arity) {
  FM_CHECK(callee != nullptr, "make_lambda_call: callee must be non-null");
  const AstNode* const* copied = nullptr;
  if (arity > 0) {
    FM_CHECK(args != nullptr, "make_lambda_call: args must be non-null when arity > 0");
    copied = CopyChildArray(arena, args, arity);
    if (copied == nullptr) {
      return nullptr;
    }
  }
  AstNode* n = arena.create<AstNode>();
  if (n == nullptr) {
    return nullptr;
  }
  n->kind_ = NodeKind::LambdaCall;
  n->data_.lambda_call.callee = callee;
  n->data_.lambda_call.args = copied;
  n->data_.lambda_call.arity = arity;
  return n;
}

AstNode* make_error_literal(Arena& arena, ErrorCode code) {
  AstNode* n = arena.create<AstNode>();
  if (n == nullptr) {
    return nullptr;
  }
  n->kind_ = NodeKind::ErrorLiteral;
  n->data_.error_literal = code;
  return n;
}

// ---------------------------------------------------------------------------
// Accessors
// ---------------------------------------------------------------------------

const Value& AstNode::as_literal() const {
  FM_CHECK(kind_ == NodeKind::Literal, "AstNode::as_literal on non-Literal");
  return data_.literal;
}

const Reference& AstNode::as_ref() const {
  FM_CHECK(kind_ == NodeKind::Ref, "AstNode::as_ref on non-Ref");
  return data_.ref;
}

std::uint32_t AstNode::as_external_ref_book_id() const {
  FM_CHECK(kind_ == NodeKind::ExternalRef, "AstNode::as_external_ref_book_id on non-ExternalRef");
  return data_.external_ref->book_id;
}

std::string_view AstNode::as_external_ref_sheet() const {
  FM_CHECK(kind_ == NodeKind::ExternalRef, "AstNode::as_external_ref_sheet on non-ExternalRef");
  return data_.external_ref->sheet;
}

const Reference& AstNode::as_external_ref_cell() const {
  FM_CHECK(kind_ == NodeKind::ExternalRef, "AstNode::as_external_ref_cell on non-ExternalRef");
  return data_.external_ref->cell;
}

std::string_view AstNode::as_structured_ref_table() const {
  FM_CHECK(kind_ == NodeKind::StructuredRef, "AstNode::as_structured_ref_table on non-StructuredRef");
  return data_.structured_ref.table;
}

std::string_view AstNode::as_structured_ref_column() const {
  FM_CHECK(kind_ == NodeKind::StructuredRef, "AstNode::as_structured_ref_column on non-StructuredRef");
  return data_.structured_ref.column;
}

StructuredRefModifier AstNode::as_structured_ref_modifier() const {
  FM_CHECK(kind_ == NodeKind::StructuredRef, "AstNode::as_structured_ref_modifier on non-StructuredRef");
  return data_.structured_ref.modifier;
}

std::string_view AstNode::as_name() const {
  FM_CHECK(kind_ == NodeKind::NameRef, "AstNode::as_name on non-NameRef");
  return data_.name;
}

UnaryOp AstNode::as_unary_op() const {
  FM_CHECK(kind_ == NodeKind::UnaryOp, "AstNode::as_unary_op on non-UnaryOp");
  return data_.unary.op;
}

const AstNode& AstNode::as_unary_operand() const {
  FM_CHECK(kind_ == NodeKind::UnaryOp, "AstNode::as_unary_operand on non-UnaryOp");
  return *data_.unary.operand;
}

BinOp AstNode::as_binary_op() const {
  FM_CHECK(kind_ == NodeKind::BinaryOp, "AstNode::as_binary_op on non-BinaryOp");
  return data_.binary.op;
}

const AstNode& AstNode::as_binary_lhs() const {
  FM_CHECK(kind_ == NodeKind::BinaryOp, "AstNode::as_binary_lhs on non-BinaryOp");
  return *data_.binary.lhs;
}

const AstNode& AstNode::as_binary_rhs() const {
  FM_CHECK(kind_ == NodeKind::BinaryOp, "AstNode::as_binary_rhs on non-BinaryOp");
  return *data_.binary.rhs;
}

const AstNode& AstNode::as_range_lhs() const {
  FM_CHECK(kind_ == NodeKind::RangeOp, "AstNode::as_range_lhs on non-RangeOp");
  return *data_.range.lhs;
}

const AstNode& AstNode::as_range_rhs() const {
  FM_CHECK(kind_ == NodeKind::RangeOp, "AstNode::as_range_rhs on non-RangeOp");
  return *data_.range.rhs;
}

std::uint32_t AstNode::as_union_arity() const {
  FM_CHECK(kind_ == NodeKind::UnionOp, "AstNode::as_union_arity on non-UnionOp");
  return data_.variadic.count;
}

const AstNode& AstNode::as_union_child(std::uint32_t i) const {
  FM_CHECK(kind_ == NodeKind::UnionOp, "AstNode::as_union_child on non-UnionOp");
  FM_CHECK(i < data_.variadic.count, "AstNode::as_union_child index out of range");
  return *data_.variadic.children[i];
}

const AstNode& AstNode::as_intersect_lhs() const {
  FM_CHECK(kind_ == NodeKind::IntersectOp, "AstNode::as_intersect_lhs on non-IntersectOp");
  return *data_.range.lhs;
}

const AstNode& AstNode::as_intersect_rhs() const {
  FM_CHECK(kind_ == NodeKind::IntersectOp, "AstNode::as_intersect_rhs on non-IntersectOp");
  return *data_.range.rhs;
}

const AstNode& AstNode::as_implicit_intersection_operand() const {
  FM_CHECK(kind_ == NodeKind::ImplicitIntersection,
           "AstNode::as_implicit_intersection_operand on non-ImplicitIntersection");
  return *data_.unary.operand;
}

std::string_view AstNode::as_call_name() const {
  FM_CHECK(kind_ == NodeKind::Call, "AstNode::as_call_name on non-Call");
  return data_.call.name;
}

std::uint32_t AstNode::as_call_arity() const {
  FM_CHECK(kind_ == NodeKind::Call, "AstNode::as_call_arity on non-Call");
  return data_.call.arity;
}

const AstNode& AstNode::as_call_arg(std::uint32_t i) const {
  FM_CHECK(kind_ == NodeKind::Call, "AstNode::as_call_arg on non-Call");
  FM_CHECK(i < data_.call.arity, "AstNode::as_call_arg index out of range");
  return *data_.call.args[i];
}

std::uint32_t AstNode::as_array_rows() const {
  FM_CHECK(kind_ == NodeKind::ArrayLiteral, "AstNode::as_array_rows on non-ArrayLiteral");
  return data_.array.rows;
}

std::uint32_t AstNode::as_array_cols() const {
  FM_CHECK(kind_ == NodeKind::ArrayLiteral, "AstNode::as_array_cols on non-ArrayLiteral");
  return data_.array.cols;
}

const AstNode& AstNode::as_array_element(std::uint32_t row, std::uint32_t col) const {
  FM_CHECK(kind_ == NodeKind::ArrayLiteral, "AstNode::as_array_element on non-ArrayLiteral");
  FM_CHECK(row < data_.array.rows && col < data_.array.cols, "AstNode::as_array_element index out of range");
  return *data_.array.elements[row * data_.array.cols + col];
}

std::uint32_t AstNode::as_lambda_param_count() const {
  FM_CHECK(kind_ == NodeKind::Lambda, "AstNode::as_lambda_param_count on non-Lambda");
  return data_.lambda.param_count;
}

std::string_view AstNode::as_lambda_param(std::uint32_t i) const {
  FM_CHECK(kind_ == NodeKind::Lambda, "AstNode::as_lambda_param on non-Lambda");
  FM_CHECK(i < data_.lambda.param_count, "AstNode::as_lambda_param index out of range");
  return data_.lambda.params[i];
}

const AstNode& AstNode::as_lambda_body() const {
  FM_CHECK(kind_ == NodeKind::Lambda, "AstNode::as_lambda_body on non-Lambda");
  return *data_.lambda.body;
}

std::uint32_t AstNode::as_let_binding_count() const {
  FM_CHECK(kind_ == NodeKind::LetBinding, "AstNode::as_let_binding_count on non-LetBinding");
  return data_.let.binding_count;
}

std::string_view AstNode::as_let_binding_name(std::uint32_t i) const {
  FM_CHECK(kind_ == NodeKind::LetBinding, "AstNode::as_let_binding_name on non-LetBinding");
  FM_CHECK(i < data_.let.binding_count, "AstNode::as_let_binding_name index out of range");
  return data_.let.names[i];
}

const AstNode& AstNode::as_let_binding_expr(std::uint32_t i) const {
  FM_CHECK(kind_ == NodeKind::LetBinding, "AstNode::as_let_binding_expr on non-LetBinding");
  FM_CHECK(i < data_.let.binding_count, "AstNode::as_let_binding_expr index out of range");
  return *data_.let.exprs[i];
}

const AstNode& AstNode::as_let_body() const {
  FM_CHECK(kind_ == NodeKind::LetBinding, "AstNode::as_let_body on non-LetBinding");
  return *data_.let.body;
}

const AstNode& AstNode::as_lambda_call_callee() const {
  FM_CHECK(kind_ == NodeKind::LambdaCall, "AstNode::as_lambda_call_callee on non-LambdaCall");
  return *data_.lambda_call.callee;
}

std::uint32_t AstNode::as_lambda_call_arity() const {
  FM_CHECK(kind_ == NodeKind::LambdaCall, "AstNode::as_lambda_call_arity on non-LambdaCall");
  return data_.lambda_call.arity;
}

const AstNode& AstNode::as_lambda_call_arg(std::uint32_t i) const {
  FM_CHECK(kind_ == NodeKind::LambdaCall, "AstNode::as_lambda_call_arg on non-LambdaCall");
  FM_CHECK(i < data_.lambda_call.arity, "AstNode::as_lambda_call_arg index out of range");
  return *data_.lambda_call.args[i];
}

ErrorCode AstNode::as_error_literal() const {
  FM_CHECK(kind_ == NodeKind::ErrorLiteral, "AstNode::as_error_literal on non-ErrorLiteral");
  return data_.error_literal;
}

// ---------------------------------------------------------------------------
// Reference helper
// ---------------------------------------------------------------------------

namespace {

// Encodes a 0-based column index as Excel column letters (0 -> "A", 25 -> "Z",
// 26 -> "AA", ..., 16383 -> "XFD"). Appends to `out`.
void AppendColumnLetters(std::string& out, std::uint32_t col) {
  // Excel uses bijective base-26: each letter is 1..26 with no zero digit.
  // We build the letters in reverse, then append them flipped.
  char buf[4];
  std::uint32_t i = 0;
  std::uint32_t v = col + 1;  // shift to 1-based for the bijective scheme.
  while (v > 0 && i < 4) {
    const std::uint32_t rem = (v - 1) % 26;
    buf[i++] = static_cast<char>('A' + rem);
    v = (v - 1) / 26;
  }
  while (i > 0) {
    out.push_back(buf[--i]);
  }
}

}  // namespace

std::string format_a1(const Reference& r) {
  std::string out;
  if (!r.sheet.empty()) {
    if (r.sheet_quoted) {
      out.push_back('\'');
      // Escape any embedded single quotes by doubling them.
      for (char c : r.sheet) {
        if (c == '\'') {
          out.push_back('\'');
        }
        out.push_back(c);
      }
      out.push_back('\'');
    } else {
      out.append(r.sheet);
    }
    out.push_back('!');
  }
  if (r.col_abs) {
    out.push_back('$');
  }
  AppendColumnLetters(out, r.col);
  if (r.row_abs) {
    out.push_back('$');
  }
  out.append(std::to_string(r.row + 1));
  return out;
}

}  // namespace parser
}  // namespace formulon
