// Copyright 2026 libraz. Licensed under the MIT License.
//
// Generic reference-transform walker over the parser AST.
//
// The walker rebuilds nodes only on the path that actually changes:
// identity walks return the original pointer. This keeps the relative-
// shift fast path (used per CF rule, per candidate cell) allocation-free
// when no reference is rewritten.

#include "parser/ast_shift.h"

#include <cstdint>
#include <optional>
#include <string_view>
#include <vector>

#include "parser/ast.h"
#include "parser/reference.h"
#include "utils/arena.h"

namespace formulon {
namespace parser {

// ---------------------------------------------------------------------------
// RefTransform default hooks
// ---------------------------------------------------------------------------

std::optional<Reference> RefTransform::apply_external(std::uint32_t /*book_id*/, std::string_view /*sheet*/,
                                                      const Reference& cell) const {
  return apply(cell);
}

std::optional<std::string_view> RefTransform::transform_external_sheet(std::uint32_t /*book_id*/,
                                                                       std::string_view /*sheet*/) const {
  return std::nullopt;
}

namespace {

// Excel's coordinate ceilings. Used by the relative-shift transform to
// detect out-of-bounds rewrites.
constexpr std::uint32_t kMaxColumn = 16384;  // XFD
constexpr std::uint32_t kMaxRow = 1048576;   // 2^20

// Builds a `#REF!` error literal node. Returns nullptr on arena failure.
AstNode* MakeRefError(Arena& arena) {
  return make_error_literal(arena, ErrorCode::Ref);
}

// Forward declaration: the recursive worker. Returns the rewritten node;
// when nothing changed it returns `&node` directly so the caller can skip
// allocation.
const AstNode* TransformNode(const AstNode& node, Arena& arena, const RefTransform& transform);

const AstNode* TransformRef(const AstNode& node, Arena& arena, const RefTransform& transform) {
  std::optional<Reference> rewritten = transform.apply(node.as_ref());
  if (!rewritten.has_value()) {
    return MakeRefError(arena);
  }
  // Identity short-circuit: avoid an allocation when the transform was a
  // no-op for this reference.
  const Reference& orig = node.as_ref();
  const Reference& nr = *rewritten;
  if (orig.sheet.data() == nr.sheet.data() && orig.sheet.size() == nr.sheet.size() &&
      orig.sheet_quoted == nr.sheet_quoted && orig.col == nr.col && orig.row == nr.row && orig.col_abs == nr.col_abs &&
      orig.row_abs == nr.row_abs && orig.is_full_col == nr.is_full_col && orig.is_full_row == nr.is_full_row) {
    return &node;
  }
  return make_ref(arena, nr);
}

const AstNode* TransformSpillRef(const AstNode& node, Arena& arena, const RefTransform& transform) {
  std::optional<Reference> rewritten = transform.apply(node.as_spill_ref());
  if (!rewritten.has_value()) {
    return MakeRefError(arena);
  }
  const Reference& orig = node.as_spill_ref();
  const Reference& nr = *rewritten;
  if (orig.sheet.data() == nr.sheet.data() && orig.sheet.size() == nr.sheet.size() &&
      orig.sheet_quoted == nr.sheet_quoted && orig.col == nr.col && orig.row == nr.row && orig.col_abs == nr.col_abs &&
      orig.row_abs == nr.row_abs && orig.is_full_col == nr.is_full_col && orig.is_full_row == nr.is_full_row) {
    return &node;
  }
  return make_spill_ref(arena, nr);
}

const AstNode* TransformExternalRef(const AstNode& node, Arena& arena, const RefTransform& transform) {
  const std::uint32_t book_id = node.as_external_ref_book_id();
  const std::string_view sheet = node.as_external_ref_sheet();
  std::optional<Reference> rewritten = transform.apply_external(book_id, sheet, node.as_external_ref_cell());
  if (!rewritten.has_value()) {
    return MakeRefError(arena);
  }
  // Sheet rename hook: when the transform supplies a new sheet name we
  // intern it into the arena so the rebuilt node owns its bytes.
  std::optional<std::string_view> new_sheet = transform.transform_external_sheet(book_id, sheet);
  const Reference& orig_cell = node.as_external_ref_cell();
  const Reference& nr = *rewritten;
  const bool cell_unchanged = orig_cell.sheet.data() == nr.sheet.data() && orig_cell.sheet.size() == nr.sheet.size() &&
                              orig_cell.sheet_quoted == nr.sheet_quoted && orig_cell.col == nr.col &&
                              orig_cell.row == nr.row && orig_cell.col_abs == nr.col_abs &&
                              orig_cell.row_abs == nr.row_abs && orig_cell.is_full_col == nr.is_full_col &&
                              orig_cell.is_full_row == nr.is_full_row;
  if (!new_sheet.has_value() && cell_unchanged) {
    return &node;
  }
  const std::string_view final_sheet = new_sheet.has_value() ? *new_sheet : sheet;
  return make_external_ref(arena, book_id, final_sheet, nr);
}

const AstNode* TransformUnary(const AstNode& node, Arena& arena, const RefTransform& transform) {
  const AstNode* operand = TransformNode(node.as_unary_operand(), arena, transform);
  if (operand == nullptr) {
    return nullptr;
  }
  if (operand == &node.as_unary_operand()) {
    return &node;
  }
  return make_unary_op(arena, node.as_unary_op(), const_cast<AstNode*>(operand));
}

const AstNode* TransformBinary(const AstNode& node, Arena& arena, const RefTransform& transform) {
  const AstNode* lhs = TransformNode(node.as_binary_lhs(), arena, transform);
  if (lhs == nullptr) {
    return nullptr;
  }
  const AstNode* rhs = TransformNode(node.as_binary_rhs(), arena, transform);
  if (rhs == nullptr) {
    return nullptr;
  }
  if (lhs == &node.as_binary_lhs() && rhs == &node.as_binary_rhs()) {
    return &node;
  }
  return make_binary_op(arena, node.as_binary_op(), const_cast<AstNode*>(lhs), const_cast<AstNode*>(rhs));
}

const AstNode* TransformRange(const AstNode& node, Arena& arena, const RefTransform& transform) {
  // Excel parses `Sheet1!A1:A10` with the sheet qualifier on the lhs
  // only — the rhs ends up as a sheet-less Ref but evaluates as if it
  // shared the lhs sheet. Transforms that key on `ref.sheet` would miss
  // the rhs without help, so synthesise the implicit sheet on the rhs
  // before walking it. A pre-transform clone keeps the input AST
  // immutable; downstream code is allowed to assume the original tree
  // is untouched.
  const AstNode& lhs_node = node.as_range_lhs();
  const AstNode& rhs_node = node.as_range_rhs();
  const AstNode* lhs = TransformNode(lhs_node, arena, transform);
  if (lhs == nullptr) {
    return nullptr;
  }
  const AstNode* rhs = nullptr;
  if (lhs_node.kind() == NodeKind::Ref && rhs_node.kind() == NodeKind::Ref && !lhs_node.as_ref().sheet.empty() &&
      rhs_node.as_ref().sheet.empty()) {
    Reference inherited = rhs_node.as_ref();
    inherited.sheet = lhs_node.as_ref().sheet;
    inherited.sheet_quoted = lhs_node.as_ref().sheet_quoted;
    AstNode* synthetic = make_ref(arena, inherited);
    if (synthetic == nullptr) {
      return nullptr;
    }
    rhs = TransformNode(*synthetic, arena, transform);
    if (rhs == nullptr) {
      return nullptr;
    }
    // Strip the synthesised sheet back off when both endpoints
    // ultimately wear the same sheet, so the formatted output keeps
    // Excel's canonical `Sheet1!A1:A10` shape (sheet on lhs only). The
    // post-transform lhs is what counts here — sheet rename rewrites
    // the sheet field, and the rhs would otherwise re-emit it
    // redundantly.
    if (rhs->kind() == NodeKind::Ref && lhs->kind() == NodeKind::Ref) {
      const Reference& lhs_ref_after = lhs->as_ref();
      Reference cleaned = rhs->as_ref();
      if (cleaned.sheet == lhs_ref_after.sheet) {
        cleaned.sheet = {};
        cleaned.sheet_quoted = false;
        AstNode* stripped = make_ref(arena, cleaned);
        if (stripped == nullptr) {
          return nullptr;
        }
        rhs = stripped;
      }
    }
  } else {
    rhs = TransformNode(rhs_node, arena, transform);
    if (rhs == nullptr) {
      return nullptr;
    }
  }
  // A range endpoint that collapsed to `#REF!` poisons the whole range
  // expression — this matches Excel's behaviour for `#REF!:#REF!` shapes.
  if (lhs->kind() == NodeKind::ErrorLiteral && lhs->as_error_literal() == ErrorCode::Ref) {
    return MakeRefError(arena);
  }
  if (rhs->kind() == NodeKind::ErrorLiteral && rhs->as_error_literal() == ErrorCode::Ref) {
    return MakeRefError(arena);
  }
  if (lhs == &lhs_node && rhs == &rhs_node) {
    return &node;
  }
  return make_range_op(arena, const_cast<AstNode*>(lhs), const_cast<AstNode*>(rhs));
}

const AstNode* TransformIntersect(const AstNode& node, Arena& arena, const RefTransform& transform) {
  const AstNode* lhs = TransformNode(node.as_intersect_lhs(), arena, transform);
  if (lhs == nullptr) {
    return nullptr;
  }
  const AstNode* rhs = TransformNode(node.as_intersect_rhs(), arena, transform);
  if (rhs == nullptr) {
    return nullptr;
  }
  if (lhs == &node.as_intersect_lhs() && rhs == &node.as_intersect_rhs()) {
    return &node;
  }
  return make_intersect_op(arena, const_cast<AstNode*>(lhs), const_cast<AstNode*>(rhs));
}

const AstNode* TransformImplicitIntersection(const AstNode& node, Arena& arena, const RefTransform& transform) {
  const AstNode* operand = TransformNode(node.as_implicit_intersection_operand(), arena, transform);
  if (operand == nullptr) {
    return nullptr;
  }
  if (operand == &node.as_implicit_intersection_operand()) {
    return &node;
  }
  return make_implicit_intersection(arena, const_cast<AstNode*>(operand));
}

const AstNode* TransformUnion(const AstNode& node, Arena& arena, const RefTransform& transform) {
  const std::uint32_t n = node.as_union_arity();
  std::vector<const AstNode*> kids;
  kids.reserve(n);
  bool changed = false;
  for (std::uint32_t i = 0; i < n; ++i) {
    const AstNode& child = node.as_union_child(i);
    const AstNode* updated = TransformNode(child, arena, transform);
    if (updated == nullptr) {
      return nullptr;
    }
    if (updated != &child) {
      changed = true;
    }
    kids.push_back(updated);
  }
  if (!changed) {
    return &node;
  }
  return make_union_op(arena, kids.data(), n);
}

const AstNode* TransformCall(const AstNode& node, Arena& arena, const RefTransform& transform) {
  const std::uint32_t n = node.as_call_arity();
  std::vector<const AstNode*> args;
  args.reserve(n);
  bool changed = false;
  for (std::uint32_t i = 0; i < n; ++i) {
    const AstNode& child = node.as_call_arg(i);
    const AstNode* updated = TransformNode(child, arena, transform);
    if (updated == nullptr) {
      return nullptr;
    }
    if (updated != &child) {
      changed = true;
    }
    args.push_back(updated);
  }
  if (!changed) {
    return &node;
  }
  return make_call(arena, node.as_call_name(), n == 0 ? nullptr : args.data(), n);
}

const AstNode* TransformArrayLiteral(const AstNode& node, Arena& arena, const RefTransform& transform) {
  const std::uint32_t rows = node.as_array_rows();
  const std::uint32_t cols = node.as_array_cols();
  const std::uint32_t total = rows * cols;
  std::vector<const AstNode*> elems;
  elems.reserve(total);
  bool changed = false;
  for (std::uint32_t r = 0; r < rows; ++r) {
    for (std::uint32_t c = 0; c < cols; ++c) {
      const AstNode& child = node.as_array_element(r, c);
      const AstNode* updated = TransformNode(child, arena, transform);
      if (updated == nullptr) {
        return nullptr;
      }
      if (updated != &child) {
        changed = true;
      }
      elems.push_back(updated);
    }
  }
  if (!changed) {
    return &node;
  }
  return make_array_literal(arena, rows, cols, elems.data());
}

const AstNode* TransformLambda(const AstNode& node, Arena& arena, const RefTransform& transform) {
  const AstNode* body = TransformNode(node.as_lambda_body(), arena, transform);
  if (body == nullptr) {
    return nullptr;
  }
  if (body == &node.as_lambda_body()) {
    return &node;
  }
  // Re-collect parameter names from the original; the lambda factory copies
  // them into the destination arena.
  const std::uint32_t pn = node.as_lambda_param_count();
  std::vector<std::string_view> params;
  params.reserve(pn);
  for (std::uint32_t i = 0; i < pn; ++i) {
    params.push_back(node.as_lambda_param(i));
  }
  return make_lambda(arena, pn == 0 ? nullptr : params.data(), pn, node.as_lambda_optional_count(),
                     const_cast<AstNode*>(body));
}

const AstNode* TransformLet(const AstNode& node, Arena& arena, const RefTransform& transform) {
  const std::uint32_t n = node.as_let_binding_count();
  std::vector<std::string_view> names;
  names.reserve(n);
  std::vector<const AstNode*> exprs;
  exprs.reserve(n);
  bool changed = false;
  for (std::uint32_t i = 0; i < n; ++i) {
    const AstNode& expr = node.as_let_binding_expr(i);
    const AstNode* updated = TransformNode(expr, arena, transform);
    if (updated == nullptr) {
      return nullptr;
    }
    if (updated != &expr) {
      changed = true;
    }
    names.push_back(node.as_let_binding_name(i));
    exprs.push_back(updated);
  }
  const AstNode* body = TransformNode(node.as_let_body(), arena, transform);
  if (body == nullptr) {
    return nullptr;
  }
  if (!changed && body == &node.as_let_body()) {
    return &node;
  }
  return make_let_binding(arena, names.data(), exprs.data(), n, const_cast<AstNode*>(body));
}

const AstNode* TransformLambdaCall(const AstNode& node, Arena& arena, const RefTransform& transform) {
  const AstNode* callee = TransformNode(node.as_lambda_call_callee(), arena, transform);
  if (callee == nullptr) {
    return nullptr;
  }
  const std::uint32_t n = node.as_lambda_call_arity();
  std::vector<const AstNode*> args;
  args.reserve(n);
  bool changed = (callee != &node.as_lambda_call_callee());
  for (std::uint32_t i = 0; i < n; ++i) {
    const AstNode& child = node.as_lambda_call_arg(i);
    const AstNode* updated = TransformNode(child, arena, transform);
    if (updated == nullptr) {
      return nullptr;
    }
    if (updated != &child) {
      changed = true;
    }
    args.push_back(updated);
  }
  if (!changed) {
    return &node;
  }
  return make_lambda_call(arena, const_cast<AstNode*>(callee), n == 0 ? nullptr : args.data(), n);
}

const AstNode* TransformNode(const AstNode& node, Arena& arena, const RefTransform& transform) {
  switch (node.kind()) {
    case NodeKind::Literal:
    case NodeKind::ErrorLiteral:
    case NodeKind::ErrorPlaceholder:
    case NodeKind::NameRef:
    case NodeKind::StructuredRef:
      return &node;
    case NodeKind::Ref:
      return TransformRef(node, arena, transform);
    case NodeKind::SpillRef:
      return TransformSpillRef(node, arena, transform);
    case NodeKind::ExternalRef:
      return TransformExternalRef(node, arena, transform);
    case NodeKind::UnaryOp:
      return TransformUnary(node, arena, transform);
    case NodeKind::BinaryOp:
      return TransformBinary(node, arena, transform);
    case NodeKind::RangeOp:
      return TransformRange(node, arena, transform);
    case NodeKind::UnionOp:
      return TransformUnion(node, arena, transform);
    case NodeKind::IntersectOp:
      return TransformIntersect(node, arena, transform);
    case NodeKind::ImplicitIntersection:
      return TransformImplicitIntersection(node, arena, transform);
    case NodeKind::Call:
      return TransformCall(node, arena, transform);
    case NodeKind::ArrayLiteral:
      return TransformArrayLiteral(node, arena, transform);
    case NodeKind::Lambda:
      return TransformLambda(node, arena, transform);
    case NodeKind::LetBinding:
      return TransformLet(node, arena, transform);
    case NodeKind::LambdaCall:
      return TransformLambdaCall(node, arena, transform);
  }
  return &node;
}

// ---------------------------------------------------------------------------
// RelativeShiftTransform
// ---------------------------------------------------------------------------

class RelativeShiftTransform final : public RefTransform {
 public:
  RelativeShiftTransform(std::int32_t row_delta, std::int32_t col_delta) noexcept
      : row_delta_(row_delta), col_delta_(col_delta) {}

  std::optional<Reference> apply(const Reference& ref) const override {
    Reference out = ref;
    // Whole-column / whole-row references shift only their non-absolute
    // axis; the absent axis stays meaningless and is left untouched.
    if (!ref.is_full_row && !ref.col_abs) {
      const std::int64_t shifted = static_cast<std::int64_t>(ref.col) + col_delta_;
      if (shifted < 0 || shifted >= static_cast<std::int64_t>(kMaxColumn)) {
        return std::nullopt;
      }
      out.col = static_cast<std::uint32_t>(shifted);
    }
    if (!ref.is_full_col && !ref.row_abs) {
      const std::int64_t shifted = static_cast<std::int64_t>(ref.row) + row_delta_;
      if (shifted < 0 || shifted >= static_cast<std::int64_t>(kMaxRow)) {
        return std::nullopt;
      }
      out.row = static_cast<std::uint32_t>(shifted);
    }
    return out;
  }

 private:
  std::int32_t row_delta_;
  std::int32_t col_delta_;
};

}  // namespace

const AstNode* shift_refs(const AstNode& root, Arena& arena, const RefTransform& transform) {
  return TransformNode(root, arena, transform);
}

const AstNode* shift_relative_refs(const AstNode& root, Arena& arena, std::int32_t row_delta, std::int32_t col_delta) {
  RelativeShiftTransform transform(row_delta, col_delta);
  return TransformNode(root, arena, transform);
}

}  // namespace parser
}  // namespace formulon
