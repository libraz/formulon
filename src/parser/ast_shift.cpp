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

std::optional<std::pair<Reference, Reference>> RefTransform::apply_range(const Reference& lhs,
                                                                         const Reference& rhs) const {
  const std::optional<Reference> rewritten_lhs = apply(lhs);
  if (!rewritten_lhs.has_value()) {
    return std::nullopt;
  }
  const std::optional<Reference> rewritten_rhs = apply(rhs);
  if (!rewritten_rhs.has_value()) {
    return std::nullopt;
  }
  return std::make_pair(*rewritten_lhs, *rewritten_rhs);
}

std::optional<RefTransform::Ref3DSheetSpan> RefTransform::apply_ref3d_span(std::string_view begin,
                                                                           std::string_view end) const {
  return Ref3DSheetSpan{begin, end};
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
  if (const AstNode* anchor = node.as_spill_ref_anchor_expr(); anchor != nullptr) {
    // A computed anchor carries its references inside the sub-expression,
    // so the rewrite belongs there; the operator itself has nothing to
    // shift.
    const AstNode* moved = TransformNode(*anchor, arena, transform);
    if (moved == nullptr) {
      return nullptr;
    }
    return moved == anchor ? &node : make_spill_ref_expr(arena, moved);
  }
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

bool SameReference(const Reference& lhs, const Reference& rhs) {
  return lhs.sheet.data() == rhs.sheet.data() && lhs.sheet.size() == rhs.sheet.size() &&
         lhs.sheet_quoted == rhs.sheet_quoted && lhs.col == rhs.col && lhs.row == rhs.row &&
         lhs.col_abs == rhs.col_abs && lhs.row_abs == rhs.row_abs && lhs.is_full_col == rhs.is_full_col &&
         lhs.is_full_row == rhs.is_full_row;
}

const AstNode* TransformRef3D(const AstNode& node, Arena& arena, const RefTransform& transform) {
  const Reference& first = node.as_ref3d_cell();
  const bool is_range = node.as_ref3d_is_range();
  const Reference& last = node.as_ref3d_cell_end();
  std::optional<Reference> rewritten_first;
  std::optional<Reference> rewritten_last;
  if (transform.preserves_ref3d_coordinates()) {
    // A 3-D reference applies one coordinate rectangle across a sheet span.
    // Excel leaves that shared tail unchanged for row/column structural
    // edits, even when the edited sheet is inside the span or owns the
    // formula. The transform can still rename the span endpoints below.
    rewritten_first = first;
    rewritten_last = last;
  } else if (is_range) {
    // A 3-D range tail is one rectangle even though its sheet span lives
    // outside the two Reference values. Give structural transforms the same
    // pair-level hook used by ordinary RangeOp nodes so deleting one tail
    // endpoint shrinks the surviving interval instead of poisoning it.
    const std::optional<std::pair<Reference, Reference>> rewritten_range = transform.apply_range(first, last);
    if (!rewritten_range.has_value()) {
      return MakeRefError(arena);
    }
    rewritten_first = rewritten_range->first;
    rewritten_last = rewritten_range->second;
  } else {
    rewritten_first = transform.apply(first);
    if (!rewritten_first.has_value()) {
      return MakeRefError(arena);
    }
    rewritten_last = last;
  }

  // Ref3D stores its workbook-local sheet span outside Reference, so the
  // span endpoints route through their own hook rather than through `apply`.
  const std::string_view begin = node.as_ref3d_sheet_begin();
  const std::string_view end = node.as_ref3d_sheet_end();
  const std::optional<RefTransform::Ref3DSheetSpan> rewritten_span = transform.apply_ref3d_span(begin, end);
  if (!rewritten_span.has_value()) {
    return MakeRefError(arena);
  }
  const std::string_view final_begin = rewritten_span->begin;
  const std::string_view final_end = rewritten_span->end;
  if (final_begin == begin && final_end == end && SameReference(first, *rewritten_first) &&
      (!is_range || SameReference(last, *rewritten_last))) {
    return &node;
  }
  if (is_range) {
    return make_ref3d_range(arena, final_begin, final_end, *rewritten_first, *rewritten_last);
  }
  return make_ref3d(arena, final_begin, final_end, *rewritten_first);
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
  const AstNode& lhs_node = node.as_range_lhs();
  const AstNode& rhs_node = node.as_range_rhs();

  // A normal A1 range is represented by two Ref nodes. Give the transform a
  // chance to rewrite the pair as one interval before falling back to the
  // recursive endpoint walk. This is important for structural deletions:
  // Excel shrinks `A1:A3` to `A1:A2` when row 1 is deleted instead of
  // poisoning the whole range because its first endpoint disappeared.
  //
  // Excel parses `Sheet1!A1:A10` with the sheet qualifier on the lhs only —
  // the rhs ends up as a sheet-less Ref but evaluates as if it shared the
  // lhs sheet. Transforms that key on `ref.sheet` would miss the rhs without
  // help, so the implicit sheet is synthesised onto the copy handed to the
  // transform and stripped again afterwards, which keeps both the input AST
  // immutable and the formatted output in Excel's canonical
  // `Sheet1!A1:A10` form.
  if (lhs_node.kind() == NodeKind::Ref && rhs_node.kind() == NodeKind::Ref) {
    const Reference& lhs_orig = lhs_node.as_ref();
    const Reference& rhs_orig = rhs_node.as_ref();
    const bool inherit_rhs_sheet = !lhs_orig.sheet.empty() && rhs_orig.sheet.empty();
    Reference rhs_for_transform = rhs_orig;
    if (inherit_rhs_sheet) {
      rhs_for_transform.sheet = lhs_orig.sheet;
      rhs_for_transform.sheet_quoted = lhs_orig.sheet_quoted;
    }

    const std::optional<std::pair<Reference, Reference>> rewritten = transform.apply_range(lhs_orig, rhs_for_transform);
    if (!rewritten.has_value()) {
      return MakeRefError(arena);
    }
    Reference rewritten_lhs = rewritten->first;
    Reference rewritten_rhs = rewritten->second;
    if (inherit_rhs_sheet && rewritten_rhs.sheet == rewritten_lhs.sheet) {
      rewritten_rhs.sheet = {};
      rewritten_rhs.sheet_quoted = false;
    }

    const bool lhs_unchanged = SameReference(lhs_orig, rewritten_lhs);
    const bool rhs_unchanged = SameReference(rhs_orig, rewritten_rhs);
    if (lhs_unchanged && rhs_unchanged) {
      return &node;
    }
    AstNode* lhs = make_ref(arena, rewritten_lhs);
    AstNode* rhs = make_ref(arena, rewritten_rhs);
    if (lhs == nullptr || rhs == nullptr) {
      return nullptr;
    }
    return make_range_op(arena, lhs, rhs);
  }

  const AstNode* lhs = TransformNode(lhs_node, arena, transform);
  if (lhs == nullptr) {
    return nullptr;
  }
  // Reaching here means at least one endpoint is not a plain `Ref`, so no
  // sheet qualifier can be inherited: the pair-level branch above returns on
  // every path and owns the `Ref`/`Ref` case exclusively.
  const AstNode* rhs = TransformNode(rhs_node, arena, transform);
  if (rhs == nullptr) {
    return nullptr;
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
    // A cross-workbook reference addresses another file's grid. Inserting
    // or deleting rows here cannot move a cell there, and Excel leaves
    // such a reference untouched for exactly that reason.
    case NodeKind::ExternalRef:
      return &node;
    case NodeKind::Ref:
      return TransformRef(node, arena, transform);
    case NodeKind::SpillRef:
      return TransformSpillRef(node, arena, transform);
    case NodeKind::Ref3D:
      return TransformRef3D(node, arena, transform);
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
  if (!ast_depth_within_limit(root, kMaxFormulaAstDepth)) {
    return MakeRefError(arena);
  }
  return TransformNode(root, arena, transform);
}

const AstNode* shift_relative_refs(const AstNode& root, Arena& arena, std::int32_t row_delta, std::int32_t col_delta) {
  if (!ast_depth_within_limit(root, kMaxFormulaAstDepth)) {
    return MakeRefError(arena);
  }
  RelativeShiftTransform transform(row_delta, col_delta);
  return TransformNode(root, arena, transform);
}

}  // namespace parser
}  // namespace formulon
