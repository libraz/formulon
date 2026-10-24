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
#include "utils/a1_ref.h"
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

std::optional<std::string_view> RefTransform::apply_name_sheet(std::string_view sheet) const {
  return sheet;
}

std::optional<std::string_view> RefTransform::apply_structured_ref(std::string_view table,
                                                                   std::string_view /*column*/) const {
  return table;
}

namespace {

// Builds a `#REF!` error literal node. Returns nullptr on arena failure.
AstNode* MakeRefError(Arena& arena) {
  return make_error_literal(arena, ErrorCode::Ref);
}

// Forward declaration: the recursive worker. Returns the rewritten node;
// when nothing changed it returns `&node` directly so the caller can skip
// allocation.
const AstNode* TransformNode(const AstNode& node, Arena& arena, const RefTransform& transform);

// Transforms a child list once, allocating scratch only after the first
// changed child. At that point the untouched prefix is copied from the
// original node and each later transformed child is appended exactly once.
// This keeps identity walks allocation-free without losing preceding siblings
// when a change appears in the middle of a variadic node.
template <typename ChildAt>
bool TransformChildList(std::uint32_t count, Arena& arena, const RefTransform& transform, ChildAt child_at,
                        std::vector<const AstNode*>* out, bool* changed) {
  if (*changed) {
    out->reserve(count);
  }
  for (std::uint32_t i = 0; i < count; ++i) {
    const AstNode& child = child_at(i);
    const AstNode* updated = TransformNode(child, arena, transform);
    if (updated == nullptr) {
      return false;
    }
    if (!*changed) {
      if (updated == &child) {
        continue;
      }
      *changed = true;
      out->reserve(count);
      for (std::uint32_t prefix = 0; prefix < i; ++prefix) {
        out->push_back(&child_at(prefix));
      }
    }
    out->push_back(updated);
  }
  return true;
}

bool SameReference(const Reference& lhs, const Reference& rhs) {
  return lhs.sheet.data() == rhs.sheet.data() && lhs.sheet.size() == rhs.sheet.size() &&
         lhs.sheet_quoted == rhs.sheet_quoted && lhs.col == rhs.col && lhs.row == rhs.row &&
         lhs.col_abs == rhs.col_abs && lhs.row_abs == rhs.row_abs && lhs.is_full_col == rhs.is_full_col &&
         lhs.is_full_row == rhs.is_full_row;
}

const AstNode* TransformRef(const AstNode& node, Arena& arena, const RefTransform& transform) {
  std::optional<Reference> rewritten = transform.apply(node.as_ref());
  if (!rewritten.has_value()) {
    return MakeRefError(arena);
  }
  // Identity short-circuit: avoid an allocation when the transform was a
  // no-op for this reference.
  const Reference& orig = node.as_ref();
  const Reference& nr = *rewritten;
  if (SameReference(orig, nr)) {
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
  if (SameReference(orig, nr)) {
    return &node;
  }
  return make_spill_ref(arena, nr);
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

const AstNode* TransformStructuredRef(const AstNode& node, Arena& arena, const RefTransform& transform) {
  const std::string_view table = node.as_structured_ref_table();
  const std::optional<std::string_view> mapped = transform.apply_structured_ref(table, node.as_structured_ref_column());
  if (!mapped.has_value()) {
    return MakeRefError(arena);
  }
  if (*mapped == table) {
    return &node;
  }
  return make_structured_ref(arena, *mapped, node.as_structured_ref_column(), node.as_structured_ref_modifier());
}

const AstNode* TransformNameRef(const AstNode& node, Arena& arena, const RefTransform& transform) {
  const std::string_view sheet = node.as_name_sheet();
  if (sheet.empty()) {
    return &node;
  }
  const std::optional<std::string_view> mapped = transform.apply_name_sheet(sheet);
  if (!mapped.has_value()) {
    return MakeRefError(arena);
  }
  if (*mapped == sheet) {
    return &node;
  }
  return make_sheet_name_ref(arena, *mapped, node.as_name(), local_sheet_needs_quoting_a1(*mapped));
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
  const std::uint32_t n = node.as_binary_arity();
  std::vector<const AstNode*> operands;
  bool changed = false;
  if (!TransformChildList(
          n, arena, transform, [&](std::uint32_t i) -> const AstNode& { return node.as_binary_operand(i); }, &operands,
          &changed)) {
    return nullptr;
  }
  if (!changed) {
    return &node;
  }
  return make_chain_with_operands(arena, node, operands.data());
}

bool IsRefError(const AstNode& node) {
  return node.kind() == NodeKind::ErrorLiteral && node.as_error_literal() == ErrorCode::Ref;
}

// The qualifier a bare reference operand of `range` evaluates under: the
// first sheet a written-out reference operand names. `Sheet2!A1:C3:E5` reads
// E5 on Sheet2 too, so a transform keyed on `ref.sheet` must see it there.
const Reference* ChainSheetOwner(const AstNode& range) noexcept {
  for (std::uint32_t i = 0; i < range.as_range_arity(); ++i) {
    const AstNode& operand = range.as_range_operand(i);
    if (operand.kind() == NodeKind::Ref && !operand.as_ref().sheet.empty()) {
      return &operand.as_ref();
    }
  }
  return nullptr;
}

// `ref` as handed to a transform: a bare operand borrows `owner`'s qualifier.
Reference OnSheetOf(const Reference& ref, const Reference* owner) noexcept {
  Reference out = ref;
  if (owner != nullptr && ref.sheet.empty()) {
    out.sheet = owner->sheet;
    out.sheet_quoted = owner->sheet_quoted;
  }
  return out;
}

// Drops a qualifier `OnSheetOf` lent to `orig` again. Transforms map a sheet
// by its name alone, so the lent qualifier came back as the owner's image
// and the operand keeps Excel's bare spelling (`Sheet1!A1:A10`).
Reference BareAgain(const Reference& orig, Reference rewritten) noexcept {
  if (orig.sheet.empty()) {
    rewritten.sheet = {};
    rewritten.sheet_quoted = false;
  }
  return rewritten;
}

// Rewrites a `:` chain token by token, as Excel tokenises it. An area token
// (`A1:A3`, `A:C`) moves as one interval through `apply_range`, so deleting
// row 1 shrinks `A1:A3` to `A1:A2` instead of poisoning it; a reference
// operand moves through `apply`; any other operand is walked as a subtree.
// Every reference operand is transformed on the sheet it evaluates on, and
// an area token whose halves land on one whole column or row becomes the
// single Ref the parser builds for that text (`A:A`, never `A:A:A:A`).
const AstNode* TransformRange(const AstNode& node, Arena& arena, const RefTransform& transform) {
  const std::uint32_t n = node.as_range_arity();
  const Reference* const chain_owner = ChainSheetOwner(node);
  std::vector<const AstNode*> operands;
  // Per output operand, the splice flag joining it to the previous one.
  std::vector<std::uint8_t> links;
  operands.reserve(n);
  links.reserve(n);
  bool changed = false;
  for (std::uint32_t i = 0; i < n;) {
    const std::uint32_t width = range_token_width(node, i);
    const AstNode& lhs_node = node.as_range_operand(i);
    i += width;
    if (lhs_node.kind() != NodeKind::Ref) {
      const AstNode* updated = TransformNode(lhs_node, arena, transform);
      if (updated == nullptr) {
        return nullptr;
      }
      // An endpoint that collapsed to `#REF!` poisons the whole range
      // expression -- this matches Excel's behaviour for `#REF!:#REF!` shapes.
      if (IsRefError(*updated)) {
        return MakeRefError(arena);
      }
      operands.push_back(updated);
      links.push_back(0U);
      changed = changed || updated != &lhs_node;
      continue;
    }
    const Reference& lhs_orig = lhs_node.as_ref();
    if (width == 1U) {
      const std::optional<Reference> rewritten = transform.apply(OnSheetOf(lhs_orig, chain_owner));
      if (!rewritten.has_value()) {
        return MakeRefError(arena);
      }
      const Reference out = BareAgain(lhs_orig, *rewritten);
      links.push_back(0U);
      if (SameReference(lhs_orig, out)) {
        operands.push_back(&lhs_node);
        continue;
      }
      AstNode* ref = make_ref(arena, out);
      if (ref == nullptr) {
        return nullptr;
      }
      ref->set_paren_depth(lhs_node.paren_depth());
      operands.push_back(ref);
      changed = true;
      continue;
    }
    // The second corner of an area token is always written bare and shares
    // the first corner's qualifier, or the chain's when the first is bare.
    const AstNode& rhs_node = node.as_range_operand(i - 1U);
    const Reference& rhs_orig = rhs_node.as_ref();
    const Reference* const token_owner = lhs_orig.sheet.empty() ? chain_owner : &lhs_orig;
    const std::optional<std::pair<Reference, Reference>> rewritten =
        transform.apply_range(OnSheetOf(lhs_orig, chain_owner), OnSheetOf(rhs_orig, token_owner));
    if (!rewritten.has_value()) {
      return MakeRefError(arena);
    }
    const Reference lhs_out = BareAgain(lhs_orig, rewritten->first);
    const Reference rhs_out = BareAgain(rhs_orig, rewritten->second);
    const bool spliced = node.as_range_spliced(i - 2U);
    if (SameReference(lhs_orig, lhs_out) && SameReference(rhs_orig, rhs_out)) {
      operands.push_back(&lhs_node);
      operands.push_back(&rhs_node);
      links.push_back(0U);
      links.push_back(spliced ? 1U : 0U);
      continue;
    }
    changed = true;
    AstNode* lhs = make_ref(arena, lhs_out);
    if (lhs == nullptr) {
      return nullptr;
    }
    operands.push_back(lhs);
    links.push_back(0U);
    if (spliced && whole_axis_pair_is_single(lhs_out, rhs_out)) {
      continue;
    }
    AstNode* rhs = make_ref(arena, rhs_out);
    if (rhs == nullptr) {
      return nullptr;
    }
    operands.push_back(rhs);
    links.push_back(spliced ? 1U : 0U);
  }
  if (!changed) {
    return &node;
  }
  if (operands.size() == 1U) {
    return operands.front();
  }
  return make_range_chain(arena, operands.data(), links.data() + 1, static_cast<std::uint32_t>(operands.size()));
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
  bool changed = false;
  if (!TransformChildList(
          n, arena, transform, [&](std::uint32_t i) -> const AstNode& { return node.as_union_child(i); }, &kids,
          &changed)) {
    return nullptr;
  }
  if (!changed) {
    return &node;
  }
  return make_union_op(arena, kids.data(), n);
}

const AstNode* TransformCall(const AstNode& node, Arena& arena, const RefTransform& transform) {
  const std::uint32_t n = node.as_call_arity();
  std::vector<const AstNode*> args;
  bool changed = false;
  if (!TransformChildList(
          n, arena, transform, [&](std::uint32_t i) -> const AstNode& { return node.as_call_arg(i); }, &args,
          &changed)) {
    return nullptr;
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
  bool changed = false;
  if (!TransformChildList(
          total, arena, transform,
          [&](std::uint32_t i) -> const AstNode& { return node.as_array_element(i / cols, i % cols); }, &elems,
          &changed)) {
    return nullptr;
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
  std::vector<const AstNode*> exprs;
  bool changed = false;
  if (!TransformChildList(
          n, arena, transform, [&](std::uint32_t i) -> const AstNode& { return node.as_let_binding_expr(i); }, &exprs,
          &changed)) {
    return nullptr;
  }
  const AstNode* body = TransformNode(node.as_let_body(), arena, transform);
  if (body == nullptr) {
    return nullptr;
  }
  if (!changed && body == &node.as_let_body()) {
    return &node;
  }
  std::vector<std::string_view> names;
  names.reserve(n);
  for (std::uint32_t i = 0; i < n; ++i) {
    names.push_back(node.as_let_binding_name(i));
  }
  if (!changed) {
    exprs.reserve(n);
    for (std::uint32_t i = 0; i < n; ++i) {
      exprs.push_back(&node.as_let_binding_expr(i));
    }
  }
  return make_let_binding(arena, names.data(), exprs.data(), n, const_cast<AstNode*>(body));
}

const AstNode* TransformLambdaCall(const AstNode& node, Arena& arena, const RefTransform& transform) {
  const AstNode* callee = TransformNode(node.as_lambda_call_callee(), arena, transform);
  if (callee == nullptr) {
    return nullptr;
  }
  // `Sheet2!Fn(3)` whose sheet is removed collapses whole to `#REF!`: an
  // error literal is no callee `(` can follow.
  if (callee->kind() == NodeKind::ErrorLiteral && node.as_lambda_call_callee().kind() == NodeKind::NameRef) {
    return callee;
  }
  const std::uint32_t n = node.as_lambda_call_arity();
  std::vector<const AstNode*> args;
  bool changed = (callee != &node.as_lambda_call_callee());
  if (!TransformChildList(
          n, arena, transform, [&](std::uint32_t i) -> const AstNode& { return node.as_lambda_call_arg(i); }, &args,
          &changed)) {
    return nullptr;
  }
  if (!changed) {
    return &node;
  }
  return make_lambda_call(arena, const_cast<AstNode*>(callee), n == 0 ? nullptr : args.data(), n);
}

const AstNode* TransformNodeBare(const AstNode& node, Arena& arena, const RefTransform& transform) {
  switch (node.kind()) {
    case NodeKind::Literal:
    case NodeKind::ErrorLiteral:
    case NodeKind::ErrorPlaceholder:
    // A cross-workbook reference addresses another file's grid. Inserting
    // or deleting rows here cannot move a cell there, and Excel leaves
    // such a reference untouched for exactly that reason.
    case NodeKind::ExternalRef:
      return &node;
    case NodeKind::StructuredRef:
      return TransformStructuredRef(node, arena, transform);
    case NodeKind::Ref:
      return TransformRef(node, arena, transform);
    case NodeKind::SpillRef:
      return TransformSpillRef(node, arena, transform);
    case NodeKind::Ref3D:
      return TransformRef3D(node, arena, transform);
    case NodeKind::NameRef:
      return TransformNameRef(node, arena, transform);
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

// A rebuilt node keeps the parentheses written around the one it replaces.
const AstNode* TransformNode(const AstNode& node, Arena& arena, const RefTransform& transform) {
  const AstNode* rebuilt = TransformNodeBare(node, arena, transform);
  if (rebuilt != nullptr && rebuilt != &node && node.paren_depth() != 0U) {
    // Every node other than `node` itself is fresh from this walk's arena.
    const_cast<AstNode*>(rebuilt)->set_paren_depth(node.paren_depth());
  }
  return rebuilt;
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
      if (shifted < 0 || shifted >= static_cast<std::int64_t>(a1::kMaxCols)) {
        return std::nullopt;
      }
      out.col = static_cast<std::uint32_t>(shifted);
    }
    if (!ref.is_full_col && !ref.row_abs) {
      const std::int64_t shifted = static_cast<std::int64_t>(ref.row) + row_delta_;
      if (shifted < 0 || shifted >= static_cast<std::int64_t>(a1::kMaxRows)) {
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
