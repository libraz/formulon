// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the relative-reference shifter. See ast_shift.h
// for the contract; this file is the recursive walk over every AST
// kind that may contain a Reference.

#include "parser/ast_shift.h"

#include <cstdint>
#include <optional>

#include "parser/ast.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace parser {

namespace {

// Apply `(row_delta, col_delta)` to a Reference. Returns nullopt when
// the shifted address is out of Sheet bounds along an axis that is
// meaningful for this reference shape.
std::optional<Reference> shift_reference(const Reference& ref, std::int32_t row_delta, std::int32_t col_delta) {
  Reference shifted = ref;

  // Whole-column (`A:A`): only the column axis is meaningful.
  if (ref.is_full_col) {
    if (!ref.col_abs) {
      const std::int64_t new_col = static_cast<std::int64_t>(ref.col) + col_delta;
      if (new_col < 0 || new_col >= static_cast<std::int64_t>(Sheet::kMaxCols)) {
        return std::nullopt;
      }
      shifted.col = static_cast<std::uint32_t>(new_col);
    }
    return shifted;
  }

  // Whole-row (`1:1`): only the row axis is meaningful.
  if (ref.is_full_row) {
    if (!ref.row_abs) {
      const std::int64_t new_row = static_cast<std::int64_t>(ref.row) + row_delta;
      if (new_row < 0 || new_row >= static_cast<std::int64_t>(Sheet::kMaxRows)) {
        return std::nullopt;
      }
      shifted.row = static_cast<std::uint32_t>(new_row);
    }
    return shifted;
  }

  // Regular A1 cell reference.
  if (!ref.col_abs) {
    const std::int64_t new_col = static_cast<std::int64_t>(ref.col) + col_delta;
    if (new_col < 0 || new_col >= static_cast<std::int64_t>(Sheet::kMaxCols)) {
      return std::nullopt;
    }
    shifted.col = static_cast<std::uint32_t>(new_col);
  }
  if (!ref.row_abs) {
    const std::int64_t new_row = static_cast<std::int64_t>(ref.row) + row_delta;
    if (new_row < 0 || new_row >= static_cast<std::int64_t>(Sheet::kMaxRows)) {
      return std::nullopt;
    }
    shifted.row = static_cast<std::uint32_t>(new_row);
  }
  return shifted;
}

AstNode* shift_node(const AstNode& node, Arena& arena, std::int32_t row_delta, std::int32_t col_delta);

// Convenience: shift a child and propagate nullptr (arena exhaustion)
// up to the caller.
AstNode* shift_child(const AstNode& child, Arena& arena, std::int32_t row_delta, std::int32_t col_delta) {
  return shift_node(child, arena, row_delta, col_delta);
}

AstNode* shift_node(const AstNode& node, Arena& arena, std::int32_t row_delta, std::int32_t col_delta) {
  switch (node.kind()) {
    case NodeKind::Literal:
      return make_literal(arena, node.as_literal());

    case NodeKind::Ref: {
      const auto shifted = shift_reference(node.as_ref(), row_delta, col_delta);
      if (!shifted.has_value()) {
        return make_error_literal(arena, ErrorCode::Ref);
      }
      return make_ref(arena, *shifted);
    }

    case NodeKind::SpillRef: {
      const auto shifted = shift_reference(node.as_spill_ref(), row_delta, col_delta);
      if (!shifted.has_value()) {
        return make_error_literal(arena, ErrorCode::Ref);
      }
      return make_spill_ref(arena, *shifted);
    }

    case NodeKind::ExternalRef: {
      // External-workbook refs may also carry relative coordinates;
      // Excel shifts them by the same rule. Workbook id and sheet
      // string are preserved verbatim.
      const auto shifted = shift_reference(node.as_external_ref_cell(), row_delta, col_delta);
      if (!shifted.has_value()) {
        return make_error_literal(arena, ErrorCode::Ref);
      }
      return make_external_ref(arena, node.as_external_ref_book_id(), node.as_external_ref_sheet(), *shifted);
    }

    case NodeKind::StructuredRef:
      // Table references are addressed by name; relative shifting does
      // not apply.
      return make_structured_ref(arena, node.as_structured_ref_table(), node.as_structured_ref_column(),
                                 node.as_structured_ref_modifier());

    case NodeKind::NameRef:
      return make_name_ref(arena, node.as_name());

    case NodeKind::UnaryOp: {
      AstNode* operand = shift_child(node.as_unary_operand(), arena, row_delta, col_delta);
      if (operand == nullptr) {
        return nullptr;
      }
      return make_unary_op(arena, node.as_unary_op(), operand);
    }

    case NodeKind::BinaryOp: {
      AstNode* lhs = shift_child(node.as_binary_lhs(), arena, row_delta, col_delta);
      if (lhs == nullptr) {
        return nullptr;
      }
      AstNode* rhs = shift_child(node.as_binary_rhs(), arena, row_delta, col_delta);
      if (rhs == nullptr) {
        return nullptr;
      }
      return make_binary_op(arena, node.as_binary_op(), lhs, rhs);
    }

    case NodeKind::RangeOp: {
      AstNode* lhs = shift_child(node.as_range_lhs(), arena, row_delta, col_delta);
      if (lhs == nullptr) {
        return nullptr;
      }
      AstNode* rhs = shift_child(node.as_range_rhs(), arena, row_delta, col_delta);
      if (rhs == nullptr) {
        return nullptr;
      }
      return make_range_op(arena, lhs, rhs);
    }

    case NodeKind::UnionOp: {
      const std::uint32_t arity = node.as_union_arity();
      AstNode** children = arena.create_array<AstNode*>(arity);
      if (children == nullptr) {
        return nullptr;
      }
      for (std::uint32_t i = 0; i < arity; ++i) {
        children[i] = shift_child(node.as_union_child(i), arena, row_delta, col_delta);
        if (children[i] == nullptr) {
          return nullptr;
        }
      }
      return make_union_op(arena, const_cast<const AstNode* const*>(children), arity);
    }

    case NodeKind::IntersectOp: {
      AstNode* lhs = shift_child(node.as_intersect_lhs(), arena, row_delta, col_delta);
      if (lhs == nullptr) {
        return nullptr;
      }
      AstNode* rhs = shift_child(node.as_intersect_rhs(), arena, row_delta, col_delta);
      if (rhs == nullptr) {
        return nullptr;
      }
      return make_intersect_op(arena, lhs, rhs);
    }

    case NodeKind::ImplicitIntersection: {
      AstNode* operand = shift_child(node.as_implicit_intersection_operand(), arena, row_delta, col_delta);
      if (operand == nullptr) {
        return nullptr;
      }
      return make_implicit_intersection(arena, operand);
    }

    case NodeKind::Call: {
      const std::uint32_t arity = node.as_call_arity();
      AstNode** args = nullptr;
      if (arity > 0) {
        args = arena.create_array<AstNode*>(arity);
        if (args == nullptr) {
          return nullptr;
        }
        for (std::uint32_t i = 0; i < arity; ++i) {
          args[i] = shift_child(node.as_call_arg(i), arena, row_delta, col_delta);
          if (args[i] == nullptr) {
            return nullptr;
          }
        }
      }
      return make_call(arena, node.as_call_name(), const_cast<const AstNode* const*>(args), arity);
    }

    case NodeKind::ArrayLiteral: {
      const std::uint32_t rows = node.as_array_rows();
      const std::uint32_t cols = node.as_array_cols();
      const std::uint32_t total = rows * cols;
      AstNode** elems = arena.create_array<AstNode*>(total);
      if (elems == nullptr) {
        return nullptr;
      }
      for (std::uint32_t r = 0; r < rows; ++r) {
        for (std::uint32_t c = 0; c < cols; ++c) {
          elems[r * cols + c] = shift_child(node.as_array_element(r, c), arena, row_delta, col_delta);
          if (elems[r * cols + c] == nullptr) {
            return nullptr;
          }
        }
      }
      return make_array_literal(arena, rows, cols, const_cast<const AstNode* const*>(elems));
    }

    case NodeKind::Lambda: {
      const std::uint32_t param_count = node.as_lambda_param_count();
      std::string_view* params = nullptr;
      if (param_count > 0) {
        params = arena.create_array<std::string_view>(param_count);
        if (params == nullptr) {
          return nullptr;
        }
        for (std::uint32_t i = 0; i < param_count; ++i) {
          params[i] = node.as_lambda_param(i);
        }
      }
      AstNode* body = shift_child(node.as_lambda_body(), arena, row_delta, col_delta);
      if (body == nullptr) {
        return nullptr;
      }
      return make_lambda(arena, params, param_count, node.as_lambda_optional_count(), body);
    }

    case NodeKind::LetBinding: {
      const std::uint32_t binding_count = node.as_let_binding_count();
      std::string_view* names = arena.create_array<std::string_view>(binding_count);
      AstNode** exprs = arena.create_array<AstNode*>(binding_count);
      if (names == nullptr || exprs == nullptr) {
        return nullptr;
      }
      for (std::uint32_t i = 0; i < binding_count; ++i) {
        names[i] = node.as_let_binding_name(i);
        exprs[i] = shift_child(node.as_let_binding_expr(i), arena, row_delta, col_delta);
        if (exprs[i] == nullptr) {
          return nullptr;
        }
      }
      AstNode* body = shift_child(node.as_let_body(), arena, row_delta, col_delta);
      if (body == nullptr) {
        return nullptr;
      }
      return make_let_binding(arena, names, const_cast<const AstNode* const*>(exprs), binding_count, body);
    }

    case NodeKind::LambdaCall: {
      AstNode* callee = shift_child(node.as_lambda_call_callee(), arena, row_delta, col_delta);
      if (callee == nullptr) {
        return nullptr;
      }
      const std::uint32_t arity = node.as_lambda_call_arity();
      AstNode** args = nullptr;
      if (arity > 0) {
        args = arena.create_array<AstNode*>(arity);
        if (args == nullptr) {
          return nullptr;
        }
        for (std::uint32_t i = 0; i < arity; ++i) {
          args[i] = shift_child(node.as_lambda_call_arg(i), arena, row_delta, col_delta);
          if (args[i] == nullptr) {
            return nullptr;
          }
        }
      }
      return make_lambda_call(arena, callee, const_cast<const AstNode* const*>(args), arity);
    }

    case NodeKind::ErrorLiteral:
      return make_error_literal(arena, node.as_error_literal());

    case NodeKind::ErrorPlaceholder:
      return make_error_placeholder(arena);
  }
  return nullptr;
}

}  // namespace

const AstNode* shift_relative_refs(const AstNode& root, Arena& arena, std::int32_t row_delta, std::int32_t col_delta) {
  return shift_node(root, arena, row_delta, col_delta);
}

}  // namespace parser
}  // namespace formulon
