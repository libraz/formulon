// Conservative AST result-shape analysis used by partial recalculation.

#include "eval/spill_potential.h"

#include <array>
#include <cstddef>
#include <string_view>
#include <vector>

#include "eval/function_registry.h"
#include "eval/tree_walker_lazy_table.h"

namespace formulon::eval {
namespace {

using parser::AstNode;
using parser::NodeKind;

bool ascii_equal_ci(std::string_view lhs, std::string_view rhs) noexcept {
  if (lhs.size() != rhs.size()) {
    return false;
  }
  for (std::size_t i = 0; i < lhs.size(); ++i) {
    const auto upper = [](char c) noexcept { return (c >= 'a' && c <= 'z') ? static_cast<char>(c - ('a' - 'A')) : c; };
    if (upper(lhs[i]) != upper(rhs[i])) {
      return false;
    }
  }
  return true;
}

template <std::size_t N>
bool in_names(std::string_view name, const std::array<std::string_view, N>& names) noexcept {
  for (std::string_view candidate : names) {
    if (ascii_equal_ci(name, candidate)) {
      return true;
    }
  }
  return false;
}

// Functions whose result is an array regardless of whether their arguments
// are scalar. This includes eager intrinsic arrays and the lazy dynamic-array
// family. The list is intentionally an over-approximation: candidate
// admission is later constrained by committed footprint geometry.
bool is_intrinsic_array(std::string_view name) noexcept {
  static constexpr std::array<std::string_view, 31> kNames = {
      "SEQUENCE",   "RANDARRAY",  "FILTER",    "SORT",     "SORTBY",    "UNIQUE",   "TAKE",     "DROP",
      "EXPAND",     "HSTACK",     "VSTACK",    "TOCOL",    "TOROW",     "WRAPCOLS", "WRAPROWS", "TRANSPOSE",
      "CHOOSECOLS", "CHOOSEROWS", "MAKEARRAY", "MAP",      "REDUCE",    "SCAN",     "BYROW",    "BYCOL",
      "FREQUENCY",  "MUNIT",      "MMULT",     "MINVERSE", "TEXTSPLIT", "GROUPBY",  "PIVOTBY"};
  return in_names(name, kNames);
}

SpillPotential combine_potential(SpillPotential lhs, SpillPotential rhs) noexcept {
  if (lhs == SpillPotential::kMaySpill || rhs == SpillPotential::kMaySpill) {
    return SpillPotential::kMaySpill;
  }
  if (lhs == SpillPotential::kNeedsRegistry || rhs == SpillPotential::kNeedsRegistry) {
    return SpillPotential::kNeedsRegistry;
  }
  return SpillPotential::kNever;
}

struct LetShape {
  std::string_view name;
  SpillPotential potential;
};

SpillPotential lookup_let_shape(std::string_view name, const std::vector<LetShape>* env) noexcept {
  if (env == nullptr) {
    return SpillPotential::kMaySpill;
  }
  for (auto it = env->rbegin(); it != env->rend(); ++it) {
    if (ascii_equal_ci(name, it->name)) {
      return it->potential;
    }
  }
  return SpillPotential::kMaySpill;
}

SpillPotential spill_potential_impl(const AstNode& root, const FunctionRegistry* registry,
                                    std::vector<LetShape>* let_env) noexcept {
  switch (root.kind()) {
    case NodeKind::Literal:
    case NodeKind::ErrorLiteral:
    case NodeKind::ErrorPlaceholder:
    case NodeKind::Ref:
    case NodeKind::ExternalRef:
      return SpillPotential::kNever;
    case NodeKind::Ref3D:
      // A 3-D single-cell tail still expands to one value per sheet when
      // the endpoint span contains multiple sheets; without workbook order
      // in this syntax-only pass, retain array potential conservatively.
      return SpillPotential::kMaySpill;
    case NodeKind::SpillRef:
    case NodeKind::StructuredRef:
    case NodeKind::RangeOp:
    case NodeKind::UnionOp:
    case NodeKind::ArrayLiteral:
    case NodeKind::Lambda:
    case NodeKind::LambdaCall:
      // A structured `Table[@col]` is the single-row implicit-intersection
      // form and cannot spill; other structured/name/lambda forms remain
      // conservative because their expansion is workbook-dependent.
      if (root.kind() == NodeKind::StructuredRef &&
          root.as_structured_ref_modifier() == parser::StructuredRefModifier::At) {
        return SpillPotential::kNever;
      }
      if (root.kind() == NodeKind::Lambda) {
        // A lambda can be invoked in a context that changes its bound
        // values' shape; retain array potential even when the current body
        // is syntactically scalar.
        return SpillPotential::kMaySpill;
      }
      if (root.kind() == NodeKind::LambdaCall) {
        // User-defined lambda result shape is registry/name-environment
        // dependent; conservatively retain it as a possible producer.
        return SpillPotential::kMaySpill;
      }
      return SpillPotential::kMaySpill;
    case NodeKind::ImplicitIntersection:
      // `@` is an explicit scalarization boundary, even when its child is an
      // intrinsic array or a range.
      return SpillPotential::kNever;
    case NodeKind::UnaryOp:
      return spill_potential_impl(root.as_unary_operand(), registry, let_env);
    case NodeKind::BinaryOp:
      return combine_potential(spill_potential_impl(root.as_binary_lhs(), registry, let_env),
                               spill_potential_impl(root.as_binary_rhs(), registry, let_env));
    case NodeKind::IntersectOp:
      return combine_potential(spill_potential_impl(root.as_intersect_lhs(), registry, let_env),
                               spill_potential_impl(root.as_intersect_rhs(), registry, let_env));
    case NodeKind::NameRef:
      return lookup_let_shape(root.as_name(), let_env);
    case NodeKind::LetBinding: {
      std::vector<LetShape> local_env = let_env == nullptr ? std::vector<LetShape>{} : *let_env;
      for (std::uint32_t i = 0; i < root.as_let_binding_count(); ++i) {
        const SpillPotential binding = spill_potential_impl(root.as_let_binding_expr(i), registry, &local_env);
        local_env.push_back({root.as_let_binding_name(i), binding});
      }
      return spill_potential_impl(root.as_let_body(), registry, &local_env);
    }
    case NodeKind::Call: {
      const std::string_view name = root.as_call_name();
      switch (find_lazy_result_shape(name)) {
        case LazyResultShape::kArray:
          return SpillPotential::kMaySpill;
        case LazyResultShape::kReduce:
        case LazyResultShape::kScalar:
          return SpillPotential::kNever;
        case LazyResultShape::kBroadcast: {
          SpillPotential result = SpillPotential::kNever;
          for (std::uint32_t i = 0; i < root.as_call_arity(); ++i) {
            result = combine_potential(result, spill_potential_impl(root.as_call_arg(i), registry, let_env));
          }
          return result;
        }
        case LazyResultShape::kNotLazy:
          break;
      }
      if (registry != nullptr) {
        if (const FunctionDef* def = registry->lookup(name); def != nullptr) {
          switch (def->result_shape) {
            case FunctionDef::ResultShape::kArray:
              return SpillPotential::kMaySpill;
            case FunctionDef::ResultShape::kReduce:
            case FunctionDef::ResultShape::kScalar:
              return SpillPotential::kNever;
            case FunctionDef::ResultShape::kBroadcast: {
              SpillPotential result = SpillPotential::kNever;
              for (std::uint32_t i = 0; i < root.as_call_arity(); ++i) {
                result = combine_potential(result, spill_potential_impl(root.as_call_arg(i), registry, let_env));
              }
              return result;
            }
            case FunctionDef::ResultShape::kAuto:
              return SpillPotential::kMaySpill;
          }
        }
        if (is_intrinsic_array(name)) {
          return SpillPotential::kMaySpill;
        }
        // Unknown/custom calls are array-capable by contract. A missing
        // registry entry is still conservative because a host can supply it
        // to the partial/full evaluation call.
        return SpillPotential::kMaySpill;
      }
      // Without a registry every eager call is registry-sensitive. This is
      // the key distinction from an unknown call: the candidate index keeps
      // it, then partial recalculation resolves the actual runtime shape.
      return SpillPotential::kNeedsRegistry;
    }
  }
  return SpillPotential::kMaySpill;
}

}  // namespace

SpillPotential spill_potential(const parser::AstNode& root) noexcept {
  return spill_potential_impl(root, nullptr, nullptr);
}

SpillPotential spill_potential(const parser::AstNode& root, const FunctionRegistry& registry) noexcept {
  return spill_potential_impl(root, &registry, nullptr);
}

bool may_produce_spill(const parser::AstNode& root) noexcept {
  return spill_potential(root, default_registry()) != SpillPotential::kNever;
}

bool may_produce_spill(const parser::AstNode& root, const FunctionRegistry& registry) noexcept {
  return spill_potential(root, registry) != SpillPotential::kNever;
}

}  // namespace formulon::eval
