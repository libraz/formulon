//
// Implementation of the name / sheet-range collection pass declared in
// `io/xlsb/ptg_writer.h`: which `BrtName` and `BrtExternSheet` entries a
// formula's Ptg encoding will resolve through.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_set>
#include <utility>
#include <vector>

#include "io/future_functions.h"
#include "io/xlsb/func_id_table.h"
#include "io/xlsb/ptg_targets.h"
#include "io/xlsb/ptg_writer.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "utils/strings.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// Recursion helper for `collect_ptg_names`: adds `name` to `names` (and
/// marks it in `seen`) unless already present.
void AddName(std::string_view name, std::vector<std::string>& names, std::unordered_set<std::string>& seen) {
  std::string owned(name);
  if (seen.insert(owned).second) {
    names.push_back(std::move(owned));
  }
}

/// Packs an XTI into a single dedupe key. Sheet indices stay below 2^19,
/// so 20 bits each keep `kXtiNoSheet` distinct.
std::uint64_t PackRangeKey(std::uint32_t book, std::int32_t itab_first, std::int32_t itab_last) {
  constexpr std::uint64_t kMask = (std::uint64_t{1} << 20) - 1U;
  return (static_cast<std::uint64_t>(book) << 40) |
         ((static_cast<std::uint64_t>(static_cast<std::uint32_t>(itab_first)) & kMask) << 20) |
         (static_cast<std::uint64_t>(static_cast<std::uint32_t>(itab_last)) & kMask);
}

/// Appends the XTI `(book, itab_first, itab_last)` to `ranges` unless
/// already present in `seen`.
void AddXti(std::uint32_t book, std::int32_t itab_first, std::int32_t itab_last, SheetRangeTable& ranges,
            std::unordered_set<std::uint64_t>& seen) {
  if (seen.insert(PackRangeKey(book, itab_first, itab_last)).second) {
    ranges.xti.push_back(XtiEntry{book, itab_first, itab_last});
  }
}

/// Recursion helper for `collect_ptg_sheet_ranges`: when both sheet indices
/// resolved, appends their span in `book`.
void AddSheetRange(std::uint32_t book, std::int32_t itab_first, std::int32_t itab_last, SheetRangeTable& ranges,
                   std::unordered_set<std::uint64_t>& seen) {
  if (itab_first < 0 || itab_last < 0) {
    return;  // Unresolvable sheet name; the encode fails later with a precise error.
  }
  AddXti(book, itab_first, itab_last, ranges, seen);
}

/// The XTI a cross-workbook reference resolves through, plus any name its
/// link's cache lacks.
void AddExternalRef(const parser::AstNode& node, SheetRangeTable& ranges, std::unordered_set<std::uint64_t>& seen) {
  const std::uint32_t position = external_link_position(ranges, node);
  if (position == 0U) {
    return;  // The encode reports the unknown book.
  }
  XlsbLinkTables& link = ranges.links[position - 1U];
  const std::string_view name = node.kind() == parser::NodeKind::NameRef ? node.as_name() : node.as_external_ref_name();
  if (!name.empty()) {
    // Saved book-scope whatever its sheet: the part records no name scope.
    if (external_name_ilbl(link, name) == 0U) {
      link.names.emplace_back(name);
    }
    AddXti(position, kXtiNoSheet, kXtiNoSheet, ranges, seen);
    return;
  }
  const int first = external_sheet_index(link, node.as_external_ref_sheet());
  const std::string_view sheet_end = node.as_external_ref_sheet_end();
  const int last = sheet_end.empty() ? first : external_sheet_index(link, sheet_end);
  AddSheetRange(position, first, last, ranges, seen);
}

enum class NameCollectMode : std::uint8_t { kPtg, kScopeResolved, kSheetQualified };

/// True when `name` matches an in-scope LET / LAMBDA parameter.
bool InParamScope(const std::vector<std::string_view>& scope, std::string_view name) {
  // Case-insensitive: see `Encoder::emit_name_ref`.
  return std::any_of(scope.begin(), scope.end(),
                     [name](std::string_view param) { return strings::case_insensitive_eq(param, name); });
}

/// Recursive worker for `collect_ptg_names` carrying the LET / LAMBDA
/// parameter names currently in scope (innermost last). A `NameRef`
/// matching an in-scope parameter resolves at encode time to that
/// parameter's hidden `_xlpm.<name>` placeholder (see
/// `Encoder::emit_name_ref`), so it must not be registered as an ordinary
/// workbook defined name. A callee with no function id (`Fn(3)`: a named
/// LAMBDA, or a name no one defined) is a name too. `mode` selects the view:
/// every unqualified name a `PtgName` or self-book `PtgNameX` needs
/// (`kPtg`); only names resolved from the formula's own scope
/// (`kScopeResolved`); or only sheet-qualified names, each added as `sheet`
/// NUL `name` (`kSheetQualified`).
void CollectNamesScoped(const parser::AstNode& node, std::vector<std::string>& names,
                        std::unordered_set<std::string>& seen, std::vector<std::string_view>& scope,
                        NameCollectMode mode, const std::vector<const parser::AstNode*>* values) {
  // Hidden `_xlfn.*` / `_xlpm.*` records and unqualified names.
  auto add = [&](std::string_view name) {
    if (mode != NameCollectMode::kSheetQualified) {
      AddName(name, names, seen);
    }
  };
  switch (node.kind()) {
    case parser::NodeKind::NameRef: {
      const std::string_view name = node.as_name();
      if (!node.as_name_sheet().empty()) {
        // Resolves through a scoped key (a definition or its sheet's stub),
        // never a LET / LAMBDA parameter or a workbook placeholder.
        if (mode == NameCollectMode::kSheetQualified) {
          AddName(std::string(node.as_name_sheet()) + '\0' + std::string(name), names, seen);
        }
        return;
      }
      if (InParamScope(scope, name)) {
        return;  // LET / LAMBDA parameter: encoded via its _xlpm. placeholder.
      }
      if (values != nullptr && std::find(values->begin(), values->end(), &node) != values->end()) {
        add(function_value_storage_name(name));
        return;
      }
      add(name);
      return;
    }
    case parser::NodeKind::Call: {
      const std::string_view name = canonical_function_name(node.as_call_name());
      if (UsesHiddenNameRoute(name)) {
        add(xlsb_hidden_function_name(name));
      } else if (lookup_func_by_name(name) == nullptr && !InParamScope(scope, node.as_call_name())) {
        add(node.as_call_name());
      }
      const std::uint32_t arity = node.as_call_arity();
      for (std::uint32_t i = 0; i < arity; ++i) {
        CollectNamesScoped(node.as_call_arg(i), names, seen, scope, mode, values);
      }
      return;
    }
    case parser::NodeKind::UnaryOp:
      CollectNamesScoped(node.as_unary_operand(), names, seen, scope, mode, values);
      return;
    case parser::NodeKind::BinaryOp:
      CollectNamesScoped(node.as_binary_lhs(), names, seen, scope, mode, values);
      CollectNamesScoped(node.as_binary_rhs(), names, seen, scope, mode, values);
      return;
    case parser::NodeKind::RangeOp:
      CollectNamesScoped(node.as_range_lhs(), names, seen, scope, mode, values);
      CollectNamesScoped(node.as_range_rhs(), names, seen, scope, mode, values);
      return;
    case parser::NodeKind::IntersectOp:
      CollectNamesScoped(node.as_intersect_lhs(), names, seen, scope, mode, values);
      CollectNamesScoped(node.as_intersect_rhs(), names, seen, scope, mode, values);
      return;
    case parser::NodeKind::UnionOp: {
      const std::uint32_t arity = node.as_union_arity();
      for (std::uint32_t i = 0; i < arity; ++i) {
        CollectNamesScoped(node.as_union_child(i), names, seen, scope, mode, values);
      }
      return;
    }
    case parser::NodeKind::ImplicitIntersection:
      add(xlsb_hidden_function_name("SINGLE"));
      CollectNamesScoped(node.as_implicit_intersection_operand(), names, seen, scope, mode, values);
      return;
    case parser::NodeKind::ArrayLiteral: {
      const std::uint32_t rows = node.as_array_rows();
      const std::uint32_t cols = node.as_array_cols();
      for (std::uint32_t r = 0; r < rows; ++r) {
        for (std::uint32_t c = 0; c < cols; ++c) {
          CollectNamesScoped(node.as_array_element(r, c), names, seen, scope, mode, values);
        }
      }
      return;
    }
    case parser::NodeKind::SpillRef: {
      // Stored as a call to the hidden `_xlfn.ANCHORARRAY` name (see
      // `Encoder::emit_spill_ref`), so it needs the same BrtName
      // registration any other future-function callee gets.
      add(xlsb_hidden_function_name("ANCHORARRAY"));
      if (const parser::AstNode* anchor = node.as_spill_ref_anchor_expr(); anchor != nullptr) {
        CollectNamesScoped(*anchor, names, seen, scope, mode, values);
      }
      return;
    }
    case parser::NodeKind::ExternalRef:
      // `[0]!Rate` is never a LET / LAMBDA parameter and, like `Sheet2!Rate`,
      // not resolved from the formula's own scope.
      if (parser::is_self_book_name_ref(node) && mode == NameCollectMode::kPtg) {
        AddName(node.as_external_ref_name(), names, seen);
      }
      return;
    case parser::NodeKind::LambdaCall: {
      CollectNamesScoped(node.as_lambda_call_callee(), names, seen, scope, mode, values);
      const std::uint32_t arity = node.as_lambda_call_arity();
      for (std::uint32_t i = 0; i < arity; ++i) {
        CollectNamesScoped(node.as_lambda_call_arg(i), names, seen, scope, mode, values);
      }
      return;
    }
    case parser::NodeKind::LetBinding: {
      add("_xlfn.LET");
      const std::uint32_t n = node.as_let_binding_count();
      const std::size_t scope_base = scope.size();
      for (std::uint32_t i = 0; i < n; ++i) {
        add(std::string("_xlpm.") + std::string(node.as_let_binding_name(i)));
        // Excel LET binds sequentially: a value expression sees only the
        // earlier bindings, so collect it before pushing this parameter.
        CollectNamesScoped(node.as_let_binding_expr(i), names, seen, scope, mode, values);
        scope.push_back(node.as_let_binding_name(i));
      }
      CollectNamesScoped(node.as_let_body(), names, seen, scope, mode, values);
      scope.resize(scope_base);
      return;
    }
    case parser::NodeKind::Lambda: {
      add("_xlfn.LAMBDA");
      const std::uint32_t n = node.as_lambda_param_count();
      const std::size_t scope_base = scope.size();
      for (std::uint32_t i = 0; i < n; ++i) {
        add(std::string("_xlpm.") + std::string(node.as_lambda_param(i)));
        scope.push_back(node.as_lambda_param(i));
      }
      CollectNamesScoped(node.as_lambda_body(), names, seen, scope, mode, values);
      scope.resize(scope_base);
      return;
    }
    // Leaves, and forms the encoder does not lower (StructuredRef): nothing
    // to collect. A future writer bundle that lowers these would extend
    // this switch alongside the corresponding `emit_*` case.
    default:
      return;
  }
}

}  // namespace

std::string sheet_scoped_name_key(std::int32_t itab, std::string_view name) {
  std::string key = std::to_string(itab);
  key.push_back('!');
  key.append(strings::to_ascii_lower(name));
  return key;
}

void collect_ptg_names(const parser::AstNode& node, std::vector<std::string>& names,
                       std::unordered_set<std::string>& seen,
                       const std::vector<const parser::AstNode*>* function_values) {
  std::vector<std::string_view> scope;
  CollectNamesScoped(node, names, seen, scope, NameCollectMode::kPtg, function_values);
}

void collect_scope_resolved_names(const parser::AstNode& node, std::vector<std::string>& names,
                                  std::unordered_set<std::string>& seen) {
  std::vector<std::string_view> scope;
  CollectNamesScoped(node, names, seen, scope, NameCollectMode::kScopeResolved, nullptr);
}

void collect_sheet_qualified_names(const parser::AstNode& node,
                                   std::vector<std::pair<std::string, std::string>>& qualified) {
  std::vector<std::string> keys;
  std::unordered_set<std::string> seen;
  std::vector<std::string_view> scope;
  CollectNamesScoped(node, keys, seen, scope, NameCollectMode::kSheetQualified, nullptr);
  for (const std::string& key : keys) {
    const std::size_t split = key.find('\0');
    qualified.emplace_back(key.substr(0, split), key.substr(split + 1));
  }
}

std::uint32_t xti_sup_book(const SheetRangeTable& table, std::uint32_t book) {
  if (book == 0U) {
    return 0U;
  }
  return book - (xti_names_self(table) ? 0U : 1U);
}

bool xti_names_self(const SheetRangeTable& table) {
  return std::any_of(table.xti.begin(), table.xti.end(), [](const XtiEntry& entry) { return entry.book == 0U; });
}

std::uint32_t external_link_position(const SheetRangeTable& table, const parser::AstNode& node) {
  std::uint32_t index = 0U;
  if (node.kind() == parser::NodeKind::NameRef) {
    // An extensionless book's book-scope name, spelled as a sheet-qualified name.
    if (table.indexer.qualifier_index == nullptr || node.as_name_sheet().empty()) {
      return 0U;
    }
    index = table.indexer.qualifier_index(table.indexer.ctx, node.as_name_sheet());
  } else {
    if (table.indexer.index == nullptr) {
      return 0U;
    }
    index = table.indexer.index(table.indexer.ctx, node.as_external_ref_path(), node.as_external_ref_book());
  }
  for (std::size_t i = 0; index != 0U && i < table.links.size(); ++i) {
    if (table.links[i].index == index) {
      return static_cast<std::uint32_t>(i + 1U);
    }
  }
  return 0U;
}

int external_sheet_index(const XlsbLinkTables& link, std::string_view sheet) {
  for (std::size_t i = 0; i < link.sheet_names.size(); ++i) {
    if (strings::case_insensitive_eq(link.sheet_names[i], sheet)) {
      return static_cast<int>(i);
    }
  }
  return -1;
}

std::uint32_t external_name_ilbl(const XlsbLinkTables& link, std::string_view name) {
  for (std::size_t i = 0; i < link.names.size(); ++i) {
    if (strings::case_insensitive_eq(link.names[i], name)) {
      return static_cast<std::uint32_t>(i + 1U);
    }
  }
  return 0U;
}

void collect_ptg_sheet_ranges(const parser::AstNode& node, const std::vector<std::string>& sheet_names,
                              SheetRangeTable& ranges, std::unordered_set<std::uint64_t>& seen) {
  switch (node.kind()) {
    case parser::NodeKind::Ref: {
      const parser::Reference& r = node.as_ref();
      if (!r.sheet.empty()) {
        const int itab = resolve_ixti(sheet_names, r.sheet);
        AddSheetRange(0U, itab, itab, ranges, seen);
      }
      return;
    }
    case parser::NodeKind::Ref3D: {
      const int itab_begin = resolve_ixti(sheet_names, node.as_ref3d_sheet_begin());
      const int itab_end = resolve_ixti(sheet_names, node.as_ref3d_sheet_end());
      AddSheetRange(0U, itab_begin, itab_end, ranges, seen);
      return;
    }
    case parser::NodeKind::NameRef:
    case parser::NodeKind::ExternalRef: {
      if (node.kind() == parser::NodeKind::ExternalRef && !parser::is_self_book_name_ref(node)) {
        AddExternalRef(node, ranges, seen);
        return;
      }
      if (node.kind() == parser::NodeKind::NameRef && external_link_position(ranges, node) != 0U) {
        AddExternalRef(node, ranges, seen);
        return;
      }
      // `Sheet1!Rate` and `[0]!Rate` encode as `PtgNameX` through the
      // book-scope entry.
      if (node.kind() == parser::NodeKind::ExternalRef || !node.as_name_sheet().empty()) {
        AddXti(0U, kXtiNoSheet, kXtiNoSheet, ranges, seen);
      }
      return;
    }
    case parser::NodeKind::LambdaCall: {
      collect_ptg_sheet_ranges(node.as_lambda_call_callee(), sheet_names, ranges, seen);
      const std::uint32_t arity = node.as_lambda_call_arity();
      for (std::uint32_t i = 0; i < arity; ++i) {
        collect_ptg_sheet_ranges(node.as_lambda_call_arg(i), sheet_names, ranges, seen);
      }
      return;
    }
    case parser::NodeKind::Call: {
      const std::uint32_t arity = node.as_call_arity();
      for (std::uint32_t i = 0; i < arity; ++i) {
        collect_ptg_sheet_ranges(node.as_call_arg(i), sheet_names, ranges, seen);
      }
      return;
    }
    case parser::NodeKind::UnaryOp:
      collect_ptg_sheet_ranges(node.as_unary_operand(), sheet_names, ranges, seen);
      return;
    case parser::NodeKind::BinaryOp:
      collect_ptg_sheet_ranges(node.as_binary_lhs(), sheet_names, ranges, seen);
      collect_ptg_sheet_ranges(node.as_binary_rhs(), sheet_names, ranges, seen);
      return;
    case parser::NodeKind::RangeOp:
      collect_ptg_sheet_ranges(node.as_range_lhs(), sheet_names, ranges, seen);
      collect_ptg_sheet_ranges(node.as_range_rhs(), sheet_names, ranges, seen);
      return;
    case parser::NodeKind::IntersectOp:
      collect_ptg_sheet_ranges(node.as_intersect_lhs(), sheet_names, ranges, seen);
      collect_ptg_sheet_ranges(node.as_intersect_rhs(), sheet_names, ranges, seen);
      return;
    case parser::NodeKind::UnionOp: {
      const std::uint32_t arity = node.as_union_arity();
      for (std::uint32_t i = 0; i < arity; ++i) {
        collect_ptg_sheet_ranges(node.as_union_child(i), sheet_names, ranges, seen);
      }
      return;
    }
    case parser::NodeKind::ImplicitIntersection:
      collect_ptg_sheet_ranges(node.as_implicit_intersection_operand(), sheet_names, ranges, seen);
      return;
    case parser::NodeKind::ArrayLiteral: {
      const std::uint32_t rows = node.as_array_rows();
      const std::uint32_t cols = node.as_array_cols();
      for (std::uint32_t r = 0; r < rows; ++r) {
        for (std::uint32_t c = 0; c < cols; ++c) {
          collect_ptg_sheet_ranges(node.as_array_element(r, c), sheet_names, ranges, seen);
        }
      }
      return;
    }
    case parser::NodeKind::LetBinding: {
      const std::uint32_t n = node.as_let_binding_count();
      for (std::uint32_t i = 0; i < n; ++i) {
        collect_ptg_sheet_ranges(node.as_let_binding_expr(i), sheet_names, ranges, seen);
      }
      collect_ptg_sheet_ranges(node.as_let_body(), sheet_names, ranges, seen);
      return;
    }
    case parser::NodeKind::Lambda:
      collect_ptg_sheet_ranges(node.as_lambda_body(), sheet_names, ranges, seen);
      return;
    default:
      return;
  }
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
