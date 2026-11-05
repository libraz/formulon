//
// Shared defined-name resolution. See `defined_name_resolve.h`.

#include "eval/defined_name_resolve.h"

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "defined_name.h"
#include "eval/builtin_names.h"
#include "eval/declared_rect.h"
#include "eval/eval_context.h"
#include "eval/external_ref.h"
#include "eval/lazy_impls.h"        // eval_node, eval_node_shaped
#include "eval/name_env_resolve.h"  // is_range_shaped_ast
#include "eval/shape_ops_lazy.h"    // eval_node_as_array
#include "external_book.h"
#include "external_link.h"
#include "parser/ast.h"
#include "sheet.h"
#include "utils/expected.h"
#include "utils/formula_text.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {

const DefinedName* find_defined_name(const Workbook& workbook, std::uint16_t current_sheet_id,
                                     std::string_view name) noexcept {
  const auto& names = workbook.defined_names();
  const DefinedName* workbook_match = nullptr;
  for (const auto& entry : names) {
    if (!strings::case_insensitive_eq(entry.name, name)) {
      continue;
    }
    // Compare in the wider type: `local_sheet_id` comes from the file and is
    // not bounded by `sheet_count()`, so narrowing it to the 16-bit sheet id
    // first would let a crafted scope (0x10000) match sheet 0.
    if (entry.local_sheet_id >= 0 && entry.local_sheet_id == static_cast<std::int32_t>(current_sheet_id)) {
      // Sheet-scoped match for the current sheet wins immediately.
      return &entry;
    }
    if (entry.local_sheet_id < 0 && workbook_match == nullptr) {
      // Latch the first workbook-scoped match but keep scanning for a
      // sheet-scoped one that should take priority.
      workbook_match = &entry;
    }
  }
  return workbook_match;
}

namespace {

// Resolves the 0-based index of `ctx.current_sheet()` within its workbook.
// Returns false (writing nothing) when either binding is absent or the sheet
// is not owned by the workbook (defensive; the two should always agree).
bool current_sheet_index(const EvalContext& ctx, std::uint16_t* out) noexcept {
  const Workbook* wb = ctx.workbook();
  const Sheet* current = ctx.current_sheet();
  if (wb == nullptr || current == nullptr) {
    return false;
  }
  for (std::size_t i = 0; i < wb->sheet_count(); ++i) {
    if (&wb->sheet(i) == current) {
      *out = static_cast<std::uint16_t>(i);
      return true;
    }
  }
  return false;
}

}  // namespace

const DefinedName* find_defined_name(const EvalContext& ctx, std::string_view name) noexcept {
  const Workbook* wb = ctx.workbook();
  std::uint16_t sheet_id = 0;
  if (wb == nullptr || wb->defined_names().empty() || !current_sheet_index(ctx, &sheet_id)) {
    return nullptr;
  }
  if (ctx.name_scope_sheet() >= 0) {
    sheet_id = static_cast<std::uint16_t>(ctx.name_scope_sheet());
  }
  return find_defined_name(*wb, sheet_id, name);
}

const DefinedName* find_self_book_defined_name(const Workbook& workbook, std::string_view name) noexcept {
  const DefinedName* local_match = nullptr;
  for (const auto& entry : workbook.defined_names()) {
    if (!strings::case_insensitive_eq(entry.name, name)) {
      continue;
    }
    if (entry.local_sheet_id < 0) {
      return &entry;
    }
    if (static_cast<std::size_t>(entry.local_sheet_id) < workbook.sheet_count() &&
        (local_match == nullptr || entry.local_sheet_id < local_match->local_sheet_id)) {
      local_match = &entry;
    }
  }
  return local_match;
}

const DefinedName* find_sheet_defined_name(const Workbook& workbook, std::string_view sheet,
                                           std::string_view name) noexcept {
  const std::size_t sheet_id = workbook.sheet_index_by_name(sheet);
  if (sheet_id >= workbook.sheet_count()) {
    return nullptr;
  }
  return find_defined_name(workbook, static_cast<std::uint16_t>(sheet_id), name);
}

const parser::AstNode* prepare_defined_name_body(const DefinedName* def, Arena& arena, const EvalContext& ctx,
                                                 DefinedNameFrame* frame, EvalContext* out_ctx, ErrorCode* out_err) {
  if (def == nullptr) {
    *out_err = ErrorCode::Name;
    return nullptr;
  }
  // Cycle guard: a name already being expanded on the active resolution chain
  // is a circular reference. Surface `#REF!` to match the cell-cycle policy in
  // `EvalContext::resolve_ref` rather than recursing until the stack blows.
  for (const DefinedNameFrame* f = ctx.defined_name_stack(); f != nullptr; f = f->prev) {
    if (f->definition == def) {
      *out_err = ErrorCode::Ref;
      return nullptr;
    }
  }
  const std::string_view src = strip_formula_prefix(def->formula);
  parser::AstNode* root = src.empty() ? nullptr : parse_formula_entry(src, arena);
  if (root == nullptr) {
    *out_err = ErrorCode::Name;
    return nullptr;
  }
  // Evaluate the body as a top-level formula: clear the using formula's
  // lexical scope (a defined name never sees LET / LAMBDA bindings) and push
  // this definition onto the cycle chain.
  *frame = DefinedNameFrame{def, ctx.defined_name_stack()};
  *out_ctx = ctx.with_name_env(nullptr).with_defined_name_frame(frame);
  // A sheet-local name's body resolves its own unqualified names in the
  // owning sheet's scope, whichever sheet uses it; a workbook name's body
  // keeps the scope it is used from.
  if (def->local_sheet_id >= 0 && static_cast<std::size_t>(def->local_sheet_id) < ctx.workbook()->sheet_count()) {
    *out_ctx = out_ctx->with_name_scope_sheet(def->local_sheet_id);
  }
  return root;
}

namespace {

// Evaluates the body of the already-located definition `def` in `ctx`.
Value evaluate_defined_name(const DefinedName* def, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx) {
  DefinedNameFrame frame;
  EvalContext def_ctx = ctx;
  ErrorCode err = ErrorCode::Name;
  const parser::AstNode* root = prepare_defined_name_body(def, arena, ctx, &frame, &def_ctx, &err);
  if (root == nullptr) {
    return Value::error(err);
  }
  // A range-shaped body (e.g. `Sheet1!$A$1:$A$5`) must surface as a
  // `Value::Array` so range-aware consumers (`SUM`, `COUNT`, `VLOOKUP`, ...)
  // and the spill committer pick up its full shape instead of collapsing it to
  // the scalar `eval_node` would produce via implicit intersection. This mirrors
  // how `INDIRECT` returns a range as an array. Scalar bodies (including a
  // single-cell `Ref`) keep the scalar `eval_node` path.
  if (is_range_shaped_ast(*root)) {
    return eval_node_as_array(*root, arena, registry, def_ctx);
  }
  return eval_node(*root, arena, registry, def_ctx);
}

// `evaluate_defined_name` keeping a whole column / row at its declared size. A
// reference or expression body reads through `eval_node_shaped`; the other
// range-shaped bodies (reference-returning calls, array literals, spill
// references) and a bounded rectangle keep the `Value` reading.
Shaped evaluate_defined_name_shaped(const DefinedName* def, Arena& arena, const FunctionRegistry& registry,
                                    const EvalContext& ctx) {
  DefinedNameFrame frame;
  EvalContext def_ctx = ctx;
  ErrorCode err = ErrorCode::Name;
  const parser::AstNode* root = prepare_defined_name_body(def, arena, ctx, &frame, &def_ctx, &err);
  if (root == nullptr) {
    return Shaped{Value::error(err), nullptr};
  }
  bool shaped = false;
  switch (root->kind()) {
    case parser::NodeKind::Ref:
    case parser::NodeKind::ExternalRef:
    case parser::NodeKind::BinaryOp:
    case parser::NodeKind::UnaryOp:
      shaped = true;
      break;
    case parser::NodeKind::RangeOp: {
      // A bounded rectangle stays an array even when it is one cell.
      parser::Reference lhs{};
      parser::Reference rhs{};
      if (declared_rect_endpoint_pair(*root, &lhs, &rhs)) {
        const Expected<DeclaredRect, ErrorCode> rect = declared_rect(lhs, rhs);
        shaped = rect && (rect.value().rows() == Sheet::kMaxRows || rect.value().cols() == Sheet::kMaxCols);
      }
      break;
    }
    default:
      break;
  }
  if (shaped) {
    return eval_node_shaped(*root, arena, registry, def_ctx);
  }
  if (is_range_shaped_ast(*root)) {
    return Shaped{eval_node_as_array(*root, arena, registry, def_ctx), nullptr};
  }
  return Shaped{eval_node(*root, arena, registry, def_ctx), nullptr};
}

}  // namespace

namespace {

// The outcome of locating a name: the definition to evaluate, or (`resolved`)
// the final value when no definition applies.
struct NameLookup {
  const DefinedName* def = nullptr;
  bool resolved = false;
  Value value = Value::blank();
};

NameLookup unresolved_name(ErrorCode code) {
  NameLookup lookup;
  lookup.resolved = true;
  lookup.value = Value::error(code);
  return lookup;
}

NameLookup lookup_name(std::string_view name, const EvalContext& ctx) {
  if (ctx.workbook() == nullptr) {
    return unresolved_name(ErrorCode::Name);
  }
  NameLookup lookup;
  lookup.def = find_defined_name(ctx, name);
  return lookup;
}

NameLookup lookup_self_book_name(std::string_view name, const EvalContext& ctx) {
  const Workbook* wb = ctx.workbook();
  if (wb == nullptr) {
    return unresolved_name(ErrorCode::Name);
  }
  NameLookup lookup;
  lookup.def = find_self_book_defined_name(*wb, name);
  return lookup;
}

NameLookup lookup_sheet_name(std::string_view sheet, std::string_view name, Arena& arena, const EvalContext& ctx) {
  const Workbook* wb = ctx.workbook();
  if (wb == nullptr) {
    return unresolved_name(ErrorCode::Name);
  }
  if (wb->sheet_index_by_name(sheet) >= wb->sheet_count()) {
    // `Src2!Name` with no sheet `Src2` is the book-scope name of a linked
    // book written without an extension, when such a link exists; with no
    // such link the name is unknown.
    if (const ExternalLinkRecord* link = wb->link_for_sheet_qualifier(sheet); link != nullptr) {
      NameLookup lookup;
      lookup.resolved = true;
      lookup.value = resolve_external_book_name(link->book, ExternalBook::kNoSheet, name, arena);
      return lookup;
    }
    return unresolved_name(ErrorCode::Name);
  }
  NameLookup lookup;
  lookup.def = find_sheet_defined_name(*wb, sheet, name);
  return lookup;
}

Value evaluate_lookup(const NameLookup& lookup, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  return lookup.resolved ? lookup.value : evaluate_defined_name(lookup.def, arena, registry, ctx);
}

Shaped evaluate_lookup_shaped(const NameLookup& lookup, Arena& arena, const FunctionRegistry& registry,
                              const EvalContext& ctx) {
  return lookup.resolved ? Shaped{lookup.value, nullptr}
                         : evaluate_defined_name_shaped(lookup.def, arena, registry, ctx);
}

}  // namespace

Value resolve_defined_name(std::string_view name, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx) {
  return evaluate_lookup(lookup_name(name, ctx), arena, registry, ctx);
}

Value resolve_self_book_defined_name(std::string_view name, Arena& arena, const FunctionRegistry& registry,
                                     const EvalContext& ctx) {
  return evaluate_lookup(lookup_self_book_name(name, ctx), arena, registry, ctx);
}

Value resolve_sheet_defined_name(std::string_view sheet, std::string_view name, Arena& arena,
                                 const FunctionRegistry& registry, const EvalContext& ctx) {
  return evaluate_lookup(lookup_sheet_name(sheet, name, arena, ctx), arena, registry, ctx);
}

Shaped resolve_defined_name_shaped(std::string_view name, Arena& arena, const FunctionRegistry& registry,
                                   const EvalContext& ctx) {
  return evaluate_lookup_shaped(lookup_name(name, ctx), arena, registry, ctx);
}

Shaped resolve_self_book_defined_name_shaped(std::string_view name, Arena& arena, const FunctionRegistry& registry,
                                             const EvalContext& ctx) {
  return evaluate_lookup_shaped(lookup_self_book_name(name, ctx), arena, registry, ctx);
}

Shaped resolve_sheet_defined_name_shaped(std::string_view sheet, std::string_view name, Arena& arena,
                                         const FunctionRegistry& registry, const EvalContext& ctx) {
  return evaluate_lookup_shaped(lookup_sheet_name(sheet, name, arena, ctx), arena, registry, ctx);
}

}  // namespace eval
}  // namespace formulon
