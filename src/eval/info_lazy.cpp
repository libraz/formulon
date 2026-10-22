//
// Implementation of the context-aware information predicates. See
// `eval/info_lazy.h` for the semantics each function advertises.
//
// The family deliberately rides the lazy seam because every entry
// inspects information the eager path would erase:
//   * ISFORMULA needs the bound Sheet's `formula_text` for the
//     referenced cell — an evaluated reference is just a `Value`.
//   * ISREF branches on the AST node kind (and, for calls, on whether
//     the call is in the reference-returning set).
//   * SHEET / SHEETS consult `Workbook::sheet_by_name` and the bound
//     current-sheet pointer held on `EvalContext`.

#include "eval/info_lazy.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>

#include "eval/a1_parse.h"
#include "eval/array_alloc.h"
#include "eval/coerce.h"
#include "eval/datetime_lazy.h"
#include "eval/eval_context.h"
#include "eval/eval_profile_scope.h"
#include "eval/formula_localize.h"
#include "eval/function_registry.h"
#include "eval/implicit_intersection.h"
#include "eval/lazy_impls.h"
#include "eval/name_env.h"
#include "eval/name_env_resolve.h"
#include "eval/omitted_arg.h"
#include "eval/range_resolvers.h"
#include "excel_locale.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/strings.h"
#include "utils/text_ops.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Hangul has no CP932 mapping, so Windows ja-JP stores each such character as the single byte '?'.
bool is_hangul(std::uint32_t cp) noexcept {
  return (cp >= 0x1100u && cp <= 0x11FFu) || (cp >= 0x3130u && cp <= 0x318Fu) || (cp >= 0xA960u && cp <= 0xA97Fu) ||
         (cp >= 0xAC00u && cp <= 0xD7FFu);
}

// Returns true when `node` is a static reference-shaped AST. Does NOT
// include reference-returning calls (`is_reference_call_name`); those need
// evaluation to confirm they produce a ref at runtime.
//
// A bare `NameRef` is intentionally NOT treated as ref-shaped here: a LET
// binding such as `=LET(x, 5, ISREF(x))` carries a scalar value, not a
// reference. Callers must first resolve a `NameRef` through
// `resolve_name_ast` so a range-bound name (`=LET(r, A1, ISREF(r))`)
// surfaces its underlying `Ref` / `RangeOp` AST, which this function then
// recognises. An unresolved (scalar-bound) name stays a `NameRef` and
// falls through to FALSE.
bool is_static_reference_shape(const parser::AstNode& node) noexcept {
  switch (node.kind()) {
    case parser::NodeKind::Ref:
    case parser::NodeKind::RangeOp:
    case parser::NodeKind::StructuredRef:
      return true;
    case parser::NodeKind::ExternalRef:
      // A cell or range on one sheet of another workbook. Neither the
      // `[Book]!Name` form nor a 3-D span is, as for a local `Ref3D`.
      return node.as_external_ref_name().empty() && node.as_external_ref_sheet_end().empty();
    default:
      return false;
  }
}

// N's coercion of one value: a number as it is, TRUE / FALSE as 1 / 0, an
// error as it is, and anything else (text, blank) as 0. Text is never parsed.
Value number_of(const Value& v) {
  switch (v.kind()) {
    case ValueKind::Number:
    case ValueKind::Error:
      return v;
    case ValueKind::Bool:
      return Value::number(v.as_boolean() ? 1.0 : 0.0);
    case ValueKind::Text:
    case ValueKind::Blank:
      return Value::number(0.0);
    case ValueKind::Array:
    case ValueKind::Ref:
    case ValueKind::Lambda:
      return Value::error(ErrorCode::Value);
  }
  return Value::error(ErrorCode::Value);
}

// T's coercion of one value: text and errors as they are, anything else as "".
Value text_of(const Value& v) {
  if (v.is_text() || v.is_error()) {
    return v;
  }
  return Value::text({});
}

// Shared body of N and T. A reference argument reads its first cell rather
// than each of its cells; any other argument is evaluated, and an array
// result is coerced element by element.
Value eval_first_cell_coercion(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                               const EvalContext& ctx, Value (*coerce)(const Value&)) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  const parser::AstNode& target = resolve_name_ast(call.as_call_arg(0), ctx.name_env());
  const parser::NodeKind kind = target.kind();
  if (kind == parser::NodeKind::Ref || kind == parser::NodeKind::RangeOp || kind == parser::NodeKind::IntersectOp ||
      kind == parser::NodeKind::NameRef ||
      (kind == parser::NodeKind::Call && is_reference_call_name(target.as_call_name()))) {
    std::string_view sheet;
    std::uint32_t top = 0;
    std::uint32_t left = 0;
    std::uint32_t bottom = 0;
    std::uint32_t right = 0;
    ErrorCode err = ErrorCode::Value;
    if (resolve_reference_rect(target, arena, registry, ctx, &sheet, &top, &left, &bottom, &right, &err)) {
      parser::Reference first{};
      first.sheet = sheet;
      first.row = top;
      first.col = left;
      return coerce(ctx.resolve_ref(first, arena, registry));
    }
  }
  const Value v = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (!v.is_array()) {
    return coerce(v);
  }
  const std::uint32_t rows = v.as_array_rows();
  const std::uint32_t cols = v.as_array_cols();
  Value* buf = nullptr;
  ArrayValue* out = allocate_array_value(rows, cols, arena, buf, kMaxDerivedArrayCells);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  const Value* cells = v.as_array_cells();
  for (std::size_t i = 0; i < static_cast<std::size_t>(rows) * cols; ++i) {
    buf[i] = coerce(cells[i]);
  }
  return Value::array(out);
}

// Locates `sheet` inside `workbook` by pointer identity and returns
// the 1-based sheet number on success, 0 on failure. Pointer identity
// is sufficient because `EvalContext` holds a raw pointer into the
// workbook's own vector; if the caller ever constructs a context with
// a sheet owned by a *different* workbook this returns 0 and SHEET
// surfaces `#N/A`, which is a safe degradation.
std::size_t find_sheet_index(const Workbook& workbook, const Sheet& sheet) noexcept {
  const std::size_t count = workbook.sheet_count();
  for (std::size_t i = 0; i < count; ++i) {
    if (&workbook.sheet(i) == &sheet) {
      return i + 1;
    }
  }
  return 0;
}

struct SingleReference {
  const Sheet* sheet = nullptr;
  std::uint32_t row = 0;
  std::uint32_t col = 0;
};

Expected<SingleReference, ErrorCode> resolve_single_reference(const parser::AstNode& node, Arena& arena,
                                                              const FunctionRegistry& registry,
                                                              const EvalContext& ctx) {
  std::string_view sheet_name;
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t bottom = 0;
  std::uint32_t right = 0;
  if (node.kind() == parser::NodeKind::Ref) {
    const parser::Reference& ref = node.as_ref();
    if (ref.is_full_col || ref.is_full_row) {
      return ErrorCode::Value;
    }
    sheet_name = ref.sheet;
    top = bottom = ref.row;
    left = right = ref.col;
  } else if (node.kind() == parser::NodeKind::Call) {
    bool is_range = false;
    ErrorCode err = ErrorCode::Value;
    if (!resolve_reference_call(node, arena, registry, ctx, &sheet_name, &top, &left, &bottom, &right, &is_range,
                                &err)) {
      return err;
    }
  } else {
    return ErrorCode::Value;
  }
  if (top != bottom || left != right) {
    return ErrorCode::Value;
  }
  const Sheet* sheet = ctx.sheet_for_qualifier(sheet_name);
  if (sheet == nullptr) {
    return ctx.current_sheet() == nullptr ? ErrorCode::Name : ErrorCode::Ref;
  }
  return SingleReference{sheet, top, left};
}

// The cell holding the formula shown at `(row, col)`: the cell itself, or
// for a cell inside a legacy array formula block, the block's anchor.
const Cell* formula_cell_at(const Sheet& sheet, std::uint32_t row, std::uint32_t col) {
  const Cell* cell = sheet.cell_at(row, col);
  if (cell != nullptr && !cell->formula_text.empty()) {
    return cell;
  }
  const std::optional<SpillFootprint> block = sheet.array_block_covering(row, col);
  return block.has_value() ? sheet.cell_at(block->anchor_row, block->anchor_col) : cell;
}

}  // namespace

Value eval_isformula_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  // Look through a LET / LAMBDA `NameRef` to its bound source AST so a
  // reference-bound name (`=LET(c, A1, ISFORMULA(c))`) inspects the cell it
  // resolves to. A scalar-bound name (`=LET(x, 5, ISFORMULA(x))`) carries no
  // AST: `resolve_name_ast` returns the `NameRef` unchanged and the
  // non-reference `#VALUE!` tail handles it.
  const parser::AstNode& arg = resolve_name_ast(call.as_call_arg(0), ctx.name_env());

  // Static references and calls that preserve a reference shape share the
  // range resolver used by ROWS/COLUMNS and CELL. This retains OFFSET and
  // CHOOSE coordinates instead of flattening their target to a Value.
  if (arg.kind() == parser::NodeKind::Ref ||
      (arg.kind() == parser::NodeKind::Call && is_reference_call_name(arg.as_call_name()))) {
    auto target = resolve_single_reference(arg, arena, registry, ctx);
    if (target) {
      const Sheet& sheet = *target.value().sheet;
      const Cell* cell = formula_cell_at(sheet, target.value().row, target.value().col);
      return Value::boolean(((cell != nullptr) && !cell->formula_text.empty()) ||
                            sheet.data_table_covering(target.value().row, target.value().col) != nullptr);
    }
    // A reference-shaped call may instead choose a scalar (for example
    // CHOOSE(1, 42)); that is not a formula argument. Real evaluation errors
    // still propagate through the existing value path below.
    if (arg.kind() == parser::NodeKind::Call) {
      const Value v = eval_node(arg, arena, registry, ctx);
      if (v.is_error()) {
        return v;
      }
      return Value::boolean(false);
    }
    return Value::error(target.error());
  }

  // Excel: `ISFORMULA(1)` and `ISFORMULA("A1")` both surface `#VALUE!`.
  // RangeOp would need array-spill to answer cell-by-cell; not in the MVP.
  return Value::error(ErrorCode::Value);
}

Value eval_formulatext_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  const parser::AstNode& arg = resolve_name_ast(call.as_call_arg(0), ctx.name_env());
  auto target = resolve_single_reference(arg, arena, registry, ctx);
  if (!target) {
    return Value::error(target.error());
  }
  const Cell* cell = formula_cell_at(*target.value().sheet, target.value().row, target.value().col);
  if (cell == nullptr || cell->formula_text.empty()) {
    // A data table's result cell shows its `TABLE()` formula, braced like a legacy array.
    if (const DataTable* table = target.value().sheet->data_table_covering(target.value().row, target.value().col);
        table != nullptr) {
      const std::string shown =
          "{" + localize_formula_text(data_table_formula_text(*table), current_eval_profile()) + "}";
      const std::string_view interned = arena.intern(shown);
      return interned.empty() ? Value::error(ErrorCode::Value) : Value::text(interned);
    }
    // No formula in the target cell (blank, value-only, or missing
    // slot). Excel 365 answers `#N/A`.
    return Value::error(ErrorCode::NA);
  }
  // Copy the formula source into the evaluation arena, stripping the
  // xlsx-only `_xlfn.` / `_xlfn._xlws.` storage prefixes so the returned
  // text matches what the user typed (and what Excel 365 displays).
  //
  // Only strip when the prefix stands at a token-start position — i.e.
  // preceded by nothing, by an ASCII non-identifier byte, or by a quote
  // boundary — and not while we are inside a `"..."` string literal.
  // This avoids corrupting user text that happens to contain `_xlfn.`.
  const std::string_view src(cell->formula_text);
  std::string stripped;
  stripped.reserve(src.size());
  constexpr std::string_view kXlws = "_xlfn._xlws.";
  constexpr std::string_view kXlfn = "_xlfn.";
  bool in_string = false;
  for (std::size_t i = 0; i < src.size();) {
    const char c = src[i];
    if (c == '"') {
      in_string = !in_string;
      stripped.push_back(c);
      ++i;
      continue;
    }
    if (!in_string) {
      const char prev = stripped.empty() ? '\0' : stripped.back();
      const bool at_token_start = prev == '\0' || !((prev >= 'A' && prev <= 'Z') || (prev >= 'a' && prev <= 'z') ||
                                                    (prev >= '0' && prev <= '9') || prev == '_' || prev == '.');
      if (at_token_start) {
        if (i + kXlws.size() <= src.size() && strings::case_insensitive_eq(src.substr(i, kXlws.size()), kXlws)) {
          i += kXlws.size();
          continue;
        }
        if (i + kXlfn.size() <= src.size() && strings::case_insensitive_eq(src.substr(i, kXlfn.size()), kXlfn)) {
          i += kXlfn.size();
          continue;
        }
      }
    }
    stripped.push_back(c);
    ++i;
  }
  std::string shown = localize_formula_text(stripped, current_eval_profile());
  // A legacy array formula shows in braces.
  if (cell->array_block_rows != 0U) {
    shown.insert(shown.begin(), '{');
    shown.push_back('}');
  }
  char* buf = static_cast<char*>(arena.allocate(shown.size(), alignof(char)));
  if (buf == nullptr) {
    return Value::error(ErrorCode::Value);
  }
  std::copy(shown.begin(), shown.end(), buf);
  return Value::text(std::string_view(buf, shown.size()));
}

Value eval_n_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx) {
  return eval_first_cell_coercion(call, arena, registry, ctx, &number_of);
}

Value eval_t_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx) {
  return eval_first_cell_coercion(call, arena, registry, ctx, &text_of);
}

Value eval_isref_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  // Look through a LET / LAMBDA `NameRef` to its bound source AST. A
  // range-shaped binding (`=LET(r, A1, ISREF(r))`) resolves to the
  // underlying `Ref` / `RangeOp` and counts as a reference; a scalar-bound
  // name (`=LET(x, 5, ISREF(x))`) carries no AST, so `resolve_name_ast`
  // returns the `NameRef` unchanged and every branch below evaluates to
  // FALSE.
  const parser::AstNode& arg = resolve_name_ast(call.as_call_arg(0), ctx.name_env());
  if (is_static_reference_shape(arg)) {
    return Value::boolean(true);
  }
  if (arg.kind() == parser::NodeKind::Call && is_reference_call_name(arg.as_call_name())) {
    // Reference-returning calls must actually produce a non-error
    // result to count as "ref". E.g. `ISREF(INDIRECT("not a ref"))`
    // is FALSE because INDIRECT surfaces `#REF!`.
    std::string_view sheet;
    std::uint32_t top = 0;
    std::uint32_t left = 0;
    std::uint32_t bottom = 0;
    std::uint32_t right = 0;
    bool is_range = false;
    ErrorCode err = ErrorCode::Value;
    return Value::boolean(resolve_reference_call(arg, arena, registry, ctx.without_dynamic_read_callback(), &sheet,
                                                 &top, &left, &bottom, &right, &is_range, &err));
  }
  return Value::boolean(false);
}

Value eval_type_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  const Value v = eval_node(call.as_call_arg(0), arena, registry, ctx);
  switch (v.kind()) {
    case ValueKind::Number:
    case ValueKind::Blank:
      return Value::number(1.0);
    case ValueKind::Text:
      return Value::number(2.0);
    case ValueKind::Bool:
      return Value::number(4.0);
    case ValueKind::Error:
      return Value::number(16.0);
    case ValueKind::Array:
      return Value::number(64.0);
    case ValueKind::Lambda:
      return Value::number(128.0);
    case ValueKind::Ref:
      break;
  }
  return Value::error(ErrorCode::Value);
}

Value eval_sheet_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity > 1U) {
    return Value::error(ErrorCode::Value);
  }
  const Sheet* current = ctx.current_sheet();
  const Workbook* wb = ctx.workbook();

  // 0-arity: number of the currently-bound sheet.
  if (arity == 0U) {
    if (current == nullptr) {
      return Value::error(ErrorCode::Value);
    }
    if (wb == nullptr) {
      // Bound only to a Sheet, no workbook. Convention: single sheet = 1.
      return Value::number(1.0);
    }
    const std::size_t idx = find_sheet_index(*wb, *current);
    if (idx == 0) {
      return Value::error(ErrorCode::Value);
    }
    return Value::number(static_cast<double>(idx));
  }

  const parser::AstNode& arg = resolve_name_ast(call.as_call_arg(0), ctx.name_env());

  // Reference AST: the answer is determined by the sheet qualifier
  // without needing to evaluate the reference.
  std::string_view sheet_qualifier;
  bool have_ref_shape = false;
  if (arg.kind() == parser::NodeKind::Ref) {
    sheet_qualifier = arg.as_ref().sheet;
    have_ref_shape = true;
  } else if (arg.kind() == parser::NodeKind::RangeOp) {
    const parser::AstNode& lhs = arg.as_range_operand(0);
    if (lhs.kind() == parser::NodeKind::Ref) {
      sheet_qualifier = lhs.as_ref().sheet;
      have_ref_shape = true;
    }
  }
  if (have_ref_shape) {
    if (sheet_qualifier.empty()) {
      // Unqualified ref — same sheet as the host context.
      if (current == nullptr) {
        return Value::error(ErrorCode::Value);
      }
      if (wb == nullptr) {
        return Value::number(1.0);
      }
      const std::size_t idx = find_sheet_index(*wb, *current);
      return idx == 0 ? Value::error(ErrorCode::Value) : Value::number(static_cast<double>(idx));
    }
    // Qualified ref — must have a workbook to look up the qualifier.
    if (wb == nullptr) {
      return Value::error(ErrorCode::NA);
    }
    const Sheet* target = wb->sheet_by_name(sheet_qualifier);
    if (target == nullptr) {
      return Value::error(ErrorCode::NA);
    }
    const std::size_t idx = find_sheet_index(*wb, *target);
    return idx == 0 ? Value::error(ErrorCode::NA) : Value::number(static_cast<double>(idx));
  }

  // Non-reference AST: evaluate the argument. If it coerces to text,
  // treat as a sheet-name lookup. Otherwise surface `#VALUE!` (or the
  // propagated error).
  const Value v = eval_node(arg, arena, registry, ctx);
  if (v.is_error()) {
    return v;
  }
  auto text_e = coerce_to_text(v);
  if (!text_e) {
    return Value::error(ErrorCode::Value);
  }
  if (wb == nullptr) {
    return Value::error(ErrorCode::NA);
  }
  const Sheet* target = wb->sheet_by_name(text_e.value());
  if (target == nullptr) {
    return Value::error(ErrorCode::NA);
  }
  const std::size_t idx = find_sheet_index(*wb, *target);
  return idx == 0 ? Value::error(ErrorCode::NA) : Value::number(static_cast<double>(idx));
}

Value eval_single_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  const parser::AstNode& arg = resolve_name_ast(call.as_call_arg(0), ctx.name_env());

  // Share the `@`-operator (NodeKind::ImplicitIntersection) projection so the
  // two spellings cannot drift on which operand shapes project or on which
  // cell they land: every reference that declares a rectangle — bounded
  // `Ref:Ref`, full-axis `Ref:Ref`, and the single `Ref` the parser folds
  // `A:A` / `1:1` into — goes through one projection. Out-of-axis formula
  // cells surface `#VALUE!`, matching the `@` operator's documented
  // behaviour.
  if (ctx.has_formula_cell()) {
    parser::Reference target{};
    switch (project_implicit_intersection(arg, ctx.formula_row(), ctx.formula_col(), &target)) {
      case IntersectionProjection::kCell:
        return ctx.resolve_ref(target, arena, registry);
      case IntersectionProjection::kNoCell:
        return Value::error(ErrorCode::Value);
      case IntersectionProjection::kNotStaticReference:
        break;
    }
    if (Value projected = Value::blank(); project_reference_result(arg, arena, registry, ctx, &projected)) {
      return projected;
    }
  } else if (arg.kind() == parser::NodeKind::RangeOp) {
    // Without a bound formula cell there is no row/col to project onto.
    return Value::error(ErrorCode::Value);
  }

  // Calls, spill references, and other non-range array expressions do not
  // retain coordinates to project. Both spellings take the top-left element.
  return implicit_intersect_value(eval_node(arg, arena, registry, ctx));
}

Value eval_sheets_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity > 1U) {
    return Value::error(ErrorCode::Value);
  }
  const Workbook* wb = ctx.workbook();

  if (arity == 0U) {
    // Total sheets in the workbook. No workbook -> degrade to 1, which
    // matches the "bound to a single Sheet" context shape.
    if (wb == nullptr) {
      return Value::number(1.0);
    }
    return Value::number(static_cast<double>(wb->sheet_count()));
  }

  const parser::AstNode& arg = call.as_call_arg(0);
  // Reference or range: MVP returns 1 (no 3D references yet). Errors
  // in the argument subtree still propagate if we evaluate.
  if (arg.kind() == parser::NodeKind::Ref || arg.kind() == parser::NodeKind::RangeOp ||
      arg.kind() == parser::NodeKind::StructuredRef) {
    return Value::number(1.0);
  }
  // Fall through: evaluate to surface errors, then reject non-references.
  // Mac Excel 365 surfaces `#N/A` here (not `#VALUE!`), matching Excel's
  // treatment of SHEETS("not a ref") as a missing-sheet lookup.
  const Value v = eval_node(arg, arena, registry, ctx);
  if (v.is_error()) {
    return v;
  }
  return Value::error(ErrorCode::NA);
}

Value eval_code_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  const Value arg = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (arg.is_error()) {
    return arg;
  }
  const ExcelProfile profile = ctx.excel_profile();
  if (profile.host == ExcelHost::kWin365 && sbcs_codepage(profile) == SbcsCodepage::kCp932) {
    auto text = coerce_to_text(arg);
    if (!text) {
      return Value::error(text.error());
    }
    if (text.value().empty()) {
      return Value::error(ErrorCode::Value);
    }
    const Utf8DecodeResult decoded = decode_first_utf8_codepoint(text.value());
    if (!decoded.valid) {
      return Value::error(ErrorCode::Value);
    }
    if (decoded.codepoint == 0x9AD9u) {
      return Value::number(38526.0);
    }
    if (decoded.codepoint > 0xFFFFu || is_hangul(decoded.codepoint)) {
      return Value::number(63.0);
    }
  }
  const FunctionDef* def = registry.lookup("CODE");
  if (def == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return def->impl(&arg, 1U, arena);
}

Value eval_lenb_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  const Value arg = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (arg.is_error()) {
    return arg;
  }
  const ExcelProfile profile = ctx.excel_profile();
  if (profile.host == ExcelHost::kWin365 && locale_facts(profile).dbcs_codepage != DbcsCodepage::kNone) {
    auto text = coerce_to_text(arg);
    if (!text) {
      return Value::error(text.error());
    }
    std::uint64_t bytes = 0;
    std::size_t i = 0;
    while (i < text.value().size()) {
      std::size_t step = 0;
      const std::uint32_t cp = decode_utf8_step(text.value(), i, &step);
      if (step == 0U) {
        break;
      }
      if (step == 1U && cp == 0xFFFDu) {
        bytes += 1U;
      } else if (cp <= 0x00FFu || (cp >= 0xFF61u && cp <= 0xFF9Fu) ||
                 (is_hangul(cp) && locale_facts(profile).dbcs_codepage == DbcsCodepage::kJis0208)) {
        bytes += 1U;
      } else {
        bytes += 2U;
      }
      i += step;
    }
    return Value::number(static_cast<double>(bytes));
  }
  const FunctionDef* def = registry.lookup("LENB");
  if (def == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return def->impl(&arg, 1U, arena);
}

Value eval_weeknum_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 2U) {
    return Value::error(ErrorCode::Value);
  }
  // Analysis-ToolPak rule: a boolean is #VALUE!, an omitted serial is #N/A, an omitted return_type is its default.
  const std::uint32_t evaluated = atp_evaluated_arity(call, 1U, 2U);
  Value args[2] = {eval_node(call.as_call_arg(0), arena, registry, ctx), Value::number(1.0)};
  for (std::uint32_t i = 0; i < evaluated; ++i) {
    if (i == 1U) {
      args[1] = eval_node(call.as_call_arg(1), arena, registry, ctx);
    }
    if (args[i].is_error()) {
      return args[i];
    }
    if (const Value atp = atp_arg_error(call.as_call_arg(i), args[i], atp_slot_required(i, 1U, 2U)); atp.is_error()) {
      return atp;
    }
  }
  // WEEKNUM is date1904-sensitive and no longer in the eager registry; reach
  // its shared calendar impl through `find_date_entry` and thread the
  // workbook epoch. An unsupported return_type is #NUM! there, as Windows
  // Excel 365 reports it (Mac Excel falls back to type 1).
  const DateEntry* entry = find_date_entry("WEEKNUM");
  if (entry == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return entry->impl(args, evaluated, arena, ctx.date1904());
}

namespace {

// Windows ja-JP writes a format's `\` as a literal ¥ rather than an escape
// (text_format.text_backslash_escape); an array of formats is rewritten per element.
Value backslash_as_yen(const Value& fmt, Arena& arena) {
  if (fmt.is_array()) {
    const ArrayValue* src = fmt.as_array();
    Value* cells = nullptr;
    ArrayValue* out = allocate_array_value(src->rows, src->cols, arena, cells, kMaxDerivedArrayCells);
    if (out == nullptr) {
      return Value::error(ErrorCode::Num);
    }
    const std::size_t count = static_cast<std::size_t>(src->rows) * src->cols;
    for (std::size_t i = 0; i < count; ++i) {
      cells[i] = backslash_as_yen(src->cells[i], arena);
    }
    return Value::array(out);
  }
  if (!fmt.is_text() || fmt.as_text().find('\\') == std::string_view::npos) {
    return fmt;
  }
  std::string rewritten;
  rewritten.reserve(fmt.as_text().size() + 4U);
  for (const char ch : fmt.as_text()) {
    if (ch == '\\') {
      rewritten += "\xC2\xA5";
    } else {
      rewritten.push_back(ch);
    }
  }
  return Value::text(arena.intern(rewritten));
}

}  // namespace

Value eval_text_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  if (call.as_call_arity() != 2U) {
    return Value::error(ErrorCode::Value);
  }
  Value args[2] = {eval_node(call.as_call_arg(0), arena, registry, ctx),
                   eval_node(call.as_call_arg(1), arena, registry, ctx)};
  if (args[0].is_error()) {
    return args[0];
  }
  if (args[1].is_error()) {
    return args[1];
  }
  const ExcelProfile profile = ctx.excel_profile();
  if (profile.host == ExcelHost::kWin365 && sbcs_codepage(profile) == SbcsCodepage::kCp932) {
    args[1] = backslash_as_yen(args[1], arena);
  }
  // TEXT is date1904-sensitive (date format codes read the workbook epoch),
  // so it is served through the shared ctx-date1904 lookup rather than the
  // eager registry; thread the workbook epoch into the impl.
  const DateEntry* entry = find_date_entry("TEXT");
  if (entry == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return invoke_date_entry(*entry, args, 2U, arena, ctx.date1904(), ctx.wall_clock());
}

}  // namespace eval
}  // namespace formulon
