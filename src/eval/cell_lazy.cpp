//
// Implementation of `CELL(info_type, [reference])`. See `cell_lazy.h`
// for the contract and the explicit list of MVP stub keys.
//
// CELL rides the lazy seam because the reference argument is optional
// and, when supplied, may be a multi-cell range whose top-left we
// extract without dereferencing the whole rectangle. The eager path
// would flatten both shapes to a `Value` before the impl runs, losing
// the address information we need for "address" / "row" / "col".

#include "eval/cell_lazy.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "eval/array_alloc.h"
#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/range_resolvers.h"
#include "io/ooxml_writer_cell.h"
#include "io/styles_reader.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// ASCII-lowercase fold for the info_type key. Excel matches CELL's first
// argument case-insensitively (`"ROW"` == `"row"` == `"Row"`); only ASCII
// is in the supported keyset so a byte-level fold suffices.
std::string ascii_tolower(std::string_view s) {
  std::string out;
  out.reserve(s.size());
  for (char c : s) {
    const auto u = static_cast<unsigned char>(c);
    if (u >= 'A' && u <= 'Z') {
      out.push_back(static_cast<char>(u + 32));
    } else {
      out.push_back(c);
    }
  }
  return out;
}

// Copies `s` into `arena` and returns a Text Value pointing at the
// arena-owned storage. Used for keys whose result is a constant string
// ("filename", "format", "prefix"). For genuinely empty strings we still
// return `Value::text({})` so we don't allocate a zero-byte buffer.
Value arena_text(Arena& arena, std::string_view s) {
  if (s.empty()) {
    return Value::text({});
  }
  char* buf = static_cast<char*>(arena.allocate(s.size(), alignof(char)));
  if (buf == nullptr) {
    return Value::error(ErrorCode::Value);
  }
  for (std::size_t i = 0; i < s.size(); ++i) {
    buf[i] = s[i];
  }
  return Value::text(std::string_view(buf, s.size()));
}

// Resolves the optional `reference` argument to the top-left cell of the
// rectangle it denotes, plus an optional sheet qualifier string when the
// reference targets a sheet other than the formula cell's sheet.
//
// Behaviour:
//   * `Ref` / `RangeOp` - delegate to `resolve_range_endpoint`, which
//     understands plain Refs and reference-returning calls (OFFSET /
//     INDIRECT). For RangeOp we resolve the lhs endpoint and then
//     normalise the rectangle's top-left as the smallest (row, col)
//     across both endpoints.
//   * Reference-returning Call (OFFSET / INDIRECT / nested CHOOSE / IF)
//     - `resolve_reference_call` produces the rectangle directly.
//   * Anything else - evaluate via `eval_node` so a subtree error
//     propagates with the correct code; if the result is a non-error
//     non-reference scalar, surface `#VALUE!` (CELL needs a reference,
//     not a Value).
//
// Returns `true` on success and writes the 0-based top-left coordinates
// to `*out_row` / `*out_col` plus the sheet qualifier (empty when the
// resolved cell is on the bound current sheet) to `*out_sheet`.
// Returns `false` on failure with the Excel error code surfaced via
// `*out_err`.
bool extract_topleft_ref(const parser::AstNode& ref_node, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, std::uint32_t* out_row, std::uint32_t* out_col,
                         std::string_view* out_sheet, Value* out_err) {
  const parser::NodeKind k = ref_node.kind();

  // Plain Ref or RangeOp: route through `resolve_range_endpoint` for the
  // lhs to get a normalised rectangle. For a RangeOp we also resolve the
  // rhs and union the rectangles so the top-left is the min(row,col)
  // pair, matching Excel's "top-left of the supplied range" rule.
  if (k == parser::NodeKind::Ref) {
    std::string_view sheet;
    std::uint32_t top = 0;
    std::uint32_t left = 0;
    std::uint32_t bottom = 0;
    std::uint32_t right = 0;
    ErrorCode err = ErrorCode::Ref;
    if (!resolve_range_endpoint(ref_node, arena, registry, ctx, &sheet, &top, &left, &bottom, &right, &err)) {
      *out_err = Value::error(err);
      return false;
    }
    *out_row = top;
    *out_col = left;
    *out_sheet = sheet;
    return true;
  }
  if (k == parser::NodeKind::RangeOp) {
    const parser::AstNode& lhs_ast = ref_node.as_range_lhs();
    const parser::AstNode& rhs_ast = ref_node.as_range_rhs();
    std::string_view lhs_sheet;
    std::string_view rhs_sheet;
    std::uint32_t lhs_top = 0;
    std::uint32_t lhs_left = 0;
    std::uint32_t lhs_bottom = 0;
    std::uint32_t lhs_right = 0;
    std::uint32_t rhs_top = 0;
    std::uint32_t rhs_left = 0;
    std::uint32_t rhs_bottom = 0;
    std::uint32_t rhs_right = 0;
    ErrorCode err = ErrorCode::Ref;
    if (!resolve_range_endpoint(lhs_ast, arena, registry, ctx, &lhs_sheet, &lhs_top, &lhs_left, &lhs_bottom, &lhs_right,
                                &err) ||
        !resolve_range_endpoint(rhs_ast, arena, registry, ctx, &rhs_sheet, &rhs_top, &rhs_left, &rhs_bottom, &rhs_right,
                                &err)) {
      *out_err = Value::error(err);
      return false;
    }
    *out_row = lhs_top < rhs_top ? lhs_top : rhs_top;
    *out_col = lhs_left < rhs_left ? lhs_left : rhs_left;
    // Prefer the lhs sheet qualifier when present; falls back to rhs's
    // qualifier so `Sheet2!A1:B2` (parser's typical RangeOp shape) and
    // the defensive `A1:Sheet2!B2` shape both surface the right sheet.
    *out_sheet = lhs_sheet.empty() ? rhs_sheet : lhs_sheet;
    return true;
  }

  // Reference-returning call: OFFSET / INDIRECT / nested CHOOSE / IF.
  // `resolve_reference_call` walks these and produces a rectangle.
  if (k == parser::NodeKind::Call) {
    std::string_view sheet;
    std::uint32_t top = 0;
    std::uint32_t left = 0;
    std::uint32_t bottom = 0;
    std::uint32_t right = 0;
    bool is_range = false;
    ErrorCode err = ErrorCode::Value;
    if (resolve_reference_call(ref_node, arena, registry, ctx, &sheet, &top, &left, &bottom, &right, &is_range, &err)) {
      *out_row = top;
      *out_col = left;
      *out_sheet = sheet;
      return true;
    }
    // Fall through: not a reference-returning call. Evaluate so any
    // subtree error propagates verbatim, then reject the scalar result.
  }

  // Anything else: evaluate the subtree and propagate a subtree error.
  // A non-error, non-reference scalar surfaces `#VALUE!` because CELL
  // requires a reference for the address / row / col / contents / type
  // info_types.
  const Value v = eval_node(ref_node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  *out_err = Value::error(ErrorCode::Value);
  return false;
}

// Resolves the top-left cell's value through the bound context. Used by
// the "contents" and "type" keys. The sheet qualifier may be empty (use
// current sheet) or non-empty (look up via workbook); when no current
// sheet is bound at all we surface `#NAME?` to mirror
// `EvalContext::resolve_ref`'s degradation rule.
Value resolve_topleft_value(std::string_view sheet, std::uint32_t row, std::uint32_t col, Arena& arena,
                            const FunctionRegistry& registry, const EvalContext& ctx) {
  parser::Reference r{};
  r.sheet = sheet;
  r.row = row;
  r.col = col;
  return ctx.resolve_ref(r, arena, registry);
}

// Resolves CELL("protect") for the cell at `(sheet, row, col)`: reads the
// cell's cell-format (`xf`) protection `locked` flag from the workbook's
// StylesTable. Excel returns 1 for a locked cell and 0 for an unlocked one.
//
// Defaults follow the OOXML schema: the effective `locked` flag defaults to
// true, so an absent cell (xf 0), an xf without a `<protection>` element, or
// a context with no styles table all resolve to locked (1). A qualified
// reference whose target sheet is missing surfaces `#REF!`.
Value resolve_cell_locked(std::string_view sheet, std::uint32_t row, std::uint32_t col, const EvalContext& ctx) {
  const Workbook* workbook = ctx.workbook();
  const Sheet* target = ctx.current_sheet();
  if (!sheet.empty()) {
    if (workbook == nullptr) {
      // Qualified reference but no workbook bound to resolve it: default to
      // locked rather than fabricating a #REF!.
      return Value::number(1.0);
    }
    target = workbook->sheet_by_name(sheet);
    if (target == nullptr) {
      return Value::error(ErrorCode::Ref);
    }
  }
  // The cell's xf index (0 = default xf when the cell is absent).
  std::uint32_t xf_index = 0;
  if (target != nullptr) {
    if (const Cell* cell = target->cell_at(row, col); cell != nullptr) {
      xf_index = cell->xf_index;
    }
  }
  bool locked = true;
  if (workbook != nullptr) {
    const std::vector<io::CellXf>& cell_xfs = workbook->styles().cell_xfs;
    if (xf_index < cell_xfs.size()) {
      locked = cell_xfs[xf_index].locked;
    }
  }
  return Value::number(locked ? 1.0 : 0.0);
}

const io::CellXf* resolve_cell_xf(std::string_view sheet_name, std::uint32_t row, std::uint32_t col,
                                  const EvalContext& ctx) {
  const Workbook* workbook = ctx.workbook();
  const Sheet* target = ctx.current_sheet();
  if (!sheet_name.empty()) {
    target = workbook == nullptr ? nullptr : workbook->sheet_by_name(sheet_name);
  }
  if (workbook == nullptr || target == nullptr) {
    return nullptr;
  }
  std::uint32_t xf_index = 0;
  if (const Cell* cell = target->cell_at(row, col); cell != nullptr) {
    xf_index = cell->xf_index;
  }
  const std::vector<io::CellXf>& xfs = workbook->styles().cell_xfs;
  return xf_index < xfs.size() ? &xfs[xf_index] : nullptr;
}

bool section_uses_color(std::string_view section) {
  for (std::size_t i = 0; i + 4U < section.size(); ++i) {
    if (section[i] != '[') {
      continue;
    }
    const std::size_t close = section.find(']', i + 1U);
    if (close == std::string_view::npos) {
      break;
    }
    const std::string_view tag = section.substr(i + 1U, close - i - 1U);
    if (tag == "Red" || tag == "Blue" || tag == "Green" || tag == "Yellow" || tag == "Magenta" || tag == "Cyan" ||
        tag.rfind("Color", 0) == 0U) {
      return true;
    }
    i = close;
  }
  return false;
}

// CELL("color") answers whether negative values use a colour. A one-section
// number format applies to all numbers; in the usual multi-section form the
// second section is the negative-number section. Semicolons in quoted text or
// bracket directives do not divide format sections.
bool number_format_colors_negative_values(std::string_view format) {
  bool quoted = false;
  std::size_t bracket_depth = 0;
  std::size_t negative_start = std::string_view::npos;
  std::size_t negative_end = format.size();
  for (std::size_t i = 0; i < format.size(); ++i) {
    const char c = format[i];
    if (c == '"') {
      quoted = !quoted;
    } else if (!quoted && c == '[') {
      ++bracket_depth;
    } else if (!quoted && c == ']' && bracket_depth != 0) {
      --bracket_depth;
    } else if (!quoted && bracket_depth == 0 && c == ';') {
      if (negative_start == std::string_view::npos) {
        negative_start = i + 1U;
      } else {
        negative_end = i;
        break;
      }
    }
  }
  if (negative_start == std::string_view::npos) {
    return section_uses_color(format);
  }
  return section_uses_color(format.substr(negative_start, negative_end - negative_start));
}

std::string_view cell_prefix(const io::CellXf* xf) {
  if (xf == nullptr) {
    return {};
  }
  if (xf->quote_prefix) {
    return "'";
  }
  switch (xf->horizontal_align) {
    case 1:
      return "\\";
    case 2:
      return "^";
    case 3:
      return "\"";
    default:
      return {};
  }
}

std::string_view number_format_for_xf(const io::CellXf* xf, const EvalContext& ctx) {
  if (xf == nullptr || ctx.workbook() == nullptr) {
    return {};
  }
  const io::StylesTable& styles = ctx.workbook()->styles();
  for (const io::NumFmtRecord& custom : styles.num_fmts) {
    if (custom.id == xf->num_fmt_id && custom.format_string_index < styles.num_fmt_strings.size()) {
      return styles.num_fmt_strings[custom.format_string_index];
    }
  }
  return io::builtin_num_fmt(xf->num_fmt_id);
}

std::string_view cell_format_code(const io::CellXf* xf) {
  if (xf == nullptr) {
    return "G";
  }
  switch (xf->num_fmt_id) {
    case 0:
      return "G";
    case 1:
      return "F0";
    case 2:
      return "F2";
    case 3:
    case 37:
      return ",0";
    case 4:
    case 39:
      return ",2";
    case 5:
      return "C0";
    case 6:
      return "C0-";
    case 7:
      return "C2";
    case 8:
      return "C2-";
    case 9:
      return "P0";
    case 10:
      return "P2";
    case 11:
      return "S2";
    case 14:
    case 22:
      return "D4";
    case 15:
      return "D1";
    case 16:
      return "D2";
    case 17:
      return "D3";
    case 18:
      return "D7";
    case 19:
      return "D6";
    case 20:
      return "D9";
    case 21:
      return "D8";
    default:
      return "G";
  }
}

// Builds CELL("width")'s 1x2 result. The first member is the effective
// column width; the second tells callers whether that width came from the
// sheet default rather than an explicit <col> override.
Value build_width_result(const Sheet& sheet, std::uint32_t col, Arena& arena) {
  double width = sheet.format_defaults().has_default_col_width ? sheet.format_defaults().default_col_width
                                                               : sheet.format_defaults().base_col_width;
  bool is_default = true;
  for (const ColumnLayout& layout : sheet.layout().columns) {
    if (layout.first <= col && col <= layout.last) {
      // A style/visibility-only span does not override the column metric.
      // Keep an explicit width="0" distinct from an omitted width.
      if (HasExplicitColumnWidth(layout)) {
        width = layout.width;
        is_default = false;
      }
    }
  }
  Value* cells = nullptr;
  ArrayValue* arr = allocate_array_value(1U, 2U, arena, cells, kMaxDerivedArrayCells);
  if (arr == nullptr) {
    return Value::error(ErrorCode::Value);
  }
  cells[0] = Value::number(width);
  cells[1] = Value::boolean(is_default);
  return Value::array(arr);
}

Value resolve_cell_width(std::string_view sheet_name, std::uint32_t col, Arena& arena, const EvalContext& ctx) {
  const Workbook* workbook = ctx.workbook();
  const Sheet* target = ctx.current_sheet();
  if (!sheet_name.empty()) {
    if (workbook == nullptr) {
      return Value::error(ErrorCode::Ref);
    }
    target = workbook->sheet_by_name(sheet_name);
  }
  if (target == nullptr) {
    return Value::error(ErrorCode::Ref);
  }
  return build_width_result(*target, col, arena);
}

// Builds the `"address"` answer: an A1 absolute address (`$Col$Row`)
// optionally prefixed with a sheet qualifier. Sheet prefixing follows
// Excel's bare-name rule (no quoting for plain identifiers); we keep
// the formatting minimal because the workbook has no filename to wrap
// in `[Workbook]` brackets.
Value build_address(Arena& arena, std::string_view sheet, std::uint32_t row, std::uint32_t col) {
  // `EncodeA1` returns an unanchored address (e.g. "A1"); we splice the
  // `$` anchors in manually because the caller wants an absolute form.
  // The helper accepts 0-based (row, col) and produces 1-based output.
  const std::string a1 = io::EncodeA1(row, col);
  // Find the boundary between the column letters and the row digits so
  // we can insert the leading `$` and the inner `$` for absolute form.
  std::size_t row_start = 0;
  while (row_start < a1.size() && a1[row_start] >= 'A' && a1[row_start] <= 'Z') {
    ++row_start;
  }
  // Compose `[sheet!]$Col$Row` directly into a std::string, then copy
  // the bytes into the arena via `arena_text`.
  std::string result;
  result.reserve(sheet.size() + a1.size() + 4);
  if (!sheet.empty()) {
    result.append(sheet);
    result.push_back('!');
  }
  result.push_back('$');
  result.append(a1.substr(0, row_start));
  result.push_back('$');
  result.append(a1.substr(row_start));
  return arena_text(arena, result);
}

// Resolves the reference argument or, when omitted, the formula cell's
// own coords. Returns `true` on success and writes the top-left into
// `*out_row` / `*out_col` plus the sheet qualifier (empty for the
// current sheet). Failure returns `false` with the surfaced Value
// (error or otherwise) in `*out_result`.
bool resolve_topleft_or_formula_cell(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                     const EvalContext& ctx, std::uint32_t* out_row, std::uint32_t* out_col,
                                     std::string_view* out_sheet, Value* out_result) {
  if (call.as_call_arity() == 1U) {
    // No reference supplied: anchor to the formula cell. Without an
    // anchor (CLI ad-hoc eval) there's nothing meaningful to answer
    // about row / col / address, so surface `#REF!` defensively.
    if (!ctx.has_formula_cell()) {
      *out_result = Value::error(ErrorCode::Ref);
      return false;
    }
    *out_row = ctx.formula_row();
    *out_col = ctx.formula_col();
    *out_sheet = std::string_view{};
    return true;
  }
  // Two-argument form: extract from the supplied reference AST.
  Value err = Value::blank();
  if (!extract_topleft_ref(call.as_call_arg(1), arena, registry, ctx, out_row, out_col, out_sheet, &err)) {
    *out_result = err;
    return false;
  }
  return true;
}

}  // namespace

Value eval_cell_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 2U) {
    return Value::error(ErrorCode::Value);
  }

  // info_type: evaluate, propagate any error, then coerce to text. An
  // Array argument fails the coercion and surfaces `#VALUE!`. An empty
  // string is also `#VALUE!` because no Excel-defined key is empty.
  const Value info_v = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (info_v.is_error()) {
    return info_v;
  }
  auto info_text = coerce_to_text(info_v);
  if (!info_text) {
    return Value::error(info_text.error());
  }
  if (info_text.value().empty()) {
    return Value::error(ErrorCode::Value);
  }
  const std::string key = ascii_tolower(info_text.value());

  // Reference-dependent protection query: reads the referenced cell's xf
  // `locked` flag. Resolves the top-left (or the formula cell for the
  // one-argument form) like the address / row / col keys below.
  if (key == "protect") {
    std::uint32_t row = 0;
    std::uint32_t col = 0;
    std::string_view sheet;
    Value early_result = Value::blank();
    if (!resolve_topleft_or_formula_cell(call, arena, registry, ctx, &row, &col, &sheet, &early_result)) {
      // The one-argument form with no bound formula cell (ad-hoc CLI eval)
      // has no cell to inspect; default to locked (the schema default).
      // The two-argument form propagates a reference-side error.
      if (call.as_call_arity() == 1U) {
        return Value::number(1.0);
      }
      return early_result;
    }
    return resolve_cell_locked(sheet, row, col, ctx);
  }

  // Reference-dependent keys: address / col / row / contents / type / width / color / prefix / format.
  // Resolve the top-left first so any reference-side error propagates
  // before we branch on the key.
  if (key == "address" || key == "col" || key == "row" || key == "contents" || key == "type" || key == "width" ||
      key == "color" || key == "prefix" || key == "format") {
    std::uint32_t row = 0;
    std::uint32_t col = 0;
    std::string_view sheet;
    Value early_result = Value::blank();
    if (!resolve_topleft_or_formula_cell(call, arena, registry, ctx, &row, &col, &sheet, &early_result)) {
      return early_result;
    }
    if (key == "address") {
      return build_address(arena, sheet, row, col);
    }
    if (key == "col") {
      return Value::number(static_cast<double>(col + 1U));
    }
    if (key == "row") {
      return Value::number(static_cast<double>(row + 1U));
    }
    if (key == "width") {
      return resolve_cell_width(sheet, col, arena, ctx);
    }
    if (key == "prefix") {
      return arena_text(arena, cell_prefix(resolve_cell_xf(sheet, row, col, ctx)));
    }
    if (key == "color") {
      return Value::number(
          number_format_colors_negative_values(number_format_for_xf(resolve_cell_xf(sheet, row, col, ctx), ctx)) ? 1.0
                                                                                                                 : 0.0);
    }
    if (key == "format") {
      return arena_text(arena, cell_format_code(resolve_cell_xf(sheet, row, col, ctx)));
    }
    // "contents" / "type" both need the resolved cell's value.
    const Value resolved = resolve_topleft_value(sheet, row, col, arena, registry, ctx);
    if (key == "contents") {
      // Excel's blank-as-zero rule for CELL("contents", blank_cell).
      if (resolved.is_blank()) {
        return Value::number(0.0);
      }
      return resolved;
    }
    // "type": "b" only for a genuinely blank cell, "l" for text, "v" for
    // everything else (number, bool, error). Errors short-circuit only
    // when they came from the reference argument itself; an error *value*
    // sitting in a cell still classifies as "v".
    //
    // A zero-length string is text, not blank: Excel reports "l" for it
    // and agrees whether it arrived as a constant or as the result of
    // `=""`. The two are indistinguishable to every predicate Excel
    // exposes, so no cell metadata is needed to tell them apart.
    if (resolved.is_blank()) {
      return arena_text(arena, "b");
    }
    if (resolved.is_text()) {
      return arena_text(arena, "l");
    }
    return arena_text(arena, "v");
  }

  // Reference-independent keys. These ignore the supplied reference
  // entirely (Excel observed: passes through without dereferencing,
  // even when the reference would otherwise error). If the reference
  // argument itself is a top-level error literal that path is rare;
  // the spec calls for fixed stubs here so we don't pre-evaluate.
  if (key == "filename") {
    // No filesystem path on the Workbook yet. Mac returns blank when the
    // workbook has never been saved; we surface empty text instead so
    // the top-level blank-as-zero rule (`evaluate()` in tree_walker)
    // does not collapse the result to 0. Oracle verification accepts the
    // xlwings "" read-back artifact via empty_string_readback.
    return arena_text(arena, "");
  }
  if (key == "parentheses") {
    // No parenthesis-format flag yet.
    return Value::number(0.0);
  }

  // Unknown info_type after lowercase fold.
  return Value::error(ErrorCode::Value);
}

}  // namespace eval
}  // namespace formulon
