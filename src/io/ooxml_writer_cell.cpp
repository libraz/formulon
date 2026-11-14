//
// Implementation of the OOXML cell/row/sheetData builder. Pure functions
// only: no zip plumbing, no I/O. The orchestrating writer in
// ooxml_writer.cpp wraps this output in the surrounding <worksheet>
// element and packages it into the .xlsx archive.
//
// Spill semantics:
//
//   * Anchor cell of a registered spill region: emit <f t="array" ref="...">
//     with ref= covering the spill footprint (the anchor cell alone when
//     the region is 1x1). That ref= is what lets Excel re-spill the
//     region on open; a bare t="array" with no ref reads back as a
//     legacy single-cell CSE array instead.
//   * Phantom cell (covered by an anchor's region but not the anchor
//     itself): written from the spill table as Excel writes it, a cached
//     value with no formula (`AppendPhantomCellXml`), so a reader that does
//     not recalculate sees the spilled values.
//
// oracle-verify: r14:spill="1" not emitted; verify against Mac Excel
// 16.108.1 if re-spill on load fails.

#include "io/ooxml_writer_cell.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "cell.h"
#include "cell_style_save_plan.h"
#include "io/data_table_xml.h"
#include "io/dynamic_array_formula.h"
#include "io/future_functions.h"
#include "io/ooxml/cell_ref_writer.h"
#include "io/ooxml/external_link_writer.h"
#include "io/ooxml/shared_strings_writer.h"
#include "io/phonetic_pr.h"
#include "io/stored_cell_error.h"
#include "io/xlsb/ptg_writer.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "io/xsd_bool.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/ast_shift.h"
#include "parser/parser.h"
#include "phonetic.h"
#include "sheet.h"
#include "sheet_layout_edit.h"
#include "style_resolve.h"
#include "utils/a1_column.h"
#include "utils/a1_ref.h"
#include "utils/arena.h"
#include "utils/formula_text.h"
#include "utils/index_sort.h"
#include "value.h"

namespace formulon {
namespace io {
namespace {

// ---------------------------------------------------------------------------
// Cell emission
// ---------------------------------------------------------------------------

// True when the explicit default style a spill phantom covers is written as
// `s="0"`. A blank's own explicit default is written bare (`<c r="..."/>`, as
// Excel does), but a phantom carries a value and reloads as a value cell,
// whose missing `s` would read as inheriting; the attribute keeps the
// explicit default across a Formulon reload. Excel itself treats an absent
// `s` and `s="0"` the same on a present cell.
bool PhantomForcesDefaultStyle(const Sheet& sheet, std::uint32_t row, std::uint32_t col, const Cell* stored) {
  const std::optional<StyleSlot> slot = sheet.cell_style_at(row, col);
  if (!slot || !slot->explicit_xf || slot->xf != 0U) {
    return false;
  }
  if (stored != nullptr && !(stored->formula_text.empty() && stored->cached_value.is_blank())) {
    return false;
  }
  const std::uint64_t end = static_cast<std::uint64_t>(col) + 1U;
  return !explicit_default_observable(sheet.layout(), row, col, end).empty();
}

// Emits an `s="N"` attribute when `xf_index` is non-zero. The default
// xf (index 0) is omitted for byte parity with Excel's writer unless `force`
// is set (see `PhantomForcesDefaultStyle`).
// The saved index of model cell xf `xf_index`, for a cell inside a merged range when `merged`; unchanged without a
// plan.
std::uint32_t SavedXf(const CellStyleSavePlan* style_plan, std::uint32_t xf_index, bool merged = false) {
  if (style_plan == nullptr) {
    return xf_index;
  }
  return merged ? style_plan->merged_cell_xf_index(xf_index) : style_plan->cell_xf_index(xf_index);
}

void AppendStyleAttr(std::string& out, std::uint32_t xf_index, bool force, bool show_phonetic = false) {
  if (xf_index != 0U || force) {
    append_xml_attr_uint(out, "s", xf_index);
  }
  if (show_phonetic) {
    emit_xsd_bool_attr(out, "ph", true);
  }
}

// The `vm=` of `(row, col)` in `rich_error_vm`, or 0 for none.
std::uint32_t VmAt(const RichErrorVmMap* rich_error_vm, std::uint32_t row, std::uint32_t col) {
  return rich_error_vm == nullptr ? 0U : rich_error_vm_at(*rich_error_vm, row, col);
}

// Emits `vm="N"` (a rich error's value metadata) when `vm` is non-zero.
void AppendVmAttr(std::string& out, std::uint32_t vm) {
  if (vm != 0U) {
    append_xml_attr_uint(out, "vm", vm);
  }
}

// Emits the <c> element for an Error value at `addr`.
void AppendErrorCellXml(std::string& out, std::string_view addr, ErrorCode code, std::uint32_t xf_index,
                        bool show_phonetic) {
  out.append("<c r=\"");
  out.append(addr);
  out.append("\"");
  AppendStyleAttr(out, xf_index, false, show_phonetic);
  out.append(" t=\"e\"><v>");
  out.append(display_name(stored_cell_error(code)));
  out.append("</v></c>");
}

// Emits the body (the <v> or <is>...</is> child) for a non-formula cell
// holding `value`. Caller has already pre-screened NaN/Inf and Blank;
// Array/Ref/Lambda are defensively downgraded to #VALUE! since they
// should never appear in cell storage in practice.
//
// On entry, `out` already contains '<c r="ADDR"' (without the closing
// quote or '>'). This function emits the closing quote, the `s=` style
// attribute for a non-zero `xf_index`, any type attribute, the body, and
// the closing '</c>'. Routing the style attribute through here rather
// than through the caller is what keeps the type attribute and the
// `<v>`/`<is>` encoding identical for styled and unstyled cells.
//
// `phonetic` is the kana annotation associated with a Text-valued cell
// (empty when none). When non-empty AND `value` is Text, the `<is>`
// block expands from `<is><t>{text}</t></is>` to
// `<is><t>{text}</t><rPh sb="S" eb="E"><t>{kana}</t></rPh>...<phoneticPr .../></is>`,
// one block per run in the order the run list holds them. The spans are
// emitted as stored rather than merged into one whole-string block:
// PHONETIC leaves the text outside every span in place, so a merged
// block would read kana over characters it does not cover.
void AppendLiteralCellBody(std::string& out, const Value& value, const std::vector<PhoneticRun>& phonetic,
                           PhoneticProperties phonetic_props, const SharedStrings* shared_strings,
                           std::uint32_t xf_index, bool show_phonetic, std::uint32_t vm) {
  out.push_back('"');
  AppendStyleAttr(out, xf_index, false, show_phonetic);
  if (value.is_number()) {
    out.append("><v>");
    append_xml_number(out, value.as_number());
    out.append("</v></c>");
    return;
  }
  if (value.is_boolean()) {
    out.append(" t=\"b\"><v>");
    out.push_back(value.as_boolean() ? '1' : '0');
    out.append("</v></c>");
    return;
  }
  if (value.is_text()) {
    if (shared_strings != nullptr) {
      out.append(" t=\"s\"><v>");
      out.append(std::to_string(shared_strings->index_of(value.as_text(), phonetic, phonetic_props)));
      out.append("</v></c>");
      return;
    }
    out.append(" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
    AppendXmlEscaped(out, value.as_text());
    out.append("</t>");
    append_phonetic_runs(out, phonetic, phonetic_props);
    out.append("</is></c>");
    return;
  }
  if (value.is_error()) {
    out.append(" t=\"e\"");
    AppendVmAttr(out, vm);
    out.append("><v>");
    out.append(display_name(stored_cell_error(value.as_error())));
    out.append("</v></c>");
    return;
  }
  // Defensive fallback for Array/Ref/Lambda: cells should never store
  // these in the post-evaluation path (arrays land in the spill table;
  // Ref/Lambda are not yet first-class cell payloads). Surface #VALUE!
  // rather than crashing the writer.
  out.append(" t=\"e\"><v>");
  out.append(display_name(ErrorCode::Value));
  out.append("</v></c>");
}

// The `t=` a formula cell's cached value needs (none for a number); a
// non-finite number is stored as the error `<v>` below writes for it.
void AppendFormulaValueType(std::string& out, const Value& cached) {
  if (cached.is_error() || (cached.is_number() && !std::isfinite(cached.as_number()))) {
    out.append(" t=\"e\"");
  } else if (cached.is_text()) {
    out.append(" t=\"str\"");
  } else if (cached.is_boolean()) {
    out.append(" t=\"b\"");
  }
}

// Appends a formula cell's cached value `<v>`.
void AppendFormulaValue(std::string& out, const Value& cv) {
  // <v>: omit when blank (Excel will recalculate on load); downgrade
  // NaN/Inf number to #NUM! text inside <v>; otherwise emit normally.
  if (cv.is_blank()) {
    // No <v> at all.
  } else if (cv.is_number()) {
    const double v = cv.as_number();
    if (std::isfinite(v)) {
      out.append("<v>");
      append_xml_number(out, v);
      out.append("</v>");
    } else {
      out.append("<v>");
      out.append(display_name(ErrorCode::Num));
      out.append("</v>");
    }
  } else if (cv.is_boolean()) {
    out.append("<v>");
    out.push_back(cv.as_boolean() ? '1' : '0');
    out.append("</v>");
  } else if (cv.is_text()) {
    // Formula cells with text results inline the string in <v> rather
    // than the <is><t> form used by literal text cells. Excel accepts
    // both shapes for formula results. `xml:space="preserve"` mirrors
    // `AppendLiteralCellBody`'s `<is><t xml:space="preserve">`: Excel
    // trims leading/trailing whitespace from a cached string value on
    // reload unless this hint is present, and a cached formula result
    // is just as much a displayed string as a literal one.
    out.append("<v xml:space=\"preserve\">");
    AppendXmlEscaped(out, cv.as_text());
    out.append("</v>");
  } else if (cv.is_error()) {
    out.append("<v>");
    out.append(display_name(stored_cell_error(cv.as_error())));
    out.append("</v>");
  }
  // Array / Ref / Lambda cached values fall through with no <v>; the
  // engine evaluates on load.
}

// A spilled (non-anchor) cell as Excel 365 stores it: its own style, the
// value typed as a formula result's, and `<f ca="1"/>` when the anchor is
// recalculated every time.
void AppendPhantomCellXml(std::string& out, std::uint32_t row, std::uint32_t col, std::uint32_t xf_index,
                          bool force_style, const Value& value, bool always_calculates, std::uint32_t vm) {
  out.append("<c r=\"");
  AppendCellRefForRef(out, row, col);
  out.append("\"");
  AppendStyleAttr(out, xf_index, force_style);
  AppendFormulaValueType(out, value);
  AppendVmAttr(out, value.is_error() ? vm : 0U);
  out.push_back('>');
  if (always_calculates) {
    out.append("<f ca=\"1\"/>");
  }
  AppendFormulaValue(out, value);
  out.append("</c>");
}

// ---------------------------------------------------------------------------
// Shared formulas
// ---------------------------------------------------------------------------

// A cell's part in a shared-formula group the save writes: the master holds
// the text and the group's `ref`, a member only the group's `si`.
struct SharedRole {
  std::uint32_t si = 0;
  bool master = false;
  MergeRange ref;
};

using SharedRoles = std::unordered_map<std::uint64_t, SharedRole>;

std::uint64_t CellKey(std::uint32_t row, std::uint32_t col) {
  return (static_cast<std::uint64_t>(row) << 32) | col;
}

// Whether the cell at `(row, col)` may sit in a shared formula: a formula
// that is no CSE block, dynamic-array formula or spill anchor.
bool SharableFormulaCell(const Sheet& sheet, std::uint32_t row, std::uint32_t col, const Cell* cell) {
  return cell != nullptr && !cell->formula_text.empty() && cell->array_block_rows == 0U &&
         !is_dynamic_array_formula(*cell) && sheet.spill_region_at_anchor(row, col) == nullptr &&
         sheet.data_table_anchored_at(row, col) == nullptr;
}

const parser::AstNode* ParseCellFormula(std::string_view formula_text, Arena& arena) {
  arena.reset();
  return parser::parse_strict(strip_formula_prefix(formula_text), arena);
}

// The role of each cell of the sheet's shared-formula groups, verified now:
// a group whose master (top-left cell) is no sharable formula, or is
// recalculated every time, is written individually; a member is shared only
// when its formula is the master's moved to its cell, as text. `si` counts
// from 0 over the groups written; a cell keeps the first group claiming it.
SharedRoles CollectSharedRoles(const Sheet& sheet, const xlsb::NameShapes& name_shapes) {
  SharedRoles roles;
  std::uint32_t next_si = 0;
  Arena master_arena;
  Arena member_arena;
  Arena shift_arena;
  for (const SharedFormulaGroup& group : sheet.shared_formula_groups()) {
    const MergeRange& ref = group.ref;
    const Cell* master = sheet.cell_at(ref.first_row, ref.first_col);
    if (!SharableFormulaCell(sheet, ref.first_row, ref.first_col, master) ||
        roles.count(CellKey(ref.first_row, ref.first_col)) != 0U) {
      continue;
    }
    const parser::AstNode* root = ParseCellFormula(master->formula_text, master_arena);
    if (root == nullptr ||
        formula_cell_always_calculates(sheet, ref.first_row, ref.first_col, *master, root, name_shapes)) {
      continue;
    }
    const std::uint32_t si = next_si++;
    roles.emplace(CellKey(ref.first_row, ref.first_col), SharedRole{si, true, ref});
    for (std::uint32_t row = ref.first_row; row <= ref.last_row; ++row) {
      for (std::uint32_t col = ref.first_col; col <= ref.last_col; ++col) {
        const Cell* cell = sheet.cell_at(row, col);
        if (!SharableFormulaCell(sheet, row, col, cell) || roles.count(CellKey(row, col)) != 0U) {
          continue;
        }
        const parser::AstNode* own = ParseCellFormula(cell->formula_text, member_arena);
        shift_arena.reset();
        const parser::AstNode* moved =
            parser::shift_relative_refs(*root, shift_arena, static_cast<std::int32_t>(row - ref.first_row),
                                        static_cast<std::int32_t>(col - ref.first_col));
        if (own != nullptr && moved != nullptr && parser::format_formula(*own) == parser::format_formula(*moved)) {
          roles.emplace(CellKey(row, col), SharedRole{si, false, ref});
        }
      }
    }
  }
  return roles;
}

// Emits the <c> element for `(row, col)` into `out`. Returns true when a
// cell was written, false when the cell was suppressed (blank literal,
// phantom of a spill region).
//
// Spill anchor handling: if `(row, col)` is anchored, the formula is
// emitted with t="array" and the cached value (cells[0]) becomes the
// <v>. Phantoms are suppressed by the caller via spill_region_covering;
// this function trusts the caller and never re-checks.
//
// `shared` is the cell's verified shared-formula role, or null. `vm` is the
// cell's rich-error value metadata, 0 for none.
bool AppendCellXml(std::string& out, const Sheet& sheet, std::uint32_t row, std::uint32_t col, const Cell& cell,
                   const SharedStrings* shared_strings, std::uint32_t dynamic_array_cm_index,
                   const xlsb::NameShapes& name_shapes, const parser::ExternalBookIndexer* indexer,
                   const SharedRole* shared, std::uint32_t vm, const CellStyleSavePlan* style_plan, bool merged) {
  const bool has_formula = !cell.formula_text.empty();
  const std::uint32_t xf = SavedXf(style_plan, cell.xf_index, merged);
  if (!CellIsEmitted(sheet, row, col, cell)) {
    return false;
  }

  // The anchor of a What-If data table carries the `<f t="dataTable">`; the
  // other result cells are plain values.
  if (const DataTable* table = has_formula ? nullptr : sheet.data_table_anchored_at(row, col)) {
    out.append("<c r=\"");
    AppendCellRefForRef(out, row, col);
    out.append("\"");
    AppendStyleAttr(out, xf, false);
    AppendFormulaValueType(out, cell.cached_value);
    out.push_back('>');
    append_data_table_formula(out, *table);
    AppendFormulaValue(out, cell.cached_value);
    out.append("</c>");
    return true;
  }

  const std::string addr = a1::encode_a1(row, col);

  // Pre-screen NaN/Inf number literals so we never half-emit a cell tag.
  // Formula cells with non-finite cached values still emit the <f>; only
  // the <v> is downgraded.
  if (!has_formula && cell.cached_value.is_number() && !std::isfinite(cell.cached_value.as_number())) {
    AppendErrorCellXml(out, addr, ErrorCode::Num, xf, cell.show_phonetic);
    return true;
  }

  // A blank stored cell (phonetic state, or a format) is written as
  // `<c r="..." s="N"/>`; an explicit default is the bare `<c r="..."/>`.
  if (!has_formula && cell.cached_value.is_blank()) {
    out.append("<c r=\"");
    out.append(addr);
    out.append("\"");
    AppendStyleAttr(out, xf, false, cell.show_phonetic);
    out.append("/>");
    return true;
  }

  if (has_formula) {
    const SpillRegion* anchored = sheet.spill_region_at_anchor(row, col);
    std::string_view formula = cell.formula_text;
    formula = strip_formula_prefix(formula);
    Arena formula_arena;
    parser::Parser formula_parser(formula, formula_arena);
    parser::AstNode* formula_root = formula_parser.parse();
    if (formula_root != nullptr && !formula_parser.errors().empty()) {
      formula_root = nullptr;
    }
    // A dynamic-array formula keeps that form when it does not spill, or
    // Excel reads it with implicit intersection (`=@...`); it needs the
    // XLDAPR entry `cm=` names, so without one the plain form stays. A legacy
    // CSE block keeps `t="array"` alone, with its own extent as `ref`.
    const bool dynamic = dynamic_array_cm_index != 0U && is_dynamic_array_formula(cell);
    const bool cse = cell.array_block_rows != 0U;
    out.append("<c r=\"");
    out.append(addr);
    out.append("\"");
    // Attributes in Excel's order: style, the value's type (which must agree
    // with the cached <v> so a save/load round trip keeps it), then `cm=`.
    // `cm=` links a dynamic-array formula to the `xl/metadata.xml` XLDAPR
    // entry (see `FindDynamicArrayCellMetadataIndex` in ooxml_writer.cpp)
    // that tells Excel this `t="array"` is a modern spill rather than a
    // legacy CSE array on reopen; emitted only when the saved package
    // carries a resolved entry for it.
    // Excel keeps `ph` on constants and blanks only.
    AppendStyleAttr(out, xf, false);
    AppendFormulaValueType(out, cell.cached_value);
    if (dynamic) {
      append_xml_attr_uint(out, "cm", dynamic_array_cm_index);
    }
    AppendVmAttr(out, cell.cached_value.is_error() ? vm : 0U);
    out.push_back('>');

    // <f> with optional t="array" ref="..." for spill anchors. Modern
    // Excel marks a dynamic-array anchor with `t="array"` plus a `ref`
    // covering the spill footprint; the `ref` is what lets Excel re-spill
    // the region on open (a bare `t="array"` reads back as a legacy
    // single-cell CSE array). The formula text always begins with '=';
    // strip it before serialisation.
    // Excel stores `ca="1"` on a formula it recalculates every time, and on a
    // dynamic-array formula whose spill is blocked; the array form adds
    // `aca="1"`.
    const bool always_calculates = formula_cell_always_calculates(sheet, row, col, cell, formula_root, name_shapes);
    if (anchored != nullptr || dynamic || cse) {
      out.append(always_calculates ? "<f t=\"array\" aca=\"1\" ref=\"" : "<f t=\"array\" ref=\"");
      out.append(a1::encode_a1(row, col));
      const std::uint32_t rows = cse ? cell.array_block_rows : anchored != nullptr ? anchored->rows : 1U;
      const std::uint32_t cols = cse ? cell.array_block_cols : anchored != nullptr ? anchored->cols : 1U;
      const std::uint32_t last_row = row + (rows > 0U ? rows - 1U : 0U);
      const std::uint32_t last_col = col + (cols > 0U ? cols - 1U : 0U);
      if (last_row != row || last_col != col) {
        out.push_back(':');
        out.append(a1::encode_a1(last_row, last_col));
      }
      out.append(always_calculates ? "\" ca=\"1\">" : "\">");
    } else if (shared != nullptr && !shared->master) {
      // A member carries no text: a reader moves the master's to this cell.
      out.append("<f t=\"shared\"");
      append_xml_attr_uint(out, "si", shared->si);
      out.append("/>");
      AppendFormulaValue(out, cell.cached_value);
      out.append("</c>");
      return true;
    } else if (shared != nullptr) {
      out.append("<f t=\"shared\" ref=\"");
      out.append(a1::encode_a1(shared->ref.first_row, shared->ref.first_col));
      if (shared->ref.last_row != shared->ref.first_row || shared->ref.last_col != shared->ref.first_col) {
        out.push_back(':');
        out.append(a1::encode_a1(shared->ref.last_row, shared->ref.last_col));
      }
      out.push_back('"');
      append_xml_attr_uint(out, "si", shared->si);
      out.push_back('>');
    } else {
      out.append(always_calculates ? "<f ca=\"1\">" : "<f>");
    }
    // Re-apply Excel's hidden storage prefixes (`_xlfn.` / `_xlfn._xlws.`
    // on the enumerated future functions, `_xlpm.` on LET / LAMBDA
    // parameters) so a real Excel reading this file resolves the modern
    // functions instead of showing #NAME?. `formula_text` was normalised
    // to the canonical formula-bar form on ingestion
    // (parser::strip_storage_prefixes). Parse it and re-serialise through the
    // storage formatter; on any parse failure fall back to the canonical
    // text unchanged.
    bool storage_emitted = false;
    if (formula_root != nullptr) {
      // A legacy formula stores no `@` where Excel's own implied one stands;
      // a CSE block intersects nothing.
      std::vector<const parser::AstNode*> implied_at;
      if (!cell.dynamic_array && !cse && anchored == nullptr) {
        implied_at = xlsb::legacy_intersections(*formula_root, name_shapes);
      }
      const std::vector<const parser::AstNode*> function_values = xlsb::function_value_refs(*formula_root, name_shapes);
      const std::string storage =
          parser::format_formula_storage(*formula_root, &storage_call_name, &implied_at, indexer, &function_values);
      // Only re-serialise when the storage form differs in substance: a
      // storage prefix, an external book's `[N]`, or a bare sheet name that
      // needs quotes. A classic formula's stored text is emitted verbatim
      // below to preserve its exact spelling.
      if (NeedsStorageSpelling(*formula_root, storage)) {
        AppendXmlEscaped(out, storage);
        storage_emitted = true;
      }
    }
    if (!storage_emitted) {
      AppendXmlEscaped(out, formula);
    }
    out.append("</f>");

    AppendFormulaValue(out, cell.cached_value);
    out.append("</c>");
    return true;
  }

  // Literal-value cell. AppendLiteralCellBody completes the opening tag
  // (closing quote + optional `s=` + optional type attribute), writes the
  // body, and closes the </c>. The cell's `phonetic_runs` are forwarded
  // so any <rPh> annotation captured at read time round-trips back
  // through the inline-string block, spans intact.
  out.append("<c r=\"");
  out.append(addr);
  AppendLiteralCellBody(out, cell.cached_value, cell.phonetic_runs, cell.phonetic_props, shared_strings, xf,
                        cell.show_phonetic, vm);
  return true;
}

// Appends OOXML-conformant `<row>` start-tag attributes derived from a
// `RowLayout` override. Caller has already emitted `<row r="N"`. Each
// override field is emitted only when it differs from the OOXML
// default (height -> none, hidden -> "0", outlineLevel -> 0); `hidden="0"`
// is never written. Excel emits `customHeight="1"` only for an explicit
// override, so an auto height stays auto across a save.
void AppendRowOverrideAttrs(std::string& out, const RowLayout& layout, const CellStyleSavePlan* style_plan) {
  if (layout.has_height || layout.height > 0.0) {
    out.append(" ht=\"");
    // Shortest round-trip spelling, the same one cell values use: a
    // recalc-save neither drifts the row metric nor respells 13.2 as
    // 13.199999999999999. Matches the column-width writer.
    append_xml_number(out, layout.height);
    out.append("\"");
    if (layout.custom_height) {
      emit_xsd_bool_attr(out, "customHeight", true);
    }
  }
  if (layout.hidden) {
    emit_xsd_bool_attr(out, "hidden", true);
  }
  if (layout.outline_level != 0U) {
    append_xml_attr_uint(out, "outlineLevel", layout.outline_level);
  }
  if (layout.collapsed) {
    emit_xsd_bool_attr(out, "collapsed", true);
  }
  if (layout.show_phonetic) {
    emit_xsd_bool_attr(out, "ph", true);
  }
  if (layout.has_style) {
    // OOXML row style is effective only with customFormat=1, which alone
    // reads as xf 0; Excel omits `s` for xf 0.
    const std::uint32_t xf = SavedXf(style_plan, layout.style_xf);
    if (xf != 0U) {
      append_xml_attr_uint(out, "s", xf);
    }
    emit_xsd_bool_attr(out, "customFormat", true);
  }
}

// Emits the <row> wrapper with all visible cells in the row. Returns true
// when at least one <c> was emitted (i.e. the <row> was actually written),
// false when the row collapsed to nothing (every cell was blank).
// `phantom_cols` lists the row's spill-covered columns other than anchors,
// ascending; a phantom wins over a stored entry in the same column. When
// `override_attrs` is non-empty it is appended to the `<row>` start-tag (between `r="N"` and the closing `>`), allowing
// the caller to merge per-row layout overrides without reshaping the body. When the row body collapses to nothing but
// `override_attrs` is non-empty, an empty self-closing `<row r="N" .../>` is still emitted so the override survives the
// round-trip. `keep_empty_row` forces that bare element even with no attributes, for a visible row under `zeroHeight`;
// `implicit_hidden` adds `hidden="1"` to a written row that has no override under `zeroHeight`.
bool AppendRowXml(std::string& out, const Sheet& sheet, std::uint32_t row, const RowCells& row_cells,
                  const std::vector<EmittedRun>& runs, const std::vector<std::uint32_t>& phantom_cols,
                  std::string_view override_attrs, bool keep_empty_row, bool implicit_hidden,
                  const SharedStrings* shared_strings, std::uint32_t dynamic_array_cm_index,
                  const xlsb::NameShapes& name_shapes, std::unordered_map<std::uint64_t, bool>& always_by_anchor,
                  const parser::ExternalBookIndexer* indexer, const SharedRoles& shared_roles,
                  const RichErrorVmMap* rich_error_vm, const CellStyleSavePlan* style_plan,
                  MergedCellIndex& merged_cells) {
  // Write straight into `out`; a row that ends up with no content is truncated away.
  const std::size_t row_start = out.size();
  out.append("<row r=\"");
  out.append(std::to_string(row + 1U));
  out.push_back('"');
  if (!override_attrs.empty()) {
    out.append(override_attrs);
  }
  if (implicit_hidden) {
    emit_xsd_bool_attr(out, "hidden", true);
  }
  out.push_back('>');
  const std::size_t body_start = out.size();
  auto emit_phantom = [&](std::uint32_t col, const Cell* stored) {
    const SpillRegion* region = sheet.spill_region_covering(row, col);
    if (region == nullptr) {
      return;
    }
    const std::uint64_t key = (static_cast<std::uint64_t>(region->anchor_row) << 32) | region->anchor_col;
    auto it = always_by_anchor.find(key);
    if (it == always_by_anchor.end()) {
      const Cell* anchor = sheet.cell_at(region->anchor_row, region->anchor_col);
      const bool always = anchor != nullptr && formula_cell_always_calculates(sheet, region->anchor_row,
                                                                              region->anchor_col, *anchor, name_shapes);
      it = always_by_anchor.emplace(key, always).first;
    }
    const std::size_t idx =
        static_cast<std::size_t>(row - region->anchor_row) * region->cols + (col - region->anchor_col);
    const Value value = idx < region->cells.size() ? region->cells[idx] : Value::blank();
    const bool force_style = PhantomForcesDefaultStyle(sheet, row, col, stored);
    AppendPhantomCellXml(out, row, col, SavedXf(style_plan, select_effective_xf(sheet, row, col).xf_index), force_style,
                         value, it->second, VmAt(rich_error_vm, row, col));
  };
  // Blank formatted columns, written as `<c r="..." s="N"/>` (bare for xf 0).
  auto run = runs.begin();
  std::uint64_t run_col = run != runs.end() ? run->first : UINT64_MAX;
  const auto advance_run_past = [&](std::uint64_t col) {
    while (run != runs.end() && run_col <= col) {
      run_col = col + 1U;
      if (run_col >= run->end) {
        ++run;
        run_col = run != runs.end() ? run->first : UINT64_MAX;
      }
    }
  };
  // Three-way merge in column order; a phantom wins over a stored cell and
  // over a run column, and a stored cell and a run never share a column.
  auto entry = row_cells.begin();
  const auto entries_end = row_cells.end();
  auto phantom = phantom_cols.begin();
  while (entry != entries_end || phantom != phantom_cols.end() || run != runs.end()) {
    const std::uint32_t entry_col = entry != entries_end ? (*entry).col : UINT32_MAX;
    const std::uint32_t phantom_col = phantom != phantom_cols.end() ? *phantom : UINT32_MAX;
    if (run_col < entry_col && run_col < phantom_col) {
      // Run columns up to the next stored cell or phantom.
      const std::uint64_t stop = std::min<std::uint64_t>({run->end, entry_col, phantom_col});
      for (std::uint64_t col = run_col; col < stop; ++col) {
        out.append("<c r=\"");
        AppendCellRefForRef(out, row, static_cast<std::uint32_t>(col));
        out.push_back('"');
        AppendStyleAttr(out, SavedXf(style_plan, run->xf, merged_cells.covers(row, static_cast<std::uint32_t>(col))),
                        false);
        out.append("/>");
      }
      advance_run_past(stop - 1U);
      continue;
    }
    if (phantom_col <= entry_col) {
      emit_phantom(phantom_col, phantom_col == entry_col ? &(*entry).cell : nullptr);
      if (phantom_col == entry_col) {
        ++entry;
      }
      advance_run_past(phantom_col);
      ++phantom;
      continue;
    }
    const RowCells::Entry e = *entry;
    const auto role = shared_roles.find(CellKey(row, e.col));
    (void)AppendCellXml(out, sheet, row, e.col, e.cell, shared_strings, dynamic_array_cm_index, name_shapes, indexer,
                        role != shared_roles.end() ? &role->second : nullptr, VmAt(rich_error_vm, row, e.col),
                        style_plan, merged_cells.covers(row, e.col));
    ++entry;
  }
  if (out.size() == body_start) {
    if (override_attrs.empty() && !keep_empty_row) {
      out.resize(row_start);
      return false;
    }
    out.pop_back();
    out.append("/>");
    return true;
  }
  out.append("</row>");
  return true;
}

}  // namespace

bool CellIsEmitted(const Sheet& sheet, std::uint32_t row, std::uint32_t col, const Cell& cell) {
  if (!cell.formula_text.empty() || !cell.cached_value.is_blank() || cell.xf_index != 0U || cell.show_phonetic) {
    return true;
  }
  return cell.has_explicit_xf &&
         !explicit_default_observable(sheet.layout(), row, col, static_cast<std::uint64_t>(col) + 1U).empty();
}

std::optional<SheetDimension> sheet_dimension(const Sheet& sheet) {
  using Span = std::optional<std::pair<std::uint32_t, std::uint32_t>>;
  const auto widen = [](Span& span, std::uint32_t lo, std::uint32_t hi) {
    span = span ? std::make_pair(std::min(span->first, lo), std::max(span->second, hi)) : std::make_pair(lo, hi);
  };
  Span rows;
  Span cols;
  std::unordered_set<std::uint32_t> phantom_rows;
  for (const CellAddress& phantom : sheet.spill_phantom_addresses()) {
    widen(rows, phantom.row, phantom.row);
    widen(cols, phantom.col, phantom.col);
    phantom_rows.insert(phantom.row);
  }
  for (const auto& [row, cells] : sheet.rows()) {
    bool written = false;
    for (const auto& [col, cell] : cells) {
      if (CellIsEmitted(sheet, row, col, cell)) {
        widen(cols, col, col);
        written = true;
      }
    }
    const RowEmitPlan plan = row_emit_plan(sheet, row, cells, phantom_rows.count(row) != 0U);
    if (!plan.runs.empty()) {
      widen(cols, plan.runs.front().first, static_cast<std::uint32_t>(plan.runs.back().end - 1U));
      written = true;
    }
    if (written || plan.as_row_format) {
      widen(rows, row, row);
    }
  }
  const SheetLayout& layout = sheet.layout();
  for (const RowLayout& entry : layout.row_overrides) {
    if (entry.row < Sheet::kMaxRows) {
      widen(rows, entry.row, entry.row);
    }
  }
  for (const ColumnLayout& span : layout.columns) {
    if (span.has_style && span.style_xf != 0U) {
      widen(cols, span.first, span.last);
    }
  }
  if (!rows && !cols) {
    return std::nullopt;
  }
  const std::pair<std::uint32_t, std::uint32_t> row_span = rows.value_or(std::make_pair(0U, 0U));
  const std::pair<std::uint32_t, std::uint32_t> col_span = cols.value_or(std::make_pair(0U, 0U));
  return SheetDimension{row_span.first, col_span.first, row_span.second, col_span.second};
}

void AppendSheetDataXml(std::string& out, const Sheet& sheet, const SharedStrings* shared_strings,
                        std::uint32_t dynamic_array_cm_index, const xlsb::NameShapes& name_shapes,
                        const parser::ExternalBookIndexer* indexer, const RichErrorVmMap* rich_error_vm,
                        const CellStyleSavePlan* style_plan) {
  // Collect populated row indices and sort ascending so the output is
  // deterministic regardless of unordered_map iteration order.
  const auto& rows_map = sheet.rows();
  const auto& row_overrides = sheet.layout().row_overrides;

  // Spilled cells live in the spill table, not in the rows; gather their
  // columns per row so each is written where Excel writes it.
  std::unordered_map<std::uint32_t, std::vector<std::uint32_t>> phantom_cols;
  for (const CellAddress& phantom : sheet.spill_phantom_addresses()) {
    phantom_cols[phantom.row].push_back(phantom.col);
  }
  for (auto& [row, cols] : phantom_cols) {
    (void)row;
    sort_ascending(cols);
  }

  // Union of populated, spilled and override rows. Rows that have only an
  // override (no cells) still need to surface so the override survives
  // a save/load round-trip.
  std::vector<std::uint32_t> row_indices;
  row_indices.reserve(rows_map.size() + phantom_cols.size() + row_overrides.size());
  for (const auto& kv : rows_map) {
    row_indices.push_back(kv.first);
  }
  for (const auto& kv : phantom_cols) {
    row_indices.push_back(kv.first);
  }
  for (const RowLayout& ro : row_overrides) {
    if (rows_map.find(ro.row) == rows_map.end()) {
      row_indices.push_back(ro.row);
    }
  }
  sort_ascending(row_indices);
  row_indices.erase(std::unique(row_indices.begin(), row_indices.end()), row_indices.end());

  const std::size_t sheet_data_start = out.size();
  out.append("<sheetData>");
  const std::size_t rows_start = out.size();
  // Sentinel empty span used when a row has no override; avoids a
  // per-row default-construction of std::string.
  static const RowCells kEmptyRow;
  static const std::vector<std::uint32_t> kNoPhantoms;
  std::unordered_map<std::uint64_t, bool> always_by_anchor;
  const SharedRoles shared_roles = CollectSharedRoles(sheet, name_shapes);
  MergedCellIndex merged_cells(style_plan != nullptr ? sheet.merges() : std::vector<MergeRange>{});
  for (std::uint32_t row : row_indices) {
    std::string override_attrs;
    const bool zero_height = sheet.format_defaults().zero_height;
    const RowLayout* row_layout = find_row_override(sheet.layout(), row);
    if (row_layout != nullptr) {
      AppendRowOverrideAttrs(override_attrs, *row_layout, style_plan);
    }
    // Under zeroHeight a present row is visible unless it says otherwise; an entry-less row is hidden.
    const bool keep_empty_row = zero_height && row_layout != nullptr && !row_layout->hidden;
    bool implicit_hidden = zero_height && row_layout == nullptr;
    auto cells_it = rows_map.find(row);
    const RowCells& row_cells = (cells_it != rows_map.end()) ? cells_it->second : kEmptyRow;
    const auto phantoms_it = phantom_cols.find(row);
    const RowEmitPlan plan = row_emit_plan(sheet, row, row_cells, phantoms_it != phantom_cols.end());
    if (plan.as_row_format) {
      // The row's explicit defaults become its own xf-0 row format.
      RowLayout converted;
      if (row_layout != nullptr) {
        converted = *row_layout;
      } else {
        converted.hidden = zero_height;
      }
      converted.row = row;
      converted.has_style = true;
      converted.style_xf = 0U;
      override_attrs.clear();
      AppendRowOverrideAttrs(override_attrs, converted, style_plan);
      implicit_hidden = false;
    }
    AppendRowXml(out, sheet, row, row_cells, plan.runs,
                 phantoms_it != phantom_cols.end() ? phantoms_it->second : kNoPhantoms, override_attrs, keep_empty_row,
                 implicit_hidden, shared_strings, dynamic_array_cm_index, name_shapes, always_by_anchor, indexer,
                 shared_roles, rich_error_vm, style_plan, merged_cells);
  }

  if (out.size() == rows_start) {
    // Self-closing form keeps the empty-sheet output identical to the
    // pre-cell-writer skeleton, which downstream tests already pin.
    out.resize(sheet_data_start);
    out.append("<sheetData/>");
    return;
  }
  out.append("</sheetData>");
}

}  // namespace io
}  // namespace formulon
