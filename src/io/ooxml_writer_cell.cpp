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
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include "cell.h"
#include "io/dynamic_array_formula.h"
#include "io/future_functions.h"
#include "io/ooxml/external_link_writer.h"
#include "io/ooxml/shared_strings_writer.h"
#include "io/phonetic_pr.h"
#include "io/stored_cell_error.h"
#include "io/xlsb/ptg_writer.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/parser.h"
#include "phonetic.h"
#include "sheet.h"
#include "style_resolve.h"
#include "utils/a1_column.h"
#include "utils/a1_ref.h"
#include "utils/arena.h"
#include "utils/index_sort.h"
#include "value.h"

namespace formulon {
namespace io {
namespace {

// ---------------------------------------------------------------------------
// Cell emission
// ---------------------------------------------------------------------------

// True when a cell's explicit default style must be written as `s="0"`: only a
// blank, formula-less cell does not own its style, so only there does the
// attribute carry meaning (it blocks row/column style inheritance).
bool ForcesDefaultStyle(const Cell& cell) {
  return cell.has_explicit_xf && cell.formula_text.empty() && cell.cached_value.is_blank();
}

// True when a row or column style would apply to `(row, col)` were its own
// style not written; an explicit default style is only observable then.
bool InheritsNonDefaultStyle(const Sheet& sheet, std::uint32_t row, std::uint32_t col) {
  const SheetLayout& layout = sheet.layout();
  for (const RowLayout& r : layout.row_overrides) {
    if (r.row == row && r.has_style && r.style_xf != 0U) {
      return true;
    }
  }
  for (const ColumnLayout& c : layout.columns) {
    if (c.first <= col && col <= c.last && c.has_style && c.style_xf != 0U) {
      return true;
    }
  }
  return false;
}

// Emits an `s="N"` attribute when `xf_index` is non-zero. The default
// xf (index 0) is omitted for byte parity with Excel's writer unless `force`
// is set (see `ForcesDefaultStyle`).
void AppendStyleAttr(std::string& out, std::uint32_t xf_index, bool force) {
  if (xf_index == 0U && !force) {
    return;
  }
  append_xml_attr_uint(out, "s", xf_index);
}

// Emits the <c> element for an Error value at `addr`.
void AppendErrorCellXml(std::string& out, std::string_view addr, ErrorCode code, std::uint32_t xf_index,
                        bool force_style) {
  out.append("<c r=\"");
  out.append(addr);
  out.append("\"");
  AppendStyleAttr(out, xf_index, force_style);
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
                           std::uint32_t xf_index, bool force_style) {
  out.push_back('"');
  AppendStyleAttr(out, xf_index, force_style);
  if (value.is_number()) {
    // Defensive: NaN / +/-Inf must never reach append_xml_number, which
    // would emit `nan`/`inf` text inside `<v>` (Excel rejects this on
    // load). The caller AppendCellXml pre-screens, but a future caller
    // path may not — downgrade to #NUM! here as a last line of defence.
    const double v = value.as_number();
    if (!std::isfinite(v)) {
      out.append(" t=\"e\"><v>");
      out.append(display_name(ErrorCode::Num));
      out.append("</v></c>");
      return;
    }
    out.append("><v>");
    append_xml_number(out, v);
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
    out.append(" t=\"e\"><v>");
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

// Emits the <c> element for `(row, col)` into `out`. Returns true when a
// cell was written, false when the cell was suppressed (blank literal,
// phantom of a spill region).
//
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
// recalculated every time (measured).
void AppendPhantomCellXml(std::string& out, std::uint32_t row, std::uint32_t col, std::uint32_t xf_index,
                          bool force_style, const Value& value, bool always_calculates) {
  out.append("<c r=\"");
  out.append(a1::encode_a1(row, col));
  out.append("\"");
  AppendStyleAttr(out, xf_index, force_style);
  AppendFormulaValueType(out, value);
  out.push_back('>');
  if (always_calculates) {
    out.append("<f ca=\"1\"/>");
  }
  AppendFormulaValue(out, value);
  out.append("</c>");
}

// Spill anchor handling: if `(row, col)` is anchored, the formula is
// emitted with t="array" and the cached value (cells[0]) becomes the
// <v>. Phantoms are suppressed by the caller via spill_region_covering;
// this function trusts the caller and never re-checks.
bool AppendCellXml(std::string& out, const Sheet& sheet, std::uint32_t row, std::uint32_t col, const Cell& cell,
                   const SharedStrings* shared_strings, std::uint32_t dynamic_array_cm_index,
                   const xlsb::NameShapes& name_shapes, const parser::ExternalBookIndexer* indexer) {
  const bool has_formula = !cell.formula_text.empty();
  if (!CellIsEmitted(cell)) {
    return false;
  }

  // Pre-screen NaN/Inf number literals so we never half-emit a cell tag.
  // Formula cells with non-finite cached values still emit the <f>; only
  // the <v> is downgraded.
  if (!has_formula && cell.cached_value.is_number() && !std::isfinite(cell.cached_value.as_number())) {
    const std::string addr = a1::encode_a1(row, col);
    AppendErrorCellXml(out, addr, ErrorCode::Num, cell.xf_index, ForcesDefaultStyle(cell));
    return true;
  }

  const std::string addr = a1::encode_a1(row, col);

  // Style-only cells (blank value, formatting attached) round-trip as a
  // bare `<c r="..." s="N"/>` shape — Excel preserves these so empty
  // formatted cells keep their visual.
  if (!has_formula && cell.cached_value.is_blank()) {
    out.append("<c r=\"");
    out.append(addr);
    out.append("\"");
    AppendStyleAttr(out, cell.xf_index, ForcesDefaultStyle(cell));
    out.append("/>");
    return true;
  }

  if (has_formula) {
    const SpillRegion* anchored = sheet.spill_region_at_anchor(row, col);
    std::string_view formula = cell.formula_text;
    if (!formula.empty() && formula.front() == '=') {
      formula.remove_prefix(1);
    }
    Arena formula_arena;
    parser::Parser formula_parser(formula, formula_arena);
    parser::AstNode* formula_root = formula_parser.parse();
    if (formula_root != nullptr && !formula_parser.errors().empty()) {
      formula_root = nullptr;
    }
    // A dynamic-array formula keeps that form when it does not spill, or
    // Excel reads it with implicit intersection (`=@...`); it needs the
    // XLDAPR entry `cm=` names, so without one the plain form stays. A spill
    // anchor that is not one (a loaded CSE block) keeps `t="array"` alone.
    const bool dynamic = dynamic_array_cm_index != 0U && is_dynamic_array_formula(cell);
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
    AppendStyleAttr(out, cell.xf_index, ForcesDefaultStyle(cell));
    AppendFormulaValueType(out, cell.cached_value);
    if (dynamic) {
      append_xml_attr_uint(out, "cm", dynamic_array_cm_index);
    }
    out.push_back('>');

    // <f> with optional t="array" ref="..." for spill anchors. Modern
    // Excel marks a dynamic-array anchor with `t="array"` plus a `ref`
    // covering the spill footprint; the `ref` is what lets Excel re-spill
    // the region on open (a bare `t="array"` reads back as a legacy
    // single-cell CSE array). The formula text always begins with '=';
    // strip it before serialisation.
    // Excel stores `ca="1"` on a formula it recalculates every time, and on a
    // dynamic-array formula whose spill is blocked; the array form adds
    // `aca="1"` (measured).
    const bool always_calculates = formula_cell_always_calculates(sheet, row, col, cell, name_shapes);
    if (anchored != nullptr || dynamic) {
      out.append(always_calculates ? "<f t=\"array\" aca=\"1\" ref=\"" : "<f t=\"array\" ref=\"");
      out.append(a1::encode_a1(row, col));
      const std::uint32_t rows = anchored != nullptr ? anchored->rows : 1U;
      const std::uint32_t cols = anchored != nullptr ? anchored->cols : 1U;
      const std::uint32_t last_row = row + (rows > 0U ? rows - 1U : 0U);
      const std::uint32_t last_col = col + (cols > 0U ? cols - 1U : 0U);
      if (last_row != row || last_col != col) {
        out.push_back(':');
        out.append(a1::encode_a1(last_row, last_col));
      }
      out.append(always_calculates ? "\" ca=\"1\">" : "\">");
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
      if (!cell.dynamic_array && anchored == nullptr) {
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
  AppendLiteralCellBody(out, cell.cached_value, cell.phonetic_runs, cell.phonetic_props, shared_strings, cell.xf_index,
                        ForcesDefaultStyle(cell));
  return true;
}

// Appends OOXML-conformant `<row>` start-tag attributes derived from a
// `RowLayout` override. Caller has already emitted `<row r="N"`. Each
// override field is emitted only when it differs from the OOXML
// default (height -> none, hidden -> "0", outlineLevel -> 0). Excel
// emits `customHeight="1"` only for an explicit override, so an auto
// height stays auto across a save.
void AppendRowOverrideAttrs(std::string& out, const RowLayout& layout) {
  if (layout.has_height || layout.height > 0.0) {
    out.append(" ht=\"");
    // Shortest round-trip spelling, the same one cell values use: a
    // recalc-save neither drifts the row metric nor respells 13.2 as
    // 13.199999999999999. Matches the column-width writer.
    append_xml_number(out, layout.height);
    out.append("\"");
    if (layout.custom_height) {
      out.append(" customHeight=\"1\"");
    }
  }
  if (layout.hidden) {
    out.append(" hidden=\"1\"");
  }
  if (layout.outline_level != 0U) {
    append_xml_attr_uint(out, "outlineLevel", layout.outline_level);
  }
  if (layout.has_style) {
    // OOXML row style is effective only with customFormat=1. Emit s even
    // when the explicit style xf is zero so style="0" survives a save.
    out.append(" s=\"");
    out.append(std::to_string(layout.style_xf));
    out.append("\" customFormat=\"1\"");
  }
}

// Emits the <row> wrapper with all visible cells in the row. Returns true
// when at least one <c> was emitted (i.e. the <row> was actually written),
// false when the row collapsed to nothing (every cell was blank).
// `phantom_cols` lists the row's spilled columns, ascending. When `override_attrs` is non-empty it is appended to the
// `<row>` start-tag (between `r="N"` and the closing `>`), allowing the
// caller to merge per-row layout overrides without reshaping the body.
// When the row body collapses to nothing but `override_attrs` is
// non-empty, an empty self-closing `<row r="N" .../>` is still emitted
// so the override survives the round-trip.
bool AppendRowXml(std::string& out, const Sheet& sheet, std::uint32_t row, const RowCells& row_cells,
                  const std::vector<std::uint32_t>& phantom_cols, std::string_view override_attrs,
                  const SharedStrings* shared_strings, std::uint32_t dynamic_array_cm_index,
                  const xlsb::NameShapes& name_shapes, std::unordered_map<std::uint64_t, bool>& always_by_anchor,
                  const parser::ExternalBookIndexer* indexer) {
  // Buffer the row body separately so we can tell whether anything ended
  // up inside the <row> wrapper before we commit to writing it.
  std::string body;
  body.reserve(row_cells.size() * 24U);
  const std::uint32_t stored = static_cast<std::uint32_t>(row_cells.size());
  auto emit_phantom = [&](std::uint32_t col) {
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
    const bool force_style =
        col < stored && ForcesDefaultStyle(row_cells[col]) && InheritsNonDefaultStyle(sheet, row, col);
    AppendPhantomCellXml(body, row, col, select_effective_xf(sheet, row, col).xf_index, force_style, value, it->second);
  };
  for (std::uint32_t col = 0; col < stored; ++col) {
    if (sheet.spill_region_covering(row, col) != nullptr) {
      emit_phantom(col);
      continue;
    }
    (void)AppendCellXml(body, sheet, row, col, row_cells[col], shared_strings, dynamic_array_cm_index, name_shapes,
                        indexer);
  }
  for (const std::uint32_t col : phantom_cols) {
    if (col >= stored) {
      emit_phantom(col);
    }
  }
  if (body.empty() && override_attrs.empty()) {
    return false;
  }
  out.append("<row r=\"");
  out.append(std::to_string(row + 1U));
  out.push_back('"');
  if (!override_attrs.empty()) {
    out.append(override_attrs);
  }
  if (body.empty()) {
    out.append("/>");
    return true;
  }
  out.push_back('>');
  out.append(body);
  out.append("</row>");
  return true;
}

}  // namespace

bool CellIsEmitted(const Cell& cell) {
  return !cell.formula_text.empty() || !cell.cached_value.is_blank() || cell.xf_index != 0U || cell.has_explicit_xf;
}

std::string BuildSheetDataXml(const Sheet& sheet, const SharedStrings* shared_strings,
                              std::uint32_t dynamic_array_cm_index, const xlsb::NameShapes& name_shapes,
                              const parser::ExternalBookIndexer* indexer) {
  // Collect populated row indices and sort ascending so the output is
  // deterministic regardless of unordered_map iteration order.
  const auto& rows_map = sheet.rows();
  // Index the per-row overrides by row index. The vector is small in
  // practice (rarely more than a few dozen entries even for hand-crafted
  // sheets), so a flat scan would also be fine; the map keeps the merge
  // O(rows + overrides).
  std::unordered_map<std::uint32_t, const RowLayout*> overrides_by_row;
  const auto& row_overrides = sheet.layout().row_overrides;
  overrides_by_row.reserve(row_overrides.size());
  for (const RowLayout& ro : row_overrides) {
    overrides_by_row.emplace(ro.row, &ro);
  }

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

  std::string body;
  body.reserve((rows_map.size() + row_overrides.size()) * 64U);
  // Sentinel empty span used when a row has no override; avoids a
  // per-row default-construction of std::string.
  static const RowCells kEmptyRow;
  static const std::vector<std::uint32_t> kNoPhantoms;
  std::unordered_map<std::uint64_t, bool> always_by_anchor;
  for (std::uint32_t row : row_indices) {
    std::string override_attrs;
    auto override_it = overrides_by_row.find(row);
    if (override_it != overrides_by_row.end()) {
      AppendRowOverrideAttrs(override_attrs, *override_it->second);
    }
    auto cells_it = rows_map.find(row);
    const RowCells& row_cells = (cells_it != rows_map.end()) ? cells_it->second : kEmptyRow;
    const auto phantoms_it = phantom_cols.find(row);
    AppendRowXml(body, sheet, row, row_cells, phantoms_it != phantom_cols.end() ? phantoms_it->second : kNoPhantoms,
                 override_attrs, shared_strings, dynamic_array_cm_index, name_shapes, always_by_anchor, indexer);
  }

  if (body.empty()) {
    // Self-closing form keeps the empty-sheet output identical to the
    // pre-cell-writer skeleton, which downstream tests already pin.
    return "<sheetData/>";
  }
  std::string out;
  out.reserve(body.size() + 24U);
  out.append("<sheetData>");
  out.append(body);
  out.append("</sheetData>");
  return out;
}

}  // namespace io
}  // namespace formulon
