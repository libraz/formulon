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
//     itself): suppress the <c> entirely. Excel reconstructs phantoms by
//     re-spilling the anchor on load.
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
#include "eval/utf8_length.h"
#include "io/future_functions.h"
#include "io/ooxml/shared_strings_writer.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/a1_column.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace io {
namespace {

// ---------------------------------------------------------------------------
// Cell emission
// ---------------------------------------------------------------------------

// Emits an `s="N"` attribute when `xf_index` is non-zero. The default
// xf (index 0) is omitted for byte parity with Excel's writer, which
// emits `<c>` without `s=` for the default-formatted majority of cells.
void AppendStyleAttr(std::string& out, std::uint32_t xf_index) {
  if (xf_index == 0U) {
    return;
  }
  out.append(" s=\"");
  out.append(std::to_string(xf_index));
  out.append("\"");
}

// True for errors representable in the legacy `t="e"` enum, i.e. writable
// as a bare `<v>#...#</v>`. The rich errors (#SPILL!, #CALC!, linked-data,
// Python, ...) require a `vm=` value-metadata attribute plus a metadata
// part; Excel rejects the whole workbook when such a code is written as a
// plain `<v>`. For those we drop the cached value and let Excel recompute
// from the formula (or emit a blank cell when there is no formula).
bool IsLegacyErrorCode(ErrorCode code) noexcept {
  switch (code) {
    case ErrorCode::Null:
    case ErrorCode::Div0:
    case ErrorCode::Value:
    case ErrorCode::Ref:
    case ErrorCode::Name:
    case ErrorCode::Num:
    case ErrorCode::NA:
    case ErrorCode::GettingData:
      return true;
    default:
      return false;
  }
}

// Emits the <c> element for an Error value at `addr`.
void AppendErrorCellXml(std::string& out, std::string_view addr, ErrorCode code, std::uint32_t xf_index) {
  out.append("<c r=\"");
  out.append(addr);
  out.append("\"");
  AppendStyleAttr(out, xf_index);
  if (!IsLegacyErrorCode(code)) {
    // Rich error with no formula: not writable as a legacy <v>; emit a
    // blank cell (keeping any style) so the workbook still opens.
    out.append("/>");
    return;
  }
  out.append(" t=\"e\"><v>");
  out.append(display_name(code));
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
// `<is><t>{text}</t><rPh sb="0" eb="N"><t>{kana}</t></rPh></is>` where
// `N` is the UTF-16 code-unit count of `value.as_text()`. The original
// document may have carried multiple `<rPh>` blocks (one per kanji
// span); we collapse them to a single block on round-trip because
// PHONETIC()'s observable behaviour (the concatenated kana) is
// preserved without per-span boundaries.
void AppendLiteralCellBody(std::string& out, const Value& value, std::string_view phonetic,
                           const SharedStrings* shared_strings, std::uint32_t xf_index) {
  out.push_back('"');
  AppendStyleAttr(out, xf_index);
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
      out.append(std::to_string(shared_strings->index_of(value.as_text(), phonetic)));
      out.append("</v></c>");
      return;
    }
    out.append(" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
    AppendXmlEscaped(out, value.as_text());
    out.append("</t>");
    if (!phonetic.empty()) {
      const std::uint32_t eb = formulon::eval::utf16_units_in(value.as_text());
      out.append("<rPh sb=\"0\" eb=\"");
      out.append(std::to_string(eb));
      out.append("\"><t xml:space=\"preserve\">");
      AppendXmlEscaped(out, phonetic);
      out.append("</t></rPh>");
    }
    out.append("</is></c>");
    return;
  }
  if (value.is_error()) {
    if (!IsLegacyErrorCode(value.as_error())) {
      // Rich error literal: not writable as a legacy <v>; emit a blank
      // cell, keeping any style already appended above.
      out.append("/>");
      return;
    }
    out.append(" t=\"e\"><v>");
    out.append(display_name(value.as_error()));
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
// Spill anchor handling: if `(row, col)` is anchored, the formula is
// emitted with t="array" and the cached value (cells[0]) becomes the
// <v>. Phantoms are suppressed by the caller via spill_region_covering;
// this function trusts the caller and never re-checks.
bool AppendCellXml(std::string& out, const Sheet& sheet, std::uint32_t row, std::uint32_t col, const Cell& cell,
                   const SharedStrings* shared_strings) {
  const bool has_formula = !cell.formula_text.empty();
  if (!CellIsEmitted(cell)) {
    return false;
  }

  // Pre-screen NaN/Inf number literals so we never half-emit a cell tag.
  // Formula cells with non-finite cached values still emit the <f>; only
  // the <v> is downgraded.
  if (!has_formula && cell.cached_value.is_number() && !std::isfinite(cell.cached_value.as_number())) {
    const std::string addr = EncodeA1(row, col);
    AppendErrorCellXml(out, addr, ErrorCode::Num, cell.xf_index);
    return true;
  }

  const std::string addr = EncodeA1(row, col);

  // Style-only cells (blank value, formatting attached) round-trip as a
  // bare `<c r="..." s="N"/>` shape — Excel preserves these so empty
  // formatted cells keep their visual.
  if (!has_formula && cell.cached_value.is_blank()) {
    out.append("<c r=\"");
    out.append(addr);
    out.append("\"");
    AppendStyleAttr(out, cell.xf_index);
    out.append("/>");
    return true;
  }

  if (has_formula) {
    const SpillRegion* anchored = sheet.spill_region_at_anchor(row, col);
    out.append("<c r=\"");
    out.append(addr);
    out.append("\"");
    AppendStyleAttr(out, cell.xf_index);

    // The `t=` attribute on the formula <c> must agree with the cached
    // <v> body so a save/load round-trip preserves the value's type.
    // The default `n` (number) is omitted; everything else is named so
    // cell_parser does not try to parse a non-numeric body as a double.
    // Non-finite numbers are downgraded to t="e" because the <v> below
    // emits #NUM! for that branch.
    const Value& cached = cell.cached_value;
    if (cached.is_error()) {
      // Only tag t="e" when the cached value will actually be written as a
      // legacy <v> below; a rich error emits no <v>, so the cell has no
      // typed body.
      if (IsLegacyErrorCode(cached.as_error())) {
        out.append(" t=\"e\">");
      } else {
        out.append(">");
      }
    } else if (cached.is_text()) {
      out.append(" t=\"str\">");
    } else if (cached.is_boolean()) {
      out.append(" t=\"b\">");
    } else if (cached.is_number() && !std::isfinite(cached.as_number())) {
      out.append(" t=\"e\">");
    } else {
      out.append(">");
    }

    // <f> with optional t="array" ref="..." for spill anchors. Modern
    // Excel marks a dynamic-array anchor with `t="array"` plus a `ref`
    // covering the spill footprint; the `ref` is what lets Excel re-spill
    // the region on open (a bare `t="array"` reads back as a legacy
    // single-cell CSE array). The formula text always begins with '=';
    // strip it before serialisation.
    if (anchored != nullptr) {
      out.append("<f t=\"array\" ref=\"");
      out.append(EncodeA1(row, col));
      const std::uint32_t last_row = row + (anchored->rows > 0U ? anchored->rows - 1U : 0U);
      const std::uint32_t last_col = col + (anchored->cols > 0U ? anchored->cols - 1U : 0U);
      if (last_row != row || last_col != col) {
        out.push_back(':');
        out.append(EncodeA1(last_row, last_col));
      }
      out.append("\">");
    } else {
      out.append("<f>");
    }
    std::string_view formula = cell.formula_text;
    if (!formula.empty() && formula.front() == '=') {
      formula.remove_prefix(1);
    }
    // Re-apply Excel's hidden storage prefixes (`_xlfn.` / `_xlfn._xlws.`
    // on the enumerated future functions, `_xlpm.` on LET / LAMBDA
    // parameters) so a real Excel reading this file resolves the modern
    // functions instead of showing #NAME?. `formula_text` was normalised
    // to the canonical formula-bar form on ingestion
    // (io::strip_storage_prefixes). Parse it and re-serialise through the
    // storage formatter; on any parse failure fall back to the canonical
    // text unchanged.
    Arena formula_arena;
    parser::Parser formula_parser(formula, formula_arena);
    parser::AstNode* formula_root = formula_parser.parse();
    bool storage_emitted = false;
    if (formula_root != nullptr && formula_parser.errors().empty()) {
      const std::string storage = parser::format_formula_storage(*formula_root, &storage_call_name);
      // Only re-serialise when a storage prefix was actually added (the
      // formula uses a future function or LET / LAMBDA). For a classic
      // formula the storage form equals the plain form, so the stored text
      // is emitted verbatim below to preserve its exact spelling.
      if (storage != parser::format_formula(*formula_root)) {
        AppendXmlEscaped(out, storage);
        storage_emitted = true;
      }
    }
    if (!storage_emitted) {
      AppendXmlEscaped(out, formula);
    }
    out.append("</f>");

    // <v>: omit when blank (Excel will recalculate on load); downgrade
    // NaN/Inf number to #NUM! text inside <v>; otherwise emit normally.
    const Value& cv = cell.cached_value;
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
      // Legacy errors round-trip as a cached <v>; rich errors (#SPILL! /
      // #CALC! / ...) are not writable there — omit the <v> and let Excel
      // recompute from the formula (writing them would make Excel reject
      // the whole workbook).
      if (IsLegacyErrorCode(cv.as_error())) {
        out.append("<v>");
        out.append(display_name(cv.as_error()));
        out.append("</v>");
      }
    }
    // Array / Ref / Lambda cached values fall through with no <v>; the
    // engine evaluates on load.

    out.append("</c>");
    return true;
  }

  // Literal-value cell. AppendLiteralCellBody completes the opening tag
  // (closing quote + optional `s=` + optional type attribute), writes the
  // body, and closes the </c>. The cell's `phonetic_text` is forwarded so
  // any <rPh> annotation captured at read time round-trips back through
  // the inline-string block.
  out.append("<c r=\"");
  out.append(addr);
  AppendLiteralCellBody(out, cell.cached_value, cell.phonetic_text, shared_strings, cell.xf_index);
  return true;
}

// Appends OOXML-conformant `<row>` start-tag attributes derived from a
// `RowLayout` override. Caller has already emitted `<row r="N"`. Each
// override field is emitted only when it differs from the OOXML
// default (height -> none, hidden -> "0", outlineLevel -> 0). Excel
// itself emits `customHeight="1"` alongside `ht`; we mirror that so a
// reload reproduces the visual size.
void AppendRowOverrideAttrs(std::string& out, const RowLayout& layout) {
  if (layout.has_height || layout.height > 0.0) {
    out.append(" ht=\"");
    // Shortest round-trip spelling, the same one cell values use: a
    // recalc-save neither drifts the row metric nor respells 13.2 as
    // 13.199999999999999. Matches the column-width writer.
    append_xml_number(out, layout.height);
    out.append("\" customHeight=\"1\"");
  }
  if (layout.hidden) {
    out.append(" hidden=\"1\"");
  }
  if (layout.outline_level != 0U) {
    out.append(" outlineLevel=\"");
    out.append(std::to_string(static_cast<unsigned int>(layout.outline_level)));
    out.push_back('"');
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
// false when the row collapsed to nothing (every cell was blank or a
// phantom). When `override_attrs` is non-empty it is appended to the
// `<row>` start-tag (between `r="N"` and the closing `>`), allowing the
// caller to merge per-row layout overrides without reshaping the body.
// When the row body collapses to nothing but `override_attrs` is
// non-empty, an empty self-closing `<row r="N" .../>` is still emitted
// so the override survives the round-trip.
bool AppendRowXml(std::string& out, const Sheet& sheet, std::uint32_t row, const RowCells& row_cells,
                  std::string_view override_attrs, const SharedStrings* shared_strings) {
  // Buffer the row body separately so we can tell whether anything ended
  // up inside the <row> wrapper before we commit to writing it.
  std::string body;
  body.reserve(row_cells.size() * 24U);
  const std::size_t col_count = row_cells.size();
  for (std::size_t i = 0; i < col_count; ++i) {
    const std::uint32_t col = static_cast<std::uint32_t>(i);
    if (sheet.spill_region_covering(row, col) != nullptr) {
      // Phantom cell: suppressed entirely; Excel re-spills from the
      // anchor on load.
      continue;
    }
    (void)AppendCellXml(body, sheet, row, col, row_cells[i], shared_strings);
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
  return !cell.formula_text.empty() || !cell.cached_value.is_blank() || cell.xf_index != 0U;
}

std::string EncodeA1(std::uint32_t row, std::uint32_t col) {
  std::string out;
  out.reserve(10U);
  if (!a1::append_column_letters(out, col)) {
    return {};
  }
  out.append(std::to_string(row + 1U));
  return out;
}

std::string BuildSheetDataXml(const Sheet& sheet, const SharedStrings* shared_strings) {
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

  // Union of populated rows and override rows. Rows that have only an
  // override (no cells) still need to surface so the override survives
  // a save/load round-trip.
  std::vector<std::uint32_t> row_indices;
  row_indices.reserve(rows_map.size() + row_overrides.size());
  for (const auto& kv : rows_map) {
    row_indices.push_back(kv.first);
  }
  for (const RowLayout& ro : row_overrides) {
    if (rows_map.find(ro.row) == rows_map.end()) {
      row_indices.push_back(ro.row);
    }
  }
  std::sort(row_indices.begin(), row_indices.end());
  row_indices.erase(std::unique(row_indices.begin(), row_indices.end()), row_indices.end());

  std::string body;
  body.reserve((rows_map.size() + row_overrides.size()) * 64U);
  // Sentinel empty span used when a row has no override; avoids a
  // per-row default-construction of std::string.
  static const RowCells kEmptyRow;
  for (std::uint32_t row : row_indices) {
    std::string override_attrs;
    auto override_it = overrides_by_row.find(row);
    if (override_it != overrides_by_row.end()) {
      AppendRowOverrideAttrs(override_attrs, *override_it->second);
    }
    auto cells_it = rows_map.find(row);
    const RowCells& row_cells = (cells_it != rows_map.end()) ? cells_it->second : kEmptyRow;
    AppendRowXml(body, sheet, row, row_cells, override_attrs, shared_strings);
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
