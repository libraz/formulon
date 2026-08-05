//
// Implementation of the OOXML cell/row/sheetData builder. Pure functions
// only: no zip plumbing, no I/O. The orchestrating writer in
// ooxml_writer.cpp wraps this output in the surrounding <worksheet>
// element and packages it into the .xlsx archive.
//
// Spill semantics:
//
//   * Anchor cell of a registered spill region: emit <f t="array">; the
//     ref= attribute is intentionally omitted (legacy CSE arrays use it,
//     dynamic arrays do not).
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
#include <cstdio>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include "cell.h"
#include "double-conversion/double-conversion.h"
#include "eval/utf8_length.h"
#include "io/ooxml/shared_strings_writer.h"
#include "io/xlsb/func_id_table.h"
#include "io/xml_escape.h"
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

// Classifies a function name into the storage prefix Excel writes for it,
// for the OOXML `<f>` re-serialisation (see the formula-emit path below).
// Worksheet-only dynamic-array functions (the original 2018 set) carry
// `_xlfn._xlws.`; every other function absent from the classic (pre-2007)
// XLSB function-id table is a post-2007 "future function" carrying
// `_xlfn.`; classic functions carry no prefix. This mirrors the
// classic-vs-future split the XLSB writer already uses.
parser::StoragePrefixKind ClassifyStoragePrefix(std::string_view canonical_name) {
  std::string upper;
  upper.reserve(canonical_name.size());
  for (char c : canonical_name) {
    if (c >= 'a' && c <= 'z') {
      c = static_cast<char>(c - 'a' + 'A');
    }
    upper.push_back(c);
  }
  static constexpr std::string_view kXlwsFunctions[] = {"FILTER", "SORT", "SORTBY", "UNIQUE"};
  for (const std::string_view name : kXlwsFunctions) {
    if (upper == name) {
      return parser::StoragePrefixKind::XlfnXlws;
    }
  }
  if (xlsb::lookup_func_by_name(upper) != nullptr) {
    return parser::StoragePrefixKind::None;
  }
  return parser::StoragePrefixKind::Xlfn;
}

// ---------------------------------------------------------------------------
// Number formatting
// ---------------------------------------------------------------------------

// Emits a double using Grisu3 shortest-roundtrip dtoa (via Google's
// double-conversion library), shaped to match Excel's own OOXML output:
//
//   * Uppercase 'E' as the exponent character (Excel-authored XLSX
//     consistently uses 'E', whereas ECMAScript / printf("%g") use 'e').
//   * EMIT_POSITIVE_EXPONENT_SIGN so the exponent on positive-magnitude
//     forms carries a '+' sign (e.g. "1E+100"). Negative exponents
//     never carry a sign (e.g. "1E-3"); this matches Excel.
//   * UNIQUE_ZERO so -0.0 collapses to "0" without a special branch.
//     The +0.0 / -0.0 fast path below is retained as a defensive
//     optimisation: it avoids the converter call entirely for the
//     overwhelmingly common case of literal zero.
//   * decimal_in_shortest_low/high mirror the ECMAScript defaults
//     (-6 / 21). Excel-authored files we surveyed switch between
//     decimal and scientific shape inside this band; the precise
//     boundary appears to depend on display-format-driven rendering
//     rather than the underlying numeric value, so the ECMAScript
//     defaults are the most defensible starting point.
//
// NaN / +/-inf are *not* handled here: the caller must short-circuit
// them to an Error cell before reaching this point. The converter is
// constructed without infinity/NaN symbols so a stray special value
// would surface as a ToShortest() failure rather than silent garbage,
// but the pre-screen in AppendCellXml / AppendLiteralCellBody makes
// that path unreachable in practice.
void AppendNumberValue(std::string& out, double v) {
  if (v == 0.0) {
    out.push_back('0');
    return;
  }
  using DC = double_conversion::DoubleToStringConverter;
  // Built once at first call; the converter is stateless and thread-safe.
  static const DC kConv(
      /*flags=*/DC::UNIQUE_ZERO | DC::EMIT_POSITIVE_EXPONENT_SIGN,
      /*infinity_symbol=*/nullptr,
      /*nan_symbol=*/nullptr,
      /*exponent_character=*/'E',
      /*decimal_in_shortest_low=*/-6,
      /*decimal_in_shortest_high=*/21,
      /*max_leading_padding_zeroes_in_precision_mode=*/0,
      /*max_trailing_padding_zeroes_in_precision_mode=*/0);
  // 32 bytes covers every shortest output: kMaxCharsEcmaScriptShortest
  // (25) plus the trailing NUL plus a comfortable margin.
  char buf[32];
  double_conversion::StringBuilder builder(buf, sizeof(buf));
  // ToShortest() only fails for NaN/Inf when no special-value symbol
  // is configured; the caller pre-screens those, so success is
  // guaranteed here. We deliberately leave the return value
  // unchecked: the project builds with -fno-exceptions and an
  // assert/log on this unreachable branch would cost more bytes than
  // it earns.
  (void)kConv.ToShortest(v, &builder);
  out.append(buf, static_cast<std::size_t>(builder.position()));
}

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
// quote or '>'). This function emits the closing quote, any type
// attribute, the body, and the closing '</c>'.
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
                           const SharedStrings* shared_strings) {
  if (value.is_number()) {
    // Defensive: NaN / +/-Inf must never reach AppendNumberValue, which
    // would emit `nan`/`inf` text inside `<v>` (Excel rejects this on
    // load). The caller AppendCellXml pre-screens, but a future caller
    // path may not — downgrade to #NUM! here as a last line of defence.
    const double v = value.as_number();
    if (!std::isfinite(v)) {
      out.append("\" t=\"e\"><v>");
      out.append(display_name(ErrorCode::Num));
      out.append("</v></c>");
      return;
    }
    out.append("\">");
    out.append("<v>");
    AppendNumberValue(out, v);
    out.append("</v></c>");
    return;
  }
  if (value.is_boolean()) {
    out.append("\" t=\"b\"><v>");
    out.push_back(value.as_boolean() ? '1' : '0');
    out.append("</v></c>");
    return;
  }
  if (value.is_text()) {
    if (shared_strings != nullptr) {
      out.append("\" t=\"s\"><v>");
      out.append(std::to_string(shared_strings->index_of(value.as_text(), phonetic)));
      out.append("</v></c>");
      return;
    }
    out.append("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
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
      // Rich error literal: not writable as a legacy <v>; emit a blank cell.
      out.append("\"/>");
      return;
    }
    out.append("\" t=\"e\"><v>");
    out.append(display_name(value.as_error()));
    out.append("</v></c>");
    return;
  }
  // Defensive fallback for Array/Ref/Lambda: cells should never store
  // these in the post-evaluation path (arrays land in the spill table;
  // Ref/Lambda are not yet first-class cell payloads). Surface #VALUE!
  // rather than crashing the writer.
  out.append("\" t=\"e\"><v>");
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
  const bool literal_blank = !has_formula && cell.cached_value.is_blank() && cell.xf_index == 0U;
  if (literal_blank) {
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
    // on future functions, `_xlpm.` on LET / LAMBDA parameters) so a real
    // Excel reading this file resolves the modern functions instead of
    // showing #NAME?. `formula_text` was normalised to the canonical
    // formula-bar form on ingestion (io::strip_storage_prefixes). Parse it
    // and re-serialise through the storage formatter; on any parse failure
    // fall back to the canonical text unchanged.
    Arena formula_arena;
    parser::Parser formula_parser(formula, formula_arena);
    parser::AstNode* formula_root = formula_parser.parse();
    bool storage_emitted = false;
    if (formula_root != nullptr && formula_parser.errors().empty()) {
      const std::string storage = parser::format_formula_storage(*formula_root, &ClassifyStoragePrefix);
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
        AppendNumberValue(out, v);
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
  // (closing quote + optional type attribute), writes the body, and
  // closes the </c>. The cell's `phonetic_text` is forwarded so any
  // <rPh> annotation captured at read time round-trips back through
  // the inline-string block.
  out.append("<c r=\"");
  out.append(addr);
  // Close the `r=` attribute, optionally inject `s=`, then re-open the
  // body without a closing quote so `AppendLiteralCellBody` can append
  // its own type attribute prefix (`\" t="..."`). Reopening a partial
  // attribute with a sentinel space rather than `"` lets the helper
  // continue using its existing prefix conventions unchanged.
  if (cell.xf_index != 0U) {
    out.append("\"");
    AppendStyleAttr(out, cell.xf_index);
    // Re-prime the helper: it expects the buffer to end immediately
    // after the address with no trailing quote (the helper writes the
    // `\"` itself). Add a placeholder `r2=""` would be wrong; instead,
    // emit a degenerate `r2=""`-free shape by undoing the helper's
    // expectation. The cleanest fix is to inline the helper's prefix
    // here.
    if (cell.cached_value.is_number()) {
      const double nv = cell.cached_value.as_number();
      if (!std::isfinite(nv)) {
        out.append(" t=\"e\"><v>");
        out.append(display_name(ErrorCode::Num));
        out.append("</v></c>");
      } else {
        out.append(" ><v>");
        AppendNumberValue(out, nv);
        out.append("</v></c>");
      }
    } else if (cell.cached_value.is_boolean()) {
      out.append(" t=\"b\"><v>");
      out.push_back(cell.cached_value.as_boolean() ? '1' : '0');
      out.append("</v></c>");
    } else if (cell.cached_value.is_text()) {
      if (shared_strings != nullptr) {
        out.append(" t=\"s\"><v>");
        out.append(std::to_string(shared_strings->index_of(cell.cached_value.as_text(), cell.phonetic_text)));
        out.append("</v></c>");
      } else {
        out.append(" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
        AppendXmlEscaped(out, cell.cached_value.as_text());
        out.append("</t>");
        if (!cell.phonetic_text.empty()) {
          const std::uint32_t eb = formulon::eval::utf16_units_in(cell.cached_value.as_text());
          out.append("<rPh sb=\"0\" eb=\"");
          out.append(std::to_string(eb));
          out.append("\"><t xml:space=\"preserve\">");
          AppendXmlEscaped(out, cell.phonetic_text);
          out.append("</t></rPh>");
        }
        out.append("</is></c>");
      }
    } else if (cell.cached_value.is_error()) {
      if (IsLegacyErrorCode(cell.cached_value.as_error())) {
        out.append(" t=\"e\"><v>");
        out.append(display_name(cell.cached_value.as_error()));
        out.append("</v></c>");
      } else {
        // Rich error literal with a style: blank cell keeping the style.
        out.append("/>");
      }
    } else {
      out.append(" t=\"e\"><v>");
      out.append(display_name(ErrorCode::Value));
      out.append("</v></c>");
    }
    return true;
  }
  AppendLiteralCellBody(out, cell.cached_value, cell.phonetic_text, shared_strings);
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
    char buf[32];
    // %.17g is round-trip safe under IEEE 754, so a recalc-save does not
    // drift the row metric. Matches the column-width writer.
    std::snprintf(buf, sizeof(buf), "%.17g", layout.height);
    out.append(buf);
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
