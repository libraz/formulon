// Copyright 2026 libraz. Licensed under the MIT License.
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
#include <cstdint>
#include <cstdio>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "sheet.h"
#include "value.h"

namespace formulon {
namespace io {
namespace {

// ---------------------------------------------------------------------------
// XML escaping
// ---------------------------------------------------------------------------

// Local copy: ooxml_writer.cpp's anonymous-namespace AppendXmlEscaped is not
// reachable here, and the two emitters can plausibly diverge later (e.g.
// shared-string interning may want a tighter escape set). 20 LOC is
// comfortably below the duplication-cost threshold.
void AppendXmlEscaped(std::string& out, std::string_view in) {
  for (char raw : in) {
    switch (raw) {
      case '&':
        out.append("&amp;");
        break;
      case '<':
        out.append("&lt;");
        break;
      case '>':
        out.append("&gt;");
        break;
      case '"':
        out.append("&quot;");
        break;
      case '\'':
        out.append("&apos;");
        break;
      default:
        out.push_back(raw);
        break;
    }
  }
}

// ---------------------------------------------------------------------------
// Number formatting
// ---------------------------------------------------------------------------

// Emits a double using %.17g (round-trip safe under IEEE 754) with two
// targeted normalisations:
//
//   * -0.0 collapses to "0" so XML diffs stay deterministic across
//     platforms that disagree on the sign-of-zero printout.
//   * +0.0 takes the same fast path.
//
// NaN / +/-inf are *not* handled here: the caller must short-circuit them to
// an Error cell before reaching this point.
//
// TODO(grisu3): switch to shortest-roundtrip dtoa once double-conversion is
// adopted.
void AppendNumberValue(std::string& out, double v) {
  if (v == 0.0) {
    out.push_back('0');
    return;
  }
  char buf[32];
  std::snprintf(buf, sizeof(buf), "%.17g", v);
  out.append(buf);
}

// ---------------------------------------------------------------------------
// Cell emission
// ---------------------------------------------------------------------------

// Emits the <c> element for an Error value at `addr`.
void AppendErrorCellXml(std::string& out, std::string_view addr, ErrorCode code) {
  out.append("<c r=\"");
  out.append(addr);
  out.append("\" t=\"e\"><v>");
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
void AppendLiteralCellBody(std::string& out, const Value& value) {
  if (value.is_number()) {
    out.append("\">");
    out.append("<v>");
    AppendNumberValue(out, value.as_number());
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
    out.append("\" t=\"inlineStr\"><is><t>");
    AppendXmlEscaped(out, value.as_text());
    out.append("</t></is></c>");
    return;
  }
  if (value.is_error()) {
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
bool AppendCellXml(std::string& out, const Sheet& sheet, std::uint32_t row, std::uint32_t col, const Cell& cell) {
  const bool has_formula = !cell.formula_text.empty();
  const bool literal_blank = !has_formula && cell.cached_value.is_blank();
  if (literal_blank) {
    return false;
  }

  // Pre-screen NaN/Inf number literals so we never half-emit a cell tag.
  // Formula cells with non-finite cached values still emit the <f>; only
  // the <v> is downgraded.
  if (!has_formula && cell.cached_value.is_number() && !std::isfinite(cell.cached_value.as_number())) {
    const std::string addr = EncodeA1(row, col);
    AppendErrorCellXml(out, addr, ErrorCode::Num);
    return true;
  }

  const std::string addr = EncodeA1(row, col);

  if (has_formula) {
    const SpillRegion* anchored = sheet.spill_region_at_anchor(row, col);
    out.append("<c r=\"");
    out.append(addr);

    // If the anchor's cached value is an Error, surface it via t="e" so
    // Excel renders the error glyph rather than a number.
    if (cell.cached_value.is_error()) {
      out.append("\" t=\"e\">");
    } else {
      out.append("\">");
    }

    // <f> with optional t="array" for spill anchors. The formula text
    // always begins with '='; strip it before serialisation.
    if (anchored != nullptr) {
      out.append("<f t=\"array\">");
    } else {
      out.append("<f>");
    }
    std::string_view formula = cell.formula_text;
    if (!formula.empty() && formula.front() == '=') {
      formula.remove_prefix(1);
    }
    AppendXmlEscaped(out, formula);
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
      // both shapes for formula results.
      out.append("<v>");
      AppendXmlEscaped(out, cv.as_text());
      out.append("</v>");
    } else if (cv.is_error()) {
      out.append("<v>");
      out.append(display_name(cv.as_error()));
      out.append("</v>");
    }
    // Array / Ref / Lambda cached values fall through with no <v>; the
    // engine evaluates on load.

    out.append("</c>");
    return true;
  }

  // Literal-value cell. AppendLiteralCellBody completes the opening tag
  // (closing quote + optional type attribute), writes the body, and
  // closes the </c>.
  out.append("<c r=\"");
  out.append(addr);
  AppendLiteralCellBody(out, cell.cached_value);
  return true;
}

// Emits the <row> wrapper with all visible cells in the row. Returns true
// when at least one <c> was emitted (i.e. the <row> was actually written),
// false when the row collapsed to nothing (every cell was blank or a
// phantom).
bool AppendRowXml(std::string& out, const Sheet& sheet, std::uint32_t row, const std::vector<Cell>& row_cells) {
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
    (void)AppendCellXml(body, sheet, row, col, row_cells[i]);
  }
  if (body.empty()) {
    return false;
  }
  out.append("<row r=\"");
  out.append(std::to_string(row + 1U));
  out.append("\">");
  out.append(body);
  out.append("</row>");
  return true;
}

}  // namespace

std::string EncodeA1(std::uint32_t row, std::uint32_t col) {
  std::string col_str;
  col_str.reserve(3U);
  std::uint32_t c = col + 1U;  // convert 0-based to 1-based
  while (c > 0U) {
    col_str.push_back(static_cast<char>('A' + ((c - 1U) % 26U)));
    c = (c - 1U) / 26U;
  }
  std::reverse(col_str.begin(), col_str.end());
  col_str.append(std::to_string(row + 1U));
  return col_str;
}

std::string BuildSheetDataXml(const Sheet& sheet) {
  // Collect populated row indices and sort ascending so the output is
  // deterministic regardless of unordered_map iteration order.
  const auto& rows_map = sheet.rows();
  std::vector<std::uint32_t> row_indices;
  row_indices.reserve(rows_map.size());
  for (const auto& kv : rows_map) {
    row_indices.push_back(kv.first);
  }
  std::sort(row_indices.begin(), row_indices.end());

  std::string body;
  body.reserve(rows_map.size() * 64U);
  for (std::uint32_t row : row_indices) {
    const auto it = rows_map.find(row);
    if (it == rows_map.end()) {
      continue;
    }
    AppendRowXml(body, sheet, row, it->second);
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
