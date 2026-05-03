// Copyright 2026 libraz. Licensed under the MIT License.
//
// `<sheetData>` walker. See sheet_reader.h for the public contract.
//
// The walker visits each `<row>`/`<c>` pair in document order. A small
// `shared_formulas` map records the master formula text per `si` index;
// slave occurrences are looked up in this map. The map is rebuilt per
// `read_sheet_data` call (per sheet) — `si` indices are sheet-local in
// OOXML, so leaking entries across sheets would be a correctness bug.
//
// Known limitations (called out in the public header):
//   * Shared formulas are reused verbatim. R1C1-style relative shift will
//     land in Bundle 2.5 / Phase 4.
//   * Cached values from `<v>` on formula cells are dropped: we let the
//     recalc engine populate them via `Workbook::recalc()`.

#include "io/sheet_reader.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <cstring>
#include <deque>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>

#include "io/cell_parser.h"
#include "io/sax_xml_reader.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

/// Master record for a shared-formula group: the formula body of the
/// first `<f t="shared" si="N" ...>` occurrence. Slave occurrences of
/// the same `si` reuse this text. The leading '=' is intentionally
/// absent (OOXML <f> contents never have it).
struct SharedFormulaMaster {
  std::string text;
};

/// Reads the `<f>` child of `c_node` and updates `formula_out`. Returns
/// `false` and surfaces an error when a slave occurrence references an
/// unknown `si`. `shared` is the per-sheet map of master formulas.
///
/// Behaviour matrix:
///   * No `<f>` -> `formula_out` left empty.
///   * `<f>BODY</f>` -> `formula_out = BODY` (no shared bookkeeping).
///   * `<f t="shared" si="N">BODY</f>` -> registers the master in
///     `shared[N]`, sets `formula_out = BODY`.
///   * `<f t="shared" si="N"/>` (no body) -> looks up master, sets
///     `formula_out` to the master text (verbatim — see file-level note).
///   * `<f t="array">BODY</f>` (CSE array) and `<f t="dataTable" ...>`
///     are accepted but treated as plain formulas: we read the body as
///     the formula text. CSE-array detail / data-table semantics will
///     land in a later bundle.
Expected<void, Error> ResolveFormula(const pugi::xml_node& c_node,
                                     std::unordered_map<std::uint32_t, SharedFormulaMaster>& shared,
                                     std::string& formula_out) {
  pugi::xml_node f_node = c_node.child("f");
  if (!f_node) {
    formula_out.clear();
    return Expected<void, Error>::Ok();
  }
  const std::string_view ftype = f_node.attribute("t").value();
  if (ftype != "shared") {
    // Plain formula (or an unhandled variant we treat as plain).
    formula_out = f_node.text().get();
    if (!formula_out.empty() && formula_out.front() == '=') {
      formula_out.erase(0, 1);
    }
    return Expected<void, Error>::Ok();
  }

  // Shared formula. `si` is required; missing/non-numeric => corrupt.
  pugi::xml_attribute si_attr = f_node.attribute("si");
  if (!si_attr) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "shared formula: <f t='shared'> missing 'si'",
                      "context=sheet_reader");
  }
  const std::int64_t si_signed = si_attr.as_llong(-1);
  if (si_signed < 0 || si_signed > 0xFFFFFFFFLL) {
    std::string ctx("context=sheet_reader si=");
    ctx.append(si_attr.value());
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "shared formula: 'si' out of range", std::move(ctx));
  }
  const auto si = static_cast<std::uint32_t>(si_signed);

  std::string body = f_node.text().get();
  if (!body.empty() && body.front() == '=') {
    body.erase(0, 1);
  }
  if (!body.empty()) {
    // Master occurrence: register and use as formula text.
    shared[si] = SharedFormulaMaster{body};
    formula_out = std::move(body);
    return Expected<void, Error>::Ok();
  }
  // Slave occurrence: look up master.
  auto it = shared.find(si);
  if (it == shared.end()) {
    std::string ctx("context=sheet_reader si=");
    ctx.append(std::to_string(si));
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "shared formula: slave references unknown si",
                      std::move(ctx));
  }
  // NOTE: this slice does NOT shift R1C1 relative refs. See sheet_reader.h.
  formula_out = it->second.text;
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<void, Error> read_sheet_data(const pugi::xml_document& sheet_doc, std::size_t sheet_index, Workbook& workbook,
                                      SheetReadContext& ctx, std::deque<std::string>& text_storage) {
  if (sheet_index >= workbook.sheet_count()) {
    std::string ctxs("context=sheet_reader sheet_index=");
    ctxs.append(std::to_string(sheet_index));
    ctxs.append(" sheet_count=");
    ctxs.append(std::to_string(workbook.sheet_count()));
    return make_error(FormulonErrorCode::kInvalidArgument, "read_sheet_data: sheet_index out of range",
                      std::move(ctxs));
  }
  pugi::xml_node worksheet = sheet_doc.child("worksheet");
  if (!worksheet) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "sheet doc: missing <worksheet> root",
                      "context=sheet_reader");
  }
  pugi::xml_node sheet_data = worksheet.child("sheetData");
  if (!sheet_data) {
    // Empty `<sheetData>` is legal (and Excel sometimes omits the element
    // entirely on a brand-new sheet); treat absence as no cells.
    return Expected<void, Error>::Ok();
  }

  std::unordered_map<std::uint32_t, SharedFormulaMaster> shared_formulas;

  for (pugi::xml_node row = sheet_data.child("row"); row; row = row.next_sibling("row")) {
    for (pugi::xml_node c = row.child("c"); c; c = c.next_sibling("c")) {
      auto parsed_or = parse_cell_element(c, text_storage);
      if (!parsed_or) {
        return parsed_or.error();
      }
      // Take a const reference rather than moving so the `string_view`
      // inside `parsed.value` (which references `text_storage`)
      // remains stable across this scope. `text_storage` is a
      // `std::deque`, so its element addresses do not move on later
      // appends.
      const ParsedCell& parsed = parsed_or.value();

      // Resolve formula text (handling shared-formula reuse).
      std::string formula_text;
      {
        auto resolved = ResolveFormula(c, shared_formulas, formula_text);
        if (!resolved) {
          return resolved.error();
        }
      }

      // The parser already routed any inline-string text through
      // `text_storage`, so the value can be stored as-is.
      const Value to_store = parsed.value;

      // Hand off to the workbook. Order matters: `set_cell_formula`
      // resets the cell's cached_value, so the cached value (if any
      // came from `<v>`) would be overwritten if we wrote it first.
      // For literal cells, just write the value.
      if (!formula_text.empty()) {
        // `Workbook::set_cell_formula` accepts both spellings, but to
        // match the parser/evaluator's expected input form (the existing
        // call sites in workbook_recalc_test.cpp pass "=A1*2") we
        // prepend '=' here.
        std::string with_eq("=");
        with_eq.append(formula_text);
        auto wf = workbook.set_cell_formula(sheet_index, parsed.row, parsed.col, std::move(with_eq));
        if (!wf) {
          return wf.error();
        }
      } else {
        // Literal cell. Skip blank-blank cells (e.g. <c r="A1"/>) to
        // keep the row map sparse and avoid spurious dirty-set entries.
        if (to_store.is_blank()) {
          // Nothing to record. Note: SST placeholders are Text("") so
          // they fall through here; the queue below still picks them up.
          if (parsed.is_sst_index) {
            ctx.pending_sst_cells.emplace_back(parsed.row, parsed.col, parsed.sst_index);
          }
          continue;
        }
        auto wv = workbook.set_cell_value(sheet_index, parsed.row, parsed.col, to_store);
        if (!wv) {
          return wv.error();
        }
      }

      if (parsed.is_sst_index) {
        ctx.pending_sst_cells.emplace_back(parsed.row, parsed.col, parsed.sst_index);
      }

      // Inline-string cells carry any <rPh> annotation directly on the
      // cell parser's output. SST-referenced cells route their phonetic
      // through the post-loop SST resolution pass, so we skip them here.
      if (!parsed.is_sst_index && !parsed.phonetic_text.empty()) {
        workbook.sheet(sheet_index).set_cell_phonetic(parsed.row, parsed.col, parsed.phonetic_text);
      }
    }
  }
  return Expected<void, Error>::Ok();
}

// ---------------------------------------------------------------------------
// View / layout helpers. Each lives in an anonymous namespace so the
// translation unit owns its parsing fences; the public driver
// `read_sheet_view_and_layout` composes them in document order.
// ---------------------------------------------------------------------------

namespace {

/// Returns true when the OOXML boolean attribute string is "true" /
/// "1" (case-insensitive). Matches the lexicon Excel emits for sheet
/// hidden / outline flags.
bool ParseXmlBool(std::string_view value) noexcept {
  if (value == "1") {
    return true;
  }
  if (value.size() != 4) {
    return false;
  }
  return (value[0] == 't' || value[0] == 'T') && (value[1] == 'r' || value[1] == 'R') &&
         (value[2] == 'u' || value[2] == 'U') && (value[3] == 'e' || value[3] == 'E');
}

/// Saturating cast from a possibly-signed C string to `std::uint8_t`.
/// Used for `outlineLevel`; OOXML caps the value at 7 but we accept up
/// to 255 defensively and clamp negatives to 0.
std::uint8_t ParseOutlineLevel(const char* text) noexcept {
  if (text == nullptr || *text == '\0') {
    return 0U;
  }
  // strtol is locale-independent for ASCII digits and doesn't pull in
  // <iostream>. The result is clamped to [0, 255] so callers see a
  // valid `uint8_t` even on garbage input.
  char* end = nullptr;
  const long n = std::strtol(text, &end, 10);
  if (end == text) {
    return 0U;
  }
  if (n <= 0) {
    return 0U;
  }
  if (n >= 255) {
    return 255U;
  }
  return static_cast<std::uint8_t>(n);
}

/// Parses `<sheetView zoomScale="...">` and `<pane state="frozen"
/// xSplit="N" ySplit="M">` into `view`. Missing or out-of-range
/// `zoomScale` falls back to `SheetView::kDefaultZoomScale`. A `<pane>`
/// element with `state != "frozen"` (or absent) leaves
/// `freeze_rows` / `freeze_cols` at zero.
void ApplySheetView(const pugi::xml_node& worksheet, SheetView& view) {
  pugi::xml_node sheet_views = worksheet.child("sheetViews");
  if (!sheet_views) {
    return;
  }
  pugi::xml_node sheet_view = sheet_views.child("sheetView");
  if (!sheet_view) {
    return;
  }
  if (pugi::xml_attribute zoom_attr = sheet_view.attribute("zoomScale"); zoom_attr) {
    const long raw = std::strtol(zoom_attr.value(), nullptr, 10);
    if (raw >= 10 && raw <= 400) {
      view.zoom_scale = static_cast<std::uint32_t>(raw);
    }
  }
  pugi::xml_node pane = sheet_view.child("pane");
  if (pane) {
    const std::string_view state = pane.attribute("state").value();
    if (state == "frozen") {
      const long y = std::strtol(pane.attribute("ySplit").value(), nullptr, 10);
      const long x = std::strtol(pane.attribute("xSplit").value(), nullptr, 10);
      if (y > 0) {
        view.freeze_rows = static_cast<std::uint32_t>(y);
      }
      if (x > 0) {
        view.freeze_cols = static_cast<std::uint32_t>(x);
      }
    }
  }
}

/// Parses `<sheetPr><tabHidden val="1"/></sheetPr>` into `view`.
/// OOXML also accepts `<sheetPr><tabColor .../>`, but only `tabHidden`
/// is a visibility-affecting flag for the worksheet part itself. The
/// workbook-level `<sheet state="hidden">` form is handled in
/// `read_ooxml`; this helper preserves any prior `tab_hidden` state so
/// the merge is OR-style.
void ApplySheetPrTabHidden(const pugi::xml_node& worksheet, SheetView& view) {
  pugi::xml_node sheet_pr = worksheet.child("sheetPr");
  if (!sheet_pr) {
    return;
  }
  // Some writers emit `tabHidden` as a direct attribute (`<sheetPr
  // tabHidden="1"/>`); others emit it as a child element with a `val`
  // attribute. Accept both shapes.
  if (pugi::xml_attribute attr = sheet_pr.attribute("tabHidden"); attr) {
    if (ParseXmlBool(attr.value())) {
      view.tab_hidden = true;
    }
  }
  if (pugi::xml_node child = sheet_pr.child("tabHidden"); child) {
    if (pugi::xml_attribute val = child.attribute("val"); val) {
      if (ParseXmlBool(val.value())) {
        view.tab_hidden = true;
      }
    } else {
      // Bare `<tabHidden/>` element with no `val` attribute is treated
      // as `val="1"` to match what Excel's older writers emit.
      view.tab_hidden = true;
    }
  }
}

/// Parses `<cols><col min max width hidden outlineLevel/></cols>` into
/// `layout.columns`. Entries that omit `width` are skipped: the layout
/// model only carries explicit width overrides, and pure
/// `customWidth=1` / `bestFit=1` markers without a stored width have no
/// observable round-trip effect.
void ApplyColumnLayouts(const pugi::xml_node& worksheet, SheetLayout& layout) {
  pugi::xml_node cols = worksheet.child("cols");
  if (!cols) {
    return;
  }
  for (pugi::xml_node col = cols.child("col"); col; col = col.next_sibling("col")) {
    pugi::xml_attribute min_attr = col.attribute("min");
    pugi::xml_attribute max_attr = col.attribute("max");
    pugi::xml_attribute width_attr = col.attribute("width");
    if (!min_attr || !max_attr) {
      continue;
    }
    const long min_v = std::strtol(min_attr.value(), nullptr, 10);
    const long max_v = std::strtol(max_attr.value(), nullptr, 10);
    if (min_v < 1 || max_v < min_v) {
      continue;
    }
    ColumnLayout entry;
    entry.first = static_cast<std::uint32_t>(min_v - 1);
    entry.last = static_cast<std::uint32_t>(max_v - 1);
    if (width_attr) {
      entry.width = std::strtod(width_attr.value(), nullptr);
    } else {
      // No explicit width override — skip. `customWidth` / `bestFit`
      // alone are not enough to round-trip through the layout model.
      continue;
    }
    if (pugi::xml_attribute hidden_attr = col.attribute("hidden"); hidden_attr) {
      entry.hidden = ParseXmlBool(hidden_attr.value());
    }
    if (pugi::xml_attribute outline_attr = col.attribute("outlineLevel"); outline_attr) {
      entry.outline_level = ParseOutlineLevel(outline_attr.value());
    }
    layout.columns.push_back(entry);
  }
}

/// Walks `<sheetData><row .../></sheetData>` collecting per-row
/// overrides (height / hidden / outline) into `layout.row_overrides`.
/// Rows that carry only `r` (the row number) are skipped — they are
/// just position markers and have no override payload.
void ApplyRowOverrides(const pugi::xml_node& worksheet, SheetLayout& layout) {
  pugi::xml_node sheet_data = worksheet.child("sheetData");
  if (!sheet_data) {
    return;
  }
  for (pugi::xml_node row = sheet_data.child("row"); row; row = row.next_sibling("row")) {
    pugi::xml_attribute r_attr = row.attribute("r");
    pugi::xml_attribute ht_attr = row.attribute("ht");
    pugi::xml_attribute hidden_attr = row.attribute("hidden");
    pugi::xml_attribute outline_attr = row.attribute("outlineLevel");
    if (!ht_attr && !hidden_attr && !outline_attr) {
      continue;
    }
    if (!r_attr) {
      continue;
    }
    const long r_v = std::strtol(r_attr.value(), nullptr, 10);
    if (r_v < 1) {
      continue;
    }
    RowLayout entry;
    entry.row = static_cast<std::uint32_t>(r_v - 1);
    if (ht_attr) {
      entry.height = std::strtod(ht_attr.value(), nullptr);
    }
    if (hidden_attr) {
      entry.hidden = ParseXmlBool(hidden_attr.value());
    }
    if (outline_attr) {
      entry.outline_level = ParseOutlineLevel(outline_attr.value());
    }
    layout.row_overrides.push_back(entry);
  }
}

}  // namespace

Expected<void, Error> read_sheet_view_and_layout(const pugi::xml_document& sheet_doc, std::size_t sheet_index,
                                                 Workbook& workbook) {
  if (sheet_index >= workbook.sheet_count()) {
    std::string ctx("context=sheet_reader.view_layout sheet_index=");
    ctx.append(std::to_string(sheet_index));
    ctx.append(" sheet_count=");
    ctx.append(std::to_string(workbook.sheet_count()));
    return make_error(FormulonErrorCode::kInvalidArgument, "read_sheet_view_and_layout: sheet_index out of range",
                      std::move(ctx));
  }
  pugi::xml_node worksheet = sheet_doc.child("worksheet");
  if (!worksheet) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "sheet doc: missing <worksheet> root",
                      "context=sheet_reader.view_layout");
  }
  Sheet& sheet = workbook.sheet(sheet_index);
  SheetView& view = sheet.mutable_view();
  SheetLayout& layout = sheet.mutable_layout();
  ApplySheetView(worksheet, view);
  ApplySheetPrTabHidden(worksheet, view);
  ApplyColumnLayouts(worksheet, layout);
  ApplyRowOverrides(worksheet, layout);
  return Expected<void, Error>::Ok();
}

// ---------------------------------------------------------------------------
// SAX-path implementation. The streaming scanner produces one
// `CellRecord` per `<c>` element; this helper translates that record
// into the same `Workbook` mutations the DOM path performs by
// delegating value decoding to `decode_cell_payload` (shared with
// `parse_cell_element`).
//
// The implementation is `#if`-guarded so WASM builds (where the
// OOXML reader's threshold pins the dispatch to DOM) do not pay the
// compile-time cost. See `src/io/sax_xml_reader.cpp` for the matching
// guard on the streaming scanner.
// ---------------------------------------------------------------------------

#if !defined(FORMULON_WASM) || defined(FORMULON_WASM_ENABLE_SAX)

namespace {

/// Per-call state passed through `SheetSaxCallbacks::user_data`. The
/// scanner's callbacks need access to the workbook, the sheet index,
/// the SST queue, and the text-storage deque; bundling them here
/// avoids `std::function` capture entirely.
struct SaxApplyState {
  std::size_t sheet_index;
  Workbook* workbook;
  SheetReadContext* ctx;
  std::deque<std::string>* text_storage;
};

/// Translates one `CellRecord` into the same workbook mutations the
/// DOM path produces.
///
/// Shared-formula and array-formula attributes (`<f t="shared" si=>`,
/// `<f t="array">`) are NOT surfaced through the SAX record, so the
/// SAX path treats every `<f>` body as plain. Sheets that rely on
/// shared formulas typically come in well under `kSaxThresholdBytes`
/// and route through the DOM path instead.
Expected<void, Error> ApplyCellRecord(const CellRecord& rec, std::size_t sheet_index, Workbook& workbook,
                                      SheetReadContext& ctx, std::deque<std::string>& text_storage) {
  const bool value_present = rec.is_inline_string || !rec.value.empty();
  auto payload_or = decode_cell_payload(rec.t, rec.value, value_present, rec.is_inline_string, text_storage);
  if (!payload_or) {
    return payload_or.error();
  }
  const ParsedCell& parsed = payload_or.value();

  if (!rec.formula.empty()) {
    std::string with_eq("=");
    with_eq.append(rec.formula);
    auto wf = workbook.set_cell_formula(sheet_index, rec.row, rec.col, std::move(with_eq));
    if (!wf) {
      return wf.error();
    }
  } else if (!parsed.value.is_blank()) {
    auto wv = workbook.set_cell_value(sheet_index, rec.row, rec.col, parsed.value);
    if (!wv) {
      return wv.error();
    }
  }
  if (parsed.is_sst_index) {
    ctx.pending_sst_cells.emplace_back(rec.row, rec.col, parsed.sst_index);
  }
  // Inline-string cells with <rPh> annotations carry their kana on the
  // SAX record. SST-referenced cells (rec.phonetic stays empty by
  // construction) route their phonetic through the post-loop SST
  // resolution pass instead — same contract the DOM path uses.
  if (!rec.phonetic.empty()) {
    workbook.sheet(sheet_index).set_cell_phonetic(rec.row, rec.col, rec.phonetic);
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> SaxOnCellTrampoline(void* user_data, const CellRecord& rec) {
  auto* st = static_cast<SaxApplyState*>(user_data);
  return ApplyCellRecord(rec, st->sheet_index, *st->workbook, *st->ctx, *st->text_storage);
}

}  // namespace

Expected<void, Error> read_sheet_data_sax(ByteSpan sheet_xml, std::size_t sheet_index, Workbook& workbook,
                                          SheetReadContext& ctx, std::deque<std::string>& text_storage) {
  if (sheet_index >= workbook.sheet_count()) {
    std::string ctxs("context=sheet_reader_sax sheet_index=");
    ctxs.append(std::to_string(sheet_index));
    ctxs.append(" sheet_count=");
    ctxs.append(std::to_string(workbook.sheet_count()));
    return make_error(FormulonErrorCode::kInvalidArgument, "read_sheet_data_sax: sheet_index out of range",
                      std::move(ctxs));
  }
  SaxApplyState state{sheet_index, &workbook, &ctx, &text_storage};
  SheetSaxCallbacks cb;
  cb.user_data = &state;
  cb.on_cell = &SaxOnCellTrampoline;
  return scan_sheet_data(sheet_xml, cb);
}

#endif  // !FORMULON_WASM || FORMULON_WASM_ENABLE_SAX

}  // namespace io
}  // namespace formulon
