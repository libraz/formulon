//
// `<sheetData>` walker. See sheet_reader.h for the public contract.
//
// The walker visits each `<row>`/`<c>` pair in document order. A small
// `shared_formulas` map records the master formula text and anchor cell
// per `si` index; slave occurrences are looked up in this map and shifted
// by their relative row/column offset. The map is rebuilt per
// `read_sheet_data` call (per sheet) — `si` indices are sheet-local in
// OOXML, so leaking entries across sheets would be a correctness bug.
//
// Known limitation (called out in the public header):
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
#include <vector>

#include "io/cell_parser.h"
#include "io/sax_xml_reader.h"
#include "io/xml_utils.h"
#include "io/xsd_bool.h"
#include "parser/ast_format.h"
#include "parser/ast_shift.h"
#include "parser/parser.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

/// Master record for a shared-formula group: the formula body of the
/// first `<f t="shared" si="N" ...>` occurrence and the cell that owns it.
/// Slave occurrences of the same `si` reuse this text after shifting every
/// relative reference by `(slave - master)`. The leading '=' is intentionally
/// absent (OOXML <f> contents never have it).
struct SharedFormulaMaster {
  std::string text;
  std::uint32_t row = 0;
  std::uint32_t col = 0;
};

std::string ShiftSharedFormulaText(const SharedFormulaMaster& master, std::uint32_t target_row,
                                   std::uint32_t target_col) {
  const std::int32_t row_delta = static_cast<std::int32_t>(target_row) - static_cast<std::int32_t>(master.row);
  const std::int32_t col_delta = static_cast<std::int32_t>(target_col) - static_cast<std::int32_t>(master.col);
  if (row_delta == 0 && col_delta == 0) {
    return master.text;
  }

  std::string source("=");
  source.append(master.text);
  Arena arena;
  parser::Parser parser(source, arena);
  parser::AstNode* root = parser.parse();
  if (root == nullptr || !parser.errors().empty()) {
    // Keep the workbook loadable for formula dialects this parser does not
    // fully understand yet. Parseable formulas still get correct Excel-style
    // shared-formula relative expansion.
    return master.text;
  }
  const parser::AstNode* shifted = parser::shift_relative_refs(*root, arena, row_delta, col_delta);
  if (shifted == nullptr) {
    return master.text;
  }
  return parser::format_formula(*shifted);
}

/// Records a dynamic-array anchor for a `<f t="array" ref="...">` whose
/// `ref` spans more than the anchor cell. Single-cell refs (legacy CSE
/// scalars) and unparseable refs are ignored. The OOXML ref must be an
/// ordered rectangle whose top-left corner is the formula cell. This keeps
/// the stored extent anchored at its origin and prevents unsigned underflow
/// when the spill region is registered below.
void RecordArrayAnchor(SheetReadContext& ctx, std::string_view ref, std::uint32_t anchor_row,
                       std::uint32_t anchor_col) {
  const std::size_t colon = ref.find(':');
  if (colon == std::string_view::npos) {
    return;  // single-cell array ref: no spill targets to protect.
  }
  auto a = parse_a1(ref.substr(0, colon));
  auto b = parse_a1(ref.substr(colon + 1));
  if (!a || !b) {
    return;
  }
  const std::uint32_t first_row = a.value().first;
  const std::uint32_t first_col = a.value().second;
  const std::uint32_t last_row = b.value().first;
  const std::uint32_t last_col = b.value().second;
  if (last_row < first_row || last_col < first_col || anchor_row != first_row || anchor_col != first_col) {
    return;
  }
  if (last_row == anchor_row && last_col == anchor_col) {
    return;  // degenerate single-cell footprint.
  }
  ctx.array_anchors.push_back(ArrayAnchor{anchor_row, anchor_col, last_row, last_col});
}

/// Registers each recorded dynamic-array anchor as a spill region so the
/// cached spill targets (bare `<v>` cells Excel writes for F7 / F8 of a
/// `=SEQUENCE(3)` in F6) do not read back as independent literals. Without
/// this, the anchor's re-spill on recalc collides with those literals and
/// surfaces `#SPILL!`. The phantom values are captured into the region and
/// the underlying non-anchor cells are blanked so recalc can overwrite the
/// footprint freely.
void RegisterArraySpills(Sheet& sheet, const std::vector<ArrayAnchor>& anchors) {
  for (const ArrayAnchor& a : anchors) {
    // Keep this defensive check here as well as in RecordArrayAnchor: this
    // helper performs unsigned extent arithmetic and may gain other callers.
    if (a.last_row < a.row || a.last_col < a.col) {
      continue;
    }
    const std::uint32_t rows = a.last_row - a.row + 1U;
    const std::uint32_t cols = a.last_col - a.col + 1U;
    if (static_cast<std::uint64_t>(rows) * cols <= 1U) {
      continue;
    }
    std::vector<Value> values;
    values.reserve(static_cast<std::size_t>(rows) * cols);
    for (std::uint32_t r = a.row; r <= a.last_row; ++r) {
      for (std::uint32_t c = a.col; c <= a.last_col; ++c) {
        const Cell* cell = sheet.cell_at(r, c);
        values.push_back(cell != nullptr ? cell->cached_value : Value::blank());
      }
    }
    // Blank the non-anchor cells: their values now live in the region as
    // phantoms, and blanking keeps them from blocking either this commit's
    // collision scan or the anchor's re-spill on recalc.
    for (std::uint32_t r = a.row; r <= a.last_row; ++r) {
      for (std::uint32_t c = a.col; c <= a.last_col; ++c) {
        if (r == a.row && c == a.col) {
          continue;
        }
        sheet.set_cell_cached_value(r, c, Value::blank());
      }
    }
    sheet.commit_spill(a.row, a.col, rows, cols, std::move(values));
  }
}

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
                                     std::unordered_map<std::uint32_t, SharedFormulaMaster>& shared, std::uint32_t row,
                                     std::uint32_t col, std::string& formula_out) {
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
    shared[si] = SharedFormulaMaster{body, row, col};
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
  formula_out = ShiftSharedFormulaText(it->second, row, col);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> ApplyParsedCell(const ParsedCell& parsed, std::string_view formula_text, std::uint32_t xf_index,
                                      std::string_view phonetic_text, std::size_t sheet_index, Workbook& workbook,
                                      SheetReadContext& ctx) {
  if (!formula_text.empty()) {
    // `Workbook::set_cell_formula` accepts both spellings, but to
    // match the parser/evaluator's expected input form (the existing
    // call sites in workbook_recalc_test.cpp pass "=A1*2") we prepend
    // '=' here.
    std::string with_eq("=");
    with_eq.append(formula_text);
    auto wf = workbook.set_cell_formula(sheet_index, parsed.row, parsed.col, std::move(with_eq));
    if (!wf) {
      return wf.error();
    }
    // Preserve Excel's cached result until a caller explicitly recalculates.
    // This is essential when a workbook uses functions Formulon does not yet
    // implement: eagerly replacing a valid loaded cache with #NAME? makes a
    // read-only inspection or save/load round-trip lose useful data.
    if (!parsed.value.is_blank() && !parsed.is_sst_index) {
      workbook.sheet(sheet_index).set_cell_cached_value(parsed.row, parsed.col, parsed.value);
    }
  } else if (parsed.value.is_blank()) {
    // Skip blank-blank cells to keep the row map sparse, unless a style
    // index exists: then the format is the payload and must materialise
    // the cell below.
    if (parsed.is_sst_index) {
      ctx.pending_sst_cells.emplace_back(parsed.row, parsed.col, parsed.sst_index);
    }
    if (xf_index == 0U) {
      return Expected<void, Error>::Ok();
    }
  } else {
    auto wv = workbook.set_cell_value(sheet_index, parsed.row, parsed.col, parsed.value);
    if (!wv) {
      return wv.error();
    }
  }

  if (parsed.is_sst_index) {
    ctx.pending_sst_cells.emplace_back(parsed.row, parsed.col, parsed.sst_index);
  }

  if (xf_index != 0U) {
    auto sx = workbook.set_cell_xf_index(sheet_index, parsed.row, parsed.col, xf_index);
    if (!sx) {
      return sx.error();
    }
  }

  if (!parsed.is_sst_index && !phonetic_text.empty()) {
    workbook.sheet(sheet_index).set_cell_phonetic(parsed.row, parsed.col, phonetic_text);
  }
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
        auto resolved = ResolveFormula(c, shared_formulas, parsed.row, parsed.col, formula_text);
        if (!resolved) {
          return resolved.error();
        }
      }

      auto applied =
          ApplyParsedCell(parsed, formula_text, parsed.xf_index, parsed.phonetic_text, sheet_index, workbook, ctx);
      if (!applied) {
        return applied.error();
      }

      // Record a dynamic-array anchor so its cached spill targets do not
      // read back as blocking literals (see `RegisterArraySpills`).
      if (pugi::xml_node f = c.child("f"); f && std::string_view(f.attribute("t").value()) == "array") {
        RecordArrayAnchor(ctx, f.attribute("ref").value(), parsed.row, parsed.col);
      }
    }
  }
  RegisterArraySpills(workbook.sheet(sheet_index), ctx.array_anchors);
  return Expected<void, Error>::Ok();
}

// ---------------------------------------------------------------------------
// View / layout helpers. Each lives in an anonymous namespace so the
// translation unit owns its parsing fences; the public driver
// `read_sheet_view_and_layout` composes them in document order.
// ---------------------------------------------------------------------------

namespace {

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
  if (sheet_view.attribute("zoomScale")) {
    const std::int32_t raw = attr_i32(sheet_view, "zoomScale", 0);
    if (raw >= 10 && raw <= 400) {
      view.zoom_scale = static_cast<std::uint32_t>(raw);
    }
  }
  // Display attributes. Three default to true in the schema, so absence
  // means "shown"; the tri-state reader applies each attribute's real
  // default rather than a blanket false.
  view.show_grid_lines = read_xsd_bool(sheet_view, "showGridLines", true);
  view.show_row_col_headers = read_xsd_bool(sheet_view, "showRowColHeaders", true);
  view.show_zeros = read_xsd_bool(sheet_view, "showZeros", true);
  view.right_to_left = read_xsd_bool(sheet_view, "rightToLeft", false);
  view.tab_selected = read_xsd_bool(sheet_view, "tabSelected", false);
  if (pugi::xml_attribute v = sheet_view.attribute("view"); v) {
    const std::string_view mode = v.value();
    // "normal" is the schema default; keep the model empty for it so a
    // default sheet stays byte-clean on re-emit.
    if (mode != "normal") {
      view.view_mode.assign(mode);
    }
  }
  pugi::xml_node pane = sheet_view.child("pane");
  if (pane) {
    if (attr_str(pane, "state") == "frozen") {
      const std::int32_t y = attr_i32(pane, "ySplit", 0);
      const std::int32_t x = attr_i32(pane, "xSplit", 0);
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
  if (attr_bool(sheet_pr, "tabHidden")) {
    view.tab_hidden = true;
  }
  if (pugi::xml_node child = sheet_pr.child("tabHidden"); child) {
    if (child.attribute("val")) {
      if (attr_bool(child, "val")) {
        view.tab_hidden = true;
      }
    } else {
      // Bare `<tabHidden/>` element with no `val` attribute is treated
      // as `val="1"` to match what Excel's older writers emit.
      view.tab_hidden = true;
    }
  }
}

/// Parses `<sheetFormatPr defaultColWidth defaultRowHeight baseColWidth/>`
/// into `defaults`. The element appears before `<cols>` in the worksheet
/// part. Absent attributes leave the corresponding fields at their
/// struct defaults; `defaultColWidth` / `defaultRowHeight` also set the
/// `has_*` flags so a consumer can distinguish an explicit `0` from an
/// absent attribute.
void ApplySheetFormatDefaults(const pugi::xml_node& worksheet, SheetFormatDefaults& defaults) {
  pugi::xml_node fmt = worksheet.child("sheetFormatPr");
  if (!fmt) {
    return;
  }
  if (pugi::xml_attribute attr = fmt.attribute("defaultColWidth"); attr) {
    defaults.default_col_width = std::strtod(attr.value(), nullptr);
    defaults.has_default_col_width = true;
  }
  if (pugi::xml_attribute attr = fmt.attribute("defaultRowHeight"); attr) {
    defaults.default_row_height = std::strtod(attr.value(), nullptr);
    defaults.has_default_row_height = true;
  }
  if (pugi::xml_attribute attr = fmt.attribute("baseColWidth"); attr) {
    defaults.base_col_width = std::strtod(attr.value(), nullptr);
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
    if (!col.attribute("min") || !col.attribute("max")) {
      continue;
    }
    const std::int32_t min_v = attr_i32(col, "min", 0);
    const std::int32_t max_v = attr_i32(col, "max", 0);
    if (min_v < 1 || max_v < min_v) {
      continue;
    }
    ColumnLayout entry;
    entry.first = static_cast<std::uint32_t>(min_v - 1);
    entry.last = static_cast<std::uint32_t>(max_v - 1);
    if (pugi::xml_attribute width_attr = col.attribute("width"); width_attr) {
      entry.width = attr_f64(col, "width");
    } else {
      // No explicit width override — skip. `customWidth` / `bestFit`
      // alone are not enough to round-trip through the layout model.
      continue;
    }
    if (pugi::xml_attribute hidden_attr = col.attribute("hidden"); hidden_attr) {
      entry.hidden = attr_bool(col, "hidden");
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
    const std::int32_t r_v = attr_i32(row, "r", 0);
    if (r_v < 1) {
      continue;
    }
    RowLayout entry;
    entry.row = static_cast<std::uint32_t>(r_v - 1);
    if (ht_attr) {
      entry.height = attr_f64(row, "ht");
    }
    if (hidden_attr) {
      entry.hidden = attr_bool(row, "hidden");
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
  ApplySheetFormatDefaults(worksheet, sheet.mutable_format_defaults());
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

std::uint32_t ParseXfIndex(std::string_view text) {
  std::uint32_t xf = 0;
  for (char c : text) {
    if (c < '0' || c > '9') {
      return 0U;
    }
    xf = (xf * 10U) + static_cast<std::uint32_t>(c - '0');
  }
  return xf;
}

/// Per-call state passed through `SheetSaxCallbacks::user_data`. The
/// scanner's callbacks need access to the workbook, the sheet index,
/// the SST queue, and the text-storage deque; bundling them here
/// avoids `std::function` capture entirely.
struct SaxApplyState {
  std::size_t sheet_index;
  Workbook* workbook;
  SheetReadContext* ctx;
  std::deque<std::string>* text_storage;
  // Shared-formula masters keyed by `si`, accumulated across the scan so
  // followers (`<f t="shared" si="N"/>`) can shift the master body into
  // their own position. Mirrors the DOM path's per-sheet `shared_formulas`
  // map (see `ResolveFormula`).
  std::unordered_map<std::uint32_t, SharedFormulaMaster> shared_formulas;
};

/// Resolves a SAX `CellRecord`'s `<f>` into the effective formula text
/// (leading `=` already stripped by the scanner), mirroring the DOM
/// `ResolveFormula`. Plain formulas pass through; shared-formula masters
/// register into `shared`; followers shift the registered master to their
/// cell. Array / dataTable formulas are treated as plain (body used
/// verbatim), matching the DOM path.
Expected<std::string, Error> ResolveSharedFromRecord(const CellRecord& rec,
                                                     std::unordered_map<std::uint32_t, SharedFormulaMaster>& shared) {
  if (rec.f_t != "shared") {
    return std::string(rec.formula);
  }
  if (rec.f_si.empty()) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "shared formula: <f t='shared'> missing 'si'",
                      "context=sheet_reader_sax");
  }
  std::uint32_t si = 0;
  for (const char ch : rec.f_si) {
    if (ch < '0' || ch > '9') {
      std::string ctx("context=sheet_reader_sax si=");
      ctx.append(rec.f_si);
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "shared formula: 'si' not a non-negative integer",
                        std::move(ctx));
    }
    si = (si * 10U) + static_cast<std::uint32_t>(ch - '0');
  }
  std::string body(rec.formula);
  if (!body.empty()) {
    // Master occurrence: register and use its body verbatim.
    shared[si] = SharedFormulaMaster{body, rec.row, rec.col};
    return body;
  }
  // Follower occurrence: shift the registered master into this cell.
  auto it = shared.find(si);
  if (it == shared.end()) {
    std::string ctx("context=sheet_reader_sax si=");
    ctx.append(std::to_string(si));
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "shared formula: slave references unknown si",
                      std::move(ctx));
  }
  return ShiftSharedFormulaText(it->second, rec.row, rec.col);
}

/// Translates one `CellRecord` into the same workbook mutations the
/// DOM path produces.
///
/// Shared-formula groups (`<f t="shared" si=>`) are resolved through
/// `shared` so followers recover the shifted master body — matching the
/// DOM path. Array / dataTable formulas are read as plain (body
/// verbatim).
Expected<void, Error> ApplyCellRecord(const CellRecord& rec, std::size_t sheet_index, Workbook& workbook,
                                      SheetReadContext& ctx, std::deque<std::string>& text_storage,
                                      std::unordered_map<std::uint32_t, SharedFormulaMaster>& shared) {
  const bool value_present = rec.is_inline_string || !rec.value.empty();
  auto payload_or = decode_cell_payload(rec.t, rec.value, value_present, rec.is_inline_string, text_storage);
  if (!payload_or) {
    return payload_or.error();
  }
  const ParsedCell& parsed = payload_or.value();
  ParsedCell cell = parsed;
  cell.row = rec.row;
  cell.col = rec.col;

  // Persist the cell's `s=` xf index when present. The SAX scanner
  // surfaces it as a string view; parse to integer here. Empty / "0"
  // collapses to the default sentinel and we skip the call to keep
  // the row map sparse.
  std::uint32_t xf = 0;
  if (!rec.s.empty()) {
    xf = ParseXfIndex(rec.s);
  }
  // Resolve shared-formula groups (plain formulas pass straight through).
  auto formula_or = ResolveSharedFromRecord(rec, shared);
  if (!formula_or) {
    return formula_or.error();
  }
  const std::string& formula_text = formula_or.value();
  // Inline-string cells with <rPh> annotations carry their kana on the
  // SAX record. SST-referenced cells (rec.phonetic stays empty by
  // construction) route their phonetic through the post-loop SST
  // resolution pass instead — same contract the DOM path uses.
  auto applied = ApplyParsedCell(cell, formula_text, xf, rec.phonetic, sheet_index, workbook, ctx);
  if (!applied) {
    return applied.error();
  }
  // Record a dynamic-array anchor so its cached spill targets do not read
  // back as blocking literals (see `RegisterArraySpills`).
  if (rec.f_t == "array" && !rec.f_ref.empty()) {
    RecordArrayAnchor(ctx, rec.f_ref, rec.row, rec.col);
  }
  return applied;
}

Expected<void, Error> SaxOnCellTrampoline(void* user_data, const CellRecord& rec) {
  auto* st = static_cast<SaxApplyState*>(user_data);
  return ApplyCellRecord(rec, st->sheet_index, *st->workbook, *st->ctx, *st->text_storage, st->shared_formulas);
}

// Captures per-row overrides on the SAX path, mirroring the DOM
// `ApplyRowOverrides`: a row contributes a `RowLayout` only when it
// carries `ht`, `hidden`, or `outlineLevel` (a bare `r=` row is a
// position marker). `customHeight` alone does not, matching the DOM path.
Expected<void, Error> SaxOnRowStartTrampoline(void* user_data, const RowRecord& rec) {
  auto* st = static_cast<SaxApplyState*>(user_data);
  if (rec.ht.empty() && rec.hidden.empty() && rec.outline_level.empty()) {
    return Expected<void, Error>::Ok();
  }
  if (rec.row_1based < 1U) {
    return Expected<void, Error>::Ok();
  }
  RowLayout entry;
  entry.row = rec.row_1based - 1U;
  if (!rec.ht.empty()) {
    entry.height = std::strtod(std::string(rec.ht).c_str(), nullptr);
  }
  if (!rec.hidden.empty()) {
    entry.hidden = parse_xml_bool(rec.hidden);
  }
  if (!rec.outline_level.empty()) {
    entry.outline_level = ParseOutlineLevel(std::string(rec.outline_level).c_str());
  }
  st->workbook->sheet(st->sheet_index).mutable_layout().row_overrides.push_back(entry);
  return Expected<void, Error>::Ok();
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
  SaxApplyState state{sheet_index, &workbook, &ctx, &text_storage, {}};
  SheetSaxCallbacks cb;
  cb.user_data = &state;
  cb.on_row_start = &SaxOnRowStartTrampoline;
  cb.on_cell = &SaxOnCellTrampoline;
  auto scanned = scan_sheet_data(sheet_xml, cb);
  if (!scanned) {
    return scanned.error();
  }
  RegisterArraySpills(workbook.sheet(sheet_index), ctx.array_anchors);
  return Expected<void, Error>::Ok();
}

#endif  // !FORMULON_WASM || FORMULON_WASM_ENABLE_SAX

namespace {

/// Decodes one A1-style range token (`A1`, `A1:B5`, or `$A$1:$B$5`)
/// into a `MergeRange`. Single-cell tokens land as `first == last`.
/// Both corners are normalised so `first <= last` componentwise.
Expected<MergeRange, Error> ParseA1RangeMerge(std::string_view ref) {
  const std::size_t colon = ref.find(':');
  if (colon == std::string_view::npos) {
    auto rc = parse_a1(ref);
    if (!rc) {
      std::string ctx("context=sheet_reader ref=");
      ctx.append(ref);
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "merge/hyperlink: ref token unparseable", std::move(ctx));
    }
    MergeRange out{};
    out.first_row = rc.value().first;
    out.first_col = rc.value().second;
    out.last_row = out.first_row;
    out.last_col = out.first_col;
    return out;
  }
  const std::string_view a = ref.substr(0, colon);
  const std::string_view b = ref.substr(colon + 1);
  auto a_rc = parse_a1(a);
  auto b_rc = parse_a1(b);
  if (!a_rc || !b_rc) {
    std::string ctx("context=sheet_reader ref=");
    ctx.append(ref);
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "merge: ref token unparseable", std::move(ctx));
  }
  const std::uint32_t r0 = a_rc.value().first;
  const std::uint32_t c0 = a_rc.value().second;
  const std::uint32_t r1 = b_rc.value().first;
  const std::uint32_t c1 = b_rc.value().second;
  MergeRange out{};
  out.first_row = (r0 < r1) ? r0 : r1;
  out.first_col = (c0 < c1) ? c0 : c1;
  out.last_row = (r0 < r1) ? r1 : r0;
  out.last_col = (c0 < c1) ? c1 : c0;
  return out;
}

/// Splits a whitespace-separated `sqref="A1 B2:C3 D4"` attribute and
/// decodes each token. Returns `kIoSheetCorrupt` on the first
/// unparseable token.
Expected<std::vector<MergeRange>, Error> ParseSqrefRanges(std::string_view sqref) {
  std::vector<MergeRange> out;
  std::size_t i = 0;
  while (i < sqref.size()) {
    while (i < sqref.size() && (sqref[i] == ' ' || sqref[i] == '\t' || sqref[i] == '\n' || sqref[i] == '\r')) {
      ++i;
    }
    const std::size_t start = i;
    while (i < sqref.size() && sqref[i] != ' ' && sqref[i] != '\t' && sqref[i] != '\n' && sqref[i] != '\r') {
      ++i;
    }
    if (start == i) {
      break;
    }
    auto r = ParseA1RangeMerge(sqref.substr(start, i - start));
    if (!r) {
      return r.error();
    }
    out.push_back(r.value());
  }
  return out;
}

}  // namespace

Expected<std::vector<MergeRange>, Error> read_merges(const pugi::xml_node& worksheet) {
  std::vector<MergeRange> out;
  if (!worksheet) {
    return out;
  }
  pugi::xml_node mc = worksheet.child("mergeCells");
  if (!mc) {
    return out;
  }
  for (pugi::xml_node m = mc.child("mergeCell"); m; m = m.next_sibling("mergeCell")) {
    const std::string_view ref = attr_str(m, "ref");
    if (ref.empty()) {
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "mergeCell: missing/empty ref",
                        "context=sheet_reader part=mergeCells");
    }
    auto r = ParseA1RangeMerge(ref);
    if (!r) {
      return r.error();
    }
    out.push_back(r.value());
  }
  return out;
}

Expected<std::vector<Hyperlink>, Error> read_hyperlinks(const pugi::xml_node& worksheet) {
  std::vector<Hyperlink> out;
  if (!worksheet) {
    return out;
  }
  pugi::xml_node node = worksheet.child("hyperlinks");
  if (!node) {
    return out;
  }
  for (pugi::xml_node h = node.child("hyperlink"); h; h = h.next_sibling("hyperlink")) {
    const std::string_view ref = attr_str(h, "ref");
    if (ref.empty()) {
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "hyperlink: missing/empty ref",
                        "context=sheet_reader part=hyperlinks");
    }
    // Hyperlinks may span a range (`ref="A1:B2"`); Excel applies the link
    // to every cell. We anchor `row`/`col` at the range's top-left and
    // preserve the full ref in `ref_span` for verbatim re-emission. A
    // single-cell ref leaves `ref_span` empty.
    const std::size_t colon = ref.find(':');
    const std::string_view anchor = colon == std::string_view::npos ? ref : ref.substr(0, colon);
    auto rc = parse_a1(anchor);
    if (!rc) {
      std::string ctx("context=sheet_reader part=hyperlinks ref=");
      ctx.append(ref);
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "hyperlink: ref anchor unparseable", std::move(ctx));
    }
    Hyperlink hl;
    hl.row = rc.value().first;
    hl.col = rc.value().second;
    if (colon != std::string_view::npos) {
      hl.ref_span.assign(ref);
    }
    // Accept both Office-namespaced ("r:id") and bare "id" attribute spellings.
    const std::string_view rid_v = attr_str(h, "r:id");
    if (!rid_v.empty()) {
      hl.rid.assign(rid_v);
    } else {
      hl.rid.assign(attr_str(h, "id"));
    }
    hl.location.assign(attr_str(h, "location"));
    hl.display.assign(attr_str(h, "display"));
    hl.tooltip.assign(attr_str(h, "tooltip"));
    out.push_back(std::move(hl));
  }
  return out;
}

void apply_hyperlink_rels(std::vector<Hyperlink>& hyperlinks,
                          const std::unordered_map<std::string, std::string>& rid_to_target) {
  for (Hyperlink& h : hyperlinks) {
    if (h.rid.empty()) {
      continue;
    }
    auto it = rid_to_target.find(h.rid);
    if (it == rid_to_target.end()) {
      continue;
    }
    h.target = it->second;
  }
}

Expected<std::vector<DataValidation>, Error> read_data_validations(const pugi::xml_node& worksheet) {
  std::vector<DataValidation> out;
  if (!worksheet) {
    return out;
  }
  pugi::xml_node dvs = worksheet.child("dataValidations");
  if (!dvs) {
    return out;
  }
  for (pugi::xml_node dv = dvs.child("dataValidation"); dv; dv = dv.next_sibling("dataValidation")) {
    DataValidation v;
    const std::string_view sqref = attr_str(dv, "sqref");
    if (sqref.empty()) {
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "dataValidation: missing/empty sqref",
                        "context=sheet_reader part=dataValidations");
    }
    auto ranges_or = ParseSqrefRanges(sqref);
    if (!ranges_or) {
      return ranges_or.error();
    }
    v.ranges = std::move(ranges_or.value());

    // type
    const std::string_view t = attr_str(dv, "type");
    if (t == "whole") {
      v.type = 1;
    } else if (t == "decimal") {
      v.type = 2;
    } else if (t == "list") {
      v.type = 3;
    } else if (t == "date") {
      v.type = 4;
    } else if (t == "time") {
      v.type = 5;
    } else if (t == "textLength") {
      v.type = 6;
    } else if (t == "custom") {
      v.type = 7;
    } else {
      v.type = 0;
    }

    // operator
    const std::string_view op = attr_str(dv, "operator");
    if (op == "notBetween") {
      v.op = 1;
    } else if (op == "equal") {
      v.op = 2;
    } else if (op == "notEqual") {
      v.op = 3;
    } else if (op == "greaterThan") {
      v.op = 4;
    } else if (op == "lessThan") {
      v.op = 5;
    } else if (op == "greaterThanOrEqual") {
      v.op = 6;
    } else if (op == "lessThanOrEqual") {
      v.op = 7;
    } else {
      v.op = 0;  // "between" / unspecified
    }

    // errorStyle
    const std::string_view es = attr_str(dv, "errorStyle");
    if (es == "warning") {
      v.error_style = 1;
    } else if (es == "information") {
      v.error_style = 2;
    } else {
      v.error_style = 0;
    }

    // boolean attributes: allowBlank defaults to true (Excel convention),
    // input/error message visibility default to false.
    if (pugi::xml_attribute ab = dv.attribute("allowBlank"); ab) {
      const std::string_view sv = ab.value();
      v.allow_blank = !(sv == "0" || sv == "false");
    }
    if (pugi::xml_attribute sim = dv.attribute("showInputMessage"); sim) {
      v.show_input_message = parse_xml_bool(sim.value());
    }
    if (pugi::xml_attribute sem = dv.attribute("showErrorMessage"); sem) {
      v.show_error_message = parse_xml_bool(sem.value());
    }
    // `showDropDown` has inverted semantics: presence with a true value
    // suppresses the arrow, so the user-facing `show_dropdown` is the
    // negation of the raw attribute (absent/false attribute => shown).
    if (pugi::xml_attribute sdd = dv.attribute("showDropDown"); sdd) {
      v.show_dropdown = !parse_xml_bool(sdd.value());
    }

    v.error_title.assign(attr_str(dv, "errorTitle"));
    v.error_message.assign(attr_str(dv, "error"));
    v.prompt_title.assign(attr_str(dv, "promptTitle"));
    v.prompt_message.assign(attr_str(dv, "prompt"));

    if (pugi::xml_node f1 = dv.child("formula1"); f1) {
      v.formula1.assign(f1.text().get());
      if (!v.formula1.empty() && v.formula1.front() == '=') {
        v.formula1.erase(0, 1);
      }
    }
    if (pugi::xml_node f2 = dv.child("formula2"); f2) {
      v.formula2.assign(f2.text().get());
      if (!v.formula2.empty() && v.formula2.front() == '=') {
        v.formula2.erase(0, 1);
      }
    }

    out.push_back(std::move(v));
  }
  return out;
}

SheetProtection read_sheet_protection(const pugi::xml_node& worksheet) {
  SheetProtection out;
  if (!worksheet) {
    return out;
  }
  const pugi::xml_node node = worksheet.child("sheetProtection");
  if (!node) {
    return out;
  }
  out.enabled = true;

  out.algorithm_name.assign(attr_str(node, "algorithmName"));
  out.hash_value.assign(attr_str(node, "hashValue"));
  out.salt_value.assign(attr_str(node, "saltValue"));
  if (pugi::xml_attribute sc = node.attribute("spinCount"); sc) {
    const long long parsed = sc.as_llong(0);
    out.spin_count = parsed < 0 ? 0U : static_cast<std::uint32_t>(parsed);
  }
  out.legacy_password.assign(attr_str(node, "password"));

  // Per-attribute XSD defaults (ECMA-376 §18.3.1.85). Eleven action flags
  // (format*, insert*, delete*, sort, autoFilter, pivotTables) default to
  // TRUE (locked); the rest default to FALSE. Reading with the wrong
  // default silently under-reports protection when Excel omits an
  // at-default attribute, so use the tri-state reader with each attribute's
  // real default rather than a blanket `false`.
  out.sheet = read_xsd_bool(node, "sheet", false);
  out.objects = read_xsd_bool(node, "objects", false);
  out.scenarios = read_xsd_bool(node, "scenarios", false);
  out.format_cells = read_xsd_bool(node, "formatCells", true);
  out.format_columns = read_xsd_bool(node, "formatColumns", true);
  out.format_rows = read_xsd_bool(node, "formatRows", true);
  out.insert_columns = read_xsd_bool(node, "insertColumns", true);
  out.insert_rows = read_xsd_bool(node, "insertRows", true);
  out.insert_hyperlinks = read_xsd_bool(node, "insertHyperlinks", true);
  out.delete_columns = read_xsd_bool(node, "deleteColumns", true);
  out.delete_rows = read_xsd_bool(node, "deleteRows", true);
  out.select_locked_cells = read_xsd_bool(node, "selectLockedCells", false);
  out.select_unlocked_cells = read_xsd_bool(node, "selectUnlockedCells", false);
  out.sort = read_xsd_bool(node, "sort", true);
  out.auto_filter = read_xsd_bool(node, "autoFilter", true);
  out.pivot_tables = read_xsd_bool(node, "pivotTables", true);

  return out;
}

}  // namespace io
}  // namespace formulon
