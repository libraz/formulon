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
// Cached-value handling (called out in the public header):
//   * A non-blank `<v>` on a formula cell is kept as the cell's value.
//     `Workbook::recalc()` replaces it with the engine's own result;
//     until then the loaded workbook reports what Excel last computed.

#include "io/sheet_reader.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <deque>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include "io/array_anchor_budget.h"
#include "io/cell_parser.h"
#include "io/cf_reader.h"
#include "io/future_functions.h"
#include "io/sax_xml_reader.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "io/xsd_bool.h"
#include "io/xsd_double.h"
#include "io/xsd_int.h"
#include "parser/ast_format.h"
#include "parser/ast_shift.h"
#include "parser/formula_prefix.h"
#include "parser/parser.h"
#include "phonetic.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/date_time.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"
#include "utils/status_macros.h"
#include "utils/structured_log.h"
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

/// Predicts how many new `Cell` slots a write to `(row, col)` will cost
/// `RowCells::ensure()`, without reaching into `RowCells` itself -- the
/// reader only sees `(row, col)` pairs in XML document order, so this
/// mirrors the run's contiguous-growth rule from the outside: a write to
/// a new row costs 1 slot, a write outside the row's already-seen
/// `[min_col, max_col]` span costs the gap it bridges (matching a styled
/// far-apart pair of cells materialising every slot between them), and a
/// write inside that span costs nothing (already materialised). Shared
/// by both the DOM and SAX read paths so each charges the same sheet-load
/// budget for the same input.
class RowGrowthTracker {
 public:
  std::uint64_t charge_for(std::uint32_t row, std::uint32_t col) {
    if (!row_.has_value() || *row_ != row) {
      row_ = row;
      min_col_ = col;
      max_col_ = col;
      return 1U;
    }
    if (col < min_col_) {
      const std::uint64_t growth = static_cast<std::uint64_t>(min_col_) - col;
      min_col_ = col;
      return growth;
    }
    if (col > max_col_) {
      const std::uint64_t growth = static_cast<std::uint64_t>(col) - max_col_;
      max_col_ = col;
      return growth;
    }
    return 0U;
  }

 private:
  std::optional<std::uint32_t> row_;
  std::uint32_t min_col_ = 0;
  std::uint32_t max_col_ = 0;
};

/// Decodes a shared-formula `si` attribute under the shared XSD
/// non-negative-integer lexer.
///
/// A malformed `si` is a hard error on both read paths: the index is what
/// binds a follower cell to its group master, so accepting a prefix or
/// defaulting it would either drop the follower's formula or attach it to
/// an unrelated group — a wrong number rather than a wrong format.
Expected<std::uint32_t, Error> ParseSharedFormulaSi(std::string_view raw) {
  std::uint32_t si = 0;
  if (!parse_xsd_nonneg_int(raw, &si)) {
    std::string ctx("context=sheet_reader si=");
    ctx.append(raw);
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "shared formula: 'si' is not a non-negative integer",
                      std::move(ctx));
  }
  return si;
}

std::string ShiftSharedFormulaText(const SharedFormulaMaster& master, std::uint32_t target_row,
                                   std::uint32_t target_col) {
  const std::int32_t row_delta = static_cast<std::int32_t>(target_row) - static_cast<std::int32_t>(master.row);
  const std::int32_t col_delta = static_cast<std::int32_t>(target_col) - static_cast<std::int32_t>(master.col);
  if (row_delta == 0 && col_delta == 0) {
    return master.text;
  }

  // `master.text` is the raw `<f>` body, which may still carry Excel's
  // `_xlfn.` / `_xlpm.` storage prefixes (the master cell's own text is
  // canonicalised later, by `Workbook::set_cell_formula`). Strip them
  // before parsing: the general Call-node path only recognises those
  // prefixes for LET / LAMBDA specially, so an un-stripped future-function
  // callee would shift and re-format with the prefix still glued to its
  // name, and the prefix bytes also count against the tokenizer's
  // formula-length cap even though Excel's own limit is measured on the
  // canonical (formula-bar) text this produces.
  std::string source("=");
  source.append(parser::strip_storage_prefixes(master.text, &has_storage_prefix));
  Arena arena(/*initial_chunk_bytes=*/4096, kMaxLoadArenaBytes);
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

/// Records a dynamic-array anchor for a `<f t="array" ref="...">`. The
/// OOXML ref must be an ordered rectangle whose top-left corner is the
/// formula cell. One-cell refs are retained because they carry dynamic-array
/// metadata that must survive an XLSB write, even with no phantom cells.
void RecordArrayAnchor(SheetReadContext& ctx, std::string_view ref, std::uint32_t anchor_row,
                       std::uint32_t anchor_col) {
  const std::size_t colon = ref.find(':');
  if (colon == std::string_view::npos) {
    auto anchor = parse_a1(ref);
    if (anchor && anchor.value().first == anchor_row && anchor.value().second == anchor_col) {
      ctx.array_anchors.push_back(ArrayAnchor{anchor_row, anchor_col, anchor_row, anchor_col});
    }
    return;
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
  ctx.array_anchors.push_back(ArrayAnchor{anchor_row, anchor_col, last_row, last_col});
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
///   * `<f t="array">BODY</f>` (CSE array) is accepted but treated as a
///     plain formula: we read the body as the formula text. CSE-array
///     detail will land in a later bundle.
///   * `<f t="dataTable" .../>` carries no body at all (its geometry
///     lives in `r1`/`r2`/`dt2D`/`dtr` attributes, which are not read),
///     so `formula_out` comes out empty and the cell falls back to its
///     cached `<v>` like any formula-less cell. `diagnostics` (when
///     non-null) counts the drop via `skipped_feature_count` so a
///     What-If data table is never silently frozen into constants with
///     no record of the loss.
Expected<void, Error> ResolveFormula(const pugi::xml_node& c_node,
                                     std::unordered_map<std::uint32_t, SharedFormulaMaster>& shared, std::uint32_t row,
                                     std::uint32_t col, std::string& formula_out, ReadDiagnostics* diagnostics) {
  pugi::xml_node f_node = c_node.child("f");
  if (!f_node) {
    formula_out.clear();
    return Expected<void, Error>::Ok();
  }
  const std::string_view ftype = f_node.attribute("t").value();
  if (ftype == "dataTable") {
    formula_out.clear();
    StructuredLog("io.sheet.formula.data_table_skip")
        .field("row", std::to_string(row))
        .field("col", std::to_string(col))
        .error_code(FormulonErrorCode::kIoSheetCorrupt)
        .warn();
    if (diagnostics != nullptr) {
      ++diagnostics->skipped_feature_count;
    }
    return Expected<void, Error>::Ok();
  }
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
  ASSIGN_OR_RETURN(const std::uint32_t si, ParseSharedFormulaSi(si_attr.value()));

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
                                      const std::vector<PhoneticRun>* phonetic_runs, PhoneticProperties phonetic_props,
                                      std::size_t sheet_index, Workbook& workbook, SheetReadContext& ctx) {
  // `decode_cell_payload` deliberately produces all ISO dates on the
  // canonical 1900 serial axis so DOM and SAX share one pure decoder. The
  // workbook epoch is applied exactly once at this shared storage boundary,
  // and only for a successfully parsed `t="d"` payload. Ordinary numeric
  // literals with the same value must keep their original serial unchanged.
  Value stored_value = parsed.value;
  if (workbook.date1904() && parsed.is_iso_date && stored_value.is_number()) {
    stored_value = Value::number(stored_value.as_number() - date_time::kDate1904EpochGap);
  }

  // A `<c s="900">` against a five-entry `<cellXfs>` names no style. Fall
  // back to the default xf so the loaded workbook stays self-consistent:
  // `fm_cell_get_xf` resolves for every cell that loaded, and a save does
  // not re-emit the dangling index for Excel to repair. This is the same
  // "a style index is cosmetic" disposition `cell_parser.cpp` applies to a
  // lexically malformed `s=`. When the package carried no styles part at
  // all there is no table to dangle against and the raw index is kept —
  // the writer's own bound check covers what it synthesises.
  const std::size_t cell_xf_count = workbook.styles().cell_xfs.size();
  if (cell_xf_count != 0U && xf_index >= cell_xf_count) {
    xf_index = 0U;
  }
  if (!formula_text.empty()) {
    // `Workbook::set_cell_formula` accepts both spellings, but to
    // match the parser/evaluator's expected input form (the existing
    // call sites in workbook_recalc_test.cpp pass "=A1*2") we prepend
    // '=' here.
    std::string with_eq("=");
    with_eq.append(formula_text);
    auto wf = workbook.set_cell_formula(sheet_index, parsed.row, parsed.col, workbook.ingest_stored_formula(with_eq));
    if (!wf) {
      return wf.error();
    }
    // Preserve Excel's cached result until a caller explicitly recalculates.
    // This is essential when a workbook uses functions Formulon does not yet
    // implement: eagerly replacing a valid loaded cache with #NAME? makes a
    // read-only inspection or save/load round-trip lose useful data.
    if (!stored_value.is_blank() && !parsed.is_sst_index) {
      workbook.sheet(sheet_index).set_cell_cached_value_borrowed(parsed.row, parsed.col, stored_value);
    }
  } else if (stored_value.is_blank()) {
    // Skip blank-blank cells to keep the row map sparse, unless a style
    // index or explicit style attribute exists: then the format metadata
    // is the payload and must materialise the cell below.
    if (parsed.is_sst_index) {
      ctx.pending_sst_cells.emplace_back(parsed.row, parsed.col, parsed.sst_index);
    }
    if (xf_index == 0U && !parsed.has_explicit_xf) {
      return Expected<void, Error>::Ok();
    }
  } else {
    // A literal lands on a freshly built sheet, whose formulas are all dirty
    // from registration, so write it straight to the sheet as the XLSB reader
    // does. Routing it through `Workbook::set_cell_value` would revisit every
    // formula watching a rectangle that covers it. Only a duplicate `<c>`
    // replacing an earlier formula needs the workbook to drop its edges.
    Sheet& sheet = workbook.sheet(sheet_index);
    const Cell* existing = sheet.cell_at(parsed.row, parsed.col);
    if (existing != nullptr && !existing->formula_text.empty()) {
      auto wv = workbook.set_cell_value(sheet_index, parsed.row, parsed.col, stored_value);
      if (!wv) {
        return wv.error();
      }
    } else {
      sheet.set_cell_value(parsed.row, parsed.col, stored_value);
    }
  }

  if (parsed.is_sst_index) {
    ctx.pending_sst_cells.emplace_back(parsed.row, parsed.col, parsed.sst_index);
  }

  if (xf_index != 0U || parsed.has_explicit_xf) {
    auto sx = workbook.set_cell_xf_index(sheet_index, parsed.row, parsed.col, xf_index);
    if (!sx) {
      return sx.error();
    }
  }

  if (!parsed.is_sst_index && phonetic_runs != nullptr && !phonetic_runs->empty()) {
    workbook.sheet(sheet_index).set_cell_phonetic_runs(parsed.row, parsed.col, *phonetic_runs);
    workbook.sheet(sheet_index).set_cell_phonetic_props(parsed.row, parsed.col, phonetic_props);
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<void, Error> RegisterArraySpills(Sheet& sheet, const std::vector<ArrayAnchor>& anchors) {
  return register_array_spills(sheet, anchors, FormulonErrorCode::kIoSheetCorrupt, "sheet_reader", "ooxml");
}

Expected<void, Error> read_sheet_data(const pugi::xml_document& sheet_doc, std::size_t sheet_index, Workbook& workbook,
                                      SheetReadContext& ctx, std::deque<std::string>& text_storage,
                                      ReadDiagnostics* diagnostics, std::uint64_t max_cell_bytes) {
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
  // Charged before every cell materialises, so a row whose two styled
  // cells sit far apart cannot force `RowCells::ensure()` to allocate
  // every slot between them unbounded — see `RowGrowthTracker`.
  ResourceBudget cell_budget(max_cell_bytes, FormulonErrorCode::kIoFileTooLarge);
  RowGrowthTracker cell_growth;

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
        auto resolved = ResolveFormula(c, shared_formulas, parsed.row, parsed.col, formula_text, diagnostics);
        if (!resolved) {
          return resolved.error();
        }
      }

      // Charge before `ApplyParsedCell` materialises: a truly empty cell
      // (matching its own no-op condition below) never calls
      // `RowCells::ensure()`, so it must not advance `cell_growth` either
      // -- otherwise a later real write in the same row would be charged
      // as if the gap it bridges were smaller than it actually is.
      const bool materializes = !formula_text.empty() || !parsed.value.is_blank() || parsed.is_sst_index ||
                                parsed.xf_index != 0U || parsed.has_explicit_xf;
      if (materializes) {
        const std::uint64_t growth = cell_growth.charge_for(parsed.row, parsed.col);
        if (growth != 0U) {
          std::string ctxs("context=sheet_reader row=");
          ctxs.append(std::to_string(parsed.row));
          ctxs.append(" col=");
          ctxs.append(std::to_string(parsed.col));
          auto charged = charge(cell_budget, growth * sizeof(Cell), std::move(ctxs));
          if (!charged) {
            return charged.error();
          }
        }
      }

      auto applied = ApplyParsedCell(parsed, formula_text, parsed.xf_index, &parsed.phonetic_runs,
                                     parsed.phonetic_props, sheet_index, workbook, ctx);
      if (!applied) {
        return applied.error();
      }

      // Record a dynamic-array anchor so its cached spill targets do not
      // read back as blocking literals (see `RegisterArraySpills`).
      if (pugi::xml_node f = c.child("f"); f && std::string_view(f.attribute("t").value()) == "array") {
        RecordArrayAnchor(ctx, f.attribute("ref").value(), parsed.row, parsed.col);
      }
      if (const pugi::xml_attribute cm = c.attribute("cm"); cm && !formula_text.empty()) {
        ctx.cell_metadata.emplace_back(parsed.row, parsed.col, cm.as_uint(0U));
      }
    }
  }
  // Spill registration is the caller's job (see `SheetReadContext::
  // array_anchors`): it must run after `ctx.pending_sst_cells` has been
  // resolved, which this function does not do.
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
  // Shared-formula masters keyed by `si`, accumulated across the scan so
  // followers (`<f t="shared" si="N"/>`) can shift the master body into
  // their own position. Mirrors the DOM path's per-sheet `shared_formulas`
  // map (see `ResolveFormula`).
  std::unordered_map<std::uint32_t, SharedFormulaMaster> shared_formulas;
  ReadDiagnostics* diagnostics;
  // Mirrors the DOM path's per-sheet cell-materialisation budget (see
  // `read_sheet_data`), so the two loaders enforce the same ceiling.
  ResourceBudget cell_budget;
  RowGrowthTracker cell_growth;
};

/// Resolves a SAX `CellRecord`'s `<f>` into the effective formula text
/// (leading `=` already stripped by the scanner), mirroring the DOM
/// `ResolveFormula`. Plain formulas pass through; shared-formula masters
/// register into `shared`; followers shift the registered master to their
/// cell. Array formulas are treated as plain (body used verbatim),
/// matching the DOM path. `<f t="dataTable">` also comes back empty here
/// (the scanner captures no body for it, same as the DOM path reading
/// no text), so the diagnostics side-effect lives in `ApplyCellRecord`
/// where `rec.f_t` is still available for the check.
Expected<std::string, Error> ResolveSharedFromRecord(const CellRecord& rec,
                                                     std::unordered_map<std::uint32_t, SharedFormulaMaster>& shared) {
  if (rec.f_t != "shared") {
    return std::string(rec.formula);
  }
  if (rec.f_si.empty()) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "shared formula: <f t='shared'> missing 'si'",
                      "context=sheet_reader_sax");
  }
  ASSIGN_OR_RETURN(const std::uint32_t si, ParseSharedFormulaSi(rec.f_si));
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
/// DOM path. Array formulas are read as plain (body verbatim).
/// `<f t="dataTable">` also reads as an empty body (matching the DOM
/// path's fallback to the cached `<v>`), but is additionally counted
/// via `diagnostics->skipped_feature_count` so the drop is never silent
/// -- see `ResolveFormula`.
Expected<void, Error> ApplyCellRecord(const CellRecord& rec, std::size_t sheet_index, Workbook& workbook,
                                      SheetReadContext& ctx, std::deque<std::string>& text_storage,
                                      std::unordered_map<std::uint32_t, SharedFormulaMaster>& shared,
                                      ReadDiagnostics* diagnostics, ResourceBudget& cell_budget,
                                      RowGrowthTracker& cell_growth) {
  const bool value_present = rec.is_inline_string || !rec.value.empty();
  ASSIGN_OR_RETURN(auto payload,
                   decode_cell_payload(rec.t, rec.value, value_present, rec.is_inline_string, text_storage));
  const ParsedCell& parsed = payload;
  ParsedCell cell = parsed;
  cell.row = rec.row;
  cell.col = rec.col;
  cell.has_explicit_xf = !rec.s.empty();

  // Persist the cell's `s=` xf index when present. The SAX scanner
  // surfaces it as a string view; parse to integer here through the same
  // lexer and with the same degrade-to-0 disposition the DOM cell parser
  // uses. An explicit non-empty `s=` remains meaningful even when it
  // normalises to the default sentinel `0`.
  std::uint32_t xf = 0;
  if (!parse_xsd_nonneg_int(rec.s, &xf)) {
    xf = 0;
  }
  if (rec.f_t == "dataTable") {
    StructuredLog("io.sheet.formula.data_table_skip")
        .field("row", std::to_string(rec.row))
        .field("col", std::to_string(rec.col))
        .error_code(FormulonErrorCode::kIoSheetCorrupt)
        .warn();
    if (diagnostics != nullptr) {
      ++diagnostics->skipped_feature_count;
    }
  }
  // Resolve shared-formula groups (plain formulas pass straight through).
  ASSIGN_OR_RETURN(auto formula, ResolveSharedFromRecord(rec, shared));
  const std::string& formula_text = formula;

  // Charge before `ApplyParsedCell` materialises -- mirrors the DOM
  // path's placement and no-op guard exactly (see `read_sheet_data`), so
  // a truly empty record never advances `cell_growth`.
  const bool materializes =
      !formula_text.empty() || !parsed.value.is_blank() || parsed.is_sst_index || xf != 0U || cell.has_explicit_xf;
  if (materializes) {
    const std::uint64_t growth = cell_growth.charge_for(rec.row, rec.col);
    if (growth != 0U) {
      std::string ctxs("context=sheet_reader_sax row=");
      ctxs.append(std::to_string(rec.row));
      ctxs.append(" col=");
      ctxs.append(std::to_string(rec.col));
      auto charged = charge(cell_budget, growth * sizeof(Cell), std::move(ctxs));
      if (!charged) {
        return charged.error();
      }
    }
  }
  // Inline-string cells with <rPh> annotations carry their kana on the
  // SAX record. SST-referenced cells (rec.phonetic stays null by
  // construction) route their phonetic through the post-loop SST
  // resolution pass instead — same contract the DOM path uses.
  auto applied = ApplyParsedCell(cell, formula_text, xf, rec.phonetic, rec.phonetic_props, sheet_index, workbook, ctx);
  if (!applied) {
    return applied.error();
  }
  // Record a dynamic-array anchor so its cached spill targets do not read
  // back as blocking literals (see `RegisterArraySpills`).
  if (rec.f_t == "array" && !rec.f_ref.empty()) {
    RecordArrayAnchor(ctx, rec.f_ref, rec.row, rec.col);
  }
  if (!rec.cm.empty() && !formula_text.empty()) {
    std::uint32_t cm = 0;
    for (const char c : rec.cm) {
      cm = (c >= '0' && c <= '9') ? cm * 10U + static_cast<std::uint32_t>(c - '0') : 0U;
    }
    ctx.cell_metadata.emplace_back(rec.row, rec.col, cm);
  }
  return applied;
}

Expected<void, Error> SaxOnCellTrampoline(void* user_data, const CellRecord& rec) {
  auto* st = static_cast<SaxApplyState*>(user_data);
  return ApplyCellRecord(rec, st->sheet_index, *st->workbook, *st->ctx, *st->text_storage, st->shared_formulas,
                         st->diagnostics, st->cell_budget, st->cell_growth);
}

// Captures per-row overrides on the SAX path, mirroring the DOM
// `ApplyRowOverrides`: a row contributes a `RowLayout` when it carries
// `ht`, `hidden`, `outlineLevel`, or effective `customFormat=1` style
// metadata. `customHeight` alone does not, matching the DOM path.
Expected<void, Error> SaxOnRowStartTrampoline(void* user_data, const RowRecord& rec) {
  auto* st = static_cast<SaxApplyState*>(user_data);
  const bool custom_format = !rec.custom_format.empty() && parse_xsd_bool(rec.custom_format, false);
  double height = 0.0;
  const bool has_height = parse_xsd_nonneg_double(rec.ht, &height);
  if (!has_height && rec.hidden.empty() && rec.outline_level.empty() && !custom_format) {
    return Expected<void, Error>::Ok();
  }
  if (rec.row_1based < 1U) {
    return Expected<void, Error>::Ok();
  }
  RowLayout entry;
  entry.row = rec.row_1based - 1U;
  if (has_height) {
    entry.height = height;
    entry.has_height = true;
    entry.custom_height = parse_xsd_bool(rec.custom_height, false);
  }
  if (!rec.hidden.empty()) {
    entry.hidden = parse_xsd_bool(rec.hidden, false);
  }
  if (!rec.outline_level.empty()) {
    entry.outline_level = parse_outline_level(std::string(rec.outline_level).c_str());
  }
  if (custom_format) {
    entry.has_style = true;
    std::uint32_t style_xf = 0;
    if (!rec.style.empty()) {
      (void)parse_xsd_nonneg_int(rec.style, &style_xf);
    }
    entry.style_xf = style_xf;
  }
  st->workbook->sheet(st->sheet_index).mutable_layout().row_overrides.push_back(entry);
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<void, Error> read_sheet_data_sax(ByteSpan sheet_xml, std::size_t sheet_index, Workbook& workbook,
                                          SheetReadContext& ctx, std::deque<std::string>& text_storage,
                                          ReadDiagnostics* diagnostics, std::uint64_t max_cell_bytes) {
  if (sheet_index >= workbook.sheet_count()) {
    std::string ctxs("context=sheet_reader_sax sheet_index=");
    ctxs.append(std::to_string(sheet_index));
    ctxs.append(" sheet_count=");
    ctxs.append(std::to_string(workbook.sheet_count()));
    return make_error(FormulonErrorCode::kInvalidArgument, "read_sheet_data_sax: sheet_index out of range",
                      std::move(ctxs));
  }
  SaxApplyState state{sheet_index,
                      &workbook,
                      &ctx,
                      &text_storage,
                      {},
                      diagnostics,
                      ResourceBudget(max_cell_bytes, FormulonErrorCode::kIoFileTooLarge),
                      {}};
  SheetSaxCallbacks cb;
  cb.user_data = &state;
  cb.on_row_start = &SaxOnRowStartTrampoline;
  cb.on_cell = &SaxOnCellTrampoline;
  auto scanned = scan_sheet_data(sheet_xml, cb);
  if (!scanned) {
    return scanned.error();
  }
  // Spill registration is the caller's job (see `SheetReadContext::
  // array_anchors`): it must run after `ctx.pending_sst_cells` has been
  // resolved, which this function does not do.
  return Expected<void, Error>::Ok();
}

#endif  // !FORMULON_WASM || FORMULON_WASM_ENABLE_SAX

}  // namespace io
}  // namespace formulon
