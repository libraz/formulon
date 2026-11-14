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
#include <memory>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include "io/array_anchor_budget.h"
#include "io/cell_parser.h"
#include "io/cf_reader.h"
#include "io/data_table_xml.h"
#include "io/future_functions.h"
#include "io/sax_xml_reader.h"
#include "io/sheet_layout_reader.h"
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
#include "sheet_layout_edit.h"
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
  /// Master body parsed once and reused for every follower; lives on the
  /// heap so the parser's views into `source` stay valid when the record moves.
  struct Parsed {
    std::string source;
    Arena arena{kDefaultArenaChunkBytes, kMaxLoadArenaBytes};
    std::unique_ptr<parser::Parser> parser;
    parser::AstNode* root = nullptr;
  };

  std::string text;
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::unique_ptr<Parsed> parsed;
  bool parse_failed = false;
};

/// Error for a follower `<f t="shared">` whose `si` has no registered master.
Error UnknownSharedSi(std::string_view context_prefix, std::uint32_t si) {
  std::string ctx("context=");
  ctx.append(context_prefix);
  ctx.append(" si=");
  ctx.append(std::to_string(si));
  return make_error(FormulonErrorCode::kIoSheetCorrupt, "shared formula: slave references unknown si", std::move(ctx));
}

/// Error for a `sheet_index` past the workbook's sheet count.
Error SheetIndexOutOfRange(std::string_view fn_name, std::string_view context_prefix, std::size_t sheet_index,
                           std::size_t sheet_count) {
  std::string ctxs("context=");
  ctxs.append(context_prefix);
  ctxs.append(" sheet_index=");
  ctxs.append(std::to_string(sheet_index));
  ctxs.append(" sheet_count=");
  ctxs.append(std::to_string(sheet_count));
  std::string message(fn_name);
  message.append(": sheet_index out of range");
  return make_error(FormulonErrorCode::kInvalidArgument, std::move(message), std::move(ctxs));
}

/// Records the `cm` (cell metadata) attribute of a formula cell; a malformed
/// value is recorded as 0.
void RecordCellMetadata(SheetReadContext& ctx, std::uint32_t row, std::uint32_t col, std::string_view cm_text) {
  std::uint32_t cm = 0;
  if (!parse_xsd_nonneg_int(cm_text, &cm)) {
    cm = 0;
  }
  ctx.cell_metadata.emplace_back(row, col, cm);
}

/// Records the `vm` (value metadata) attribute of a cell; a malformed value is
/// recorded as 0.
void RecordValueMetadata(SheetReadContext& ctx, std::uint32_t row, std::uint32_t col, std::string_view vm_text) {
  std::uint32_t vm = 0;
  if (!parse_xsd_nonneg_int(vm_text, &vm)) {
    vm = 0;
  }
  ctx.value_metadata.push_back(RichErrorCell{row, col, vm});
}

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

std::string ShiftSharedFormulaText(SharedFormulaMaster& master, std::uint32_t target_row, std::uint32_t target_col) {
  const std::int32_t row_delta = static_cast<std::int32_t>(target_row) - static_cast<std::int32_t>(master.row);
  const std::int32_t col_delta = static_cast<std::int32_t>(target_col) - static_cast<std::int32_t>(master.col);
  if (row_delta == 0 && col_delta == 0) {
    return master.text;
  }

  if (master.parse_failed) {
    return master.text;
  }
  if (master.parsed == nullptr) {
    // `master.text` is the raw `<f>` body, which may still carry Excel's
    // `_xlfn.` / `_xlpm.` storage prefixes (the master cell's own text is
    // canonicalised later, by `Workbook::set_cell_formula`). Strip them
    // before parsing: the general Call-node path only recognises those
    // prefixes for LET / LAMBDA specially, so an un-stripped future-function
    // callee would shift and re-format with the prefix still glued to its
    // name, and the prefix bytes also count against the tokenizer's
    // formula-length cap even though Excel's own limit applies to the
    // canonical (formula-bar) text this produces.
    auto parsed = std::make_unique<SharedFormulaMaster::Parsed>();
    parsed->source.assign("=");
    parsed->source.append(parser::strip_storage_prefixes(master.text, &has_storage_prefix));
    parsed->parser = std::make_unique<parser::Parser>(parsed->source, parsed->arena);
    parsed->root = parsed->parser->parse();
    if (parsed->root == nullptr || !parsed->parser->errors().empty()) {
      // Keep the workbook loadable for formula dialects this parser does not
      // fully understand yet. Parseable formulas still get correct Excel-style
      // shared-formula relative expansion.
      master.parse_failed = true;
      return master.text;
    }
    master.parsed = std::move(parsed);
  }
  // Shifted nodes go to a per-follower arena so the master's stays bounded.
  Arena shift_arena(kDefaultArenaChunkBytes, kMaxLoadArenaBytes);
  const parser::AstNode* shifted = parser::shift_relative_refs(*master.parsed->root, shift_arena, row_delta, col_delta);
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

/// Models the `<f t="dataTable">` at `(row, col)` as a table of `sheet`. A
/// table whose `ref` does not parse stays a bare cached value, and
/// `diagnostics` (when non-null) counts the drop via `skipped_feature_count`.
void RegisterDataTable(Sheet& sheet, std::uint32_t row, std::uint32_t col, const DataTableAttrs& attrs,
                       ReadDiagnostics* diagnostics) {
  DataTable table;
  if (parse_data_table(attrs, &table)) {
    // The anchor must stay a stored cell for the writer to find it.
    const Cell* anchor = sheet.cell_at(table.ref.first_row, table.ref.first_col);
    if (anchor == nullptr || anchor->cached_value.is_blank()) {
      sheet.set_cell_value(table.ref.first_row, table.ref.first_col, Value::number(0.0));
    }
    sheet.mutable_data_tables().push_back(table);
    return;
  }
  StructuredLog("io.sheet.formula.data_table_skip")
      .field("row", std::to_string(row))
      .field("col", std::to_string(col))
      .error_code(FormulonErrorCode::kIoSheetCorrupt)
      .warn();
  if (diagnostics != nullptr) {
    ++diagnostics->skipped_feature_count;
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
///   * `<f t="array">BODY</f>` (CSE array) is accepted but treated as a
///     plain formula: we read the body as the formula text. CSE-array
///     detail will land in a later bundle.
///   * `<f t="dataTable" .../>` carries no body (its geometry lives in the
///     `ref`/`r1`/`r2`/`dt2D`/`dtr` attributes, read by `RegisterDataTable`),
///     so `formula_out` comes out empty and the cell keeps its cached `<v>`.
Expected<void, Error> ResolveFormula(const pugi::xml_node& c_node,
                                     std::unordered_map<std::uint32_t, SharedFormulaMaster>& shared, std::uint32_t row,
                                     std::uint32_t col, std::string& formula_out) {
  pugi::xml_node f_node = c_node.child("f");
  if (!f_node) {
    formula_out.clear();
    return Expected<void, Error>::Ok();
  }
  const std::string_view ftype = f_node.attribute("t").value();
  if (ftype == "dataTable") {
    formula_out.clear();
    return Expected<void, Error>::Ok();
  }
  if (ftype != "shared") {
    // Plain formula (or an unhandled variant we treat as plain).
    ReadFormulaElementText(formula_out, f_node);
    return Expected<void, Error>::Ok();
  }

  // Shared formula. `si` is required; missing/non-numeric => corrupt.
  pugi::xml_attribute si_attr = f_node.attribute("si");
  if (!si_attr) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "shared formula: <f t='shared'> missing 'si'",
                      "context=sheet_reader");
  }
  ASSIGN_OR_RETURN(const std::uint32_t si, ParseSharedFormulaSi(si_attr.value()));

  std::string body;
  ReadFormulaElementText(body, f_node);
  if (!body.empty()) {
    // Master occurrence: register and use as formula text.
    shared[si] = SharedFormulaMaster{body, row, col, nullptr, false};
    formula_out = std::move(body);
    return Expected<void, Error>::Ok();
  }
  // Slave occurrence: look up master.
  auto it = shared.find(si);
  if (it == shared.end()) {
    return UnknownSharedSi("sheet_reader", si);
  }
  formula_out = ShiftSharedFormulaText(it->second, row, col);
  return Expected<void, Error>::Ok();
}

/// Stores one parsed `<c>` and charges `budget` for what it adds (see
/// `sheet_load_write_bytes`); `context_prefix` opens the error context of an
/// exceeded budget. A blank, formula-less, non-phonetic `<c>` is a style run,
/// written for every xf including 0 (`Sheet::load_style_run`); the
/// explicit-default rule runs once the layout is applied.
Expected<void, Error> ApplyParsedCell(const ParsedCell& parsed, std::string_view formula_text, std::uint32_t xf_index,
                                      const std::vector<PhoneticRun>* phonetic_runs, PhoneticProperties phonetic_props,
                                      std::size_t sheet_index, Workbook& workbook, SheetReadContext& ctx,
                                      ResourceBudget& budget, std::string_view context_prefix) {
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
  Sheet& target = workbook.sheet(sheet_index);
  const bool content = !formula_text.empty() || !stored_value.is_blank() || parsed.is_sst_index || parsed.show_phonetic;
  if (!content && !target.has_cell(parsed.row, parsed.col)) {
    RETURN_IF_ERROR(charge_sheet_load(budget, sheet_load_write_bytes(target, parsed.row, parsed.col, false),
                                      context_prefix, parsed.row, parsed.col));
    const std::size_t added = target.load_style_run(parsed.row, parsed.col, xf_index);
    return charge_sheet_load(budget, added * kSheetLoadRunBytes, context_prefix, parsed.row, parsed.col);
  }
  RETURN_IF_ERROR(charge_sheet_load(budget, sheet_load_write_bytes(target, parsed.row, parsed.col, content),
                                    context_prefix, parsed.row, parsed.col));
  if (!formula_text.empty()) {
    // `Workbook::set_cell_formula` accepts both spellings, but to
    // match the parser/evaluator's expected input form (the existing
    // call sites in workbook_recalc_test.cpp pass "=A1*2") we prepend
    // '=' here.
    std::string with_eq("=");
    with_eq.append(formula_text);
    auto wf = workbook.set_cell_formula(sheet_index, parsed.row, parsed.col,
                                        workbook.ingest_stored_formula(with_eq, Workbook::FormulaSource::kEntered),
                                        Workbook::FormulaSource::kRendered);
    if (!wf) {
      return std::move(wf.error());
    }
    // Preserve Excel's cached result until a caller explicitly recalculates.
    // This is essential when a workbook uses functions Formulon does not yet
    // implement: eagerly replacing a valid loaded cache with #NAME? makes a
    // read-only inspection or save/load round-trip lose useful data.
    if (!stored_value.is_blank() && !parsed.is_sst_index) {
      workbook.sheet(sheet_index).set_cell_cached_value_borrowed(parsed.row, parsed.col, stored_value);
    }
  } else if (stored_value.is_blank()) {
    // A phonetic blank, or a formatted duplicate of a stored coordinate: the
    // format lands on the stored cell below. A bare duplicate changes nothing.
    if (xf_index == 0U && !parsed.has_explicit_xf && !parsed.show_phonetic) {
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
        return std::move(wv.error());
      }
    } else {
      sheet.set_cell_value(parsed.row, parsed.col, stored_value);
    }
  }

  if (parsed.is_sst_index) {
    ctx.pending_sst_cells.emplace_back(parsed.row, parsed.col, parsed.sst_index);
  }

  // Before the style, so a phonetic blank is stored first and takes its xf
  // as a cell rather than as a run the cell then splits.
  if (parsed.show_phonetic) {
    workbook.sheet(sheet_index).set_cell_show_phonetic(parsed.row, parsed.col, true);
  }

  if (xf_index != 0U || parsed.has_explicit_xf) {
    auto sx = workbook.set_cell_xf_index(sheet_index, parsed.row, parsed.col, xf_index);
    if (!sx) {
      return std::move(sx.error());
    }
  }

  if (!parsed.is_sst_index && phonetic_runs != nullptr && !phonetic_runs->empty()) {
    workbook.sheet(sheet_index).set_cell_phonetic_runs(parsed.row, parsed.col, *phonetic_runs);
    workbook.sheet(sheet_index)
        .set_cell_phonetic_props(parsed.row, parsed.col, phonetic_props, parsed.phonetic_pr_present);
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

std::uint64_t sheet_load_write_bytes(const Sheet& sheet, std::uint32_t row, std::uint32_t col, bool stores_cell) {
  const auto it = sheet.rows().find(row);
  if (it == sheet.rows().end()) {
    return kSheetLoadRowBytes + (stores_cell ? kSheetLoadCellBytes : 0U);
  }
  return stores_cell && it->second.find(col) == nullptr ? kSheetLoadCellBytes : 0U;
}

Expected<void, Error> charge_sheet_load(ResourceBudget& budget, std::uint64_t bytes, std::string_view context_prefix,
                                        std::uint32_t row, std::uint32_t col) {
  if (!budget.would_exceed(bytes)) {
    return budget.consume(bytes);
  }
  std::string ctxs(context_prefix);
  ctxs.append(" row=");
  ctxs.append(std::to_string(row));
  ctxs.append(" col=");
  ctxs.append(std::to_string(col));
  return charge(budget, bytes, std::move(ctxs));
}

Expected<void, Error> RegisterArraySpills(Sheet& sheet, const std::vector<ArrayAnchor>& anchors) {
  return register_array_spills(sheet, anchors, FormulonErrorCode::kIoSheetCorrupt, "sheet_reader", "ooxml");
}

Expected<void, Error> read_sheet_data(const pugi::xml_document& sheet_doc, std::size_t sheet_index, Workbook& workbook,
                                      SheetReadContext& ctx, std::deque<std::string>& text_storage,
                                      ReadDiagnostics* diagnostics, std::uint64_t max_cell_bytes) {
  if (sheet_index >= workbook.sheet_count()) {
    return SheetIndexOutOfRange("read_sheet_data", "sheet_reader", sheet_index, workbook.sheet_count());
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
  // Per-sheet load ceiling, charged by `ApplyParsedCell`.
  ResourceBudget cell_budget(max_cell_bytes, "sheet_load_bytes");

  for (pugi::xml_node row = sheet_data.child("row"); row; row = row.next_sibling("row")) {
    for (pugi::xml_node c = row.child("c"); c; c = c.next_sibling("c")) {
      auto parsed_or = parse_cell_element(c, text_storage);
      if (!parsed_or) {
        return std::move(parsed_or.error());
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
          return std::move(resolved.error());
        }
      }

      auto applied =
          ApplyParsedCell(parsed, formula_text, parsed.xf_index, &parsed.phonetic_runs, parsed.phonetic_props,
                          sheet_index, workbook, ctx, cell_budget, "context=sheet_reader");
      if (!applied) {
        return std::move(applied.error());
      }

      // Record a dynamic-array anchor so its cached spill targets do not
      // read back as blocking literals (see `RegisterArraySpills`).
      if (pugi::xml_node f = c.child("f"); f && std::string_view(f.attribute("t").value()) == "array") {
        RecordArrayAnchor(ctx, f.attribute("ref").value(), parsed.row, parsed.col);
      }
      if (!formula_text.empty() && parse_xsd_bool(c.child("f").attribute("ca").value(), false)) {
        ctx.recalc_marked.push_back(CellAddress{parsed.row, parsed.col});
      }
      if (pugi::xml_node f = c.child("f"); f && std::string_view(f.attribute("t").value()) == "dataTable") {
        const DataTableAttrs attrs{f.attribute("ref").value(),  f.attribute("r1").value(),  f.attribute("r2").value(),
                                   f.attribute("dt2D").value(), f.attribute("dtr").value(), f.attribute("del1").value(),
                                   f.attribute("del2").value(), f.attribute("ca").value(),  f.attribute("aca").value()};
        RegisterDataTable(workbook.sheet(sheet_index), parsed.row, parsed.col, attrs, diagnostics);
      }
      if (const pugi::xml_attribute cm = c.attribute("cm"); cm && !formula_text.empty()) {
        RecordCellMetadata(ctx, parsed.row, parsed.col, cm.value());
      }
      if (const pugi::xml_attribute vm = c.attribute("vm"); vm) {
        RecordValueMetadata(ctx, parsed.row, parsed.col, vm.value());
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
    shared[si] = SharedFormulaMaster{body, rec.row, rec.col, nullptr, false};
    return body;
  }
  // Follower occurrence: shift the registered master into this cell.
  auto it = shared.find(si);
  if (it == shared.end()) {
    return UnknownSharedSi("sheet_reader_sax", si);
  }
  return ShiftSharedFormulaText(it->second, rec.row, rec.col);
}

/// Translates one `CellRecord` into the same workbook mutations the
/// DOM path produces.
///
/// Shared-formula groups (`<f t="shared" si=>`) are resolved through
/// `shared` so followers recover the shifted master body — matching the
/// DOM path. Array formulas are read as plain (body verbatim).
/// `<f t="dataTable">` reads as an empty body (matching the DOM path's
/// fallback to the cached `<v>`) and is modelled by `RegisterDataTable`.
Expected<void, Error> ApplyCellRecord(const CellRecord& rec, std::size_t sheet_index, Workbook& workbook,
                                      SheetReadContext& ctx, std::deque<std::string>& text_storage,
                                      std::unordered_map<std::uint32_t, SharedFormulaMaster>& shared,
                                      ReadDiagnostics* diagnostics, ResourceBudget& cell_budget) {
  const bool value_present = rec.is_inline_string || !rec.value.empty();
  ASSIGN_OR_RETURN(auto payload,
                   decode_cell_payload(rec.t, rec.value, value_present, rec.is_inline_string, text_storage));
  ParsedCell& cell = payload;
  cell.row = rec.row;
  cell.col = rec.col;
  cell.has_explicit_xf = !rec.s.empty();
  cell.phonetic_pr_present = rec.phonetic_pr_present;
  // Excel drops a formula cell's phonetic flag on load.
  cell.show_phonetic = parse_xsd_bool(rec.ph, false) && !rec.has_formula;

  // Persist the cell's `s=` xf index when present. The SAX scanner
  // surfaces it as a string view; parse to integer here through the same
  // lexer and with the same degrade-to-0 disposition the DOM cell parser
  // uses. An explicit non-empty `s=` remains meaningful even when it
  // normalises to the default sentinel `0`.
  std::uint32_t xf = 0;
  if (!parse_xsd_nonneg_int(rec.s, &xf)) {
    xf = 0;
  }
  // Resolve shared-formula groups (plain formulas pass straight through).
  ASSIGN_OR_RETURN(auto formula, ResolveSharedFromRecord(rec, shared));
  const std::string& formula_text = formula;

  // Inline-string cells with <rPh> annotations carry their kana on the
  // SAX record. SST-referenced cells (rec.phonetic stays null by
  // construction) route their phonetic through the post-loop SST
  // resolution pass instead — same contract the DOM path uses.
  auto applied = ApplyParsedCell(cell, formula_text, xf, rec.phonetic, rec.phonetic_props, sheet_index, workbook, ctx,
                                 cell_budget, "context=sheet_reader_sax");
  if (!applied) {
    return std::move(applied.error());
  }
  // Record a dynamic-array anchor so its cached spill targets do not read
  // back as blocking literals (see `RegisterArraySpills`).
  if (rec.f_t == "array" && !rec.f_ref.empty()) {
    RecordArrayAnchor(ctx, rec.f_ref, rec.row, rec.col);
  }
  if (!formula_text.empty() && parse_xsd_bool(rec.f_ca, false)) {
    ctx.recalc_marked.push_back(CellAddress{rec.row, rec.col});
  }
  if (rec.f_t == "dataTable") {
    RegisterDataTable(workbook.sheet(sheet_index), rec.row, rec.col,
                      DataTableAttrs{rec.f_ref, rec.f_r1, rec.f_r2, rec.f_dt2d, rec.f_dtr, rec.f_del1, rec.f_del2,
                                     rec.f_ca, rec.f_aca},
                      diagnostics);
  }
  if (!rec.cm.empty() && !formula_text.empty()) {
    RecordCellMetadata(ctx, rec.row, rec.col, rec.cm);
  }
  if (!rec.vm.empty()) {
    RecordValueMetadata(ctx, rec.row, rec.col, rec.vm);
  }
  return applied;
}

Expected<void, Error> SaxOnCellTrampoline(void* user_data, const CellRecord& rec) {
  auto* st = static_cast<SaxApplyState*>(user_data);
  return ApplyCellRecord(rec, st->sheet_index, *st->workbook, *st->ctx, *st->text_storage, st->shared_formulas,
                         st->diagnostics, st->cell_budget);
}

// Captures per-row overrides on the SAX path, mirroring the DOM
// `ApplyRowOverrides`: a row contributes a `RowLayout` when it carries
// `ht`, `hidden`, `outlineLevel`, or effective `customFormat=1` style
// metadata. `customHeight` alone does not, matching the DOM path. Under
// `zeroHeight` every present row contributes, as a visible entry.
Expected<void, Error> SaxOnRowStartTrampoline(void* user_data, const RowRecord& rec) {
  auto* st = static_cast<SaxApplyState*>(user_data);
  const bool custom_format = !rec.custom_format.empty() && parse_xsd_bool(rec.custom_format, false);
  double height = 0.0;
  const bool has_height = parse_xsd_nonneg_double(rec.ht, &height);
  const bool show_phonetic = parse_xsd_bool(rec.ph, false);
  const bool collapsed = parse_xsd_bool(rec.collapsed, false);
  const bool default_hidden = st->workbook->sheet(st->sheet_index).format_defaults().zero_height;
  if (!has_height && rec.hidden.empty() && rec.outline_level.empty() && !custom_format && !show_phonetic &&
      !collapsed && !default_hidden) {
    return Expected<void, Error>::Ok();
  }
  if (rec.row_1based < 1U) {
    return Expected<void, Error>::Ok();
  }
  RowLayout entry;
  entry.row = rec.row_1based - 1U;
  entry.show_phonetic = show_phonetic;
  entry.collapsed = collapsed;
  if (has_height) {
    entry.height = height;
    entry.has_height = true;
    entry.custom_height = parse_xsd_bool(rec.custom_height, false);
  }
  if (!rec.hidden.empty()) {
    entry.hidden = parse_xsd_bool(rec.hidden, false);
  }
  if (!rec.outline_level.empty()) {
    entry.outline_level = parse_outline_level(rec.outline_level);
  }
  if (custom_format) {
    entry.has_style = true;
    std::uint32_t style_xf = 0;
    if (!rec.style.empty()) {
      (void)parse_xsd_nonneg_int(rec.style, &style_xf);
    }
    entry.style_xf = style_xf;
  }
  insert_row_override(st->workbook->sheet(st->sheet_index).mutable_layout(), entry);
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<void, Error> read_sheet_data_sax(ByteSpan sheet_xml, std::size_t sheet_index, Workbook& workbook,
                                          SheetReadContext& ctx, std::deque<std::string>& text_storage,
                                          ReadDiagnostics* diagnostics, std::uint64_t max_cell_bytes) {
  if (sheet_index >= workbook.sheet_count()) {
    return SheetIndexOutOfRange("read_sheet_data_sax", "sheet_reader_sax", sheet_index, workbook.sheet_count());
  }
  // `<sheetFormatPr>` precedes the rows; row visibility needs its zeroHeight before the scan.
  {
    const std::string_view text(reinterpret_cast<const char*>(sheet_xml.data), sheet_xml.size);
    const std::size_t open = text.find("<sheetFormatPr");
    const std::size_t close = open == std::string_view::npos ? open : text.find('>', open);
    if (close != std::string_view::npos) {
      std::string tag(text.substr(open, close - open));
      if (!tag.empty() && tag.back() == '/') {
        tag.pop_back();
      }
      tag.append("/>");
      pugi::xml_document fmt_doc;
      if (fmt_doc.load_buffer(tag.data(), tag.size())) {
        workbook.sheet(sheet_index).mutable_format_defaults().zero_height =
            sheet_format_hides_rows(fmt_doc.document_element());
      }
    }
  }
  SaxApplyState state{
      sheet_index, &workbook, &ctx, &text_storage, {}, diagnostics, ResourceBudget(max_cell_bytes, "sheet_load_bytes")};
  SheetSaxCallbacks cb;
  cb.user_data = &state;
  cb.on_row_start = &SaxOnRowStartTrampoline;
  cb.on_cell = &SaxOnCellTrampoline;
  auto scanned = scan_sheet_data(sheet_xml, cb);
  if (!scanned) {
    return std::move(scanned.error());
  }
  // Spill registration is the caller's job (see `SheetReadContext::
  // array_anchors`): it must run after `ctx.pending_sst_cells` has been
  // resolved, which this function does not do.
  return Expected<void, Error>::Ok();
}

#endif  // !FORMULON_WASM || FORMULON_WASM_ENABLE_SAX

}  // namespace io
}  // namespace formulon
