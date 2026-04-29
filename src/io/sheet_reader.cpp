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

#include <cstddef>
#include <cstdint>
#include <deque>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>

#include "io/cell_parser.h"
#include "pugixml.hpp"
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
    }
  }
  return Expected<void, Error>::Ok();
}

}  // namespace io
}  // namespace formulon
