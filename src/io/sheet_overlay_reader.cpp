//
// Presentation-overlay readers for one worksheet: merged cells, hyperlinks,
// data validations and sheet protection. See sheet_overlay_reader.h for the
// public contract.

#include "io/sheet_overlay_reader.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include "io/cell_parser.h"
#include "io/cf_reader.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "io/xsd_bool.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/structured_log.h"

namespace formulon {
namespace io {

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

/// Emits the WARN diagnostic for one skipped presentation-overlay entry.
/// Merges, hyperlinks and data validations are optional worksheet
/// metadata carrying no cell value, so a single malformed reference
/// drops that entry and load continues — the same disposition
/// `read_conditional_formats` applies to a malformed `<conditionalFormatting>`
/// block. Genuine cell-data corruption still fails the sheet.
void SkipOverlayEntry(std::string_view part, std::string_view reason, std::string_view ref,
                      ReadDiagnostics* diagnostics) {
  StructuredLog("io.sheet.overlay.skip")
      .field("part", part)
      .field("reason", reason)
      .field("ref", ref)
      .error_code(FormulonErrorCode::kIoSheetCorrupt)
      .warn();
  if (diagnostics != nullptr) {
    ++diagnostics->skipped_feature_count;
  }
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

Expected<std::vector<MergeRange>, Error> read_merges(const pugi::xml_node& worksheet, ReadDiagnostics* diagnostics) {
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
      SkipOverlayEntry("mergeCells", "ref attribute missing or empty", ref, diagnostics);
      continue;
    }
    auto r = ParseA1RangeMerge(ref);
    if (!r) {
      SkipOverlayEntry("mergeCells", "ref unparseable", ref, diagnostics);
      continue;
    }
    out.push_back(r.value());
  }
  return out;
}

Expected<std::vector<Hyperlink>, Error> read_hyperlinks(const pugi::xml_node& worksheet, ReadDiagnostics* diagnostics) {
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
      SkipOverlayEntry("hyperlinks", "ref attribute missing or empty", ref, diagnostics);
      continue;
    }
    // Hyperlinks may span a range (`ref="A1:B2"`); Excel applies the link
    // to every cell. Keep both corners in the model so structural edits can
    // shrink/shift the complete rectangle and the writer can regenerate the
    // A1 reference from numeric coordinates.
    auto range_or = ParseA1RangeMerge(ref);
    if (!range_or) {
      SkipOverlayEntry("hyperlinks", "ref unparseable", ref, diagnostics);
      continue;
    }
    Hyperlink hl;
    hl.row = range_or.value().first_row;
    hl.col = range_or.value().first_col;
    hl.last_row = range_or.value().last_row;
    hl.last_col = range_or.value().last_col;
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

Expected<std::vector<DataValidation>, Error> read_data_validations(const pugi::xml_node& worksheet,
                                                                   ReadDiagnostics* diagnostics) {
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
      SkipOverlayEntry("dataValidations", "sqref attribute missing or empty", sqref, diagnostics);
      continue;
    }
    auto ranges_or = ParseSqrefRanges(sqref);
    if (!ranges_or) {
      SkipOverlayEntry("dataValidations", "sqref unparseable", sqref, diagnostics);
      continue;
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

    // Boolean attributes default to false: Excel omits allowBlank when it
    // is off (the .xlsb twin's fAllowBlank bit is clear) and writes "1"
    // when it is on.
    v.allow_blank = false;
    if (pugi::xml_attribute ab = dv.attribute("allowBlank"); ab) {
      v.allow_blank = parse_xml_bool(ab.value());
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

    // `AppendOoxmlTextUnescaped` decodes the OOXML `_xHHHH_` control-
    // character escapes the writer emits for this slot (`AppendXmlEscaped`
    // in ooxml/sheet_xml_builder.cpp), on top of `.text().get()`'s
    // ordinary XML entity decoding.
    if (pugi::xml_node f1 = dv.child("formula1"); f1) {
      v.formula1.clear();
      AppendOoxmlTextUnescaped(v.formula1, f1.text().get());
      if (!v.formula1.empty() && v.formula1.front() == '=') {
        v.formula1.erase(0, 1);
      }
      v.formula1 = canonical_feature_formula(v.formula1);
    }
    if (pugi::xml_node f2 = dv.child("formula2"); f2) {
      v.formula2.clear();
      AppendOoxmlTextUnescaped(v.formula2, f2.text().get());
      if (!v.formula2.empty() && v.formula2.front() == '=') {
        v.formula2.erase(0, 1);
      }
      v.formula2 = canonical_feature_formula(v.formula2);
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
    // Saturate at both ends rather than truncating. A `static_cast` of a
    // value above 2^32-1 wraps, and a value that wraps to exactly 0 is then
    // dropped by the writer's non-zero guard, so the saved file would carry
    // a hash and salt with no iteration count at all. This surface promises
    // verbatim round-trip, so an unrepresentable count must degrade to the
    // nearest representable one instead of disappearing.
    const long long parsed = sc.as_llong(0);
    out.spin_count = parsed < 0 ? 0U
                     : parsed > static_cast<long long>(std::numeric_limits<std::uint32_t>::max())
                         ? std::numeric_limits<std::uint32_t>::max()
                         : static_cast<std::uint32_t>(parsed);
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
