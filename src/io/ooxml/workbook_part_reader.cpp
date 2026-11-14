#include "io/ooxml/workbook_part_reader.h"

#include <algorithm>
#include <cstdint>
#include <iterator>
#include <string>
#include <string_view>

#include "calc_settings.h"
#include "io/ooxml/raw_child_sweep.h"
#include "io/ooxml_defs.h"
#include "io/xml_utils.h"
#include "io/xsd_bool.h"
#include "io/xsd_double.h"
#include "sheet_passthrough.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace ooxml {

namespace {

/// The element child of `node` when it has exactly one, else a null node.
pugi::xml_node sole_element_child(const pugi::xml_node& node) {
  pugi::xml_node found;
  for (pugi::xml_node child = node.first_child(); child; child = child.next_sibling()) {
    if (child.type() != pugi::node_element) {
      continue;
    }
    if (found) {
      return pugi::xml_node();
    }
    found = child;
  }
  return found;
}

/// Whether `node` is the `mc:AlternateContent` Excel writes around
/// `x15ac:absPath`: one `mc:Choice Requires="x15"` holding only an `absPath`
/// in the x15ac namespace. Any other shape is retained verbatim.
bool IsAbsPathAlternateContent(const pugi::xml_node& node) {
  if (node.type() != pugi::node_element || local_name(node) != "AlternateContent") {
    return false;
  }
  const pugi::xml_node choice = sole_element_child(node);
  if (!choice || local_name(choice) != "Choice" || std::string_view(choice.attribute("Requires").value()) != "x15") {
    return false;
  }
  const pugi::xml_node abs_path = sole_element_child(choice);
  return abs_path && local_name(abs_path) == "absPath" &&
         find_namespace_binding(abs_path, qname_prefix(abs_path.name())) == kNsX15ac;
}

}  // namespace

void apply_workbook_part_settings(const pugi::xml_node& wb_root, Workbook& wb) {
  // <calcPr> — workbook-level calc mode and iterative-calc options. The
  // element is optional; absence means Excel defaults (auto + iterative
  // off). When present, accept the documented `calcMode` values
  // (`auto` / `manual` / `autoNoTable`) and the iterative trio
  // (`iterate`, `iterateCount`, `iterateDelta`). Unknown calcMode
  // strings fall back to `auto` rather than failing the load.
  if (pugi::xml_node calc_pr = wb_root.child("calcPr"); calc_pr) {
    const std::string_view calc_mode_attr = calc_pr.attribute("calcMode").value();
    if (calc_mode_attr == "manual") {
      wb.set_calc_mode(Workbook::CalcMode::kManual);
    } else if (calc_mode_attr == "autoNoTable") {
      wb.set_calc_mode(Workbook::CalcMode::kAutoNoTable);
    } else {
      wb.set_calc_mode(Workbook::CalcMode::kAuto);
    }
    IterativeOptions opts;
    if (pugi::xml_attribute iterate = calc_pr.attribute("iterate"); iterate) {
      opts.enabled = parse_xml_bool(iterate.value());
    }
    if (pugi::xml_attribute count = calc_pr.attribute("iterateCount"); count) {
      const long long parsed = count.as_llong(static_cast<long long>(kDefaultMaxIterations));
      // Clamp into Excel's own dialog range. The upper bound matters more
      // than the lower one: the file decides how much work the first
      // `recalc()` performs, the solver has no wall-clock limit, and the
      // cancellation callback is null unless the host opted in. Without
      // this, `iterateCount="4294967295"` on a two-cell cycle is an
      // unrecoverable hang rather than a slow load.
      opts.max_iterations = parsed < 1                                           ? 1U
                            : parsed > static_cast<long long>(kMaxIterationsCap) ? kMaxIterationsCap
                                                                                 : static_cast<std::uint32_t>(parsed);
    }
    // The solver stops once the largest change falls below `max_change`.
    // A NaN tolerance makes that comparison false forever, so the workbook
    // silently burns the whole iteration budget and reports the
    // unconverged values; a negative one can never be reached either.
    double delta_value = 0.0;
    if (parse_xsd_nonneg_double(attr_str(calc_pr, "iterateDelta"), &delta_value)) {
      opts.max_change = delta_value;
    }
    wb.set_iterative_options(opts);

    CalcOptions& calc = wb.mutable_calc_options();
    calc = CalcOptions{};
    calc.full_precision = read_xsd_bool(calc_pr, "fullPrecision", true);
    calc.calc_on_save = read_xsd_bool(calc_pr, "calcOnSave", true);
    calc.ref_mode = attr_str(calc_pr, "refMode") == "R1C1" ? RefMode::kR1C1 : RefMode::kA1;
    calc.concurrent_calc = read_xsd_bool(calc_pr, "concurrentCalc", true);
    calc.force_full_calc = read_xsd_bool(calc_pr, "forceFullCalc", false);
    calc.full_calc_on_load = read_xsd_bool(calc_pr, "fullCalcOnLoad", false);
    calc.calc_completed = read_xsd_bool(calc_pr, "calcCompleted", true);
    if (calc_pr.attribute("calcId")) {
      calc.has_calc_id = true;
      calc.calc_id = attr_u32(calc_pr, "calcId", 0U);
    }
    if (calc_pr.attribute("concurrentManualCount")) {
      calc.has_concurrent_manual_count = true;
      calc.concurrent_manual_count = attr_u32(calc_pr, "concurrentManualCount", 0U);
    }
  }

  // Workbook-level elements captured raw for verbatim re-emission:
  // `<fileVersion>`, `<fileSharing>`, `<workbookPr>`, `<workbookProtection>`,
  // `<bookViews>`, and trailing `<extLst>`. Without this,
  // the writer regenerates only `<sheets>` / `<definedNames>` / `<calcPr>`
  // / `<pivotCaches>` and silently drops the date system, tab-selection
  // state, and workbook protection. `<workbookPr date1904>` additionally
  // seeds the model-level `date1904` flag the date-serial conversions read.
  if (pugi::xml_node file_version = wb_root.child("fileVersion"); file_version) {
    wb.set_file_version_xml(raw_xml(file_version));
  }
  if (pugi::xml_node file_sharing = wb_root.child("fileSharing"); file_sharing) {
    wb.set_file_sharing_xml(raw_xml(file_sharing));
  }
  if (pugi::xml_node workbook_pr = wb_root.child("workbookPr"); workbook_pr) {
    wb.set_workbook_pr_xml(raw_xml(workbook_pr));
    // Excel emits the attribute as `date1904`; some legacy producers use
    // the bare `1904` spelling. Both default to false when absent.
    const bool from_date1904 = read_xsd_bool(workbook_pr, "date1904", false);
    const bool from_legacy = read_xsd_bool(workbook_pr, "1904", false);
    wb.set_date1904(from_date1904 || from_legacy);
  }
  if (pugi::xml_node workbook_protection = wb_root.child("workbookProtection"); workbook_protection) {
    wb.set_workbook_protection_xml(raw_xml(workbook_protection));
  }
  if (pugi::xml_node book_views = wb_root.child("bookViews"); book_views) {
    wb.set_book_views_xml(raw_xml(book_views));
  }
  // The Compatibility Version stamp, kept as read (absent unless the part carries one); the raw <extLst> keeps it too
  // and the writer regenerates only that ext from the model.
  std::optional<std::uint32_t> compat_set_version;
  std::optional<std::uint32_t> compat_warn_below;
  if (pugi::xml_node ext_lst = wb_root.child("extLst"); ext_lst) {
    wb.set_workbook_ext_lst_xml(raw_xml(ext_lst));
    for (pugi::xml_node ext : ext_lst.children()) {
      if (std::string_view(ext.attribute("uri").value()) != kCompatibilityVersionExtUri) {
        continue;
      }
      for (pugi::xml_node version : ext.children()) {
        if (local_name(version) != "version") {
          continue;
        }
        compat_set_version = parse_xml_u32_attr_strict(version.attribute("setVersion"));
        compat_warn_below = parse_xml_u32_attr_strict(version.attribute("warnBelowVersion"));
      }
    }
  }
  wb.restore_compatibility_stamp(compat_set_version, compat_warn_below);

  // Everything else is retained by exclusion, so an element no release has
  // modelled survives too. A name outside the sequence (`mc:AlternateContent`)
  // takes the slot of the child before it.
  static constexpr std::string_view kModelled[] = {
      "fileVersion",        "fileSharing",  "workbookPr", "workbookProtection", "bookViews", "sheets",
      "externalReferences", "definedNames", "calcPr",     "pivotCaches",        "extLst",
  };
  CaptureUnmodelledChildren(
      wb_root, workbook_child::slot_of, workbook_child::kCount,
      [](std::string_view name) {
        return std::find(std::begin(kModelled), std::end(kModelled), name) != std::end(kModelled);
      },
      wb.mutable_raw_children());
  // The writer regenerates `x15ac:absPath` from `Workbook::location()`, which
  // a load never sets: the captured element is dropped, not passed through.
  for (pugi::xml_node child = wb_root.first_child(); child; child = child.next_sibling()) {
    if (!IsAbsPathAlternateContent(child)) {
      continue;
    }
    const std::string xml = raw_xml(child);
    WorksheetRawExtensions& raw = wb.mutable_raw_children();
    raw.erase(std::remove_if(raw.begin(), raw.end(), [&xml](const WorksheetRawChild& c) { return c.xml == xml; }),
              raw.end());
  }
}

}  // namespace ooxml
}  // namespace io
}  // namespace formulon
